import io
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import streamlit as st
import imageio_ffmpeg
from pptx import Presentation


st.set_page_config(
    page_title="PPT Notes Extractor & MP4 to MP3",
    page_icon="🎓",
    layout="wide",
)

st.title("PPT Notes Extractor & MP4 to MP3 Converter")
st.write("Extract speaker notes from PowerPoint files or convert MP4 videos to MP3 audio.")


# -------------------------------------------------------------------
# PowerPoint speaker notes extraction
# -------------------------------------------------------------------
def _local_name(tag):
    """Return an XML tag's local name without its namespace."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _extract_text_from_notes_xml(notes_xml_bytes):
    """
    Extract readable text from a PowerPoint notes XML part.
    Includes ordinary DrawingML text and text nodes nested in math markup.
    """
    root = ET.fromstring(notes_xml_bytes)
    paragraphs = []

    # PowerPoint notes commonly store visible text in a:p paragraphs.
    for paragraph in root.iter():
        if _local_name(paragraph.tag) != "p":
            continue

        pieces = []
        for node in paragraph.iter():
            name = _local_name(node.tag)
            if name == "t" and node.text:
                pieces.append(node.text)
            elif name == "tab":
                pieces.append("\t")
            elif name == "br":
                pieces.append("\n")

        text = "".join(pieces).strip()
        if text:
            paragraphs.append(text)

    # If no paragraph text was found, fall back to collecting text nodes.
    if not paragraphs:
        pieces = []
        for node in root.iter():
            if _local_name(node.tag) == "t" and node.text:
                pieces.append(node.text)
        fallback = " ".join(pieces).strip()
        if fallback:
            paragraphs.append(fallback)

    # Remove adjacent duplicate lines while preserving order.
    cleaned = []
    for item in paragraphs:
        item = re.sub(r"[ \t]+", " ", item).strip()
        if item and (not cleaned or item != cleaned[-1]):
            cleaned.append(item)

    return "\n".join(cleaned).strip()


def extract_ppt_notes(pptx_file):
    """
    Return a dictionary with notes and per-slide extraction status.
    Slides with no readable notes are marked for review, because empty notes
    cannot always be distinguished from notes containing unsupported content.
    """
    result = {
        "notes_text": "",
        "total_slides": 0,
        "extracted_slides": [],
        "empty_slides": [],
        "failed_slides": [],
    }

    pptx_bytes = pptx_file.getvalue()
    presentation = Presentation(io.BytesIO(pptx_bytes))
    result["total_slides"] = len(presentation.slides)

    # Map slide number to notesSlide XML using the PPTX package relationships.
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as archive:
            names = set(archive.namelist())
            slide_notes = {}

            # PowerPoint's standard notes part names are notesSlides/notesSlideN.xml.
            notes_files = [
                name for name in names
                if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", name)
            ]

            # The notes-slide number generally corresponds to the slide number,
            # but use slide relationship files when available for accuracy.
            for slide_number in range(1, result["total_slides"] + 1):
                slide_path = f"ppt/slides/slide{slide_number}.xml"
                rel_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"

                try:
                    if slide_path not in names:
                        result["failed_slides"].append(slide_number)
                        continue

                    notes_path = None
                    if rel_path in names:
                        rel_root = ET.fromstring(archive.read(rel_path))
                        for rel in rel_root:
                            if rel.attrib.get("Type", "").endswith("/notesSlide"):
                                target = rel.attrib.get("Target", "")
                                if target:
                                    # Relationship targets are relative to ppt/slides/.
                                    if target.startswith("/"):
                                        notes_path = target.lstrip("/")
                                    else:
                                        notes_path = str(
                                            Path("ppt/slides") / target
                                        ).replace("\\", "/")
                                        # Normalize ../ paths without relying on filesystem.
                                        parts = []
                                        for part in notes_path.split("/"):
                                            if part == "..":
                                                if parts:
                                                    parts.pop()
                                            elif part != ".":
                                                parts.append(part)
                                        notes_path = "/".join(parts)
                                    break

                    if not notes_path or notes_path not in names:
                        result["empty_slides"].append(slide_number)
                        continue

                    notes = _extract_text_from_notes_xml(archive.read(notes_path))

                    # Remove common PowerPoint notes placeholders if they appear alone.
                    if notes.strip().lower() in {
                        "click to add notes",
                        "click to edit master text styles",
                    }:
                        notes = ""

                    if notes:
                        result["extracted_slides"].append(slide_number)
                        slide_notes[slide_number] = notes
                    else:
                        result["empty_slides"].append(slide_number)

                except Exception:
                    result["failed_slides"].append(slide_number)

            output_blocks = []
            for slide_number in range(1, result["total_slides"] + 1):
                if slide_number in slide_notes:
                    output_blocks.append(
                        f"SLIDE {slide_number}\n{slide_notes[slide_number]}"
                    )

            result["notes_text"] = "\n\n".join(output_blocks)

    except zipfile.BadZipFile as exc:
        raise ValueError(
            "This file does not appear to be a valid, unencrypted .pptx file."
        ) from exc
    except Exception as exc:
        raise ValueError(f"Could not read the PowerPoint file: {exc}") from exc

    return result


# -------------------------------------------------------------------
# MP4 to MP3 conversion
# -------------------------------------------------------------------
def convert_mp4_to_mp3(uploaded_video):
    """Convert uploaded MP4 audio to MP3 using the FFmpeg executable."""
    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
    video_bytes = uploaded_video.getvalue()

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = Path(temp_dir) / "input.mp4"
        output_path = Path(temp_dir) / "output.mp3"
        input_path.write_bytes(video_bytes)

        command = [
            ffmpeg_exe,
            "-hide_banner",
            "-loglevel", "error",
            "-y",
            "-i", str(input_path),
            "-vn",
            "-codec:a", "libmp3lame",
            "-q:a", "2",
            str(output_path),
        ]

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=600,
        )

        if completed.returncode != 0:
            detail = (completed.stderr or "FFmpeg could not convert this file.").strip()
            raise ValueError(detail[-1800:])

        if not output_path.exists() or output_path.stat().st_size == 0:
            raise ValueError(
                "No audio was produced. The MP4 may not contain an audio track."
            )

        return output_path.read_bytes()


# -------------------------------------------------------------------
# App tabs
# -------------------------------------------------------------------
ppt_tab, audio_tab = st.tabs(["📊 PPT Speaker Notes", "🎵 MP4 to MP3"])

with ppt_tab:
    st.subheader("Extract PowerPoint Speaker Notes")
    uploaded_pptx = st.file_uploader(
        "Upload PowerPoint file",
        type=["pptx"],
        key="pptx_upload",
    )

    if uploaded_pptx is not None:
        st.caption(f"File: {uploaded_pptx.name}")

        if st.button("Extract Speaker Notes", key="extract_notes_button"):
            try:
                with st.spinner("Reading slides and extracting notes..."):
                    result = extract_ppt_notes(uploaded_pptx)

                st.success("PowerPoint processing completed.")

                col1, col2, col3 = st.columns(3)
                col1.metric("Total slides", result["total_slides"])
                col2.metric("Slides with notes", len(result["extracted_slides"]))
                col3.metric(
                    "Slides to review",
                    len(result["empty_slides"]) + len(result["failed_slides"]),
                )

                if result["empty_slides"]:
                    slide_list = ", ".join(map(str, result["empty_slides"]))
                    st.warning(
                        "No readable speaker notes were extracted from slide(s): "
                        f"{slide_list}. These slides may have no notes, or may contain "
                        "content such as equations that needs manual review."
                    )

                if result["failed_slides"]:
                    slide_list = ", ".join(map(str, result["failed_slides"]))
                    st.error(
                        f"An error occurred while reading slide(s): {slide_list}. "
                        "Try opening and resaving the presentation in PowerPoint."
                    )

                if result["notes_text"].strip():
                    st.text_area(
                        "Extracted speaker notes",
                        value=result["notes_text"],
                        height=400,
                    )
                    safe_name = Path(uploaded_pptx.name).stem
                    st.download_button(
                        "Download notes as TXT",
                        data=result["notes_text"].encode("utf-8"),
                        file_name=f"{safe_name}_speaker_notes.txt",
                        mime="text/plain",
                    )
                else:
                    st.info(
                        "No readable speaker notes were found. Check the slide review "
                        "list above and verify the notes in PowerPoint."
                    )

            except Exception as exc:
                st.error(
                    "Could not process this PowerPoint file. Please check that it is "
                    "a valid, unencrypted PPTX file."
                )
                st.caption(f"Details: {exc}")

with audio_tab:
    st.subheader("Convert MP4 Video to MP3 Audio")
    uploaded_mp4 = st.file_uploader(
        "Upload MP4 video",
        type=["mp4"],
        key="mp4_upload",
    )

    if uploaded_mp4 is not None:
        st.caption(f"File: {uploaded_mp4.name}")

        if st.button("Convert to MP3", key="convert_mp3_button"):
            try:
                with st.spinner("Extracting audio... this may take a little while."):
                    mp3_bytes = convert_mp4_to_mp3(uploaded_mp4)

                safe_name = Path(uploaded_mp4.name).stem
                st.success("Audio conversion completed.")
                st.download_button(
                    "Download MP3",
                    data=mp3_bytes,
                    file_name=f"{safe_name}.mp3",
                    mime="audio/mpeg",
                )
            except subprocess.TimeoutExpired:
                st.error("The conversion took too long. Try a shorter video.")
            except Exception as exc:
                st.error("Could not convert this MP4 file.")
                st.caption(f"Details: {exc}")
