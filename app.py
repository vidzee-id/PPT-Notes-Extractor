import io
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import imageio_ffmpeg
import streamlit as st
from pptx import Presentation


st.set_page_config(
    page_title="Learning Media Toolkit",
    page_icon="🎓",
    layout="wide",
)

# Friendly, lightweight styling. Uses standard Streamlit components.
st.markdown(
    """
    <style>
      .stApp {
        background: linear-gradient(180deg, #f5f8ff 0%, #ffffff 42%);
      }
      .block-container {
        max-width: 1120px;
        padding-top: 1.8rem;
        padding-bottom: 3rem;
      }
      .hero {
        background: linear-gradient(115deg, #2447a8 0%, #4677dc 58%, #6d8fe8 100%);
        border-radius: 22px;
        padding: 1.6rem 1.8rem;
        color: white;
        margin-bottom: 1.25rem;
        box-shadow: 0 10px 28px rgba(36, 71, 168, 0.15);
      }
      .hero h1 {
        color: white;
        font-size: 2rem;
        line-height: 1.2;
        margin: 0 0 .45rem 0;
      }
      .hero p {
        color: rgba(255,255,255,.92);
        margin: 0;
        font-size: 1rem;
      }
      .tool-card {
        background: white;
        border: 1px solid #e5eaf5;
        border-radius: 18px;
        padding: 1.1rem 1.25rem;
        margin: .2rem 0 1rem 0;
        box-shadow: 0 4px 16px rgba(27, 50, 100, .045);
      }
      .tool-icon {
        display: inline-flex;
        width: 46px;
        height: 46px;
        border-radius: 14px;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
        background: #edf2ff;
        margin-bottom: .45rem;
      }
      .tool-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #1d2b53;
        margin: .1rem 0 .25rem 0;
      }
      .tool-description {
        color: #5c6780;
        font-size: .95rem;
        margin: 0;
      }
      .small-note {
        color: #667085;
        font-size: .88rem;
      }
      div[data-testid="stFileUploader"] {
        border: 1px dashed #b9c8ee;
        border-radius: 14px;
        padding: .45rem .6rem;
        background: #fbfcff;
      }
      .stButton > button, .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        min-height: 2.7rem;
      }
      div[data-testid="stMetric"] {
        background: #f7f9ff;
        border: 1px solid #e6ebfa;
        padding: .8rem;
        border-radius: 12px;
      }
      div[data-testid="stTabs"] button {
        border-radius: 10px 10px 0 0;
        font-weight: 600;
      }
      hr { border-color: #e8edf6; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <h1>🎓 Learning Media Toolkit</h1>
      <p>Two handy tools for your content workflow — extract PowerPoint speaker notes or turn video into audio.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------------
# PowerPoint speaker notes extraction
# -------------------------------------------------------------------
def _local_name(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _extract_text_from_notes_xml(notes_xml_bytes):
    root = ET.fromstring(notes_xml_bytes)
    paragraphs = []

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

    if not paragraphs:
        pieces = []
        for node in root.iter():
            if _local_name(node.tag) == "t" and node.text:
                pieces.append(node.text)
        fallback = " ".join(pieces).strip()
        if fallback:
            paragraphs.append(fallback)

    cleaned = []
    for item in paragraphs:
        item = re.sub(r"[ \t]+", " ", item).strip()
        if item and (not cleaned or item != cleaned[-1]):
            cleaned.append(item)

    return "\n".join(cleaned).strip()


def extract_ppt_notes(pptx_file):
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

    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as archive:
            names = set(archive.namelist())
            slide_notes = {}

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
                                    if target.startswith("/"):
                                        notes_path = target.lstrip("/")
                                    else:
                                        notes_path = str(
                                            Path("ppt/slides") / target
                                        ).replace("\\", "/")
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

            output_blocks = [
                f"SLIDE {number}\n{slide_notes[number]}"
                for number in range(1, result["total_slides"] + 1)
                if number in slide_notes
            ]
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
            detail = (
                completed.stderr or "FFmpeg could not convert this file."
            ).strip()
            raise ValueError(detail[-1800:])

        if not output_path.exists() or output_path.stat().st_size == 0:
            raise ValueError(
                "No audio was produced. The MP4 may not contain an audio track."
            )

        return output_path.read_bytes()


# -------------------------------------------------------------------
# Main tools
# -------------------------------------------------------------------
ppt_tab, audio_tab = st.tabs(
    ["📝  PPT Speaker Notes", "🎧  MP4 to MP3"]
)

with ppt_tab:
    st.markdown(
        """
        <div class="tool-card">
          <div class="tool-icon">📝</div>
          <div class="tool-title">PowerPoint Speaker Notes Extractor</div>
          <p class="tool-description">Pull speaker notes into a clean text file, with a quick report of slides that may need checking.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_pptx = st.file_uploader(
        "Choose a PowerPoint file",
        type=["pptx"],
        help="Upload a .pptx file. Password-protected files are not supported.",
        key="pptx_upload",
    )

    if uploaded_pptx is not None:
        st.caption(f"📄 Selected file: {uploaded_pptx.name}")

        if st.button("✨ Extract speaker notes", key="extract_notes_button", type="primary"):
            try:
                with st.spinner("Reading slides and collecting notes…"):
                    result = extract_ppt_notes(uploaded_pptx)

                st.success("Done! Your PowerPoint has been processed.")

                col1, col2, col3 = st.columns(3)
                col1.metric("📑 Total slides", result["total_slides"])
                col2.metric("✅ Notes extracted", len(result["extracted_slides"]))
                col3.metric(
                    "🔎 Slides to review",
                    len(result["empty_slides"]) + len(result["failed_slides"]),
                )

                if result["empty_slides"]:
                    slide_list = ", ".join(map(str, result["empty_slides"]))
                    st.warning(
                        "No readable notes were extracted from slide(s): "
                        f"**{slide_list}**. They may have no notes, or may contain "
                        "equations or other content that needs manual review."
                    )

                if result["failed_slides"]:
                    slide_list = ", ".join(map(str, result["failed_slides"]))
                    st.error(
                        f"Could not read slide(s): **{slide_list}**. "
                        "Try opening and resaving the presentation in PowerPoint."
                    )

                if result["notes_text"].strip():
                    st.markdown("#### 📋 Extracted notes")
                    st.text_area(
                        "Review the notes below before downloading",
                        value=result["notes_text"],
                        height=360,
                        label_visibility="collapsed",
                    )
                    safe_name = Path(uploaded_pptx.name).stem
                    st.download_button(
                        "⬇️ Download notes as TXT",
                        data=result["notes_text"].encode("utf-8"),
                        file_name=f"{safe_name}_speaker_notes.txt",
                        mime="text/plain",
                        type="primary",
                    )
                else:
                    st.info(
                        "No readable speaker notes were found. Please check the "
                        "slide review information above and verify the notes in PowerPoint."
                    )

            except Exception as exc:
                st.error(
                    "We couldn't process this PowerPoint file. Please check that it "
                    "is a valid, unencrypted PPTX file."
                )
                with st.expander("Show technical details"):
                    st.code(str(exc))

    else:
        st.markdown(
            '<p class="small-note">💡 Tip: speaker notes are different from the text visible on the slides.</p>',
            unsafe_allow_html=True,
        )


with audio_tab:
    st.markdown(
        """
        <div class="tool-card">
          <div class="tool-icon">🎧</div>
          <div class="tool-title">MP4 to MP3 Converter</div>
          <p class="tool-description">Extract the audio track from a video and save it as an MP3 file.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_mp4 = st.file_uploader(
        "Choose an MP4 video",
        type=["mp4"],
        help="Upload an MP4 file that contains an audio track.",
        key="mp4_upload",
    )

    if uploaded_mp4 is not None:
        size_mb = uploaded_mp4.size / (1024 * 1024)
        st.caption(f"🎬 Selected file: {uploaded_mp4.name}  ·  {size_mb:.1f} MB")

        if st.button("🎵 Convert to MP3", key="convert_mp3_button", type="primary"):
            try:
                with st.spinner("Extracting audio… larger videos may take a few minutes."):
                    mp3_bytes = convert_mp4_to_mp3(uploaded_mp4)

                safe_name = Path(uploaded_mp4.name).stem
                st.success("Your MP3 is ready!")
                st.download_button(
                    "⬇️ Download MP3",
                    data=mp3_bytes,
                    file_name=f"{safe_name}.mp3",
                    mime="audio/mpeg",
                    type="primary",
                )
            except subprocess.TimeoutExpired:
                st.error("The conversion took too long. Try a shorter video.")
            except Exception as exc:
                st.error("We couldn't convert this MP4 file.")
                with st.expander("Show technical details"):
                    st.code(str(exc))

    else:
        st.markdown(
            '<p class="small-note">💡 Tip: the MP3 will usually be much smaller than the original video.</p>',
            unsafe_allow_html=True,
        )

st.markdown("---")
st.markdown(
    '<p class="small-note" style="text-align:center;">Made for a smoother content workflow · PPT notes & audio tools</p>',
    unsafe_allow_html=True,
)
