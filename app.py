
import os
import subprocess
import tempfile
from io import BytesIO

import streamlit as st
from pptx import Presentation
import imageio_ffmpeg


# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="Learning Content Utilities",
    page_icon="🎓",
    layout="wide"
)

st.title("🎓 Learning Content Utilities")
st.write(
    "Tools for extracting PowerPoint speaker notes "
    "and converting video recordings into MP3 audio."
)


# --------------------------------------------------
# HELPER: EXTRACT SPEAKER NOTES
# --------------------------------------------------

def extract_ppt_notes(pptx_file):
    """
    Extract speaker notes from each PowerPoint slide.
    Skip empty notes and the default placeholder text.
    Continue if an individual slide has inaccessible notes.
    """

    prs = Presentation(pptx_file)
    notes_list = []
    skipped_slides = []

    for slide_number, slide in enumerate(prs.slides, start=1):
        try:
            notes_slide = slide.notes_slide
            notes_frame = notes_slide.notes_text_frame

            if notes_frame is None:
                skipped_slides.append(slide_number)
                continue

            text = notes_frame.text.strip()

            # Ignore empty/default PowerPoint notes.
            if not text:
                continue

            if text.lower() in (
                "click to add notes",
                "click to edit master text styles"
            ):
                continue

            notes_list.append(
                f"Slide {slide_number}\n{text}"
            )

        except (AttributeError, ValueError):
            skipped_slides.append(slide_number)
            continue

    return "\n\n".join(notes_list), skipped_slides, len(prs.slides)


# --------------------------------------------------
# HELPER: CONVERT MP4 TO MP3 USING FFMPEG
# --------------------------------------------------

def convert_mp4_to_mp3(input_path, output_path):
    """
    Extract audio from an MP4 file using FFmpeg.
    Raises an error if the file has no audio or conversion fails.
    """

    ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()

    command = [
        ffmpeg_path,
        "-y",
        "-i", input_path,
        "-map", "0:a:0",
        "-vn",
        "-codec:a", "libmp3lame",
        "-q:a", "2",
        output_path
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=3600
    )

    if result.returncode != 0:
        error_message = result.stderr or "Unknown FFmpeg error."
        raise RuntimeError(error_message[-2500:])

    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("The MP3 file could not be created.")


# --------------------------------------------------
# TABS
# --------------------------------------------------

tab1, tab2 = st.tabs([
    "📄 PPT Notes Extractor",
    "🎥 MP4 to MP3 Extractor"
])


# ==================================================
# TAB 1: PPT NOTES EXTRACTOR
# ==================================================

with tab1:

    st.header("📄 PPT Notes Extractor")

    st.write(
        "Upload a PowerPoint file to extract its speaker notes "
        "into a downloadable text file."
    )

    uploaded_pptx = st.file_uploader(
        "Upload PowerPoint file",
        type=["pptx"],
        key="pptx_upload"
    )

    if uploaded_pptx is not None:

        st.write(f"**File:** {uploaded_pptx.name}")

        if st.button(
            "Extract Speaker Notes",
            key="extract_notes_button"
        ):

            try:
                with st.spinner("Extracting speaker notes..."):

                    uploaded_pptx.seek(0)

                    notes_text, skipped_slides, total_slides = (
                        extract_ppt_notes(uploaded_pptx)
                    )

                if notes_text.strip():

                    st.success(
                        f"Notes extracted from a presentation "
                        f"containing {total_slides} slides."
                    )

                    st.text_area(
                        "Extracted Speaker Notes",
                        value=notes_text,
                        height=400,
                        key="extracted_notes_preview"
                    )

                    st.download_button(
                        label="⬇️ Download Notes as TXT",
                        data=notes_text.encode("utf-8"),
                        file_name=(
                            os.path.splitext(uploaded_pptx.name)[0]
                            + "_notes.txt"
                        ),
                        mime="text/plain",
                        key="download_notes_button"
                    )

                else:
                    st.warning(
                        "No speaker notes were found. "
                        "Check whether the PowerPoint contains notes."
                    )

                if skipped_slides:
                    st.warning(
                        "Notes could not be accessed on slide(s): "
                        + ", ".join(map(str, skipped_slides))
                        + ". Other slides were processed normally."
                    )

            except Exception as e:
                st.error(
                    "Could not process this PowerPoint file. "
                    "Please check that it is a valid, unencrypted "
                    "PPTX file."
                )
                st.exception(e)


# ==================================================
# TAB 2: MP4 TO MP3 EXTRACTOR
# ==================================================

with tab2:

    st.header("🎥 MP4 to MP3 Extractor")

    st.write(
        "Upload a video recording to extract its audio "
        "and download it as an MP3 file."
    )

    st.info(
        "Maximum upload size depends on your Streamlit configuration. "
        "Large videos may take longer to upload and process."
    )

    uploaded_video = st.file_uploader(
        "Upload MP4 video",
        type=["mp4"],
        key="mp4_upload"
    )

    if uploaded_video is not None:

        file_size_mb = len(uploaded_video.getbuffer()) / (1024 * 1024)

        st.write(f"**File:** {uploaded_video.name}")
        st.write(f"**Size:** {file_size_mb:.2f} MB")

        if st.button(
            "Extract Audio",
            key="extract_audio_button"
        ):

            with tempfile.TemporaryDirectory() as temp_dir:

                input_path = os.path.join(temp_dir, "input.mp4")
                output_path = os.path.join(temp_dir, "extracted_audio.mp3")

                try:
                    with st.spinner(
                        "Extracting audio. Please wait..."
                    ):

                        uploaded_video.seek(0)

                        with open(input_path, "wb") as input_file:
                            input_file.write(uploaded_video.getbuffer())

                        convert_mp4_to_mp3(input_path, output_path)

                        with open(output_path, "rb") as mp3_file:
                            mp3_data = mp3_file.read()

                    st.success("Audio extraction completed!")

                    st.audio(mp3_data, format="audio/mpeg")

                    original_name = os.path.splitext(
                        uploaded_video.name
                    )[0]

                    st.download_button(
                        label="⬇️ Download MP3",
                        data=mp3_data,
                        file_name=original_name + ".mp3",
                        mime="audio/mpeg",
                        key="download_mp3_button"
                    )

                except subprocess.TimeoutExpired:
                    st.error(
                        "Processing took too long. "
                        "Try a shorter video or process it locally."
                    )

                except RuntimeError as e:
                    error_text = str(e)

                    if (
                        "matches no streams" in error_text.lower()
                        or "does not contain any stream" in error_text.lower()
                    ):
                        st.error(
                            "No audio track was found in this video. "
                            "Please upload a video that contains audio."
                        )
                    else:
                        st.error(
                            "Audio extraction failed. "
                            "The video may be damaged or use an "
                            "unsupported codec."
                        )
                        with st.expander("Technical details"):
                            st.code(error_text)

                except Exception as e:
                    st.error("An unexpected error occurred.")
                    st.exception(e)
