import streamlit as st
from pptx import Presentation
from io import StringIO
import tempfile


st.title("PowerPoint Notes Extractor")

uploaded_file = st.file_uploader(
    "Upload a PowerPoint file",
    type=["pptx"]
)

if uploaded_file:

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".pptx"
    ) as tmp:

        tmp.write(uploaded_file.read())
        ppt_path = tmp.name

    prs = Presentation(ppt_path)

    all_notes = []

    for slide_num, slide in enumerate(prs.slides, start=1):

        if slide.has_notes_slide:

            notes_slide = slide.notes_slide

            notes_text = notes_slide.notes_text_frame.text.strip()

            if (
                notes_text and
                notes_text.lower() != "click to add notes"
            ):

                all_notes.append(
                    f"SLIDE {slide_num}\n\n{notes_text}"
                )

    output_text = "\n\n" + ("-" * 50) + "\n\n".join(all_notes)

    st.success(
        f"Extracted notes from {len(all_notes)} slides."
    )

    st.download_button(
        label="Download TXT File",
        data=output_text,
        file_name="presentation_notes.txt",
        mime="text/plain"
    )