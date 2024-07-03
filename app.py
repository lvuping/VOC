import streamlit as st
import tempfile
import os
import sys
from io import BytesIO
import email
from email.iterators import typed_subpart_iterator

# Import the load function from msg2eml.py
from msg2eml import load


def convert_msg_to_eml(msg_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp_file:
        tmp_file.write(msg_file.getvalue())
        tmp_file_path = tmp_file.name

    try:
        eml_content = load(tmp_file_path)
        return eml_content
    finally:
        os.unlink(tmp_file_path)


def parse_eml(eml_content):
    msg = email.message_from_bytes(eml_content.as_bytes())

    parsed_content = f"From: {msg['From']}\n"
    parsed_content += f"To: {msg['To']}\n"
    parsed_content += f"Subject: {msg['Subject']}\n"
    parsed_content += f"Date: {msg['Date']}\n\n"

    attachments = []

    def process_part(part, is_inline=False):
        content_type = part.get_content_type()
        filename = part.get_filename()
        if content_type == "text/plain":
            return part.get_payload(decode=True).decode("utf-8", errors="ignore")
        elif content_type.startswith("image/"):
            img_data = part.get_payload(decode=True)
            if is_inline:
                return f"\n[Inline Image: {filename}]\n"
            else:
                attachments.append(
                    (filename, img_data, content_type, part.get("Content-ID"))
                )
                return ""
        elif filename:
            attachments.append(
                (
                    filename,
                    part.get_payload(decode=True),
                    content_type,
                    part.get("Content-ID"),
                )
            )
            return ""
        return ""

    if msg.is_multipart():
        for part in msg.walk():
            if part.is_multipart():
                continue
            is_inline = part.get("Content-Disposition", "").startswith("inline")
            parsed_content += process_part(part, is_inline)
    else:
        parsed_content += process_part(msg)

    return parsed_content, attachments


st.title("MSG to EML Converter")

uploaded_file = st.file_uploader("Choose a MSG file", type="msg")

if uploaded_file is not None:
    st.write("File uploaded successfully!")

    # Convert MSG to EML
    eml_content = convert_msg_to_eml(uploaded_file)

    # Parse and display EML content
    parsed_eml, attachments = parse_eml(eml_content)
    st.subheader("EML Content:")
    st.text_area("", value=parsed_eml, height=300)

    # Display attachments
    if attachments:
        st.subheader("Attachments:")
        for filename, content, content_type, content_id in attachments:
            if content_type.startswith("image/"):
                st.image(content, caption=f"{filename} (Content-ID: {content_id})")
            else:
                st.download_button(
                    label=f"Download {filename} (Content-ID: {content_id})",
                    data=content,
                    file_name=filename,
                    mime=content_type,
                )

    # Provide download button for the entire EML file
    eml_filename = os.path.splitext(uploaded_file.name)[0] + ".eml"
    st.download_button(
        label="Download EML file",
        data=eml_content.as_bytes(),
        file_name=eml_filename,
        mime="message/rfc822",
    )
else:
    st.write("Please upload a MSG file to convert.")
