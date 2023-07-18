import os
import zipfile
import streamlit as st
import shutil
import textract

def convert_doc_to_txt(input_dir, output_dir):
    # Check if the output directory exists, create it if not
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Get a list of all files in the input directory
    files = os.listdir(input_dir)

    # Iterate over the files and convert .doc and .docx files to .txt
    for file in files:
        if file.endswith(('.doc', '.docx')):
            # For .doc and .docx files, use textract library to extract text
            doc_path = os.path.join(input_dir, file)
            text = textract.process(doc_path, encoding='utf-8')

            # Create the output .txt file path
            txt_filename = os.path.splitext(file)[0] + '.txt'
            txt_path = os.path.join(output_dir, txt_filename)

            # Save the extracted text as .txt
            with open(txt_path, 'wb') as txt_file:
                txt_file.write(text)

            print(f"Converted {file} to {txt_filename}")

def process_zip_files(zip_file):
    # Create a directory to extract the zip contents
    output_dir = "extracted_files"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Extract the contents of the zip file
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

    # Process the extracted files and convert .doc and .docx to .txt
    convert_doc_to_txt(output_dir, output_dir)

    # Create a new zip file containing the processed .txt files
    processed_zip_file = "processed_files.zip"
    with zipfile.ZipFile(processed_zip_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        for root, _, files in os.walk(output_dir):
            for file in files:
                if file.endswith('.txt'):
                    txt_path = os.path.join(root, file)
                    zip_out.write(txt_path, os.path.relpath(txt_path, output_dir))

    # Remove the temporary extracted directory
    shutil.rmtree(output_dir)

    return processed_zip_file

def main():
    st.set_page_config(
        page_title="FileAlchemy",
        page_icon="📑",
    )
    st.title("FileAlchemy")

    uploaded_file = st.file_uploader("Choose a zip file", type="zip")
    st.write("Note: Upload a zip file containing .doc and .docx files")

    if uploaded_file is not None:
        # Create a temporary directory to process the zip file
        temp_dir = "temp_zip"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # Save the uploaded zip file to the temporary directory
        zip_path = os.path.join(temp_dir, "uploaded.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.snow()
        # Process the zip file and get the processed zip file path
        processed_zip_file = process_zip_files(zip_path)
        st.success('This is a success message!', icon="✅")
        # Offer the processed zip file for download
        st.download_button(
            label="Download Processed Zip File",
            data=open(processed_zip_file, "rb").read(),
            file_name="processed_files.zip",
        )

        # Clean up the temporary directory
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    main()
