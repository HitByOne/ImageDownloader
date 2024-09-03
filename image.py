import streamlit as st
import pandas as pd
import requests
import io
import zipfile

def download_images_to_zip(df, zip_filename):
    # Create an in-memory ZIP file
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        # Iterate through each row in the DataFrame
        for index, row in df.iterrows():
            item_number = str(row['Item'])  # Replace 'Item' with the actual header name
            image_url = row['Image']  # Replace 'Image' with the actual header name

            # Request the image from the URL
            try:
                response = requests.get(image_url)
                response.raise_for_status()

                # Save the image in the ZIP file
                image_data = response.content
                zipf.writestr(f"{item_number}.jpg", image_data)

                st.write(f"Downloaded and added to ZIP: {item_number}.jpg")
            except requests.exceptions.RequestException as e:
                st.error(f"Failed to download {image_url}: {e}")

    # Seek to the beginning of the BytesIO object
    zip_buffer.seek(0)
    
    return zip_buffer

def main():
    st.title("Excel Image Downloader and Zipper")

    # Add instructions for the user
    st.markdown("""
    **Instructions:**
    - The uploaded Excel file must contain two columns:
        1. **Item**: The name of the file you want for each image.
        2. **Image**: The URL link to the image you want to download.
    """)

    # Step 1: Upload Excel file
    uploaded_file = st.file_uploader("Select an Excel file", type="xlsx")
    
    # Step 2: Enter the ZIP file name
    zip_filename = st.text_input("Enter the name for the ZIP file", "images.zip")

    if st.button("Start Download"):
        if uploaded_file is not None:
            try:
                # Load the Excel file
                df = pd.read_excel(uploaded_file)

                # Start the download process
                zip_buffer = download_images_to_zip(df, zip_filename)
                
                # Provide the download link
                st.download_button(
                    label="Download ZIP",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip"
                )
            except Exception as e:
                st.error(f"Error processing the Excel file: {e}")
        else:
            st.error("Please upload an Excel file.")

if __name__ == "__main__":
    main()
