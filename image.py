import streamlit as st
import pandas as pd
import requests
import os
import zipfile

def download_images_to_zip(df, save_dir, zip_filename):
    # Ensure the directory exists for saving images
    os.makedirs(save_dir, exist_ok=True)

    # Create a ZIP file to store images
    zip_path = os.path.join(save_dir, zip_filename)
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        # Iterate through each row in the DataFrame
        for index, row in df.iterrows():
            item_number = str(row['Item'])  # Replace 'Item' with the actual header name
            image_url = row['Image']  # Replace 'Image' with the actual header name

            # Request the image from the URL
            try:
                response = requests.get(image_url)
                response.raise_for_status()

                # Save the image temporarily with the item number as the file name
                image_path = os.path.join(save_dir, f"{item_number}.jpg")
                with open(image_path, 'wb') as file:
                    file.write(response.content)

                # Add the image to the ZIP file
                zipf.write(image_path, arcname=f"{item_number}.jpg")

                # Remove the temporary image file
                os.remove(image_path)

                st.write(f"Downloaded and added to ZIP: {item_number}.jpg")
            except requests.exceptions.RequestException as e:
                st.error(f"Failed to download {image_url}: {e}")

    st.success(f"Image download and ZIP creation process completed. ZIP file saved at: {zip_path}")

def main():
    st.title("Excel Image Downloader and Zipper")

    # Step 1: Upload Excel file
    uploaded_file = st.file_uploader("Select an Excel file", type="xlsx")
    
    # Step 2: Select a directory to save images and ZIP file
    save_dir = st.text_input("Enter the directory to save images and ZIP file")

    # Step 3: Enter the ZIP file name
    zip_filename = st.text_input("Enter the name for the ZIP file", "images.zip")

    if st.button("Start Download"):
        if uploaded_file is not None and save_dir:
            # Load the Excel file
            try:
                df = pd.read_excel(uploaded_file)

                # Start the download process
                download_images_to_zip(df, save_dir, zip_filename)
            except Exception as e:
                st.error(f"Error processing the Excel file: {e}")
        else:
            st.error("Please upload an Excel file and enter a directory path.")

if __name__ == "__main__":
    main()

