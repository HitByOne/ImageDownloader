import streamlit as st
import pandas as pd
import requests
import io
import zipfile

def download_images_to_zip(df, zip_filename):
    zip_buffer = io.BytesIO()
    errors = []

    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for index, row in df.iterrows():
            item_number = str(row.get('Item', f"row_{index}"))
            image_url = str(row.get('Image', '')).strip()

            if not image_url:
                errors.append(f"{item_number}: Missing image URL")
                continue

            # --- Try to download image ---
            for attempt, url_variant in enumerate([image_url, image_url.replace("http://", "https://")]):
                try:
                    response = requests.get(url_variant, timeout=10)
                    response.raise_for_status()
                    image_data = response.content
                    zipf.writestr(f"{item_number}_{attempt+1}.jpg", image_data)
                    st.write(f"✅ {item_number} downloaded from {url_variant}")
                    break
                except requests.exceptions.RequestException as e:
                    if attempt == 1:  # both HTTP & HTTPS failed
                        st.warning(f"⚠️ Skipped {item_number}: {e}")
                        errors.append(f"{item_number}: {url_variant} failed ({e})")

        # Add log file if there were failures
        if errors:
            error_log = "\n".join(errors)
            zipf.writestr("error_log.txt", error_log)

    zip_buffer.seek(0)
    return zip_buffer


def main():
    st.title("📸 Excel Image Downloader and Zipper")

    st.markdown("""
    **Instructions:**
    1. Upload an Excel file containing columns:
       - `Item` → file name base for the image  
       - `Image` → image URL
    2. Enter a name for your ZIP file.
    3. Click **Start Download**.
    """)

    uploaded_file = st.file_uploader("Select an Excel file", type=["xlsx"])
    zip_filename = st.text_input("Enter ZIP file name", "images.zip")

    if st.button("Start Download"):
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file)
                zip_buffer = download_images_to_zip(df, zip_filename)

                st.success("✅ Download process completed.")
                st.download_button(
                    label="Download ZIP File",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip"
                )
            except Exception as e:
                st.error(f"Error processing the file: {e}")
        else:
            st.error("Please upload an Excel file.")

if __name__ == "__main__":
    main()
