import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import time
from pathlib import Path
from urllib.parse import urlparse
from typing import Tuple, List, Dict
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Supported image formats
SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.tiff'}

def get_image_extension(response) -> str:
    """Detect image format from Content-Type header or default to jpg"""
    content_type = response.headers.get('content-type', '').lower()
    
    format_map = {
        'image/jpeg': '.jpg',
        'image/png': '.png',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/bmp': '.bmp',
        'image/tiff': '.tiff',
    }
    
    for mime_type, ext in format_map.items():
        if mime_type in content_type:
            return ext
    
    return '.jpg'  # Default fallback

def download_image(url: str, timeout: int = 10, max_retries: int = 3) -> Tuple[bytes, str, str]:
    """
    Download an image with retry logic and format detection.
    
    Returns:
        Tuple of (image_data, extension, status_message)
    """
    url_variants = [url]
    
    # Add HTTP/HTTPS variant if not already present
    if url.startswith('http://'):
        url_variants.append(url.replace('http://', 'https://'))
    elif url.startswith('https://'):
        url_variants.append(url.replace('https://', 'http://'))
    
    for variant_idx, current_url in enumerate(url_variants):
        for attempt in range(1, max_retries + 1):
            try:
                response = requests.get(current_url, timeout=timeout, allow_redirects=True)
                response.raise_for_status()
                
                extension = get_image_extension(response)
                image_data = response.content
                
                if len(image_data) == 0:
                    raise ValueError("Empty image data received")
                
                status = f"✅ Downloaded successfully from {current_url}"
                return image_data, extension, status
                
            except requests.exceptions.Timeout:
                status_msg = f"Timeout (attempt {attempt}/{max_retries})"
                if attempt == max_retries and variant_idx == len(url_variants) - 1:
                    return None, None, f"⚠️ Failed after {max_retries} attempts: Connection timeout"
                    
            except requests.exceptions.ConnectionError as e:
                status_msg = f"Connection error (attempt {attempt}/{max_retries})"
                if attempt == max_retries and variant_idx == len(url_variants) - 1:
                    return None, None, f"⚠️ Connection failed after {max_retries} attempts"
                    
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    return None, None, f"⚠️ Image not found (404)"
                elif e.response.status_code == 403:
                    return None, None, f"⚠️ Access denied (403)"
                else:
                    return None, None, f"⚠️ HTTP Error {e.response.status_code}"
                    
            except ValueError as e:
                return None, None, f"⚠️ Invalid image data: {str(e)}"
                
            except Exception as e:
                if attempt == max_retries and variant_idx == len(url_variants) - 1:
                    return None, None, f"⚠️ Failed: {str(e)[:100]}"
    
    return None, None, "⚠️ All download attempts failed"

def download_images_to_zip(df: pd.DataFrame, zip_filename: str, progress_callback=None) -> io.BytesIO:
    """
    Download images from URLs in DataFrame and create a ZIP file.
    
    Args:
        df: DataFrame with 'Item' and 'Image' columns
        zip_filename: Output ZIP filename
        progress_callback: Function to call with (current, total) for progress tracking
    
    Returns:
        BytesIO object containing the ZIP file
    """
    zip_buffer = io.BytesIO()
    errors = []
    successes = []
    
    total_rows = len(df)
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for index, row in df.iterrows():
            # Update progress
            if progress_callback:
                progress_callback(index + 1, total_rows)
            
            item_number = str(row.get('Item', f"row_{index}")).strip()
            if not item_number:
                item_number = f"row_{index}"
            
            # Sanitize filename
            item_number = "".join(c if c.isalnum() or c in ('_', '-') else '_' for c in item_number)
            
            image_url = str(row.get('Image', '')).strip()
            
            if not image_url or image_url.lower() in ('nan', 'none', ''):
                error_msg = f"{item_number}: Missing or empty image URL"
                errors.append(error_msg)
                continue
            
            # Download image
            image_data, extension, status_msg = download_image(image_url)
            
            if image_data:
                # Generate unique filename
                filename = f"{item_number}{extension}"
                
                try:
                    zipf.writestr(filename, image_data)
                    successes.append((item_number, status_msg))
                except Exception as e:
                    error_msg = f"{item_number}: Failed to add to ZIP: {str(e)}"
                    errors.append(error_msg)
            else:
                # status_msg already contains the error
                errors.append(f"{item_number}: {status_msg}")
            
            # Small delay to prevent rate limiting
            time.sleep(0.1)
    
    # Create detailed log file
    log_content = generate_log_report(successes, errors, total_rows)
    
    # Add log to ZIP
    with zipfile.ZipFile(zip_buffer, 'a') as zipf:
        zipf.writestr("download_report.txt", log_content)
    
    zip_buffer.seek(0)
    return zip_buffer, successes, errors

def generate_log_report(successes: List[Tuple], errors: List[str], total_rows: int) -> str:
    """Generate a detailed download report"""
    report = []
    report.append("=" * 70)
    report.append("IMAGE DOWNLOAD REPORT")
    report.append("=" * 70)
    report.append("")
    report.append(f"Total Items Processed: {total_rows}")
    report.append(f"Successfully Downloaded: {len(successes)}")
    report.append(f"Failed Downloads: {len(errors)}")
    report.append(f"Success Rate: {(len(successes) / total_rows * 100):.1f}%")
    report.append("")
    report.append("-" * 70)
    report.append("SUCCESSFUL DOWNLOADS")
    report.append("-" * 70)
    
    if successes:
        for item, status in successes:
            report.append(f"✓ {item}: {status}")
    else:
        report.append("No successful downloads")
    
    report.append("")
    report.append("-" * 70)
    report.append("FAILED DOWNLOADS")
    report.append("-" * 70)
    
    if errors:
        for error in errors:
            report.append(f"✗ {error}")
    else:
        report.append("No failures")
    
    report.append("")
    report.append("=" * 70)
    
    return "\n".join(report)

def validate_dataframe(df: pd.DataFrame) -> Tuple[bool, str]:
    """Validate that DataFrame has required columns"""
    if df.empty:
        return False, "Excel file is empty"
    
    required_columns = ['Item', 'Image']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        available = ", ".join(df.columns.tolist())
        return False, f"Missing columns: {', '.join(missing_columns)}. Available: {available}"
    
    return True, "Valid"

def main():
    st.set_page_config(
        page_title="Image Downloader & Zipper",
        page_icon="📸",
        layout="wide"
    )
    
    st.title("📸 Excel Image Downloader and Zipper")
    
    # Create two columns for better layout
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### Instructions
        1. **Prepare your Excel file** with these columns:
           - `Item` → name/identifier for the image  
           - `Image` → full image URL (HTTP or HTTPS)
        
        2. **Upload** the Excel file (.xlsx or .csv supported)
        
        3. **Configure** retry settings and ZIP filename
        
        4. **Start Download** and monitor progress in real-time
        """)
    
    with col2:
        st.info("""
        **✨ Features**
        - 🔄 Auto retry logic
        - 📊 Progress tracking
        - 🎨 Auto format detection
        - 📋 Detailed reports
        - 🛡️ Error logging
        """)
    
    st.divider()
    
    # File upload section
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader(
            "Select an Excel or CSV file",
            type=["xlsx", "csv"],
            help="File must contain 'Item' and 'Image' columns"
        )
    
    with col2:
        zip_filename = st.text_input(
            "ZIP filename",
            value="images.zip",
            help="Name for your output ZIP file"
        )
    
    # Settings section
    with st.expander("⚙️ Advanced Settings"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            max_retries = st.slider(
                "Max retries per URL",
                min_value=1,
                max_value=5,
                value=3,
                help="Number of retry attempts for failed downloads"
            )
        
        with col2:
            timeout = st.slider(
                "Request timeout (seconds)",
                min_value=5,
                max_value=30,
                value=10,
                help="Seconds to wait before timing out a request"
            )
        
        with col3:
            st.markdown("")
            st.markdown("")
            st.info(f"📊 Timeout: {timeout}s | Retries: {max_retries}")
    
    st.divider()
    
    # Process button
    if st.button("🚀 Start Download", type="primary", use_container_width=True):
        if uploaded_file is None:
            st.error("❌ Please upload an Excel or CSV file first.")
            st.stop()
        
        try:
            # Read file
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # Validate
            is_valid, message = validate_dataframe(df)
            if not is_valid:
                st.error(f"❌ {message}")
                st.stop()
            
            st.success(f"✅ Loaded {len(df)} items from file")
            st.divider()
            
            # Progress section
            progress_bar = st.progress(0)
            status_text = st.empty()
            progress_info = st.empty()
            
            def update_progress(current, total):
                progress = current / total
                progress_bar.progress(progress)
                status_text.text(f"Processing: {current}/{total} items...")
                progress_info.write(f"**Progress:** {current}/{total} ({progress*100:.1f}%)")
            
            st.info("⏳ Downloading images... This may take a moment depending on file size.")
            
            # Download images
            zip_buffer, successes, errors = download_images_to_zip(
                df,
                zip_filename,
                progress_callback=update_progress
            )
            
            progress_bar.progress(1.0)
            status_text.text("✅ Download process completed!")
            
            st.divider()
            
            # Results summary
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Items", len(df))
            with col2:
                st.metric("✅ Successful", len(successes), delta=f"{(len(successes)/len(df)*100):.1f}%")
            with col3:
                st.metric("❌ Failed", len(errors), delta=f"{(len(errors)/len(df)*100):.1f}%")
            
            # Download button
            st.download_button(
                label="📥 Download ZIP File",
                data=zip_buffer,
                file_name=zip_filename,
                mime="application/zip",
                use_container_width=True
            )
            
            st.divider()
            
            # Detailed results
            if successes:
                with st.expander(f"✅ Successful Downloads ({len(successes)})", expanded=False):
                    for item, status in successes:
                        st.success(f"**{item}**: {status}")
            
            if errors:
                with st.expander(f"❌ Failed Downloads ({len(errors)})", expanded=True):
                    for error in errors:
                        st.error(error)
            
            st.info("📋 Full report included in download_report.txt inside the ZIP file")
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            logger.exception("Error in download_images_to_zip")

if __name__ == "__main__":
    main()
