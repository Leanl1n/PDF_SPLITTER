import os
import re
import shutil
import tempfile
import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
import base64
import zipfile
from io import BytesIO

def extract_text_from_page(pdf_reader, page_number):
    """
    Extract text from a specific page of a PDF
    
    Args:
        pdf_reader: PdfReader object
        page_number: Page number to extract text from
    
    Returns:
        str: Extracted text from the page
    """
    try:
        page = pdf_reader.pages[page_number]
        return page.extract_text()
    except Exception as e:
        st.error(f"Error extracting text from page {page_number+1}: {e}")
        return ""

def extract_name_from_text(text):
    """
    Extract a name from certificate text
    
    Args:
        text (str): Text from certificate
    
    Returns:
        str: Extracted name, or None if not found
    """
    # Clean input text first
    text = text.replace('\x00', '')
    if not text:
        return None
    
    # Extract name using various methods
    name = None
    
    # Method 1: Look for name after "CERTIFICATE OF COMPLETION"
    # This appears to be the most reliable pattern for these certificates
    match = re.search(r'CERTIFICATE OF COMPLETION\s+([A-Za-z\s\.\-]+)(?=\s+This certificate)', text, re.DOTALL)
    if match:
        name = match.group(1).strip()
        # Clean up any extra whitespace or newlines
        name = re.sub(r'\s+', ' ', name).strip()
        
        # Avoid capturing titles or headers
        if not any(title in name for title in ["Chief Operating", "Operations Manager"]):
            name = name
    
    # Method 2: Look for lines between CERTIFICATE OF COMPLETION and "This certificate"
    if not name:
        lines = text.split('\n')
        certificate_index = -1
        this_certificate_index = -1
        
        for i, line in enumerate(lines):
            if 'CERTIFICATE OF COMPLETION' in line:
                certificate_index = i
            if 'This certificate' in line:
                this_certificate_index = i
                break
        
        if certificate_index != -1 and this_certificate_index != -1 and certificate_index < this_certificate_index:
            # Check the lines between these markers
            for i in range(certificate_index + 1, this_certificate_index):
                candidate = lines[i].strip()
                if candidate and not candidate.startswith("Chief") and not candidate.endswith("Officer") and not "Manager" in candidate:
                    name = candidate
                    break
    
    # Method 3: Look for a name pattern with special handling for Sta./Sto. abbreviations
    if not name:
        # Enhanced pattern to capture names with Sta./Sto. abbreviations
        patterns = [
            # Pattern for names with Sta./Sto. abbreviations
            r'([A-Z][a-z]+)\s+(Sta\.|Sto\.)\s+([A-Z][a-z]+)',
            # Regular name patterns
            r'([A-Z][a-z]+|[A-Z]{2,})\s+((?:[A-Za-z]\.?\s+)?(?:[a-z]{1,3}\s+)?[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)'
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                candidate = match.group(0).strip()  # Use the full match
                context = text[max(0, match.start() - 20):min(len(text), match.end() + 20)]
                
                # Skip if this looks like a title
                if any(title in candidate for title in ["Chief Operating", "Operations Manager"]):
                    continue
                    
                # Good candidates are not near these keywords
                if "Officer" not in context and "Manager" not in context:
                    name = candidate
                    break
            
            if name:
                break
    
    # Method 4: Try the presented to pattern as a last resort
    if not name:
        match = re.search(r'presented to\s+(.*?)\s+for successfully', text, re.DOTALL)
        if match:
            name = match.group(1).strip()
    
    # If we found a name, clean it up
    if name:
        # SPECIFIC CLEANUP: Remove the word "COMPLETION" (case insensitive)
        name = re.sub(r'COMPLETION', '', name, flags=re.IGNORECASE)
        name = re.sub(r'CERTIFICATE', '', name, flags=re.IGNORECASE)
        
        # Remove "This" if present
        name = re.sub(r'\bThis\b', '', name)
        
        # Clean up any extra spaces that resulted from the removals
        name = re.sub(r'\s+', ' ', name).strip()
    
    return name

def format_name_for_filename(name):
    """
    Format a name for filename (LastName, FirstName)
    Handles special cases like name prefixes and initials
    
    Args:
        name (str): Full name with possible middle initial
        
    Returns:
        str: Formatted name for filename
    """
    if not name:
        return "unknown"
    
    # Remove problematic characters
    clean_name = re.sub(r'[\n\r\t\\/:*?"<>|]', '', name)
    
    # SPECIFIC CLEANUP: Remove the word "COMPLETION" (case insensitive)
    clean_name = re.sub(r'COMPLETION', '', clean_name, flags=re.IGNORECASE)
    clean_name = re.sub(r'CERTIFICATE', '', clean_name, flags=re.IGNORECASE)
    
    # Remove "This" if present
    clean_name = re.sub(r'\bThis\b', '', clean_name)
    
    # Clean up any extra spaces that resulted from the removals
    clean_name = re.sub(r'\s+', ' ', clean_name).strip()
    
    # Split the name into parts
    name_parts = clean_name.split()
    
    if len(name_parts) < 2:
        return clean_name
    
    # List of common name prefixes that should be kept with the last name
    name_prefixes = ['de', 'del', 'dela', 'della', 'des', 'di', 'du', 'el', 'la', 'le', 
                     'van', 'von', 'der', 'den', 'das', 'dos', 'da', 'do', 'san', 'st',
                     'sta.', 'sto.', 'sta', 'sto']
    
    # First name is always the first part
    first_name = name_parts[0]
    
    # Special case for names with Sta./Sto.
    if len(name_parts) >= 3:
        if name_parts[1].lower() in ['sta.', 'sto.', 'sta', 'sto']:
            first_name = name_parts[0]
            prefix = name_parts[1] 
            last_part = " ".join(name_parts[2:])
            last_name = f"{prefix} {last_part}"
            return f"{last_name}, {first_name}"
    
    # Handle last name with prefixes:
    # Case 1: "Godwin de Guzman" -> last_name = "de Guzman"
    # Case 2: "John van der Wal" -> last_name = "van der Wal"
    # Case 3: "Rafa Sta. Ana" -> last_name = "Sta. Ana"
    last_name = ""
    
    if len(name_parts) >= 3:
        # Check for Sta./Sto. pattern (e.g., "Rafa Sta. Ana")
        if name_parts[-2].lower().replace('.', '') in ['sta', 'sto']:
            last_name = f"{name_parts[-2]} {name_parts[-1]}"
        # Check if second-to-last part is a prefix
        elif name_parts[-2].lower() in name_prefixes:
            # Use last two parts as last name
            last_name = f"{name_parts[-2]} {name_parts[-1]}"
        # Check if third-to-last is a prefix (for double prefixes like "van der")
        elif len(name_parts) >= 4 and name_parts[-3].lower() in name_prefixes and name_parts[-2].lower() in name_prefixes:
            # Use last three parts as last name
            last_name = f"{name_parts[-3]} {name_parts[-2]} {name_parts[-1]}"
        else:
            # No prefix detected, use the last part
            last_name = name_parts[-1]
    else:
        # Just two parts, use the last part
        last_name = name_parts[-1]
    
    # Format as "LastName, FirstName"
    formatted_name = f"{last_name}, {first_name}"
    
    return formatted_name

def split_pdf_with_names(uploaded_file, temp_dir, organize_into_folders=False):
    """
    Split a PDF into individual pages and name each page using extracted text
    
    Args:
        uploaded_file: Streamlit's UploadedFile object
        temp_dir: Directory to save files temporarily
        organize_into_folders: Whether to organize files by person
        
    Returns:
        tuple: (List of file paths, total pages)
    """
    # Save the uploaded file to a temporary location
    pdf_path = os.path.join(temp_dir, "temp.pdf")
    with open(pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # Create a directory for the split pages
    pages_dir = os.path.join(temp_dir, "pages")
    os.makedirs(pages_dir, exist_ok=True)
    
    # List to store the output files
    output_files = []
    
    try:
        # Open the PDF
        pdf = PdfReader(pdf_path)
        total_pages = len(pdf.pages)
        st.info(f"Processing PDF with {total_pages} pages...")
        
        # Setup progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Process each page
        for i in range(total_pages):
            # Update progress
            progress = (i + 1) / total_pages
            progress_bar.progress(progress)
            status_text.text(f"Processing page {i+1} of {total_pages}...")
            
            # Extract text from the page
            text = extract_text_from_page(pdf, i)
            
            # Try to extract a name from the text
            name = extract_name_from_text(text)
            
            # Make sure "This" doesn't get into the name
            if name and "This" in name:
                name = name.replace("This", "").strip()
            
            # Create output filename
            if name:
                # Clean and format the name for the filename
                formatted_name = format_name_for_filename(name)
                
                # Format page number as 3-digit number
                page_num = str(i+1).zfill(3)
                
                # Final filename format: "001 Moreno, Michelle.pdf"
                output_filename = f"{page_num} {formatted_name}.pdf"
                st.write(f"Extracted name from page {i+1}: {name}")
            else:
                # Fallback to generic name if no name found
                page_num = str(i+1).zfill(3)
                output_filename = f"{page_num} unknown.pdf"
                st.write(f"No name found in page {i+1}, using generic filename")
            
            # Create full output path
            output_path = os.path.join(pages_dir, output_filename)
            
            # Create a new PDF with just this page
            output = PdfWriter()
            output.add_page(pdf.pages[i])
            
            # Save the individual page
            with open(output_path, "wb") as output_file:
                output.write(output_file)
            
            output_files.append(output_path)
        
        # If organizing into folders, do that
        if organize_into_folders:
            # Create a directory for the organized files
            organized_dir = os.path.join(temp_dir, "organized")
            os.makedirs(organized_dir, exist_ok=True)
            
            st.info("Organizing files into folders by person...")
            
            # Track created folders to avoid duplicates
            created_folders = set()
            organized_files = []
            
            # Process each PDF file
            for pdf_file in os.listdir(pages_dir):
                if pdf_file.lower().endswith('.pdf'):
                    # Extract the person's name from the filename
                    # Expected format: "001 LastName, FirstName.pdf"
                    match = re.match(r'\d+\s+(.*?)\.pdf$', pdf_file)
                    
                    if match:
                        # Get the name part without the numbering
                        formatted_name = match.group(1).strip()
                        
                        # Skip unknown files
                        if formatted_name.lower() == "unknown":
                            continue
                        
                        # Create a folder for this person if it doesn't exist
                        folder_name = formatted_name  # Use the formatted name as folder name
                        folder_path = os.path.join(organized_dir, folder_name)
                        
                        if folder_name not in created_folders:
                            os.makedirs(folder_path, exist_ok=True)
                            created_folders.add(folder_name)
                        
                        # Copy the PDF file to the person's folder, removing the numbering
                        source_file = os.path.join(pages_dir, pdf_file)
                        # New filename without the numbering
                        new_filename = f"{formatted_name}.pdf"
                        dest_file = os.path.join(folder_path, new_filename)
                        
                        # Copy the file
                        shutil.copy2(source_file, dest_file)
                        organized_files.append(dest_file)
            
            st.success(f"Successfully organized files into {len(created_folders)} folders")
            
            # Return the organized files instead
            return organized_files, total_pages
        
        # Complete the progress bar
        progress_bar.progress(1.0)
        status_text.text(f"Completed processing {total_pages} pages!")
        
        return output_files, total_pages
        
    except Exception as e:
        st.error(f"Error processing PDF: {e}")
        return [], 0

def create_download_zip(file_paths, temp_dir):
    """
    Create a ZIP file containing all processed PDFs
    
    Args:
        file_paths: List of file paths to include
        temp_dir: Temporary directory
        
    Returns:
        str: Path to the ZIP file
    """
    # Create a ZIP file
    zip_path = os.path.join(temp_dir, "certificate_pages.zip")
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in file_paths:
            # Add each file to the ZIP
            arcname = os.path.basename(file)
            zipf.write(file, arcname=arcname)
    
    return zip_path

def get_binary_file_downloader_html(bin_file, file_label='File'):
    """
    Generate HTML code for file download
    """
    with open(bin_file, 'rb') as f:
        data = f.read()
    
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/zip;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download {file_label}</a>'
    return href

def main():
    st.set_page_config(
        page_title="Certificate PDF Splitter",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 Certificate PDF Splitter")
    st.write("Upload a PDF file with certificates to split it into individual pages and extract names.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
    
    # Option to organize files
    organize_files = st.checkbox("Organize files into folders by person", value=True)
    
    if uploaded_file is not None:
        # Create a temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Process button
            if st.button("Process PDF"):
                with st.spinner("Processing PDF..."):
                    # Split the PDF and get the file paths
                    file_paths, total_pages = split_pdf_with_names(uploaded_file, temp_dir, organize_files)
                    
                    if file_paths:
                        st.success(f"Successfully processed {total_pages} pages from the PDF.")
                        
                        # Create a ZIP file for download
                        zip_path = create_download_zip(file_paths, temp_dir)
                        
                        # Provide a download link
                        st.markdown(get_binary_file_downloader_html(zip_path, 'Processed Files (ZIP)'), unsafe_allow_html=True)
                    else:
                        st.error("Failed to process the PDF.")

if __name__ == "__main__":
    main()
