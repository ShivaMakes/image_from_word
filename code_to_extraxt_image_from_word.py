import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import re

def split_merged_numbers(text):
    """Attempt to split merged numbers based on a pattern of 8-10 digit sequences."""
    matches = re.findall(r'\d{8,10}', text)
    return matches if matches else [text]  # Return matches or original text if nothing found

def extract_images_from_docx_with_names(docx_path, output_folder, names_list, cleanup=False):
    # Clean names from Excel-style copy (strip spaces, ignore empty lines)
    names = [line.strip() for line in names_list if line.strip()]

    # Unzip the docx
    unzip_folder = os.path.splitext(docx_path)[0] + '_unzipped'
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_folder)

    document_xml_path = os.path.join(unzip_folder, 'word', 'document.xml')
    rels_path = os.path.join(unzip_folder, 'word', '_rels', 'document.xml.rels')
    media_folder = os.path.join(unzip_folder, 'word', 'media')

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    if not os.path.exists(rels_path):
        raise FileNotFoundError(f"Could not find document.xml.rels at {rels_path}")

    # Parse relationships
    rels_tree = ET.parse(rels_path)
    rels_root = rels_tree.getroot()
    rid_to_file = {}
    for rel in rels_root:
        rid = rel.attrib.get('Id')
        target = rel.attrib.get('Target')
        if target and target.startswith('media/'):
            rid_to_file[rid] = os.path.join(media_folder, os.path.basename(target))

    # Parse document.xml to find all rId image references in the order they appear
    doc_tree = ET.parse(document_xml_path)
    doc_root = doc_tree.getroot()

    image_refs = []
    for elem in doc_root.iter():
        embed = elem.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if embed:
            image_refs.append(embed)

    # Notify if names are more than images and log which names couldn't be used to rename an image
    if len(names) > len(image_refs):
        used_names_count = len(image_refs)
        names_without_images = names[used_names_count:]
        messagebox.showinfo("Notice", f"There are more names ({len(names)}) than images ({len(image_refs)}). The names that couldn't be assigned to an image will be logged.")
        log_path = os.path.join(output_folder, 'names_without_images_log.txt')
        with open(log_path, 'w') as log_file:
            log_file.write("These names were not assigned to any image (no image found for them):\n")
            log_file.write("\n".join(names_without_images))

    # Export and rename images according to names list in serial order
    extracted_files = []
    for count, rid in enumerate(image_refs):
        if rid in rid_to_file:
            source_file = rid_to_file[rid]
            ext = os.path.splitext(source_file)[1]
            if count < len(names):
                dest_file = os.path.join(output_folder, f'{names[count]}{ext}')
            else:
                dest_file = os.path.join(output_folder, f'Extra_Image_{count + 1}{ext}')
            shutil.copy2(source_file, dest_file)
            extracted_files.append(dest_file)
            print(f"Extracted and renamed: {dest_file}")

    print(f"All available images extracted and renamed to: {output_folder}")

    if cleanup:
        shutil.rmtree(unzip_folder)
        print(f"Cleaned up temporary files at: {unzip_folder}")

    return extracted_files

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("Select DOCX", "Please select the DOCX file")
    docx_path = filedialog.askopenfilename(title="Select DOCX file", filetypes=[("Word Documents", "*.docx")])

    messagebox.showinfo("Select Output Folder", "Select the output folder for extracted images")
    output_folder = filedialog.askdirectory(title="Select Output Folder")

    names_input = simpledialog.askstring("Image Names", "Paste the list of image names (copied from Excel, in correct serial order):")
    if not names_input:
        messagebox.showerror("Error", "No names provided.")
        exit()

    # Detect if numbers are merged and split them correctly
    raw_names = names_input.strip().replace('\t', '\n').split('\n')
    corrected_names = []
    for name in raw_names:
        corrected_names.extend(split_merged_numbers(name))

    # Preview extracted names before proceeding
    preview_text = "\n".join(corrected_names[:20])  # Show first 20 names for confirmation
    confirm = messagebox.askyesno("Confirm Names", f"Detected the following names:\n{preview_text}\n...\nProceed with these names?")
    if not confirm:
        exit()

    cleanup_confirm = messagebox.askyesno("Cleanup", "Delete temporary files after extraction?")

    extract_images_from_docx_with_names(docx_path, output_folder, corrected_names, cleanup_confirm)

    messagebox.showinfo("Done", f"All available images extracted and renamed to: {output_folder}")
