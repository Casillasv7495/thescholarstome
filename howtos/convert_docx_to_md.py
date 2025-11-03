import docx2md
import os
import shutil
import re

# Folders
docx_folder = "docx_files"
processed_folder = "docx_processed"
media_folder = "media"
folder_prefix = "howtos/"  # folder prefix for Markdown links

# Ensure folders exist
os.makedirs(docx_folder, exist_ok=True)
os.makedirs(processed_folder, exist_ok=True)
os.makedirs(media_folder, exist_ok=True)

current_folder = os.getcwd()

for filename in os.listdir(docx_folder):
    if filename.lower().endswith(".docx"):
        docx_path = os.path.join(docx_folder, filename)
        base_name = os.path.splitext(filename)[0]

        md_file = os.path.join(current_folder, f"{base_name}.md")

        # Convert DOCX → Markdown
        markdown_text = docx2md.do_convert(
            docx_path,
            target_dir=media_folder,
            use_md_table=False
        )

        # Flatten any subfolders docx2md created
        for root, dirs, files in os.walk(media_folder):
            for dir_name in dirs:
                subfolder = os.path.join(root, dir_name)
                for i, img in enumerate(sorted(os.listdir(subfolder)), start=1):
                    ext = os.path.splitext(img)[1]
                    new_name = f"{base_name}_{i}{ext}"  # rename with DOCX prefix
                    new_path = os.path.join(media_folder, new_name)
                    shutil.move(os.path.join(subfolder, img), new_path)
                    # Update Markdown links with folder_prefix
                    markdown_text = markdown_text.replace(f"{img}", f"{folder_prefix}media/{new_name}")
                    markdown_text = markdown_text.replace(f"{dir_name}/{img}", f"{folder_prefix}media/{new_name}")
                os.rmdir(subfolder)

        # Rename remaining files directly in media_folder that weren’t prefixed
        for i, img in enumerate(sorted(os.listdir(media_folder)), start=1):
            if img.startswith(f"{base_name}_"):
                continue
            ext = os.path.splitext(img)[1]
            new_name = f"{base_name}_{i}{ext}"
            old_path = os.path.join(media_folder, img)
            new_path = os.path.join(media_folder, new_name)
            os.rename(old_path, new_path)
            # Update Markdown links with folder_prefix
            markdown_text = markdown_text.replace(img, f"{folder_prefix}media/{new_name}")

        # Replace <img src="..."> HTML tags with Markdown syntax and folder_prefix
        def replace_img_tag(match):
            src = match.group(1)
            src_name = os.path.basename(src)
            # Set only width to keep aspect ratio
            width = 400
            return f'<img src="{folder_prefix}media/{src_name}" alt="{base_name}" width="{width}">'

        markdown_text = re.sub(r'<img src="([^"]+)"[^>]*>', replace_img_tag, markdown_text)

        # Save Markdown
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(markdown_text)

        # Move processed DOCX
        shutil.move(docx_path, os.path.join(processed_folder, filename))

        print(f"Processed '{filename}' → '{md_file}' with images in '{media_folder}'")

print("All DOCX files converted with images prefixed by DOCX name and Markdown links updated with folder prefix.")
