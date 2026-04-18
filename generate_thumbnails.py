import os
import comtypes.client
from pdf2image import convert_from_path

# Folder paths
base_folder = r'C:\Users\User\Downloads\files webpage\DJ_Talim'
thumbnails_folder = r'C:\Users\User\Downloads\files webpage\thumbnails'

# Create thumbnails folder if not exists
os.makedirs(thumbnails_folder, exist_ok=True)

def get_thumbnail_path(file_path):
    rel_path = os.path.relpath(file_path, base_folder)
    thumb_path = os.path.join(thumbnails_folder, rel_path + '.jpg')
    os.makedirs(os.path.dirname(thumb_path), exist_ok=True)
    return thumb_path

def generate_pdf_thumbnail(file_path, thumb_path):
    try:
        images = convert_from_path(file_path, size=(200, 200))
        for i, img in enumerate(images):
            page_thumb = thumb_path.replace('.jpg', f'_{i+1}.jpg')
            img.save(page_thumb, 'JPEG')
        print(f'Generated {len(images)} thumbnails for {file_path}')
    except Exception as e:
        print(f'Error generating PDF thumbnail for {file_path}: {e}')

def generate_office_to_pdf(file_path, pdf_path, app_name):
    try:
        app = comtypes.client.CreateObject(app_name)
        app.Visible = False
        doc = app.Documents.Open(file_path)
        doc.ExportAsFixedFormat(pdf_path, 17)
        doc.Close()
        app.Quit()
        return True
    except Exception as e:
        print(f'Error exporting {app_name} to PDF for {file_path}: {e}')
        return False

# Walk through all files
for root, dirs, files in os.walk(base_folder):
    for file in files:
        file_path = os.path.join(root, file)
        ext = file.split('.')[-1].lower()
        thumb_base = get_thumbnail_path(file_path).replace('.jpg', '')
        if any(os.path.exists(f'{thumb_base}_{i+1}.jpg') for i in range(10)):  # Check if any exist
            continue
        if ext == 'pdf':
            generate_pdf_thumbnail(file_path, thumb_base + '.jpg')
        elif ext in ['pptx', 'docx']:
            pdf_path = thumb_base + '.pdf'
            app_name = 'PowerPoint.Application' if ext == 'pptx' else 'Word.Application'
            if generate_office_to_pdf(file_path, pdf_path, app_name):
                generate_pdf_thumbnail(pdf_path, thumb_base + '.jpg')
                os.remove(pdf_path)  # Clean up PDF
        # For mp4, skip

print('Thumbnail generation complete.')