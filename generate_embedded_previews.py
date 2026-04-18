import os
import base64
from io import BytesIO
from pdf2image import convert_from_path
import comtypes.client

# Folder paths
base_folder = r'C:\Users\User\Downloads\files webpage\DJ_Talim'
html_file = r'C:\Users\User\Downloads\files webpage\index.html'
output_html = r'C:\Users\User\Downloads\files webpage\index_with_previews.html'

# Read the original HTML
with open(html_file, 'r', encoding='utf-8') as f:
    html_content = f.read()

# Function to get base64 of image
def image_to_base64(img):
    buffer = BytesIO()
    img.save(buffer, format='JPEG')
    return base64.b64encode(buffer.getvalue()).decode('utf-8')

# Function to generate first page thumb for PDF
def get_pdf_thumb_base64(file_path):
    try:
        images = convert_from_path(file_path, first_page=1, last_page=1, size=(200, 200))
        return image_to_base64(images[0])
    except:
        return None

# Function to generate first page thumb for Office
def get_office_thumb_base64(file_path, app_name):
    try:
        pdf_path = file_path + '.temp.pdf'
        app = comtypes.client.CreateObject(app_name)
        app.Visible = False
        doc = app.Documents.Open(file_path)
        doc.ExportAsFixedFormat(pdf_path, 17)
        doc.Close()
        app.Quit()
        b64 = get_pdf_thumb_base64(pdf_path)
        os.remove(pdf_path)
        return b64
    except:
        return None

# Function to get thumb for file
def get_thumb_base64(file_path, ext):
    if ext == 'pdf':
        return get_pdf_thumb_base64(file_path)
    elif ext == 'pptx':
        return get_office_thumb_base64(file_path, 'PowerPoint.Application')
    elif ext == 'docx':
        return get_office_thumb_base64(file_path, 'Word.Application')
    elif ext in ['jpg', 'jpeg', 'png']:
        from PIL import Image
        try:
            img = Image.open(file_path)
            img.thumbnail((200, 200))
            return image_to_base64(img)
        except:
            return None
    return None

# Collect all files
files_data = []
for root, dirs, files in os.walk(base_folder):
    for file in files:
        file_path = os.path.join(root, file)
        rel_path = os.path.relpath(file_path, os.path.dirname(base_folder))
        ext = file.split('.')[-1].lower()
        thumb_b64 = get_thumb_base64(file_path, ext)
        files_data.append({
            'rel_path': rel_path,
            'ext': ext,
            'thumb_b64': thumb_b64,
            'size': os.path.getsize(file_path)
        })

# Now, generate the HTML with embedded previews
# For simplicity, replace the buildFileList in the script tag

# Find the buildFileList function
start = html_content.find('function buildFileList() {')
end = html_content.find('}', start) + 1
while html_content[end] != '}':
    end += 1
end += 1

# Generate the new function
new_build = '''
function buildFileList() {
  const djTalimFiles = [
'''

for data in files_data:
    shorturl = os.path.splitext(data['rel_path'])[0].replace('/', '_').replace('\\', '_')
    longurl = data['rel_path']
    size = data['size']
    thumb = f'data:image/jpeg;base64,{data["thumb_b64"]}' if data['thumb_b64'] else None
    new_build += f'    {{ shorturl: "{shorturl}", longurl: "{longurl}", size: {size}, thumb: "{thumb}" }},\n'

new_build += '''
  ];

  return djTalimFiles;
}
'''

# Replace
html_content = html_content.replace(html_content[start:end], new_build)

# Now, update the renderTable to use the thumb
# Find the previewContent part
preview_start = html_content.find('let previewContent = \'\';')
preview_end = html_content.find('html += `</td>`;', preview_start) + len('html += `</td>`;')

preview_code = '''
        let previewContent = '';
        if (row.thumb) {
          previewContent = `<img src="${row.thumb}" alt="Preview" style="max-width: 200px; max-height: 200px; border: 1px solid #ccc;">`;
        } else {
          previewContent = `<div style="max-width: 300px;"><strong>${row.typeDisplay}</strong><br>Size: ${sizeFormatted}<br>Path: ${row.longurl}</div>`;
        }
'''

html_content = html_content.replace(html_content[preview_start:preview_end], preview_code)

# Write the new HTML
with open(output_html, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f'Generated {output_html} with embedded previews.')