import zipfile
import xml.etree.ElementTree as ET
import re
import os

def extract_modern_comments(pptx_path):
    comments = []
    
    with zipfile.ZipFile(pptx_path, 'r') as pptx:
        comment_files = [f for f in pptx.namelist() if f.startswith('ppt/comments/modernComment') and f.endswith('.xml')]
        slide_number_pattern = re.compile(r'modernComment_(\w+)_')
        
        for comment_file in comment_files:
            xml_content = pptx.read(comment_file)
            root = ET.fromstring(xml_content)
            
            slide_match = slide_number_pattern.search(comment_file)
            slide_number = int(slide_match.group(1), 16) if slide_match else float('inf')
            
            namespaces = {
                'p188': "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
                'p': "http://schemas.openxmlformats.org/presentationml/2006/main",
                'a': "http://schemas.openxmlformats.org/drawingml/2006/main"
            }
            
            for cm in root.findall('.//p188:cm', namespaces):
                text_elem = cm.find('.//a:t', namespaces)
                if text_elem is None:
                    text_elem = cm.find('.//p:text', namespaces)
                
                comment_text = text_elem.text if text_elem is not None else "No text found"
                
                comments.append({
                    'slide': slide_number,
                    'text': comment_text
                })
    
    return sorted(comments, key=lambda x: x['slide'])

# Get all .pptx files in the current directory
pptx_files = [f for f in os.listdir('.') if f.endswith('.pptx')]

for pptx_file in pptx_files:
    print(f"\nProcessing file: {pptx_file}")
    print("=" * 50)
    
    try:
        extracted_comments = extract_modern_comments(pptx_file)

        if not extracted_comments:
            print("No comments were found in this presentation.")
        else:
            print(f"Total comments extracted: {len(extracted_comments)}")
            for comment in extracted_comments:
                print(f"{comment['text']}")
                print("---")
    except Exception as e:
        print(f"An error occurred while processing {pptx_file}: {e}")

print("\nAll PowerPoint files processed.")
