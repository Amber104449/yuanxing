import sys
sys.stdout.reconfigure(encoding='utf-8')
from xml.etree import ElementTree as ET

# Parse comments.xml
tree = ET.parse(r'c:\Users\adone\OneDrive\桌面\AI练习\wendang\fangan\unpacked\word\comments.xml')
root = tree.getroot()

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

with open(r'c:\Users\adone\OneDrive\桌面\AI练习\wendang\fangan\comments.txt', 'w', encoding='utf-8') as f:
    for comment in root.findall('w:comment', ns):
        cid = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
        author = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author')
        date = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date')
        
        texts = []
        for t in comment.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            if t.text:
                texts.append(t.text)
        
        f.write(f'Comment {cid} by {author} ({date}):\n')
        f.write('  ' + '\n  '.join(texts) + '\n\n')

print("Done")
