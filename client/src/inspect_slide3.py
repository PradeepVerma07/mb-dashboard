import xml.etree.ElementTree as ET

tree = ET.parse('client/public/assets/ppt/slides/slide3.xml')
root = tree.getroot()

for pic in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}pic'):
    name = pic.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr').get('name')
    xfrm = pic.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
    if xfrm is not None:
        off = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off')
        ext = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
        x = int(off.get('x')) / 914400
        y = int(off.get('y')) / 914400
        w = int(ext.get('cx')) / 914400
        h = int(ext.get('cy')) / 914400
        print(f'{name}: x={x:.2f}", y={y:.2f}", w={w:.2f}", h={h:.2f}"')
3