import xml.etree.ElementTree as ET

ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
tree = ET.parse(r'C:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\Color\LogicDocs\xlsx_extracted\xl\worksheets\sheet1.xml')
root = tree.getroot()

with open(r'c:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\parsed_formulas.txt', 'w') as f:
    for c in root.findall('.//x:c', ns):
        f_el = c.find('x:f', ns)
        if f_el is not None:
            v_el = c.find('x:v', ns)
            v_text = v_el.text if v_el is not None else ''
            f.write(f'{c.get("r")}: {f_el.text} (val: {v_text})\n')
