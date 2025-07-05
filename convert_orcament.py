import re

input_path = 'IRRIG/ORCAMENT.FRM'
output_path = 'IRRIG/ORCAMENT_NOVO.FRM'

with open(input_path, 'r', encoding='cp1252', errors='ignore') as f:
    text = f.read()

# Replace custom controls with VB6 equivalents
replacements = {
    'IRRIG.GPainel': 'Frame',
    'IRRIG.GBotao': 'CommandButton',
    'IRRIG.GTexto': 'TextBox',
    'IRRIG.GLabel': 'Label',
    'IRRIG.GListV': 'ListView',
    'IRRIG.GCpMM': 'TextBox',
}

for old, new in replacements.items():
    text = text.replace(old, new)

# Replace LoadGasString calls with placeholder strings (ASCII)
text = re.sub(r'LoadGasString\([^\)]+\)', '"TODO"', text)

# Comment out any remaining IRRIG.* function calls
text = re.sub(r'IRRIG\.\w+', "'TODO", text)

with open(output_path, 'w', encoding='cp1252', errors='ignore') as f:
    f.write(text)
