import glob
import re

files = glob.glob('*.doc')
filename = files[0]
with open(filename, 'rb') as f:
    raw = f.read()

print("File:", filename, "Size:", len(raw))

# Parse RTF text and tables
# In RTF, SPSS outputs tables with \trowd, \cell, \row
# Let's see all text by properly handling \'xx or \uXXXX
def parse_rtf(data):
    # Match tokens
    tokens = re.findall(rb"\\([a-zA-Z]+)(-?\d+)? ?|\\\'([0-9a-fA-F]{2})|\\([{}\\])|([^{}\\]+)", data)
    out = []
    for tag, num, hex_char, esc, text in tokens:
        if hex_char:
            out.append(bytes.fromhex(hex_char.decode('ascii')))
        elif esc:
            out.append(esc)
        elif text:
            out.append(text)
        elif tag in [b'par', b'row']:
            out.append(b'\n')
        elif tag == b'cell' or tag == b'tab':
            out.append(b'\t')
    
    full_bytes = b''.join(out)
    # try decoding with big5/cp950 or utf-8 or cp1252
    for enc in ['cp950', 'big5', 'utf-8', 'latin1']:
        try:
            decoded = full_bytes.decode(enc)
            print(f"Decoded successfully with {enc}!")
            return decoded
        except Exception as e:
            pass
    return full_bytes.decode('latin1', errors='ignore')

txt = parse_rtf(raw)
with open('parsed_output.txt', 'w', encoding='utf-8') as f:
    f.write(txt)

print("Parsed output saved to parsed_output.txt")
