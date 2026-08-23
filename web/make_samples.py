#!/usr/bin/env python3
"""Generate tiny sample documents for the wasm demo.

Outputs: web/samples/sample.docx and web/samples/sample.pdf.
No third-party libraries — only the stdlib (zip + manual PDF xref).
"""
import zipfile
import os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples")
os.makedirs(OUT, exist_ok=True)

# ---------------------------------------------------------------------------
# DOCX (minimal OOXML package)
# ---------------------------------------------------------------------------
content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Sample Report</w:t></w:r></w:p>
    <w:p><w:r><w:t>Hello from batdoc compiled to WebAssembly.</w:t></w:r></w:p>
    <w:p><w:r><w:t>This text was extracted in the browser, no server involved.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""

with zipfile.ZipFile(os.path.join(OUT, "sample.docx"), "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document)

# ---------------------------------------------------------------------------
# PDF (one page, one line of text, hand-built xref)
# ---------------------------------------------------------------------------
def pdf_object(n, body):
    return f"{n} 0 obj\n{body}\nendobj\n".encode()

stream = b"BT /F1 24 Tf 72 720 Td (Hello from batdoc wasm!) Tj ET"

objs = []
objs.append(pdf_object(1, "<< /Type /Catalog /Pages 2 0 R >>"))
objs.append(pdf_object(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>"))
objs.append(pdf_object(
    3,
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
    "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
))
objs.append(pdf_object(4, f"<< /Length {len(stream)} >>\nstream\n".encode() + stream + b"\nendstream"))
objs.append(pdf_object(5, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"))

pdf = bytearray(b"%PDF-1.4\n")
offsets = []
for o in objs:
    offsets.append(len(pdf))
    pdf += o

xref_pos = len(pdf)
pdf += f"xref\n0 {len(objs)+1}\n".encode()
pdf += b"0000000000 65535 f \n"
for off in offsets:
    pdf += f"{off:010d} 00000 n \n".encode()
pdf += f"trailer\n<< /Size {len(objs)+1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode()

with open(os.path.join(OUT, "sample.pdf"), "wb") as f:
    f.write(pdf)

print(f"wrote {os.path.join(OUT, 'sample.docx')}")
print(f"wrote {os.path.join(OUT, 'sample.pdf')}")
