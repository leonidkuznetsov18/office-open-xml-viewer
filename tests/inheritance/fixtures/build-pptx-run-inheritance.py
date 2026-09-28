"""Build a minimal OPC deck that exercises the production PPTX parser."""
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'


def shape(i, y, text, field=False, no_fill=False, highlight=False):
    local = '<a:noFill/>' if no_fill else ''
    if highlight:
        local += '<a:highlight><a:srgbClr val="F5DD22"/></a:highlight>'
    run = ('<a:fld id="{0FBB9679-570D-4509-9600-E07789CF24BB}" type="datetime">'
           if field else '<a:r>')
    run += f'<a:rPr>{local}</a:rPr><a:t>{text}</a:t>'
    run += '</a:fld>' if field else '</a:r>'
    style = ('<a:lstStyle><a:lvl1pPr><a:defRPr sz="4800">'
             '<a:solidFill><a:srgbClr val="D21D54"/></a:solidFill>'
             '<a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></a:lstStyle>')
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="Inheritance {i}"/>'
            '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="600000" y="{y}"/>'
            '<a:ext cx="9000000" cy="1300000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/>'
            f'{style}<a:p><a:pPr><a:defRPr b="1"/></a:pPr>{run}</a:p></p:txBody></p:sp>')

shapes = (shape(2, 500000, 'RUN INHERITS RED')
          + shape(3, 2300000, 'FIELD INHERITS RED', field=True)
          + shape(4, 4100000, 'INVISIBLE GLYPHS', no_fill=True, highlight=True))
slide = (f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>'
         '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
         f'<p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>')
parts = {
    '[Content_Types].xml': f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
    'ppt/presentation.xml': f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{R}/slide" Target="slides/slide1.xml"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
}
out = Path('tests/inheritance/fixtures/pptx-run-inheritance.pptx')
with ZipFile(out, 'w', ZIP_DEFLATED) as package:
    for path, content in sorted(parts.items()):
        item = ZipInfo(path, (2020, 1, 1, 0, 0, 0))
        item.compress_type = ZIP_DEFLATED
        package.writestr(item, content)
print(out, out.stat().st_size)
