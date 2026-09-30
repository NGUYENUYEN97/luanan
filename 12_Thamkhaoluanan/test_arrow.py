from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(1), Inches(1), Inches(3), Inches(3))
# Get the connection shape properties
cxnSp = connector.element
spPr = cxnSp.spPr
# spPr should have <a:ln> child
ln = spPr.ln
if ln is not None:
    tailEnd = parse_xml(r'<a:tailEnd type="triangle" w="med" len="med" %s/>' % nsdecls('a'))
    ln.append(tailEnd)
prs.save('test_arrow.pptx')
print("Success")
