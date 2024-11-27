from spire.doc import *
from spire.doc.common import *

doc=Document()

section = doc.AddSection()

section.PageSetup.PageSize = PageSize.A4()
section.PageSetup.Margins.All = 36.1

titlepara=section.AddParagraph()
titlepara.AppendText("Title")

hstyle=ParagraphStyle(doc)
hstyle.Name = "headingStyle"
hstyle.CharacterFormat.Bold=True
hstyle.CharacterFormat.FontName = "Times New Roman"
hstyle.CharacterFormat.FontSize = 18
doc.Styles.Add(hstyle)

titlepara.ApplyStyle("headingStyle")
titlepara.Format.HorizontalAlignment = HorizontalAlignment.Center
titlepara.Format.AfterSpacing = 10


namepara=section.AddParagraph()
namepara.AppendText("\nAtharva Saraf (0009-0006-3842-5508)\n")

addrpara=section.AddParagraph()
addrpara.AppendText("School of Engineering\nAjeenkya D Y Patil University\nCharoli BK. Via Lohegaon, Pune - 412105, Maharashtra, India\n")

mailpara=section.AddParagraph()
mailpara.AppendText("atharvasaraf2004@gmail.com\n")

pstyle=ParagraphStyle(doc)
pstyle.Name = "paraStyle"
pstyle.CharacterFormat.FontName = "Times New Roman"
pstyle.CharacterFormat.FontSize = 12
doc.Styles.Add(pstyle)

namepara.ApplyStyle("paraStyle")
mailpara.ApplyStyle("paraStyle")
addrpara.ApplyStyle("paraStyle")

namepara.Format.HorizontalAlignment = HorizontalAlignment.Center
namepara.Format.AfterSpacing = 10
mailpara.Format.HorizontalAlignment = HorizontalAlignment.Center
mailpara.Format.AfterSpacing = 10
addrpara.Format.HorizontalAlignment = HorizontalAlignment.Center
addrpara.Format.AfterSpacing = 10


footer = section.HeadersFooters.Footer

footerParagraph = footer.AddParagraph()
footerParagraph.AppendField("page number",FieldType.FieldPage)
footerParagraph.AppendText('/')
footerParagraph.AppendField("page count",FieldType.FieldNumPages)
footerParagraph.Format.HorizontalAlignment=HorizontalAlignment.Center
                               
print('done')
doc.SaveToFile(r"C:\Users\athar\Desktop\Projects\Python Projects\Word formater\test.docx",FileFormat.Docx2019)
doc.Close()
