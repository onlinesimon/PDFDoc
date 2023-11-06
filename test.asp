<!-- #include virtual="/PDFDoc.class.asp" -->
<%

dim PDF
set PDF = new PDFDoc

PDF.Format = "A4"
PDF.Orientation = "PORTRAIT"

PDF.Open

PDF.Title = "Test page"
PDF.Creator = "PDFDoc"
PDF.Author = "Simon Beal"

PDF.SetFont = "Arial","",8
PDF.SetDrawColour 150,150,150

dim x,y

x = PDF.Page.Width / 2
y = PDF.Page.Height / 2

dim t
t = "Hello, world!"

dim w
w = PDF.Font.GetWidth(t)

'// CENTER TEXT ON PAGE

PDF.Text x - (w / 2), y, t

PDF.Close
PDF.Publish

set PDF = nothing

%>
