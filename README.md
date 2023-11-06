# PDFDoc
"PDFDoc" PDF class (Classic ASP)

Include the class at the beginning of your code:

<code><! -- #include virtual="/_inc/_class/_newPDF.asp" -- ></code><br>
(Note the added space character between the "! –" and "-- >", these spaces will need removing in your code.)

Initialise PDF document class: 

<code>Dim PDF
Set PDF = New PDFDoc</code>

Set page size: 

<code>PDF.Format = "A4"</code><br>
(Available sizes: A5, A4, A3, A2, LETTER, LEGAL.)

Set page orientation: 

<code>PDF.Orientation = "PORTRAIT"</code><br>
(Available orientations: PORTRAIT, LANDSCAPE.)

Create document:

<code>PDF.Open
PDF.Title = "Page title"
PDF.Creator = "Website or company name"
PDF.Author = "Name"</code>

Set font, font size and colour: 

<code>PDF.SetFont "Arial","",8</code>

Set draw colour (RGB): 

<code>PDF.SetDrawColour 150,150,150</code>

Get page dimensions: 

<code>Dim x, y
x = PDF.Page.Width
y = PDF.Page.Height</code>

Get width of a string of text: 

<code>Dim text, width
text = "Hello, world!"
width = PDF.Font.GetWidth(text)</code>

Write text centered on page: 

<code>PDF.Text (PDF.Page.Width / 2) – (width / 2), PDF.Page.Height / 2, text</code>

Example of drawing graphics: 

<code>Dim x1,y1,x2,y2
x1 = 10
y1 = 10
x2 = 50
y2 = 50
PDF.SetLineWidth(0.1)
PDF.Line x1, y1, x2, y2</code>

Add a new page to document: 

<code>PDF.AddPage</code>

Draw a barcode: 

<code>x = 50
y = 100
width = 50
text = "ABC12345"
Page.Code39 x, y, width, text</code>

Close and publish the PDF document: 

<code>PDF.Close
PDF.Publish</code>

Other available functions within this class include: 

<code>PDF.Image fiile, x, y, width
PDF.StartTransform
PDF.Rotate angle, x, y
PDF.Skew angle, x, y
PDF.EndTransform
PDF.Box x, y, width, style
PDF.Line x, y, x2, y2
PDF.SetFont font, style, size
PDF.Bold true/false
PDF.SetFontSize n
PDF.SetTextColour r,g,b
PDF.SetDrawColour r,g,b
PDF.Orientation = "PORTRAIT/LANDSCAPE"
PDF.Format = "A3/A4/A5/LEGAL/LETTER/BARCODE"
PDF.AddPage
PDF.Code39 x, y, width, text
n = PDF.Font.GetWidth(text)
PDF.Paragraph(text, width, line-height)
PDF.Centered = true/false</code>

KNOWN ISSUE: <br>
UTF8 PAGES CAUSE ISSUE WITH ISO-8897- IMAGE

