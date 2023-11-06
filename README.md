# PDFDoc
"PDFDoc" PDF class (Classic ASP)

Include the class at the beginning of your code:

```
<!-- #include virtual="/_inc/_class/_newPDF.asp" -->
```

Initialise PDF document class: 

```
Dim PDF
Set PDF = New PDFDoc
```

Set page size: 

```
PDF.Format = "A4"
```
(Available sizes: A5, A4, A3, A2, LETTER, LEGAL.)

Set page orientation: 

```
PDF.Orientation = "PORTRAIT"
```
(Available orientations: PORTRAIT, LANDSCAPE.)

Create document:

```
PDF.Open
PDF.Title = "Page title"
PDF.Creator = "Website or company name"
PDF.Author = "Name"
```

Set font, font size and colour: 

```
PDF.SetFont "Arial","",8
```

Set draw colour (RGB): 

```
PDF.SetDrawColour 150,150,150
```

Get page dimensions: 

```
Dim x, y
x = PDF.Page.Width
y = PDF.Page.Height
```

Get width of a string of text: 

```
Dim text, width
text = "Hello, world!"
width = PDF.Font.GetWidth(text)
```

Write text centered on page: 

```
PDF.Text (PDF.Page.Width / 2) – (width / 2), PDF.Page.Height / 2, text
```

Example of drawing graphics: 

```
Dim x1,y1,x2,y2
x1 = 10
y1 = 10
x2 = 50
y2 = 50
PDF.SetLineWidth(0.1)
PDF.Line x1, y1, x2, y2
```

Add a new page to document: 

```
PDF.AddPage
```

Draw a barcode: 

```
x = 50
y = 100
width = 50
text = "ABC12345"
Page.Code39 x, y, width, text
```

Close and publish the PDF document: 

```
PDF.Close
PDF.Publish
```

Other available functions within this class include: 

```
PDF.Image fiile, x, y, width
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
PDF.Orientation = "PORTRAIT" '// or LANDSCAPE
PDF.Format = "A4" '// or A3/A4/A5/LEGAL/LETTER/BARCODE
PDF.AddPage
PDF.Code39 x, y, width, text
n = PDF.Font.GetWidth(text)
PDF.Paragraph(text, width, line-height)
PDF.Centered = true '// or false
```

KNOWN ISSUE: <br>
UTF8 PAGES CAUSE ISSUE WITH ISO-8897- IMAGE

