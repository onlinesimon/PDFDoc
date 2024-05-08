<!-- #include virtual="/PDFDoc.class.asp" -->
<%

dim PDF
set PDF = new PDFDoc

PDF.Format = "A4"
PDF.Orientation = "LANDSCAPE"

PDF.Open

PDF.Title = "Example paragraph"
PDF.Creator = "PDFDoc"
PDF.Author = "Simon Beal"

PDF.SetFont "Arial","",10
PDF.SetDrawColour 150,150,150

'// SET TOP-LEFT POSITION

PDF.Page.X = PDF.Page.Width / 4
PDF.Page.Y = PDF.Page.Height / 4

'// EXAMPLE TEXT

dim lineHeight
lineHeight = 5.5

PDF.Bold true
PDF.Text PDF.Page.X, PDF.Page.Y, "EXAMPLE PARAGRAPHS (UNJUSTIFIED AND JUSTIFIED)"
PDF.Page.Y = PDF.Page.Y + (lineHeight * 2)
PDF.Bold false

'// FIRST PARAGRAPH

dim t
t = "Ah, those beloved ill-begotten websites of the wild frontier days - franken-pages stitched from Geocities templates and bagpipe-Smash Mouth midis. We'll never forget thy visionary markup, thy boundary-pushing abuse of the blink tag and animated construction gifs. What cutting-edge webmastery they represented! Sure, by today's standards those dopey fan pages look like they were designed by particularly un-web-savvy mattress-mites from Hampstedance. But who among us doesn't feel a warm nostalgic flatulence revisiting the erratic marquees, garish colours, and made-with-an-etch-a-sketch navigation schemes of the web's toddlerhood?"

n = PDF.Paragraph(t, PDF.Page.Width / 2, lineHeight)
PDF.Page.Y = PDF.Page.Y + lineHeight

'// 2ND (JUSTIFIED) PARAGRAPH

t = "Those pioneering pages now exist as lifeless relics in the ammonic cloud of web archeology - petrified forests of broken links and squatted domains overgrown with popup-weeds. Yet we salute their eternal buffer, for allowing us to appreciate just how unpleasantly surprised our current selves would be at getting a faceful of juddery java over dial-up. So raise a self-hotlinking cup to these fallen disaster areas! Their improper-nested spirit guides us still, like indecipherable browser errors seeking the Ultimate Answer."

PDF.Justified = true
n = PDF.Paragraph(t, PDF.Page.Width / 2, lineHeight)

PDF.Close
PDF.Publish

set PDF = nothing

%>
