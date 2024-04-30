<!-- #include virtual="/PDFDoc.class.asp" -->
<%

dim PDF
set PDF = new PDFDoc

PDF.Format = "A4"
PDF.Orientation = "PORTRAIT"

PDF.Open

PDF.Title = "Example polar graph page"
PDF.Creator = "PDFDoc"
PDF.Author = "Simon Beal"

PDF.SetFont "Arial","",8
PDF.SetDrawColour 150,150,150

dim x,y

x = PDF.Page.Width / 2
y = PDF.Page.Height / 2

const GapR = 15 '// DEFAULT

dim sizeR
sizeR = 10 '// DEFAULT

dim maxRadius
select case PDF.Orientation
	case "PORTRAIT"
		maxRadius = PDF.Page.Width
	case else
		maxRadius = PDF.Page.Height
end select

dim r
for r = sizeR to x - GapR step sizeR
	PDF.Ellipse x,y,r,r,"D"
next

if r > (maxRadius / 2) - GapR then
	'// CORRECT RADIUS (TOO BIG)'
	r = r - sizeR
else
	'// OUTSIDE RADIAL'
	PDF.Ellipse x,y,r,r,"D"
end if

Const Rad = 0.01745329 '// EQUALS 2 * Pi / 360

dim x1,y1,x2,y2

dim angle
for angle = 0 to 357.5 step 2.5

	if angle mod 10 = 0 then
		'// 10-DEG INTERVAL

		x1 = x + r * cos(angle * Rad)
		y1 = y + r * sin(angle * Rad)

		x2 = x + sizeR * cos(angle * Rad)
		y2 = y + sizeR * sin(angle * Rad)

		if angle mod 30 then
			PDF.SetLineWidth(0.1)
		else
			'// 30-DEG INTERVAL

			call PrintAngle(angle)
			PDF.SetLineWidth(0.4)
			x2 = x
			y2 = y

		end if

		PDF.Line x1,y1,x2,y2

	else

		if angle mod 5 = 0 then
			'// IN-BETWEEN 5-DEG INTERVAL

			x1 = x + r * cos(angle * Rad)
			y1 = y + r * sin(angle * Rad)

			x2 = x + (r - (sizeR*4)) * cos(angle * Rad)
			y2 = y + (r - (sizeR*4)) * sin(angle * Rad)

		else
			'// IN-BETWEEN 2.5-DEG INTERVAL

			x1 = x + r * cos(angle * Rad)
			y1 = y + r * sin(angle * Rad)

			x2 = x + (r - sizeR) * cos(angle * Rad)
			y2 = y + (r - sizeR) * sin(angle * Rad)

		end if

		PDF.SetLineWidth(0.1)
		PDF.Line x1,y1,x2,y2

		if angle mod 15 = 0 then
			call PrintAngle(angle)
		end if

	end if

next

PDF.Close
PDF.Publish

set PDF = nothing

sub PrintAngle(byref angle)

	dim x3, y3, theta, w

	x3 = x - (r + 3) * cos(angle * Rad)
	y3 = y - (r + 3) * sin(angle * Rad)

	theta = angle + 270

	if theta >= 360 then
		theta = theta - 360
	end if

	w = PDF.Font.GetWidth(theta)
	PDF.Text x3 - (w / 2), y3 + 0.6, theta

end sub

%>
