<%

'// AUTHOR:	Simon Beal (onlinesimon@outlook.com)'
'//	LAST UPDATE: 2021-06-24 / CREATED: 2018-09-17'

'// VERSION: 1.0.2

'// REQUIRES: 	_inc\adovbs.inc'

'// USAGE: 							FUNCTIONS: 
'// 	dim PDF 						.Image file, x, y, width
'// 	set PDF = new PDFDoc 			.StartTransform
'// 	PDF.Open 						->	.Rotate angle, x, y
'// 	PDF.Text x, y, text 			->	.Skew angle, x, y
'// 	PDF.Close 						.EndTransform
'// 	PDF.Publish 					.Box x, y, width, style
'// 									.Line x,y, x2, y2
'// 									.SetFont font, style, size'
'// 									.Bold true/false'
'// 									.SetFontSize n
'// 									.SetTextColour r,g,b
'// 									.SetDrawColour r,g,b
'// 									.Orientation = "PORTRAIT/LANDSCAPE"'
'// 									.Format = "A3/A4/A5/LEGAL/LETTER/BARCODE"'
'// 									.AddPage
'// 									.Code39 x, y, width, text'
'// 									.Font.GetWidth(text)'
'// 									.Paragraph(text, width, line-height)'
'// 									.Centered = True/False'

'// KNOWN ISSUE: 	UTF8 PAGES CAUSE ISSUE WITH ISO-8897- IMAGE'

class PDFDoc

	public CurrentPage
	public Pages(99)
	public State
	public Title
	public Author
	public Keywords
	public Creator
	public LastH
	public LineWidth
	public LayoutMode
	public Orientation
	public Unit
	public Format
	public Offsets(99)
	public Filters
	public Compress
	public Colour
	public Doc
	public Font
	public Margin
	public Page
	public ZoomMode
	public Barcode

	private DefOrientation
	private OrientationChanges(99)
	private CurOrientation
	private k
	private n
	private FormatArray
	private Buffer
	private InFooter
	private ThisImage

	public LastImageHeightMM
	public Centered
	public Justified
	public ToSave

	public sub SetXY(byval x, byval y)

		Page.X = x
		Page.Y = y

	end sub

	public function GetY()

		GetY = Page.Y

	end function

	public function Justify(byval Phrase, byval intWidth, byval intLineHeight)

		if len(Phrase) > 0 then

			dim w
			w = Font.GetWidth(Phrase)

			if w >= intWidth then

				'// PRINT PHRASE AS-IS'
				Text Page.X, Page.Y, Phrase

			else

				'// JUSTIFY WORDS'
				dim Wording
				Wording = split(Phrase)

				dim Gap
				Gap = (intWidth - w) / ubound(Wording)
				Gap = Gap + Font.GetWidth(" ")

				dim x
				dim y

				x = Page.X

				for each Word in Wording

					y = Page.Y
					Text x, y, Word
					x = x + Font.GetWidth(Word) + Gap

				next
			end if
		end if

	end function

	public function Paragraph(byval strPara, intWidth, intLineHeight)

		'// OUTPUT NEXT LINE OF TEXT'

		dim strOut
		dim intTotalWidth
		dim intWordWidth
		dim intLines
		dim arrPara

		arrPara = split(strPara,chr(32))

		for each item in arrPara

			intWordWidth = Font.GetWidth(item & chr(32))

			if intTotalWidth + intWordWidth >= intWidth then

				'// EXCEEDS MAXIMUM WIDTH, SO OUTPUT'
				if Centered = false and Justified = false then

					Text Page.X, Page.Y, strOut

				else

					if Justified = false then
						if len(strOut) > 0 then

							w = Font.GetWidth(strOut)
							Text (Page.Width / 2) - (w / 2), Page.Y, strOut

						end if
					else

						if len(strOut) > 0 then

							Justify strOut, intWidth, intLineHeight

						end if
					end if
				end if

				Page.Y = Page.Y + intLineHeight

				strOut = ""
				intTotalWidth = 0
				intLines = intLines + 1

			end if

			strOut = strOut & item & chr(32)
			intTotalWidth = intTotalWidth + intWordWidth

		next

		if len(strOut) > 0 then

			'// OUTPUT LAST LINE'

			strOut = replace(strOut,vbcrlf,"")
			strOut = trim(strOut)

			if Centered = false then

				Text Page.X, Page.Y, strOut

			else

				if len(strOut) > 0 then

					w = Font.GetWidth(strOut)
					Text (Page.Width / 2) - (w / 2), Page.Y, strOut

				end if

			end if

			Page.Y = Page.Y + intLineHeight
			intLines = intLines + 1

		end if

		Paragraph = intLines

	end function

	public sub Publish()

		dim i
		for i = 1 to len(PDF.GetBuffer)
			response.binarywrite chrB(asc(mid(PDF.GetBuffer,i,1)))
		next

	end sub

	public sub PublishAsFile(byval Filename)

		'// OUTPUTS TEXT NOT BINARY / WILL NOT WORK'

		dim TempFile
		TempFile = replace(Filename, ".pdf", ".tmp")

		Const adTypeBinary = 1
		Const adTypeText = 2
		Const adSaveCreateOverWrite = 2

		Dim TextStream
		Set TextStream = CreateObject("ADODB.Stream")

		'// SAVE AS TEXT (TEMP)'

		TextStream.Type = adTypeText
		TextStream.Open

		TextStream.WriteText PDF.GetBuffer
		TextStream.SaveToFile TempFile, adSaveCreateOverWrite

		TextStream.Close

		'// READ TEXT / WRITE BINARY'

		TextStream.Type = adTypeBinary
		TextStream.Open

		TextStream.LoadFromFile TempFile

		dim BinaryStream
		Set BinaryStream = CreateObject("ADODB.Stream")

		BinaryStream.Type = adTypeBinary
		BinaryStream.Open

		n = TextStream.Read

		BinaryStream.Close
		TextStream.Close

	end sub

	public sub Bold(byval IsBold)

		select case IsBold
			case true
				SetFont Font.Name, "B", Font.SizePt
			case else
				SetFont Font.Name, "", Font.SizePt
		end select

	end sub

	private sub Output()

		if State < 3 then call Close

	end sub

	private sub Header()
	end sub

	private sub Footer()
	end sub

	public function CreatePDF()

		CreatePDF = true

		CurrentPage = 0
		State = 0
		n = 2
		Buffer = ""
		InFooter = false

		Font.Underline = false
		Font.Style = ""
		Font.SizePt = 12

		if len(Orientation) = 0 then Orientation = "P"
		if len(Unit) = 0 then Unit = "mm"
		if len(Format) = 0 then Format = "A4"

		ws = 0

		select case lcase(Unit)
			case "pt"
				k = 1
			case "mm"
				k = 72/25.4
			case "cm"
				k = 72/2.54
			case "in"
				k = 72
			case else
'				'// "Incorrect Unit"
				CreatePDF = false
		end select

		select case ucase(Format)
			case "A2"
				FormatArray = array(1190.55,1683.78)
			case "A3"
				FormatArray = array(841.89,1190.55)
			case "A4"
				FormatArray = array(595.28,841.89)
			case "A5"
				FormatArray = array(420.94,595.28)
			case "LETTER"
				FormatArray = array(612,792)
			case "LEGAL"
				FormatArray = array(612,1008)
			case "BARCODE"
				FormatArray = array(70.86,155.9)
			case "B3"
				FormatArray = array(1000.63,1417.32)
			case "B4"
				FormatArray = array(708.66,1000.63)
			case "B5"
				FormatArray = array(498.90,708.66)
			case "C3"
				FormatArray = array(918.42,1298.27)
			case "C4"
				FormatArray = array(639.13,918.42)
			case "C5"
				FormatArray = array(459.21,649.13)
			case "ANSI A"
				FormatArray = array(612.28,790.87)
			case "ANSI B"
				FormatArray = array(790.87,1223.57)
			case "ANSI C"
				FormatArray = array(1223.57,1584.57)
			case "ANSI D"
				FormatArray = array(1584.57,2449.13)
			case else
				'// "Unknown page format"
				CreatePDF = false
		end select

		Doc.Pt.Width = FormatArray(0)
		Doc.Pt.Height = FormatArray(1)

		Doc.Width = Doc.Pt.Width / k
		Doc.Height = Doc.Pt.Height / k

		select case ucase(Orientation)
			case "P","PORTRAIT"
				DefOrientation = "P"
				Page.Pt.Width = Doc.Pt.Width
				Page.Pt.Height = Doc.Pt.Height
			case "L","LANDSCAPE"
				DefOrientation = "L"
				Page.Pt.Width = Doc.Pt.Height
				Page.Pt.Height = Doc.Pt.Width
			case else
				'// "Incorrect orientation"
				CreatePDF = false
		end select

		Margin.Top = 28.35 / k
		Margin.Bottom = 28.35 / k
		Margin.LeftX = 28.35 / k
		Margin.RightX = 28.35 / k

		Margin.C = Margin.Top / 10
		LineWidth = .567 / k

	end function

	public sub Open()

		call CreatePDF
		call BeginDoc
		call AddPage

	end sub

	public sub Close()

		if not ToSave = true then
			Response.ContentType = "application/pdf"
		end if

		if CurrentPage = 0 then AddPage

		InFooter = true
		call Footer
		InFooter = false

		call EndPage()
		call EndDoc()

	end sub

	public sub SetFont(byval tmpFamily, byval tmpStyle, byval tmpSize)

		dim CharWidths(99)

		if len(tmpFamily) = 0 then tmpFamily = "Arial"
		if len(tmpStyle) = 0 then tmpStyle = ""
		if len(tmpSize) = 0 then tmpSize = 1 'DEFAULT'

		Font.Name = tmpFamily

		select case ucase(tmpFamily)
			case "ARIAL"
				Font.Family = "helvetica"
			case "SYMBOL"
				Font.Family = "zapfdingbats"
				Font.Style = ""
			case else
				Font.Family = Font.FontFamily
		end select

		if instr(ucase(tmpStyle),"U") > 0 then
			Font.Underline = true
			tmpStyle = replace(tmpStyle,"U","")
		else
			Font.Underline = false
		end if

		if tmpStyle = "IB" then tmpStyle = "BI"

		if tmpSize = 0 then tmpSize = Font.SizePt

		dim tmpKey
		tmpKey = Font.Family & tmpStyle

		Font.Family = tmpFamily
		Font.Style = tmpStyle
		Font.SizePt = tmpSize
		Font.Size = tmpSize / k

		if Font.Style = "B" then
			Font.Current = 2
		else
			Font.Current = 1
		end if

		if CurrentPage > 0 then

			Out("BT /F" & Font.Current & " " & formatnumber(Font.SizePt,2,-1,0,0) & " Tf ET")

		end if

	end sub

	public sub SetFontSize(byval tmpSize)

		if tmpSize <> Font.SizePt then

			Font.SizePt = tmpSize
			Font.Size = tmpSize / k

			if CurrentPage > 0 then
				Out("BT /F" & Font.Current & " " & formatnumber(Font.SizePt,2,-1,0,0) & " Tf ET")
			end if

		end if

	end sub

	public sub SetLineStyle(byval W, byval Cap, byval Jn, byval Dash, byval Phase)

		dim str
		str = ""

		if not isnumeric(W) then W = 1

		if W > 0 then str = W & " w"

		select case Cap
			case 0
				str = str & " butt J"
			case 1
				str = str & " round J"
			case 2
				str = str & " square J"
		end select

		select case Jn
			case 0
				str = str & " miter j"
			case 1
				str = str & " round j"
			case 2
				str = str & " bevel j"
		end select

		if len(Dash) > 0 then

			str = str & " ["

			dim i
			for i = 1 to len(Dash)

				str = str & " " & mid(Dash,i,1)

			next

			str = str & " ] " & Phase & " d"

		end if

		str = str & vblf
		Out(str)

	end sub

	public function AddPage()

		if CurrentPage > 0 then
			InFooter = true
			call Footer
			InFooter = false
			call EndPage
		end if

		i = BeginPage()
		Out("2 J")

		Out(formatnumber(LineWidth * k,2,-1,0,0) & " w")

		call SetFont(Font.Family, Font.Style, Font.SizePt)

		if Colour.Flag = true then

			if Colour.Draw <> "" then Out(Colour.Draw)
			if Colour.Fill <> "" then Out(Colour.Fill)

		end if

		call Header

	end function

	public sub Text(byval tmpX, byval tmpY, byval tmpText)

		if tmpX = -1 then tmpX = Page.X
		if tmpY = -1 then tmpY = Page.Y

		tmpText = replace(tmpText, "\\","\\\\")
		tmpText = replace(tmpText, ")","\)")
		tmpText = replace(tmpText, "(","\(")

		xs = "BT " & formatnumber(tmpX * k,2,-1,0,0) & " " & formatnumber((Page.Height - tmpY) * k,2,-1,0,0) & " Td (" & tmpText & ") Tj ET"

		if Colour.Flag = true then
			xs = "q " & Colour.Text & " " & xs & " Q"
		end if

		Out(xs)

	end sub

	public sub SetTextColour(byval R, byval G, byval B)

		if not isnumeric(R) then R = 0
		if not isnumeric(G) then G = 0
		if not isnumeric(B) then B = 0

		if R = 0 and G = 0 and B = 0 then
			Colour.Text = "0 g"
		else

			if R > 0 then R = formatnumber(R / 255,4,-1,0,0)
			if G > 0 then G = formatnumber(G / 255,4,-1,0,0)
			if B > 0 then B = formatnumber(B / 255,4,-1,0,0)

			Colour.Text = R & " " & G & " " & B & " rg"

		end if

		Colour.Flag = true

	end sub

	public sub SetDrawColour(byval R, byval G, byval B)

		if not isnumeric(R) then R = 0
		if not isnumeric(G) then G = 0
		if not isnumeric(B) then B = 0

		if R = 0 and G = 0 and B = 0 then
			Colour.Draw = "0 G"
		else

			if R > 0 then R = formatnumber(R / 255,4,-1,0,0)
			if G > 0 then G = formatnumber(G / 255,4,-1,0,0)
			if B > 0 then B = formatnumber(B / 255,4,-1,0,0)

			Colour.Draw = R & " " & G & " " & B & " RG"

		end if

		if CurrentPage > 0 then
			Out(Colour.Draw)
		end if

	end sub

	public sub SetFillColour(byval R, byval G, byval B)

		if not isnumeric(R) then R = 0
		if not isnumeric(G) then G = 0
		if not isnumeric(B) then B = 0

		if R = 0 and G = 0 and B = 0 then
			Colour.Fill = "0 g"
		else

			if R > 0 then R = formatnumber(R / 255,4,-1,0,0)
			if G > 0 then G = formatnumber(G / 255,4,-1,0,0)
			if B > 0 then B = formatnumber(B / 255,4,-1,0,0)

			Colour.Fill = R & " " & G & " " & B & " rg"

		end if

		if CurrentPage > 0 then
			Out(Colour.Fill)
		end if

	end sub

	public sub Line(byval X1, byval Y1, byval X2, byval Y2)

		X1 = formatnumber(X1 * k,2,0,0,0)
		Y1 = formatnumber((Page.Height - Y1) * k,2,0,0,0)

		X2 = formatnumber(X2 * k,2,0,0,0)
		Y2 = formatnumber((Page.Height - Y2) * k,2,0,0,0)

		Out(X1 & " " & Y1 & " m " & X2 & " " & Y2 & " l S")

	end sub

	public sub SetLineWidth(byval W)

		LineWidth = W

		if CurrentPage > 0 then
			Out(LineWidth * k & " w")
		end if

	end sub

	private sub BeginDoc()

		State = 1
		Out("%PDF-1.3")

	end sub

	private function BeginPage()

		CurrentPage = CurrentPage + 1

		Pages(CurrentPage) = ""
		State = 2

		Page.X = Margin.LeftX
		Page.Y = Margin.Top

		LastH = 0

		FontFamily = ""

		if Reorientate = "" then 
			Reorientate = DefOrientation
		else
			if Reorientate <> DefOrientation then
				OrientationChanges(CurrentPage) = true
			end if
		end if

		if Reorientate <> CurOrientation then

			if Reorientate = "P" then
				Page.Pt.Width = Doc.Pt.Width
				Page.Pt.Height = Doc.Pt.Height
				Page.Width = Doc.Width
				Page.Height = Doc.Height
			else
				Page.Pt.Width = Doc.Pt.Height
				Page.Pt.Height = Doc.Pt.Width
				Page.Width = Doc.Height
				Page.Height = Doc.Width
			end if

			PageBreakTrigger = Margin.Bottom
			CurOrientation = Reorientate

		end if

	end function

	private sub Out(byval Code)

		if State = 2 then
			Pages(CurrentPage) = Pages(CurrentPage) & Code & vblf
		else
			Buffer = Buffer & Code & vblf
		end if

	end sub

	private sub EndPage()
		State = 1
	end sub

	private sub PutPages()

		dim Temp
		set Temp = new PDF_Coords

		if DefOrientation = "P" then
			Temp.Pt.Width = Doc.Pt.Width
			Temp.Pt.Height = Doc.Pt.Height
		else
			Temp.Pt.Width = Doc.Pt.Height
			Temp.Pt.Height = Doc.Pt.Width
		end if

		Filters = ""

		for i = 1 to CurrentPage

			call NewObj
			Out("<</Type /Page")
			Out("/Parent 1 0 R")

			if OrientationChanges(i) = true then
				Out("/MediaBox [0 0 " & Temp.Pt.Height & " " & Temp.Pt.Width)
			end if

			Out("/Resources 2 0 R")

			Out("/Contents " & n + 1 & " 0 R>>")
			Out("endobj")

			if len(Compress) > 0 then
				xp = Pages(i) 'gzcompress'
			else
				xp = Pages(i)
			end if

			call NewObj
			Out("<<" & Filters & "/Length " & len(xp) & ">>")
			call PutStream(xp)
			Out("endobj")

		next

		Offsets(1) = len(Buffer)
		Out("1 0 obj")
		Out("<</Type /Pages")
		xkids = "/Kids ["
		for xi = 0 to CurrentPage - 1
			xkids = xkids & (3 + 2 * xi) & " 0 R "
		next
		Out(xkids & "]")
		Out("/Count " & CurrentPage)
		Out("/MediaBox [0 0 " & Temp.Pt.Width & " " & Temp.Pt.Height & "]")
		Out(">>")
		Out("endobj")

		set Temp = nothing

	end sub

	private sub EndDoc()

		call PutPages
		call PutResources
		call NewObj
		Out("<<")
		call PutInfo
		Out(">>")
		Out("endobj")

		ZoomMode = "FULLPAGE"

		call NewObj
		Out("<<")
		call PutCatalog
		Out(">>")
		Out("endobj")

		xo = len(Buffer)
		Out("xref")
		Out("0 " & n + 1)
		Out("0000000000 65535 f ")

		for xi = 1 to n
			Out(stringPadLeft(Offsets(xi),"0",10) & " 00000 n ")
		next

		Out("trailer")

		Out("<<")
		call PutTrailer
		Out(">>")
		Out("startxref")
		Out(xo)
		Out("%%EOF")
		State = 3

	end sub

	private function stringPadLeft (byval strValue, byval strPadchar, byval intLength)

		if len(strValue) > intLength then intLength = len(strValue)
		stringPadLeft = string(intLength - Len(strValue), strPadchar) & strValue

	end function

	private sub PutStream(byval xs)

		Out("stream")
		Out(xs)
		Out("endstream")

	end sub

	public function GetBuffer()

		GetBuffer = Buffer

	end function

	private sub NewObj()

		n = n + 1

		Offsets(n) = len(Buffer)
		Out(n & " 0 obj")

	end sub

	private sub PutResources()

		call PutFonts
		call PutImages

		Offsets(2) = len(Buffer)

		Out("2 0 obj")
		Out("<</ProcSet [/PDF /Text /ImageB /ImageC /ImageI]")

		Out("/Font <<")
		Out("/F1 " & Font.ObjNumber(1) & " 0 R")
		Out("/F2 " & Font.ObjNumber(2) & " 0 R")
		Out(">>")

		Out("/XObject <<")

		if ThisImage.Total > 0 then

			dim i
			for i = 1 to ThisImage.Total
				Out("/I" & i & " " & ThisImage.ObjNumber(i) & " 0 R")
			next

		end if

		Out(">>")
		Out(">>")
		Out("endobj")

	end sub

	private sub PutFonts()

		xnf = n

		for xk = 1 to 2

			call NewObj

			Font.ObjNumber(xk) = n

			Out("<</Type /Font")
			Out("/BaseFont /" & Font.ThisFonts(xk,conFont_name))

			if Font.ThisFonts(xk,conFont_type) = "core" then
				
				Out("/Subtype /Type1")
				Out("/Encoding /WinAnsiEncoding")

			else

				Out("/Subtype /" & Font.ThisFonts(xk,conFont_type))
				Out("/FirstChar 32")
				Out("/LastChar 255")
				Out("/Widths " & n + 1 & " 0 R")
				Out("/FontDescriptor " & n + 2 & " 0 R")

			end if

			Out(">>")
			Out("endobj")

		next

	end sub

	private sub PutImages()

		if ThisImage.Total > 0 then

			dim i
			for i = 1 to ThisImage.Total

				call NewObj
				ThisImage.ObjNumber(i) = n

				Out(ThisImage.Images(i))
				Out("endobj")

			next

		end if

	end sub

	private sub PutInfo()

		Out("/Producer (PDFDoc by Simon Beal [inspired by FPDF])")

		if len(Title) > 0 then Out("/Title " & Title)
		if len(Author) > 0 then Out("/Author " & Author)
		if len(Keywords) > 0 then Out("/Keywords " & Keywords)
		if len(Creator) > 0 then Out("/Creator " & Creator)

		Out("/CreationDate (D:" & year(now) & month(now) & day(now) & hour (now) & minute(now) & second(now) & ")")

	end sub

	private sub PutCatalog()

		Out("/Type /Catalog")
		Out("/Pages 1 0 R")

		select case ucase(ZoomMode)
			case "FULLPAGE"
				Out("/OpenAction [3 0 R /Fit]")
			case "FULLWIDTH"
				Out("/OpenAction [3 0 R /FitH null]")
			case "REAL"
				Out("/OpenAction [3 0 R /XYZ null null 1]")
			case else
				if isnumeric(ZoomMode) then
					Out("/OpenAction [3 0 R /XYZ null null " & ZoomMode / 100 & "]")
				end if
		end select

		select case ucase(LayoutMode)
			case "SINGLE"
				Out("/PageLayout /SinglePage")
			case "CONTINUOUS"
				Out("/PageLayout /OneColumn")
			case "TWO"
				Out("/PageLayout /TwoColumnLeft")
			case else
				'// "No LayoutMode"
		end select

	end sub

	private sub PutTrailer()

		Out("/Size " & n + 1)
		Out("/Root " & n & " 0 R")
		Out("/Info " & n - 1 & " 0 R")

	end sub

	public sub Image(byval tmpFile, byval tmpX, byval tmpY, byval tmpWidth)

		dim tmpHeight

		if ThisImage.Open(tmpFile) then

			ThisImage.Total = ThisImage.Total + 1
			tmpHeight = tmpWidth * ThisImage.Height / ThisImage.Width
			LastImageHeightMM = tmpHeight

			Out("q " & formatnumber(tmpWidth * k,2,0,0,0) & _
				" 0 0 " & formatnumber(tmpHeight * k,2,0,0,0) & _
				" " & formatnumber(tmpX * k,2,0,0,0) & _
				" " & formatnumber((Page.Height - (tmpY + tmpHeight)) * k,2,0,0,0) & _
				" cm " & _
				"/I" & ThisImage.Total & _
				" Do Q")

			ThisImage.Images(ThisImage.Total) = "<</Type /XObject" & vblf & _
												"/Subtype /Image" & vblf & _
												"/Width " & ThisImage.Width & vblf & _
												"/Height " & ThisImage.Height & vblf & _
												"/ColorSpace /" & ThisImage.ColourSpace & vblf & _
												"/BitsPerComponent " & ThisImage.Bits & vblf & _
												"/Filter /DCTDecode" & vblf & _
												"/Length " & ThisImage.Size & _
												">>" & vblf

			ThisImage.Images(ThisImage.Total) = ThisImage.Images(ThisImage.Total) & _
												"stream" & vblf & ThisImage.Data & vblf & _
												"endstream"

		end if

	end sub

	public sub Rotate(byval Angle, byval aX, byval aY)
	
		if not isnumeric(aX) then aX = Page.X
		if not isnumeric(aY) then aY = Page.Y

		aY = (Page.Height - aY) * k
		aX = aX * k

		dim Transform(5)
		Transform(0) = cos(DegToRad(Angle))
		Transform(1) = sin(DegToRad(Angle))
		Transform(2) = - Transform(1)
		Transform(3) = Transform(0)
		Transform(4) = aX + Transform(1) * aY - Transform(0) * aX
		Transform(5) = aY - Transform(0) * aY - Transform(1) * aX

		dim tmp
		dim i

		for i = 0 to 5
			tmp = tmp & formatnumber(Transform(i),5,-1,0,0) & " "
		next

		Out(tmp & "cm")

	end sub

	public sub Skew(byval AngleX, byval AngleY, byval aX, byval aY)

		if not isnumeric(aX) then aX = Page.X
		if not isnumeric(aY) then aY = Page.Y

		if AngleX >= -90 and AngleX <=90 and AngleY >= -90 and AngleY <= 90 then
			
			aX = aX * k
			aY = (Page.Height - aY) * k

			dim Transform(5)
			Transform(0) = 1
			Transform(1) = tan(DegToRad(AngleY))
			Transform(2) = tan(DegToRad(AngleX))
			Transform(3) = 1
			Transform(4) = 0 'Transform(2) * aY
			Transform(5) = 0 'Transform(1) * aX

			dim tmp
			dim i

			for i = 0 to 5
				tmp = tmp & formatnumber(Transform(i),5,-1,0,0) & " "
			next

			Out(tmp & "cm")

		end if

	end sub

	public sub Ellipse(byval cx, byval cy, byval rx, byval ry, byval style)

		dim op

		dim M_SQRT2
		M_SQRT2 = sqr(2)

		select case ucase(style)
			case "F"
				op = "f"
			case "FD", "DF"
				op = "B"
			case else
				op = "S"
		end select

		X1 = 4/3*(M_SQRT2-1)*rx
		Y1 = 4/3*(M_SQRT2-1)*ry

		dim p1,p2,p3

		dim h
		h = Page.Height

		'// MOVE
		p1 = formatnumber((cx+rx)*k,2,0,0,0) & " " & formatnumber((h-cy)*k,2,0,0,0) & " m"

		Out(p1)

		'// BEZIER CURVES

		p1 = formatnumber((cx+rx)*k,2,0,0,0) & " " & formatnumber((h-(cy-Y1))*k,2,0,0,0) & " "
		p2 = formatnumber((cx+X1)*k,2,0,0,0) & " " & formatnumber((h-(cy-ry))*k,2,0,0,0) & " "
		p3 = formatnumber(cx*k,2,0,0,0) & " " & formatnumber((h-(cy-ry))*k,2,0,0,0) & " c"

		Out(p1 & p2 & p3)

		p1 = formatnumber((cx-X1)*k,2,0,0,0) & " " & formatnumber((h-(cy-ry))*k,2,0,0,0) & " "
		p2 = formatnumber((cx-rx)*k,2,0,0,0) & " " & formatnumber((h-(cy-Y1))*k,2,0,0,0) & " "
		p3 = formatnumber((cx-rx)*k,2,0,0,0) & " " & formatnumber((h-cy)*k,2,0,0,0) & " c"

		Out(p1 & p2 & p3)

		p1 = formatnumber((cx-rx)*k,2,0,0,0) & " " & formatnumber((h-(cy+Y1))*k,2,0,0,0) & " "
		p2 = formatnumber((cx-X1)*k,2,0,0,0) & " " & formatnumber((h-(cy+ry))*k,2,0,0,0) & " "
		p3 = formatnumber(cx*k,2,0,0,0) & " " & formatnumber((h-(cy+ry))*k,2,0,0,0) & " c"

		Out(p1 & p2 & p3)

		p1 = formatnumber((cx+X1)*k,2,0,0,0) & " " & formatnumber((h-(cy+ry))*k,2,0,0,0) & " "
		p2 = formatnumber((cx+rx)*k,2,0,0,0) & " " & formatnumber((h-(cy+Y1))*k,2,0,0,0) & " "
		p3 = formatnumber((cx+rx)*k,2,0,0,0) & " " & formatnumber((h-cy)*k,2,0,0,0) & " c " & op

		Out(p1 & p2 & p3)

	end sub

	public sub StartTransform()
		Out("q")
	end sub

	public sub EndTransform()
		Out("Q")
	end sub

	private function DegToRad(byval Num)

		if isnumeric(Num) then
			DegToRad = Num / 57.2957795131 '// (180 / 3.14159265359)
		else
			DegToRad = 0
		end if

	end function

	public sub Box(byval xx, byval xy, byval xw, byval xh, byval xstyle)

		dim xOp

		if xstyle = "F" then
			xOp = "f"
		elseif xstyle = "FD" or xstyle = "DF" then
			xOp = "B"
		else
			xOp = "S"
		end if

		Out(formatnumber(xx * k,2,-1,0,0) & " " & _
			formatnumber((Page.Height - xy) * k,2,-1,0,0) & " " & _
			formatnumber(xw * k,2,-1,0,0) & " " & _
			formatnumber(xh * k,2,-1,0,0) & " re " & _
			xOp)

	end sub

	public sub Code39(byval xPos, byval yPos, byval MaxWidth, byval Code)

		if len(trim(Code)) > 0 then

			dim OriginalX
			OriginalX = xPos

			dim OriginalY
			OriginalY = yPos

			dim OriginalCode
			OriginalCode = Code

			dim Wide
			dim Narrow
			dim Gap

			Wide = Barcode.Baseline

			if len(OriginalCode) >= 15 then 
				Narrow = Barcode.Baseline / 2.5
			else
				Narrow = Barcode.Baseline / 3
			end if

			Gap = Narrow

			dim LineWidth

			SetFillColour 0,0,0
			Code = "*" & ucase(Code) & "*"

			dim xChar
			dim bar
			dim i
			dim Seg

			dim passloop
			dim beginpass

			if MaxWidth = 0 then
				beginpass = 2
			else
				beginpass = 1
			end if

			for passloop = beginpass to 2
				for i = 1 to len(Code)

					xChar = asc(mid(Code,i,1))
					Seg = Barcode.CharCode(xChar)

					for bar = 0 to 8

						if mid(Seg,bar+1,1) = "n" then
							LineWidth = Narrow
						else
							LineWidth = Wide
						end if

						if bar mod 2 = 0 then

							if passloop = 2 then

								Box formatnumber(xPos,5,-1,0,0), _
									formatnumber(yPos,5,-1,0,0), _
									formatnumber(LineWidth,5,-1,0,0), _
									formatnumber(Barcode.Height,5,-1,0,0), "F"

							end if
						end if

						xPos = xPos + LineWidth + Gap

					next

					xPos = xPos + Gap

'					if Barcode.Trick = true then
'						tmp = getRandomRange(0,Barcode.Height - 1)
'						yPos = OriginalY + (tmp - (Barcode.Height / 2))
'					end if

				next

				if passloop = 1 then

					w = xPos - OriginalX

					if MaxWidth <> w then

						if MaxWidth > w then
							scale = w / MaxWidth
							w = w - OriginalX
							OriginalX = (Page.Width / 2) - (w / 2)
						else
							scale = MaxWidth / w
						end if

						Barcode.Baseline = Barcode.Baseline * scale

					end if

					xPos = OriginalX
					Wide = Barcode.Baseline

					if len(OriginalCode) >= 15 then 
						Narrow = Barcode.Baseline / 2.5
					else
						Narrow = Barcode.Baseline / 3
					end if

					Gap = Narrow

				end if
			next
		end if

		Barcode.Reset

	end sub

	Private Sub Class_Initialize

		set Colour = new PDF_Colours
		set Doc = new PDF_Coords
		set Font = new PDF_Fonts
		set Margin = new PDF_Margins
		set Page = new PDF_Coords
		set ThisImage = new PDF_Image
		set Barcode = new PDF_Barcode

		call Reset()

	End Sub

	public sub Reset()

		CurrentPage = 0

		dim i
		for i = 1 to 99
			Pages(i) = ""
			OrientationChanges(i) = false
		next

		DefOrientation = "P"
		LayoutMode = "CONTINUOUS"

	end sub

end class

class PDF_Barcode

	public Trick
	public Height
	public Baseline
	public CharCode(255)

	Private Sub Class_Initialize

		call Reset()

		CharCode(asc("0")) = "nnnwwnwnn"
		CharCode(asc("1")) = "wnnwnnnnw"
		CharCode(asc("2")) = "nnwwnnnnw"
		CharCode(asc("3")) = "wnwwnnnnn"
		CharCode(asc("4")) = "nnnwwnnnw"
		CharCode(asc("5")) = "wnnwwnnnn"
		CharCode(asc("6")) = "nnwwwnnnn"
		CharCode(asc("7")) = "nnnwnnwnw"
		CharCode(asc("8")) = "wnnwnnwnn"
		CharCode(asc("9")) = "nnwwnnwnn"
		CharCode(asc("A")) = "wnnnnwnnw"
		CharCode(asc("B")) = "nnwnnwnnw"
		CharCode(asc("C")) = "wnwnnwnnn"
		CharCode(asc("D")) = "nnnnwwnnw"
		CharCode(asc("E")) = "wnnnwwnnn"
		CharCode(asc("F")) = "nnwnwwnnn"
		CharCode(asc("G")) = "nnnnnwwnw"
		CharCode(asc("H")) = "wnnnnwwnn"
		CharCode(asc("I")) = "nnwnnwwnn"
		CharCode(asc("J")) = "nnnnwwwnn"
		CharCode(asc("K")) = "wnnnnnnww"
		CharCode(asc("L")) = "nnwnnnnww"
		CharCode(asc("M")) = "wnwnnnnwn"
		CharCode(asc("N")) = "nnnnwnnww"
		CharCode(asc("O")) = "wnnnwnnwn" 
		CharCode(asc("P")) = "nnwnwnnwn"
		CharCode(asc("Q")) = "nnnnnnwww"
		CharCode(asc("R")) = "wnnnnnwwn"
		CharCode(asc("S")) = "nnwnnnwwn"
		CharCode(asc("T")) = "nnnnwnwwn"
		CharCode(asc("U")) = "wwnnnnnnw"
		CharCode(asc("V")) = "nwwnnnnnw"
		CharCode(asc("W")) = "wwwnnnnnn"
		CharCode(asc("X")) = "nwnnwnnnw"
		CharCode(asc("Y")) = "wwnnwnnnn"
		CharCode(asc("Z")) = "nwwnwnnnn"
		CharCode(asc("-")) = "nwnnnnwnw"
		CharCode(asc(".")) = "wwnnnnwnn"
		CharCode(asc(" ")) = "nwwnnnwnn"
		CharCode(asc("*")) = "nwnnwnwnn"
		CharCode(asc("$")) = "nwnwnwnnn"
		CharCode(asc("/")) = "nwnwnnnwn"
		CharCode(asc("+")) = "nwnnnwnwn"
		CharCode(asc("%")) = "nnnwnwnwn"

	End Sub

	public sub Reset()

		Trick = false
		Baseline = 0.5
		Height = 10

	end sub

end class

class PDF_Margins

	public Top
	public LeftX
	public RightX
	public Bottom
	public C

	Private Sub Class_Initialize

		call Reset()

	End Sub

	public sub Reset()

		Top = 0
		LeftX = 0
		RightX = 0
		Bottom = 0
		C = 0

	end sub

end class

class PDF_Colours

	public Fill
	public Text
	public Draw
	public Flag

	Private Sub Class_Initialize

		call Reset()

	End Sub

	public sub Reset()

		Flag = false
		Fill = "0 g"
		Text = "0 g"
		Draw = "0 G"

	end sub

end class

const conFont_i = 1
const conFont_type = 2
const conFont_name = 3
const conFont_up = 4
const conFont_ut = 5
const conFont_cw = 6

class PDF_Fonts

	public Name
	public SizePt
	public Family
	public Style
	public Size
	public FontFamily
	public Current
	public Core(14,2)
	public Underline

	public ObjNumber(2)

	public Diffs(99)

	public ThisFonts(14,6)
	public Max

	private Helvetica
	private HelveticaB
	private HelveticaI
	private HelveticaBI

	public function GetWidth(byval Phrase)

		if len(Phrase) > 0 then

			dim i
			dim Characters(256)

			select case Current
				case 1 '"Helvetica"
					for i = 1 to ubound(Helvetica)
						Characters(i) = Helvetica(i)
					next
				case 2 '"HelveticaB"
					for i = 1 to ubound(HelveticaB)
						Characters(i) = HelveticaB(i)
					next
				case else
					for i = 1 to ubound(Helvetica)
						Characters(i) = Helvetica(i)
					next
			end select

			for i = 1 to len(Phrase)
				GetWidth = GetWidth + Characters(asc(mid(Phrase,i,1)))
			next

			GetWidth = (GetWidth * Size) / 1000

		else

			GetWidth = 0

		end if

	end function

	Private Sub Class_Initialize

		dim i
		for i = 1 to 6
			ThisFonts(i, conFont_type) = "core"
			ThisFonts(i, conFont_up) = -100
			ThisFonts(i, conFont_ut) = 50
		next

		ThisFonts(1, conFont_name) = "Helvetica"
		ThisFonts(2, conFont_name) = "Helvetica-Bold"

		Core(1,1) = "courier"
		Core(1,2) = "Courier"

		Core(2,1) = "courierB"
		Core(2,2) = "Courier-Bold"

		Core(3,1) = "courierI"
		Core(3,2) = "Courier-Oblique"

		Core(4,1) = "courierBI"
		Core(4,2) = "Courier-BoldOblique"

		Core(5,1) = "helvetica"
		Core(5,2) = "Helvetica"

		Core(6,1) = "helveticaB"
		Core(6,2) = "Helivetica-Bold"

		Core(7,1) = "helveticaI"
		Core(7,2) = "Helivetica-Oblique"

		Core(8,1) = "helveticaBI"
		Core(8,2) = "Helvetica-BoldOblique"

		Core(9,1) = "times"
		Core(9,2) = "Times-Roman"

		Core(10,1) = "timesB"
		Core(10,2) = "Times-Bold"

		Core(11,1) = "timesI"
		Core(11,2) = "Times-Italic"

		Core(12,1) = "timesBI"
		Core(12,2) = "Times-BoldItalic"

		Core(13,1) = "symbol"
		Core(13,2) = "Symbol"

		Core(14,1) = "zapfdingbats"
		Core(14,2) = "ZapfDingbats"

		Helvetica = array( _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 355, 556, 556, 889, 667, 191, _
			333, 333, 389, 584, 278, 333, 278, 278, _
			556, 556, 556, 556, 556, 556, 556, 556, _
			556, 556, 278, 278, 584, 584, 584, 556, _
			1015, 667, 667, 722, 722, 667, 611, 778, _
			722, 278, 500, 667, 556, 833, 722, 778, _
			667, 778, 722, 667, 611, 722, 667, 944, _
			667, 667, 611, 278, 278, 278, 469, 556, _
			333, 556, 556, 500, 556, 556, 278, 556, _
			556, 222, 222, 500, 222, 833, 556, 556, _
			556, 556, 333, 500, 278, 556, 500, 722, _
			500, 500, 500, 334, 260, 334, 584, 350, _
			556, 350, 222, 556, 333, 1000, 556, 556, _
			333, 1000, 667, 333, 1000, 350, 611, 350, _
			350, 222, 222, 333, 333, 350, 556, 1000, _
			333, 1000, 500, 333, 944, 350, 500, 667, _
			278, 333, 556, 556, 556, 556, 260, 556, _
			333, 737, 370, 556, 584, 333, 737, 333, _
			400, 584, 333, 333, 333, 556, 537, 278, _
			333, 333, 365, 556, 834, 834, 834, 611, _
			667, 667, 667, 667, 667, 667, 1000, 722, _
			667, 667, 667, 667, 278, 278, 278, 278, _
			722, 722, 778, 778, 778, 778, 778, 584, _
			778, 722, 722, 722, 722, 667, 667, 611, _
			556, 556, 556, 556, 556, 556, 889, 500, _
			556, 556, 556, 556, 278, 278, 278, 278, _
			556, 556, 556, 556, 556, 556, 556, 584, _
			611, 556, 556, 556, 556, 500, 556, 500)

		HelveticaB = array( _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 333, 474, 556, 556, 889, 722, 238, _
			333, 333, 389, 584, 278, 333, 278, 278, _
			556, 556, 556, 556, 556, 556, 556, 556, _
			556, 556, 333, 333, 584, 584, 584, 611, _
			975, 722, 722, 722, 722, 667, 611, 778, _
			722, 278, 556, 722, 611, 833, 722, 778, _
			667, 778, 722, 667, 611, 722, 667, 944, _
			667, 667, 611, 333, 278, 333, 584, 556, _
			333, 556, 611, 556, 611, 556, 333, 611, _
			611, 278, 278, 556, 278, 889, 611, 611, _
			611, 611, 389, 556, 333, 611, 556, 778, _
			556, 556, 500, 389, 280, 389, 584, 350, _
			556, 350, 278, 556, 500, 1000, 556, 556, _
			333, 1000, 667, 333, 1000, 350, 611, 350, _
			350, 278, 278, 500, 500, 350, 556, 1000, _
			333, 1000, 556, 333, 944, 350, 500, 667, _
			278, 333, 556, 556, 556, 556, 280, 556, _
			333, 737, 370, 556, 584, 333, 737, 333, _
			400, 584, 333, 333, 333, 611, 556, 278, _
			333, 333, 365, 556, 834, 834, 834, 611, _
			722, 722, 722, 722, 722, 722, 1000, 722, _
			667, 667, 667, 667, 278, 278, 278, 278, _
			722, 722, 778, 778, 778, 778, 778, 584, _
			778, 722, 722, 722, 722, 667, 667, 611, _
			556, 556, 556, 556, 556, 556, 889, 556, _
			556, 556, 556, 556, 278, 278, 278, 278, _
			611, 611, 611, 611, 611, 611, 611, 584, _
			611, 611, 611, 611, 611, 556, 611, 556)

		HelveticaI = array( _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 355, 556, 556, 889, 667, 191, _
			333, 333, 389, 584, 278, 333, 278, 278, _
			556, 556, 556, 556, 556, 556, 556, 556, _
			556, 556, 278, 278, 584, 584, 584, 556, _
			1015, 667, 667, 722, 722, 667, 611, 778, _
			722, 278, 500, 667, 556, 833, 722, 778, _
			667, 778, 722, 667, 611, 722, 667, 944, _
			667, 667, 611, 278, 278, 278, 469, 556, _
			333, 556, 556, 500, 556, 556, 278, 556, _
			556, 222, 222, 500, 222, 833, 556, 556, _
			556, 556, 333, 500, 278, 556, 500, 722, _
			500, 500, 500, 334, 260, 334, 584, 350, _
			556, 350, 222, 556, 333, 1000, 556, 556, _
			333, 1000, 667, 333, 1000, 350, 611, 350, _
			350, 222, 222, 333, 333, 350, 556, 1000, _
			333, 1000, 500, 333, 944, 350, 500, 667, _
			278, 333, 556, 556, 556, 556, 260, 556, _
			333, 737, 370, 556, 584, 333, 737, 333, _
			400, 584, 333, 333, 333, 556, 537, 278, _
			333, 333, 365, 556, 834, 834, 834, 611, _
			667, 667, 667, 667, 667, 667, 1000, 722, _
			667, 667, 667, 667, 278, 278, 278, 278, _
			722, 722, 778, 778, 778, 778, 778, 584, _
			778, 722, 722, 722, 722, 667, 667, 611, _
			556, 556, 556, 556, 556, 556, 889, 500, _
			556, 556, 556, 556, 278, 278, 278, 278, _
			556, 556, 556, 556, 556, 556, 556, 584, _
			611, 556, 556, 556, 556, 500, 556, 500)

		HelveticaBI = array( _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 278, 278, 278, 278, 278, 278, 278, _
			278, 333, 474, 556, 556, 889, 722, 238, _
			333, 333, 389, 584, 278, 333, 278, 278, _
			556, 556, 556, 556, 556, 556, 556, 556, _
			556, 556, 333, 333, 584, 584, 584, 611, _
			975, 722, 722, 722, 722, 667, 611, 778, _
			722, 278, 556, 722, 611, 833, 722, 778, _
			667, 778, 722, 667, 611, 722, 667, 944, _
			667, 667, 611, 333, 278, 333, 584, 556, _
			333, 556, 611, 556, 611, 556, 333, 611, _
			611, 278, 278, 556, 278, 889, 611, 611, _
			611, 611, 389, 556, 333, 611, 556, 778, _
			556, 556, 500, 389, 280, 389, 584, 350, _
			556, 350, 278, 556, 500, 1000, 556, 556, _
			333, 1000, 667, 333, 1000, 350, 611, 350, _
			350, 278, 278, 500, 500, 350, 556, 1000, _
			333, 1000, 556, 333, 944, 350, 500, 667, _
			278, 333, 556, 556, 556, 556, 280, 556, _
			333, 737, 370, 556, 584, 333, 737, 333, _
			400, 584, 333, 333, 333, 611, 556, 278, _
			333, 333, 365, 556, 834, 834, 834, 611, _
			722, 722, 722, 722, 722, 722, 1000, 722, _
			667, 667, 667, 667, 278, 278, 278, 278, _
			722, 722, 778, 778, 778, 778, 778, 584, _
			778, 722, 722, 722, 722, 667, 667, 611, _
			556, 556, 556, 556, 556, 556, 889, 556, _
			556, 556, 556, 556, 278, 278, 278, 278, _
			611, 611, 611, 611, 611, 611, 611, 584, _
			611, 611, 611, 611, 611, 556, 611, 556)

		call Reset()

	End Sub

	public sub Reset()

		Max = 0
		FontFamily = "helvetica"
		Current = 1

	end sub

end class

class PDF_Coords

	public X
	public Y
	public Height
	public Width
	public Pt
	public LineHeight

	Private Sub Class_Initialize

		set Pt = new PDF_PtSizes

		call Reset()

	End Sub

	public sub Reset()

		X = 0
		Y = 0
		Height = 0
		Width = 0

	end sub

end class

class PDF_PtSizes

	public Height
	public Width

	Private Sub Class_Initialize

		call Reset()

	End Sub

	public sub Reset()

		Height = 0
		Width = 0

	end sub

end class

const conImage_TEM 	= &H01
const conImage_SOF 	= &Hc0
const conImage_DHT 	= &Hc4
const conImage_JPGA = &Hc8
const conImage_DAC 	= &Hcc
const conImage_RST 	= &Hd0
const conImage_SOI 	= &Hd8
const conImage_EOI 	= &Hd9
const conImage_SOS 	= &Hda
const conImage_DQT 	= &Hdb
const conImage_DNL 	= &Hdc
const conImage_DRI 	= &Hdd
const conImage_DHP 	= &Hde
const conImage_EXPx = &Hdf
const conImage_APP 	= &He0
const conImage_JPG 	= &Hf0
const conImage_COM 	= &Hfe

class PDF_Image

	public Buffer
	public Binary
	public Width
	public Height
	public MIME
	public Channels
	public Bits
	public Size
	public Extension
	public ID
	Public Marker
	public Length
	public Total
	public Data
	public ColourSpace
	public ObjNumber(99)
	public Images(99)

	public function Open(byval Filename)

		Open = false
		call Reset

		Set Buffer = CreateObject("ADODB.Stream")

		Buffer.CharSet = "ISO-8859-1"
		Buffer.Type = 2 'adTypeText '(2)'
		Buffer.Open
		Buffer.LoadFromFile(Server.MapPath(Filename))
		Buffer.Position = 0

		Size = Buffer.Size

		dim i

		if len(Filename) > 0 then

			i = instrrev(Filename,".")
			if i > 0 then Extension = lcase(mid(Filename,i + 1))

		end if

		select case Extension
			case "jpg","jpeg"
				MIME = "image/jpeg"
				call ParseJpeg
			case "png"
				MIME = "image/png"
		end select

		Buffer.Position = 0
		Data = Buffer.ReadText

		Open = true

	end function

	private sub ParseJpeg()

		ID = 2
		dim rdByte
		dim n

		rdByte = Read(1,10)

		if rdByte = &Hff then

			dim Skip
			Skip = false

			while Buffer.Position < Buffer.Size and Skip = false

				n = n + 1

				Marker = Read(1,10)

				while Marker = &Hff
					Marker = Read(1,10)
				wend

				select case cint(Marker)
					case conImage_DHP, conImage_SOF+0, conImage_SOF+1, conImage_SOF+2, _
						 conImage_SOF+3, conImage_SOF+5, conImage_SOF+6, conImage_SOF+7, _
						 conImage_SOF+9, conImage_SOF+10, conImage_SOF+11, conImage_SOF+13, _
						 conImage_SOF+14, conImage_SOF+15

						Length = Read(2,10)
						Bits = Read(1,10)
						Height = Read(2,10)
						Width = Read(2,10)
						Channels = Read(1,10)

						Skip = true

					case conImage_APP+0, conImage_APP+1, conImage_App+2, conImage_APP+3, _
						 conImage_APP+4, conImage_APP+5, conImage_APP+6, conImage_APP+7, _
						 conImage_APP+8, conImage_APP+9, conImage_APP+10, conImage_APP+11, _
						 conImage_APP+12, conImage_APP+13, conImage_APP+14, conImage_APP+15, _
						 conImage_DRI, conImage_SOS, conImage_DHT, conImage_DAC, conImage_DNL, _
						 conImage_EXPx

						h = Read(2,10)
						i = clng(h) - 2

						Buffer.Position = Buffer.Position + i

				end select

			wend
		end if

		select case Channels
			case 3
				ColourSpace = "DeviceRGB"
			case 4
				ColourSpace = "DeviceCMYK"
			case else
				ColourSpace = "DeviceGrey"
		end select

	end sub

	private function Read(byval nB, byval Radix)

		dim res
		dim i

		if Radix = "string" then
			Buffer.ReadText(nb)
		else

			res = 0

			for i = 1 to nB

				ch = asc(Buffer.ReadText(1))

				if nb = 2 then
					select case i
						case 2
							res = res + ch
						case 1
							ch = ch * 256
							res = res + ch
					end select
				else
					res = res + ch
				end if

			next

			if res > 255 then res = toAscii(res)

		end if

		Read = res

	end function

	private function ToAscii(byval Code)

		select case Code
			case 8364
				ToAscii = 128
			case 8218
				ToAscii = 130
			case 402
				ToAscii = 131
			case 8222
				ToAscii = 132
			case 8230
				ToAscii = 133
			case 8224
				ToAscii = 134
			case 8225
				ToAscii = 135
			case 710
				ToAscii = 136
			case 8240
				ToAscii = 137
			case 352
				ToAscii = 138
			case 8249
				ToAscii = 139
			case 338
				ToAscii = 140
			case 381
				ToAscii = 142
			case 8216
				ToAscii = 145
			case 8217
				ToAscii = 146
			case 8220
				ToAscii = 147
			case 8221
				ToAscii = 148
			case 8226
				ToAscii = 149
			case 8211
				ToAscii = 150
			case 8212
				ToAscii = 151
			case 732
				ToAscii = 152
			case 8482
				ToAscii = 153
			case 353
				ToAscii = 154
			case 8250
				ToAscii = 155
			case 339
				ToAscii = 156
			case 382
				ToAscii = 158
			case 376
				ToAscii = 159
			case else
				'// "Error ASCII code"
				ToAscii = Code
		end select

	end function

	Private Sub Class_Initialize

		Total = 0
		call Reset()

	End Sub

	public sub Reset()

		Height = -1
		Width = -1
		Size = -1
		Marker = ""
		ID = 2
		MIME = ""
		ColourSpace = ""

	end sub

end class

%>