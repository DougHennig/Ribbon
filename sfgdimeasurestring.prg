text to lcText noshow
This is line 1
This is line 2
This is line 3
endtext
loMeasure = createobject('SFGDIMeasureString')
lnWidth = loMeasure.GetWidth(lcText, 'Tahoma', 10)
messagebox('Width: ' + transform(lnWidth) + chr(13) + ;
	'Height: ' + transform(loMeasure.nHeight) + chr(13) + ;
	'Chars: ' + transform(loMeasure.nChars) + chr(13) + ;
	'Lines: ' + transform(loMeasure.nLines))

lnWidth = loMeasure.GetWidth('Date Required', 'Arial', 10, 'B')
messagebox('Width: ' + transform(lnWidth) + chr(13) + ;
	'Height: ' + transform(loMeasure.nHeight) + chr(13) + ;
	'Chars: ' + transform(loMeasure.nChars) + chr(13) + ;
	'Lines: ' + transform(loMeasure.nLines))

loMeasure.SetSize(100, 22)
lnWidth = loMeasure.GetWidth(lcText, 'Tahoma', 10)
messagebox('Width: ' + transform(lnWidth) + chr(13) + ;
	'Height: ' + transform(loMeasure.nHeight) + chr(13) + ;
	'Chars: ' + transform(loMeasure.nChars) + chr(13) + ;
	'Lines: ' + transform(loMeasure.nLines))


*==============================================================================
* Class:			SFGDIMeasureString
* Based on:			Custom
* Purpose:			Calculates the dimensions of a string
* Author:			Doug Hennig
* Last revision:	03/26/2024
* Notes:			1.	Before calling MeasureString or GetWidth, you can
*						modify the settings of the oFormat, oFont, oSize, and
*						oGDI members as necessary.
*					2.	Scaling calculations adapted from Antonio Lopes'
*						DPIAwareManager: https://github.com/atlopes/DPIAwareManager
*==============================================================================

#define ccCR						chr(13)
	&& carriage return
#define ccLF						chr(10)
	&& line feed
#define ccCRLF						chr(13) + chr(10)
	&& carriage return + line feed
#define cnERR_ARGUMENT_INVALID		11
	&& Function argument value, type, or count is invalid
#define cnERR_TOO_FEW_ARGS			1229
	&& Too few arguments
#define cnFACTOR					104.166            
	&& cnFACTOR is the number of report units per pixel: report units per inch
	&& (10,000) divided by pixels per inch (96)
#define DPI_STANDARD				96
	&& standard screen DPI
#define DPI_STANDARD_SCALE			100
	&& standard scaling
#define DPI_MAX_SCALE				300
	&& max scaling
#define DC_LOGPIXELSX				88
	&& value to get X pixels

define class SFGDIMeasureString as Custom
	oGDI        = .NULL.
		&& a reference to a System.Drawing.Graphics object
	oFormat     = .NULL.
		&& a reference to a System.Drawing.StringFormat object
	oFont       = .NULL.
		&& a reference to a System.Drawing.Font object
	oSize       = .NULL.
		&& a reference to a System.Drawing.Size object
	oStringSize = .NULL.
		&& a reference to a System.Drawing.Size object for the size of the text
	nChars      = 0
		&& the number of characters fitted in the bounding box
	nLines      = 0
		&& the number of lines in the bounding box
	nWidth      = 0
		&& the width of the bounding box
	nHeight     = 0
		&& the height of the bounding box
	lGetDPIForWindow = .F.
		&& can we use GetDpiForWindow

	function Init
		if type('_screen.System.Drawing') <> 'O'
			do System
		endif type('_screen.System.Drawing') <> 'O'
		with _screen.System.Drawing

* Create the helper objects we need so we don't have to do it later (for
* performance reasons).

			This.oGDI    = .Graphics.FromHwnd(_screen.HWnd)
			This.oFormat = .StringFormat.New()

* Use anti-aliasing: it seems to give more accurate results.

			This.oGDI.TextRenderingHint = .Drawing2D.SmoothingMode.AntiAlias

* If the size of the layout box isn't specified, use a very large box so size
* isn't a factor.

			This.SetSize(100000, 100000)
		endwith

* Declare some API functions and determine which function we'll use to determine scaling.

		declare long GetWindowDC in Win32API ;
			long hWnd
		declare long ReleaseDC in Win32API ;
			long hWnd, long hDC
		declare long GetDeviceCaps in Win32API  ;
			long hDC, integer CapIndex
		try
			declare integer GetDpiForWindow in Win32API ;
				long hWnd
			This.lGetDPIForWindow = .T.
		catch
		endtry
	endfunc

* Nuke member objects.

	function Destroy
		store .NULL. to This.oGDI, This.oFormat, This.oFont, This.oSize, ;
			This.oStringSize
	endfunc

* Determine the dimensions of the bounding box for the specified string.

	function MeasureString(tcString, tcFontName, tnFontSize, tuStyle)
		local luStyle, ;
			lnChars, ;
			lnLines, ;
			lcString, ;
			lcFontName, ;
			lnFontSize, ;
			lcStyle, ;
			lnWidth

* Ensure the parameters are passed correctly.

		do case
			case vartype(tcString) <> 'C'
				error cnERR_ARGUMENT_INVALID
				return
			case pcount() > 1 and (vartype(tcFontName) <> 'C' or ;
				empty(tcFontName) or vartype(tnFontSize) <> 'N' or ;
				not between(tnFontSize, 1, 128))
				error cnERR_ARGUMENT_INVALID
				return
			case pcount() = 4 and not vartype(tuStyle) $ 'CN'
				error cnERR_ARGUMENT_INVALID
				return
		endcase
		luStyle = NULL
		do case

* Set up the font object if the font and size were specified.

			case pcount() > 1
				luStyle = iif(pcount() = 4, tuStyle, '')
				This.SetFont(tcFontName, tnFontSize, luStyle)

* If no font or size were specified, bug out with an error.

			case vartype(This.oFont) <> 'O'
				error cnERR_TOO_FEW_ARGS
				return
		endcase
		with This

* Initialize output variables used in GdipMeasureString.

			lnChars = 0
			lnLines = 0

* If the string has CHR(13) but not CHR(10), insert them. Also, GdipMeasureString
* seems to measure strings with CR a little narrower than needed, so add a little
* fudge character to make it wider.

			if ccCR $ tcString and not ccLF $ tcString
				lcString = strtran(tcString, ccCR, 'i' + ccCRLF)
			else
				lcString = tcString
			endif ccCR $ tcString ...

* Call MeasureString to get the dimensions of the bounding box for the
* specified string.

			.oStringSize = .oGDI.MeasureString(lcString, .oFont, .oSize, ;
				.oFormat, @lnChars, @lnLines)
			.nChars      = lnChars
			.nLines      = lnLines
			.nWidth      = .oStringSize.Width
			.nHeight     = .oStringSize.Height

* If scaling is set to more than 100, that means the application is DPI aware,
* in which case GdipMeasureString is incorrect because it isn't DPI aware. So,
* we use the old fashioned method of calculating the text width.

			if .GetMonitorDPIScale() > 100
				lcFontName = .oFont.FontFamily.Name
				lnFontSize = .oFont.SizeInPoints
				lcStyle    = iif(isnull(luStyle), iif(.oFont.Bold, 'B', '') + ;
					iif(.oFont.Italic, 'I', ''), luStyle)
				lnWidth    = txtwidth(lcString, lcFontName, lnFontSize, lcStyle) * ;
					fontmetric(6, lcFontName, lnFontSize, lcStyle)
				.nWidth    = max(.nWidth, lnWidth)
			endif .GetMonitorDPIScale() > 100
		endwith
	endfunc

* Determine scaling. Returns a percentage relative to 96 DPI (the standard).

	function GetMonitorDPIScale()
		local lnhWnd, ;
			lnDPIX, ;
			lnDPIY, ;
			lnhDC
		lnhWnd = _screen.hWnd
		try
			if This.lGetDPIForWindow
				lnDPIX = GetDpiForWindow(lnhWnd)
			else
				lnhDC  = GetWindowDC(lnhWnd)
				lnDPIX = GetDeviceCaps(lnhDC, DC_LOGPIXELSX)
				ReleaseDC(lnhWnd, lnhDC)
			endif This.lGetDPIForWindow
		catch
			lnDPIX = DPI_STANDARD
		endtry
		return min(max(int(lnDPIX/DPI_STANDARD * DPI_STANDARD_SCALE), DPI_STANDARD_SCALE), DPI_MAX_SCALE)
	endfunc

* Return the width of the specified string.

	function GetWidth(tcString, tcFontName, tnFontSize, tuStyle)
		local lnReturn
		with This
			do case
				case pcount() < 2
					.MeasureString(tcString)
				case pcount() < 4
					.MeasureString(tcString, tcFontName, tnFontSize)
				otherwise
					.MeasureString(tcString, tcFontName, tnFontSize, tuStyle)
			endcase
			lnReturn = .nWidth
		endwith
		return lnReturn
	endfunc

* Return the width of the specified string in FRUs.

	function GetFRUWidth(tcString, tcFontName, tnFontSize, tuStyle)
		local lnReturn
		with This
			.SetSize(100000, 100000)
			do case
				case pcount() < 2
					.MeasureString(tcString)
				case pcount() < 4
					.MeasureString(tcString, tcFontName, tnFontSize)
				otherwise
					.MeasureString(tcString, tcFontName, tnFontSize, tuStyle)
			endcase
			lnReturn = ceiling(.nWidth) * cnFACTOR
		endwith
		return lnReturn
	endfunc

* Get the height of the specified text within a given width.

	function GetHeight(tcString, tnWidth, tcFontName, tnFontSize, tuStyle)
		local lnReturn
		with This
			.SetSize(tnWidth, 100000)
			do case
				case pcount() < 3
					.MeasureString(tcString)
				case pcount() < 5
					.MeasureString(tcString, tcFontName, tnFontSize)
				otherwise
					.MeasureString(tcString, tcFontName, tnFontSize, tuStyle)
			endcase
			lnReturn = ceiling(.nHeight)
		endwith
		return lnReturn
	endfunc

* Set the dimensions of the layout box.

	function SetSize(tnWidth, tnHeight)
		if vartype(tnWidth) = 'N' and tnWidth >= 0 and ;
			vartype(tnHeight) = 'N' and tnHeight >=0
			This.oSize = _screen.System.Drawing.SizeF.New(tnWidth, tnHeight)
		else
			error cnERR_ARGUMENT_INVALID
		endif vartype(tnWidth) = 'N' ...
	endfunc

* Set the font object to the specified font name, size, and style.

	function SetFont(tcFontName, tnFontSize, tuStyle)
		local lnStyle, ;
			lcStyle
		do case
			case pcount() < 2
				error cnERR_ARGUMENT_INVALID
				return
			case pcount() >= 2 and (vartype(tcFontName) <> 'C' or ;
				empty(tcFontName) or vartype(tnFontSize) <> 'N' or ;
				not between(tnFontSize, 1, 128))
				error cnERR_ARGUMENT_INVALID
				return
			case pcount() = 3 and not vartype(tuStyle) $ 'CN'
				error cnERR_ARGUMENT_INVALID
				return
		endcase
		with _screen.System.Drawing
			if vartype(tuStyle) = 'N'
				lnStyle = tuStyle
			else
				lcStyle = iif(vartype(tuStyle) = 'C', tuStyle, '')
				do case
					case lcStyle = 'BI'
						lnStyle = .FontStyle.BoldItalic
					case lcStyle = 'B'
						lnStyle = .FontStyle.Bold 
					case lcStyle = 'I'
						lnStyle = .FontStyle.Italic
					otherwise
						lnStyle = .FontStyle.Regular
				endcase
			endif vartype(tuStyle) = 'N'
			This.oFont = .Font.New(tcFontName, tnFontSize, lnStyle)
		endwith
	endfunc
enddefine
