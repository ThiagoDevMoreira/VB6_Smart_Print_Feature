Option Explicit
'
'# cLayout
' This class provides a library of "components" to compose the print layout.
' Designed to be used by printing strategies (strategy pattern).
'
Private LayoutPrinter As Printer
Private IsA4 As Boolean
Private IsA5 As Boolean
Private IsReceipt As Boolean

Private Const VALUE_FIELD_WIDTH = 15

Private Const DEFAULT_LEFT_MARGIN_RECEIPT As Integer = 1
Private Const DEFAULT_LEFT_MARGIN_A5 As Integer = 1
Private Const DEFAULT_LEFT_MARGIN_A4 As Integer = 4
Private Const EXTENDED_LEFT_MARGIN_A4 As Integer = 20

Private Const DEFAULT_FONT_SIZE_RECEIPT As Integer = 8
Private Const DEFAULT_FONT_SIZE_A5 As Integer = 9
Private Const DEFAULT_FONT_SIZE_A4 As Integer = 11
Private DefaultFontSize As Integer

Private Const DEFAULT_FONT_NAME As String = "Courier New"
Private Const DEFAULT_FONT_BOLD_RECEIPT As Boolean = True
Private Const DEFAULT_FONT_BOLD_A4_A5 As Boolean = False

'## Core methods
Private Sub Class_Initialize()
On Error GoTo EH
   Set LayoutPrinter = Printer
   LayoutPrinter.Orientation = vbPRORPortrait
   LayoutPrinter.ScaleMode = vbCentimeters
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.Class_Initialize" & vbCrLf & Err.Description, , Err.Source
End Sub

Public Sub SetFormat(ByRef pFormat As PrintFormatEnum)
On Error GoTo EH
   IsA4 = False
   IsA5 = False
   IsReceipt = False
   
   Select Case pFormat
      Case formatoA4
         IsA4 = True
         DefaultFontSize = DEFAULT_FONT_SIZE_A4
      Case formatoA5
         IsA5 = True
         ' LayoutPrinter.PaperSize = vbPRPSA5 'Apparently this causes issues on some machines (client Sirlene)
         LayoutPrinter.CurrentY = 0.5
         DefaultFontSize = DEFAULT_FONT_SIZE_A5
      Case formatoCupom
         IsReceipt = True
         DefaultFontSize = DEFAULT_FONT_SIZE_RECEIPT
         ' LayoutPrinter.PaperSize = vbPRPSA4 'Apparently this causes issues on some machines (client Sirlene)
         LayoutPrinter.CurrentY = 0.5
   End Select
   ApplyDefaultStyle
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.SetFormat" & vbCrLf & Err.Description, , Err.Source
End Sub

'## Component library methods
Public Sub PrintLogo()
On Error GoTo EH
   '--- validate
   If ShowLogoStr <> 1 Then Exit Sub
   
   '--- define variables
   Dim Img As PictureBox: Set Img = frmStatus.Picture5
   Dim ImgWidth As Integer: ImgWidth = Img.Width
   Dim ImgHeight As Integer: ImgHeight = Img.Height
   Dim i As Integer

   '--- print
   If IsReceipt Or IsA5 Then
      LayoutPrinter.PaintPicture Img.Picture, (ImgWidth / 2), 0, ImgWidth, ImgHeight
      For i = 1 To 10 ' blank lines
         LayoutPrinter.Print
      Next i
   ElseIf IsA4 Then
      Printer.PaintPicture Img.Picture, 200, 250, Img.Width, Img.Height
   End If
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.PrintLogo" & vbCrLf & Err.Description, , Err.Source
End Sub

Public Sub PrintCompanyName()
On Error GoTo EH
   '--- reset layout
   ApplyDefaultStyle
   LayoutPrinter.FontBold = True
   LayoutPrinter.FontSize = 12
   
   '--- define variables
   Dim mainName As String: mainName = StoreNameStr
   Dim secondaryName As String: secondaryName = StoreName2Str
   
   '--- print
   If IsReceipt Or IsA5 Then
      PrintCentered mainName
      If Len(secondaryName) > 0 Then PrintCentered secondaryName
   ElseIf IsA4 Then
      If ShowLogoStr = 1 Then
         PrintBlankLine
         LayoutPrinter.Print Space$(EXTENDED_LEFT_MARGIN_A4) & mainName
         If Len(secondaryName) > 0 Then
            LayoutPrinter.Print Space$(EXTENDED_LEFT_MARGIN_A4) & secondaryName
         End If
      Else
         PrinterCentered StoreNameStr
         If Len(secondaryName) > 0 Then PrinterCentered StoreName2Str
      End If
   End If
   ApplyDefaultStyle
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.PrintCompanyName" & vbCrLf & Err.Description, , Err.Source
End Sub

Public Sub PrintCompanyAddress()
On Error GoTo EH
   '--- validate inputs
   If Len(StoreAddressStr) = 0 _
   Or Len(StoreNumber) = 0 _
   Or Len(StoreDistrictStr) = 0 _
   Or Len(StoreCityStr) = 0 _
   Or Len(StoreStateStr) = 0 _
   Or Len(StoreZipStr) = 0 _
   Then Exit Sub
   
   '--- define variables
   Dim AddressNumber As String
   Dim DistrictCityState As String

   '--- format text
   AddressNumber = StoreAddressStr & ", " & StoreNumber & " | " & StoreDistrictStr
   DistrictCityState = StoreCityStr & " | " & StoreStateStr & " | " & StoreZipStr

   '--- print
   If IsReceipt Or IsA5 Then
      PrinterCentered AddressNumber
      PrinterCentered DistrictCityState
   ElseIf IsA4 Then
      If ShowLogoStr = 1 Then
         LayoutPrinter.Print Space$(EXTENDED_LEFT_MARGIN_A4) & AddressNumber
         LayoutPrinter.Print Space$(EXTENDED_LEFT_MARGIN_A4) & DistrictCityState
      Else
         PrinterCentered AddressNumber
         PrinterCentered DistrictCityState
      End If
   End If
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.PrintCompanyAddress" & vbCrLf & Err.Description, , Err.Source
End Sub

Public Sub PrintSeparator()
On Error GoTo EH
   '--- reset layout
   ApplyDefaultStyle
   LayoutPrinter.FontSize = 8
   
   '--- define
   Dim LineLength As Integer
   LineLength = GetCharactersPerLine(LayoutPrinter.FontSize)
   
   '--- print
   LayoutPrinter.Print String$(LineLength, "-")

   ApplyDefaultStyle
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.PrintSeparator" & vbCrLf & Err.Description, , Err.Source
End Sub

Public Sub PrintBlankLine()
On Error GoTo EH
   '--- reset layout
   ApplyDefaultStyle
   LayoutPrinter.FontSize = 8
   
   '--- print
   LayoutPrinter.Print 'blank line

   ApplyDefaultStyle
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.PrintBlankLine" & vbCrLf & Err.Description, , Err.Source
End Sub

Public Sub PrintSubtotal(ByRef pSubtotal As Double)
On Error GoTo EH
   '--- reset layout
   ApplyDefaultStyle
   
   '--- set font size
   Dim FontSize As Integer

   If IsReceipt Or IsA5 Then
      FontSize = 9
   ElseIf IsA4 Then
      FontSize = 11
   End If

   '--- print
   LayoutPrinter.FontBold = True
   LayoutPrinter.FontSize = FontSize
   LayoutPrinter.Print WriteRightAlignedValue("SUBTOTAL:", pSubtotal, FontSize)
   
   '--- reset
   ApplyDefaultStyle
   Exit Sub
EH:
   Screen.MousePointer = 0
   MsgBox "cLayout.PrintSubtotal" & vbCrLf & Err.Description, , Err.Source
End Sub

' Writes text with line breaks if the text is too long for the line width
Public Sub WriteWithLineBreaks(ByVal Text As String, Optional ByVal LeftMargin As Integer = -1)
On Error GoTo EH
    Dim Margin As Integer
    Dim MaxChars As Integer
    Dim RemainingText As String
    Dim CurrentLine As String
    
    ' Determine the left margin
    If LeftMargin = -1 Then
        If IsReceipt Then
            Margin = DEFAULT_LEFT_MARGIN_RECEIPT
        ElseIf IsA5 Then
            Margin = DEFAULT_LEFT_MARGIN_A5
        ElseIf IsA4 Then
            Margin = DEFAULT_LEFT_MARGIN_A4
        End If
    Else
        Margin = LeftMargin
    End If
    
    ' Calculate the max number of characters per line based on font size
    MaxChars = GetCharactersPerLine(LayoutPrinter.FontSize) - Margin
    
    RemainingText = Text
    
    Do While Len(RemainingText) > MaxChars
        ' Find last space within MaxChars to break line nicely
        Dim BreakPos As Integer
        BreakPos = InStrRev(Left(RemainingText, MaxChars), " ")
        
        If BreakPos = 0 Then
            BreakPos = MaxChars
        End If
        
        CurrentLine = Left(RemainingText, BreakPos)
        LayoutPrinter.Print Space(Margin) & Trim(CurrentLine)
        RemainingText = Mid(RemainingText, BreakPos + 1)
    Loop
    
    ' Print the last line
    LayoutPrinter.Print Space(Margin) & RemainingText
    Exit Sub
EH:
    Screen.MousePointer = 0
    MsgBox "cLayout.WriteWithLineBreaks" & vbCrLf & Err.Description, , Err.Source
End Sub

' Writes centered text in the current line
Public Sub PrintCentered(ByVal Text As String)
On Error GoTo EH
    Dim Margin As Integer
    Dim LineLength As Integer
    Dim SpaceCount As Integer
    
    LineLength = GetCharactersPerLine(LayoutPrinter.FontSize)
    
    SpaceCount = (LineLength - Len(Text)) \ 2
    If SpaceCount < 0 Then SpaceCount = 0
    
    Margin = 0 ' Reset margin for centered text
    
    LayoutPrinter.Print Space(SpaceCount) & Text
    Exit Sub
EH:
    Screen.MousePointer = 0
    MsgBox "cLayout.PrintCentered" & vbCrLf & Err.Description, , Err.Source
End Sub

' Returns the number of characters per line based on the current font size
Public Function GetCharactersPerLine(ByVal FontSize As Integer) As Integer
    On Error GoTo EH
    Dim Characters As Integer
    
    If IsReceipt Then
        Characters = 40
    ElseIf IsA5 Then
        Characters = 80
    ElseIf IsA4 Then
        Characters = 120
    Else
        Characters = 80 ' default fallback
    End If
    
    ' Adjust characters by font size roughly (smaller font = more chars)
    Characters = Characters * (DEFAULT_FONT_SIZE_A4 / FontSize)
    
    GetCharactersPerLine = Characters
    Exit Function
EH:
    Screen.MousePointer = 0
    MsgBox "cLayout.GetCharactersPerLine" & vbCrLf & Err.Description, , Err.Source
    GetCharactersPerLine = 80
End Function

' Helper to format value as right-aligned string with label and value
Private Function WriteRightAlignedValue(ByVal LabelText As String, ByVal Value As Double, ByVal FontSize As Integer) As String
    Dim TotalWidth As Integer
    Dim FormattedValue As String
    Dim LabelLength As Integer
    Dim SpacesNeeded As Integer
    
    TotalWidth = GetCharactersPerLine(FontSize)
    FormattedValue = Format(Value, "#,##0.00")
    LabelLength = Len(LabelText)
    
    SpacesNeeded = TotalWidth - LabelLength - Len(FormattedValue)
    If SpacesNeeded < 0 Then SpacesNeeded = 0
    
    WriteRightAlignedValue = LabelText & Space(SpacesNeeded) & FormattedValue
End Function

' Applies the default font style to the printer
Private Sub ApplyDefaultStyle()
    With LayoutPrinter
        .FontName = DEFAULT_FONT_NAME
        .FontBold = IIf(IsReceipt, DEFAULT_FONT_BOLD_RECEIPT, DEFAULT_FONT_BOLD_A4_A5)
        .FontSize = DefaultFontSize
    End With
End Sub
