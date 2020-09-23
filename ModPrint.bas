Attribute VB_Name = "ModPrint"
'ModPrint
'by -=renyi[ace]=-
'February 26, 2001
'
'If you're using this module, pls credit me by leaving this few lines with the module, :)
'If you need more just email me and I'll add them.
'Email : Kry00@hotmail.com
'ICQ : 7640843

Option Explicit

Public Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Public HorizontalMargin As Single
Public VerticalMargin As Single

Global Const m_Top = 25
Global Const m_Bottom = 25
Global m_Page As Integer
Global m_HeaderText As String
Global m_PageHeader As String
Global sDate As String
Global sPage As String
Global sHeader As String

Global wdApp As Word.Application
Global oPrinter As Object
Global PrintUseWord As Boolean

Private Declare Function RegOpenKey Lib _
"advapi32" Alias "RegOpenKeyA" (ByVal hKey _
As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As _
String, lpReserved As Long, lptype As _
Long, lpData As Any, lpcbData As Long) _
As Long

Private Declare Function RegCloseKey& Lib _
"advapi32" (ByVal hKey&)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Sub DoNewPage()
    'Prints page header
    cPrint.pFooter
    cPrint.pNewPage
    cPrint.pPrint
    cPrint.pCenter MDIMasterKid.cmbChildren.Text
    cPrint.FontName = "Ariel"
    cPrint.FontSize = 10
    cPrint.pPrint
    cPrint.pBox 0.5, 0.5, cPrint.GetPaperWidth - 1, cPrint.GetPaperHeight - 1, , , vbFSTransparent
    cPrint.pPrint
End Sub
Public Function GetRegString(hKey As Long, _
strSubKey As String, strValueName As _
String) As String
Dim strSetting As String
Dim lngDataLen As Long
Dim lngRes As Long
If RegOpenKey(hKey, strSubKey, _
lngRes) = ERROR_SUCCESS Then
   strSetting = Space(255)
   lngDataLen = Len(strSetting)
   If RegQueryValueEx(lngRes, _
   strValueName, ByVal 0, _
   REG_EXPAND_SZ, ByVal strSetting, _
   lngDataLen) = ERROR_SUCCESS Then
      If lngDataLen > 1 Then
      GetRegString = Left(strSetting, lngDataLen - 1)
   End If
End If

If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
   MsgBox "RegCloseKey Failed: " & _
   strSubKey, vbCritical
End If
End If
End Function
Public Function IsAppPresent(strSubKey$, strValueName$) As Boolean
    IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)))
End Function
Public Sub PrintFront()
Dim rsMyRecord As Recordset
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    On Error GoTo errPrintFront
    With frmFrames
        .rsClipArt.Recordset.MoveFirst
        Do While Not .rsClipArt.Recordset.EOF
            If CLng(.rsClipArt.Recordset.Fields("LineNo")) = CLng(rsMyRecord.Fields("SectionPicID")) Then Exit Do
        .rsClipArt.Recordset.MoveNext
        Loop
        
        cPrint.FontTransparent = True
        cPrint.pPrintPicture .Image1.Picture, , , , , , False
        cPrint.FontBold = True
        cPrint.pPrint
        cPrint.pCenter MDIMasterKid.cmbChildren.Text
        cPrint.FontBold = False
        cPrint.FontSize = 16
        cPrint.FontItalic = True
        cPrint.FontBold = True
        cPrint.pCenter sHeader
        cPrint.FontBold = False
        cPrint.FontItalic = False
        cPrint.FontSize = 10
        cPrint.pPrint
        cPrint.pPrint
        cPrint.pPrint
    End With
    Unload frmFrames
    rsMyRecord.Close
    Exit Sub
    
errPrintFront:
    Beep
    MsgBox Err.Description, vbExclamation, "Print SectionPage"
    Err.Clear
End Sub

Public Function PrintPictureToFitPage(Prn As Printer, Pic As Picture) As Boolean
Const vbHiMetric As Integer = 8
Dim PicRatio      As Double
Dim PrnWidth      As Double
Dim PrnHeight     As Double
Dim PrnRatio      As Double
Dim PrnPicWidth   As Double
Dim PrnPicHeight  As Double

    On Error GoTo errorHandler
    
    ' *** Determine if picture should be printed in
    'landscape or portrait and set the orientation
    If Pic.Height >= Pic.Width Then
       Prn.Orientation = vbPRORPortrait ' Taller than wide
    Else
       Prn.Orientation = vbPRORLandscape ' Wider than tall
    End If
    
    ' *** Calculate device independent Width to Height ratio for picture
    PicRatio = Pic.Width / Pic.Height
    
    ' *** Calculate the dimentions of the printable area in HiMetric
    PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
    PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
    
    ' *** Calculate device independent Width to Height ratio for printer
    PrnRatio = PrnWidth / PrnHeight
    
    ' *** Scale the output to the printable area
    If PicRatio >= PrnRatio Then
       ' *** Scale picture to fit full width of printable area
       PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, _
           Prn.ScaleMode)
       PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, _
           vbHiMetric, Prn.ScaleMode)
    Else
       ' *** Scale picture to fit full height of printable area
       PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, _
           Prn.ScaleMode)
       PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, _
           vbHiMetric, Prn.ScaleMode)
    End If
    
    ' *** Print the picture using the PaintPicture method
    Prn.PaintPicture Pic, 0, 0, PrnPicWidth, PrnPicHeight
    Prn.EndDoc
    PrintPictureToFitPage = True
    Exit Function

errorHandler:
    PrintPictureToFitPage = False
End Function
Public Sub WriteHeader(sName As String)
    On Error Resume Next
    Set wdApp = New Word.Application
    wdApp.Application.Visible = True
    wdApp.Application.WindowState = wdWindowStateMaximize
    wdApp.Caption = sName
    With wdApp
        .Documents.Add DocumentType:=wdNewBlankDocument
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        .Selection.TypeText Text:=MDIMasterKid.cmbChildren.Text & "  -  " & Format(CDate(Now), "dd.mm.yyyy")
        .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        .ActiveWindow.Selection.Font.Bold = True
        .ActiveWindow.Selection.Font.Shadow = True
        .ActiveWindow.Selection.Font.Size = 18
        .Selection.TypeText Text:=sName
        .ActiveWindow.Selection.Font.Bold = False
        .ActiveWindow.Selection.Font.Shadow = False
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.Tables.Add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
        .Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=168.45, RulerStyle:=wdAdjustNone
        .Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=276.4, RulerStyle:=wdAdjustNone
    End With
    With wdApp.Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
End Sub


