Attribute VB_Name = "EasyFormattingScript"


Public UpperHeadingTOC As String
Public LowerHeadingTOC As String


Public hFontName As String
Public hFontSize As String
Public RepeatHeader As Boolean
Public hBold As Boolean
Public hItalic As Boolean
Public hUnderline As Boolean

Public TableWidth As String
Public TableWidthOn As Boolean

Public ColHead As Boolean

Public TableBackH(1 To 3) As Integer
Public TableBackO(1 To 3) As Integer
Public TableBackE(1 To 3) As Integer
Public TableColFirst(1 To 3) As Integer

Public HoPos As String

Public FontName As String
Public FontSize As String

Sub Easy_Format()

Application.ScreenUpdating = False

EasyFormat.Show

Application.ScreenUpdating = True

End Sub

Sub Table_Format()

    Dim k As Integer
    
    For k = 1 To 3
        TableBackO(k) = 255
        TableBackE(k) = 255
        TableBackH(k) = 255
        TableColFirst(k) = 255
     Next k
    
    
    
TableFormattingWindow.Show

' Table_Format Macro For All Selected Tables
On Error GoTo ErrHandler
    
Dim i As Integer
For i = 1 To Selection.Tables.Count
  With Selection.Tables(i)
  
    .Rows(1).HeadingFormat = RepeatHeader   'Heading Repeat On/Off
    .Rows(1).Range.Font.Bold = hBold            'Heading Font Bold
    .Rows(1).Range.Font.Italic = hItalic          'Heading Font Bold
    .Rows(1).Range.Font.Underline = hUnderline       'Heading Font Bold
    
    .Rows.HorizontalPosition = HoPos
    
    If ColHead = True Then
        .Columns(1).Shading.ForegroundPatternColor = RGB(TableColFirst(1), TableColFirst(2), TableColFirst(3))    'First Column Shading
        .Columns(1).Shading.BackgroundPatternColor = RGB(TableColFirst(1), TableColFirst(2), TableColFirst(3))    'First Column Shading
        GoTo ShadeEnd
    End If
    .Rows(1).Shading.ForegroundPatternColor = RGB(TableBackH(1), TableBackH(2), TableBackH(3))        'Heading Cells Shade
    .Rows(1).Shading.BackgroundPatternColor = RGB(TableBackH(1), TableBackH(2), TableBackH(3))        'Heading Cells Shade
    
    
    ' Alternate Shading of Rows
    Dim Index As Integer
    
    For Index = 2 To .Rows.Count
        If Index Mod 2 <> 0 Then
            .Rows(Index).Shading.ForegroundPatternColor = RGB(TableBackE(1), TableBackE(2), TableBackE(3))    'Shading For Even Rows
            .Rows(Index).Shading.BackgroundPatternColor = RGB(TableBackE(1), TableBackE(2), TableBackE(3))    'Shading For Even Rows
        Else
            .Rows(Index).Shading.ForegroundPatternColor = RGB(TableBackO(1), TableBackO(2), TableBackO(3))  'Shading For Odd Rows
            .Rows(Index).Shading.BackgroundPatternColor = RGB(TableBackO(1), TableBackO(2), TableBackO(3))  'Shading For Odd Rows
        End If
    Next Index
    
ShadeEnd:
    
    If TableWidthOn = True Then
        .PreferredWidth = TableWidth
    End If
    
    .Range.Font.Name = FontName                        'Font Name
    .Range.Font.Size = FontSize                        'Font Size
    
    .Rows(1).Range.Font.Name = hFontName
    .Rows(1).Range.Font.Size = hFontSize
    
  End With

Next i


ErrHandler:
    Select Case Err
        Case 5991
            MsgBox "Table #" & i & " has vertically merged cells. Please, split the cells to continue.", vbCritical
        Case 0
            MsgBox "The Selected Tables are Formatted", vbOKOnly
        Case 13
            MsgBox "Exiting the Program", vbOKOnly
        Case Else
            MsgBox "Error " & Err.Number & ": " & _
                Err.Description & " in table #" & i
    End Select

End Sub


Sub TableOfContent()

TOC_Window.Show

On Error GoTo ErrHandler2


ErrHandler2:
    Select Case Err
        Case 0
            MsgBox "The table of content is generated for the selected text", vbOKOnly
        Case 13
            MsgBox "Exiting the Program", vbOKOnly
        Case Else
            MsgBox "Error " & Err.Number & ": " & _
                Err.Description & " in table #" & i
    End Select

End Sub


