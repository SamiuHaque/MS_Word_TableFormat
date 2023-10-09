VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableFormattingWindow 
   Caption         =   "Table Formatting Tool"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13065
   OleObjectBlob   =   "TableFormattingWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableFormattingWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CheckBox4_Click()
    If CheckBox4.Value = True Then
        ComboBox3.Value = ComboBox2.Value
        ComboBox3.Locked = True
    Else
        ComboBox3.Locked = False
    End If
End Sub

Private Sub CheckBox5_Click()
    If CheckBox5.Value = True Then
        TextBox12.Locked = False
        TextBox11.Locked = False
        TextBox10.Locked = False
    End If
End Sub

Private Sub CheckBox6_Click()
    If CheckBox6.Value = True Then
        TextBox7.Locked = False
        TextBox8.Locked = False
        TextBox9.Locked = False
    End If
End Sub

Private Sub CheckBox7_Click()
    TextBox15.Locked = False
    TextBox14.Locked = False
    TextBox13.Locked = False
End Sub

Private Sub CheckBox8_Click()
    TextBox18.Locked = False
    TextBox17.Locked = False
    TextBox16.Locked = False
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
        TextBox4.Text = TextBox1.Text
        TextBox4.Locked = True
    Else
        TextBox4.Locked = False
    End If
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3.Value = True Then
        TextBox3.Text = TextBox2.Text
        TextBox3.Locked = True
    Else
        TextBox3.Locked = False
    End If
End Sub

Private Sub ComboBox4_Change()
    For i = 1 To Selection.Tables.Count
        With Selection.Tables(i)
            If ComboBox4.ListIndex = 0 Then
                .AutoFitBehavior wdAutoFitWindow
                TextBox5.Locked = True
                Label12.Visible = False
                TextBox5.Visible = False
            ElseIf ComboBox4.ListIndex = 1 Then
                TextBox5.Locked = False
                TableWidthOn = True
                Label12.Visible = True
                TextBox5.Visible = True
                .PreferredWidthType = wdPreferredWidthPercent
            End If
        End With
    Next i
End Sub

Private Sub CommandButton1_Click()
    
    FontName = TextBox1.Text
    FontSize = TextBox2.Text
    
    hFontName = TextBox4.Text
    hFontSize = TextBox3.Text
    RepeatHeader = CheckBox1.Value
    
    TableWidth = TextBox5.Text
    
    HoPos = TextBox6.Text
    
    If CheckBox6.Value = True Then
        TableBackO(1) = TextBox7.Text
        TableBackO(2) = TextBox8.Text
        TableBackO(3) = TextBox9.Text
    End If
    
    If CheckBox5.Value = True Then
        TableBackE(1) = TextBox12.Text
        TableBackE(2) = TextBox11.Text
        TableBackE(3) = TextBox10.Text
    End If
    
    If CheckBox7.Value = True Then
        TableBackH(1) = TextBox15.Text
        TableBackH(2) = TextBox14.Text
        TableBackH(3) = TextBox13.Text
    End If
    
    If CheckBox8.Value = True Then
        ColHead = True
        CheckBox6.Value = False
        CheckBox5.Value = False
        CheckBox7.Value = False
        
        TableColFirst(1) = TextBox18.Text
        TableColFirst(2) = TextBox17.Text
        TableColFirst(3) = TextBox16.Text
    End If
    
    If ComboBox1.ListIndex = 0 Then
        hBold = False
        hItalic = False
        hUnderline = False
        
    ElseIf ComboBox1.ListIndex = 1 Then
        hBold = True
        hItalic = False
        hUnderline = False
        
    ElseIf ComboBox1.ListIndex = 2 Then
        hBold = False
        hItalic = True
        hUnderline = False
        
    ElseIf ComboBox1.ListIndex = 3 Then
        hBold = False
        hItalic = False
        hUnderline = True
        
    ElseIf ComboBox1.ListIndex = 4 Then
        hBold = True
        hItalic = True
        hUnderline = False
        
    ElseIf ComboBox1.ListIndex = 5 Then
        hBold = True
        hItalic = False
        hUnderline = True
        
    ElseIf ComboBox1.ListIndex = 6 Then
        hBold = False
        hItalic = True
        hUnderline = True
        
    ElseIf ComboBox1.ListIndex = 7 Then
        hBold = True
        hItalic = True
        hUnderline = True
        
    End If
    
    
    For i = 1 To Selection.Tables.Count
        With Selection.Tables(i)
            If ComboBox2.ListIndex = 0 Then
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ElseIf ComboBox2.ListIndex = 1 Then
                .Range.ParagraphFormat.Alignment = wdAlignParagraphDistribute
            ElseIf ComboBox2.ListIndex = 2 Then
                .Range.ParagraphFormat.Alignment = wdAlignParagraphJustify
            ElseIf ComboBox2.ListIndex = 3 Then
                .Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            ElseIf ComboBox2.ListIndex = 4 Then
                .Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            End If
       
    
            If ComboBox3.ListIndex = 0 Then
                .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            ElseIf ComboBox3.ListIndex = 1 Then
                .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphDistribute
            ElseIf ComboBox3.ListIndex = 2 Then
                .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphJustify
            ElseIf ComboBox3.ListIndex = 3 Then
                .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            ElseIf ComboBox3.ListIndex = 4 Then
                .Rows(1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            End If
            
        End With
    Next i
        
    Unload Me
End Sub

Private Sub CommandButton2_Click()
Dim AnswerYes As String

AnswerYes = MsgBox("Do you wish to exit?", vbQuestion + vbYesNo, "Exit Window")

If AnswerYes = vbYes Then
    Unload Me
End If

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub CommandButton3_Click()

TableFormattingWindow.Hide
Easy_Format

End Sub

Private Sub UserForm_Initialize()
    
    With ComboBox1
        .AddItem "Regular"
        .AddItem "Bold"
        .AddItem "Italic"
        .AddItem "Underline"
        .AddItem "Bold & Italic"
        .AddItem "Bold & Underline"
        .AddItem "Italic & Underline"
        .AddItem "Bold Italic Underline"
    End With
    
    With ComboBox2
        .AddItem "Center"
        .AddItem "Distribute"
        .AddItem "Justify"
        .AddItem "Left"
        .AddItem "Right"
    End With
    
    With ComboBox3
        .AddItem "Center"
        .AddItem "Distribute"
        .AddItem "Justify"
        .AddItem "Left"
        .AddItem "Right"
    End With
    
    With ComboBox4
        .AddItem "Auto Fit to Window"
        .AddItem "Set Width Percentage"
    End With
    
    TextBox5.Locked = True
    Label12.Visible = False
    TextBox5.Visible = False
    TableWidthOn = False
    
    ColHead = False
    
    TextBox7.Locked = True
    TextBox8.Locked = True
    TextBox9.Locked = True
    
    TextBox12.Locked = True
    TextBox11.Locked = True
    TextBox10.Locked = True
    
    TextBox15.Locked = True
    TextBox14.Locked = True
    TextBox13.Locked = True
    
    TextBox18.Locked = True
    TextBox17.Locked = True
    TextBox16.Locked = True
       
End Sub
