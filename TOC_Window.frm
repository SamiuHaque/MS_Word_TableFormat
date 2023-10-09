VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TOC_Window 
   Caption         =   "Generate Table of Contents"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9165.001
   OleObjectBlob   =   "TOC_Window.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TOC_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

UpperHeadingTOC = TextBox3.Text
LowerHeadingTOC = TextBox5.Text

Set myRange = ActiveDocument.Range(0, 0)
ActiveDocument.TablesOfContents.Add _
 Range:=myRange, _
 UseFields:=False, _
 UseHeadingStyles:=True, _
 LowerHeadingLevel:=LowerHeadingTOC, _
 UpperHeadingLevel:=UpperHeadingTOC, _
 AddedStyles:="myStyle, yourStyle"
 


With ActiveDocument
.TablesOfContents(1).Range.Font.Name = TextBox1.Text
.TablesOfContents(1).Range.Font.Size = TextBox2.Text
.TablesOfContents(1).Range.ParagraphFormat.LineSpacing = TextBox4.Text
End With

If ComboBox1.ListIndex = 0 Then
    ActiveDocument.TablesOfContents(1).TabLeader = wdTabLeaderDashes
ElseIf ComboBox1.ListIndex = 1 Then
    ActiveDocument.TablesOfContents(1).TabLeader = wdTabLeaderDots
ElseIf ComboBox1.ListIndex = 2 Then
    ActiveDocument.TablesOfContents(1).TabLeader = wdTabLeaderHeavy
ElseIf ComboBox1.ListIndex = 3 Then
    ActiveDocument.TablesOfContents(1).TabLeader = wdTabLeaderLines
ElseIf ComboBox1.ListIndex = 4 Then
    ActiveDocument.TablesOfContents(1).TabLeader = wdTabLeaderMiddleDot
ElseIf ComboBox1.ListIndex = 5 Then
    ActiveDocument.TablesOfContents(1).TabLeader = wdTabLeaderSpaces
End If
   

TOC_Window.Hide

End Sub

Private Sub CommandButton2_Click()
Dim AnswerYes As String

AnswerYes = MsgBox("Do you wish to exit?", vbQuestion + vbYesNo, "Exit Window")

If AnswerYes = vbYes Then
    Unload Me
End If
End Sub

Private Sub CommandButton3_Click()

TOC_Window.Hide
Easy_Format

End Sub

Private Sub UserForm_Initialize()

With ComboBox1
        .AddItem "Dashes"
        .AddItem "Dots"
        .AddItem "Heavy"
        .AddItem "Lines"
        .AddItem "Middle Dot"
        .AddItem "Spaces"
    End With

End Sub

