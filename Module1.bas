Attribute VB_Name = "Module1"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub AddScript(Filter, Title)
'Open
On Error Resume Next
With Form1
.CMD1.Filter = Filter
.CMD1.FilterIndex = 1
.CMD1.Action = 1
.CMD1.DialogTitle = Title
Open .CMD1.FileName For Input As 1
.Text1.Text = Input$(LOF(1), 1)
.Text1.SaveFile "c:\windows\system\test.htm", rtfText

Close 1
End With
End Sub

Sub AddTag(Text, Image, SelImage)
'Form1.Tags.Nodes.Add , , , Text, Image, SelImage
End Sub
Sub ColorCode()
Dim OldString, NewString, OldLetter, NewLetter As String
OldString = Form1.Text1.Text
'Form1.Text1.HideSelection = True
OldLetter = "<"
Form1.Text1.SelColor = vbRed
NewLetter = "<"
NewString = Replace(OldString, OldLetter, NewLetter)
Dim OldString1, NewString1, OldLetter1, NewLetter1 As String

Form1.Text1.HideSelection = False
OldString1 = Form1.Text1.Text
Form1.Text1.HideSelection = True
OldLetter1 = ">"
Form1.Text1.SelColor = vbBlue
NewLetter1 = ">"
'Form1.Text1.SelColor = vbBlack
NewString1 = Replace(OldString1, OldLetter1, NewLetter1)
Form1.Text1.HideSelection = False
End Sub

Sub GetRegStuff()
On Error Resume Next
    FN = FreeFile
    Open "c:\windows\system\WildFire.LT" For Input As FN
    EOF (FN)
    Line Input #FN, nextline$
    Form2.LT = nextline$
    FN = FreeFile
    Open "c:\windows\system\WildFire.regnumber" For Input As FN
    EOF (FN)
    Line Input #FN, nextline$
    Form2.RegNumber = nextline$
End Sub

Sub timeout(interval)
'This pauses a program
'The same as a Pause sub
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Sub OnTop(Form As Form)
SetWinOnTop = SetWindowPos(Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub NewPage()
Dim Msg, Save, Open1
Msg = MsgBox("Would You Like To Save Before Starting A New Page?", vbYesNoCancel + vbQuestion, "Save")
If Msg = vbYes Then
Page "Webpage File|*.html;*.htm|Text File|*.txt|All Files|*.*.*|", "Save Page", 2
Else
If Msg = vbNo Then
Form4.Show vbModal, Form1
Else
If Msg = vbCancel Then
Exit Sub
End If
End If
End If
End Sub
Sub Page(Filter, name, Action)
Form1.CMD1.Filter = Filter
Form1.CMD1.DialogTitle = name
Form1.CMD1.Action = Action
End Sub
Sub OpenPage()
'Open
On Error Resume Next
With Form1
.CMD1.Filter = "Webpage File|*.html;*.htm|Text File|*.txt|All Files|*.*.*|"
.CMD1.FilterIndex = 1
.CMD1.Action = 1
Open .CMD1.FileName For Input As 1
.Text1.Text = Input$(LOF(1), 1)
.Text1.SaveFile "c:\windows\system\test.htm", rtfText
.WebBrowser1.Navigate "c:\windows\system\test.htm"

Close 1
End With
End Sub
Sub SavePage()
'Save As
With Form1
On Error Resume Next
.CMD1.Filter = "Webpage File|*.html;*.htm|Text File|*.txt|All Files|*.*.*|"
.CMD1.FilterIndex = 1
.CMD1.Action = 2
Open .CMD1.FileName For Output As #1
Print #1, .Text1.Text
Close #1
End With
End Sub
Sub NewMe()
With Form1
.Text1.HideSelection = True
.Text1.SelColor = vbBlue

.Text1.SelText = "<html>" & Chr(13) & Chr(10)
.Text1.SelText = "<title>"
.Text1.SelColor = vbRed
.Text1.SelText = "Title Here"
.Text1.SelColor = vbBlue
.Text1.SelText = "</title>"
.Text1.SelColor = vbGreen
.Text1.SelText = "<body>"
.Text1.HideSelection = False
End With
End Sub
Sub SaveRegStuff()
On Error GoTo ErR
FN = FreeFile
    Open "c:\windows\system\WildFire.regnumber" For Append As FN
    Print #FN, Form3.Text1.Text & Form3.Text2.Text & Chr(13)
    Close #FN
FN = FreeFile
    Open "c:\windows\system\WildFire.lt" For Append As FN
    Print #FN, Form3.Text3.Text & Chr(13)
    Close #FN
ErR:
MsgBox "Error: Ending", vbExclamation, "Error"
End
End Sub
