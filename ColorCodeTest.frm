VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "WildFire Page Editor"
   ClientHeight    =   4020
   ClientLeft      =   1896
   ClientTop       =   1452
   ClientWidth     =   7512
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7512
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   2640
      Top             =   1080
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList TagIcons 
      Left            =   5160
      Top             =   120
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":0000
            Key             =   "Not-Seld"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":0114
            Key             =   "Seld"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7512
      _ExtentX        =   13250
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "Icons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New Webpage"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open Webpage"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save Webpage"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut Selected Text"
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy Selected Text"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste Text That Is In The Clipboard"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "left"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "center"
            Object.ToolTipText     =   "Align Center"
            ImageKey        =   "center"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "right"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "font"
            Object.ToolTipText     =   "Font Control"
            ImageKey        =   "font"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Icons 
      Left            =   3240
      Top             =   1680
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":0228
            Key             =   "copy"
            Object.Tag             =   "&copy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":076C
            Key             =   "cut"
            Object.Tag             =   "&cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":0CB0
            Key             =   "help"
            Object.Tag             =   "&help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":11F4
            Key             =   "new"
            Object.Tag             =   "&new"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":1738
            Key             =   "open"
            Object.Tag             =   "&open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":1C7C
            Key             =   "paste"
            Object.Tag             =   "&paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":21C0
            Key             =   "preview"
            Object.Tag             =   "&preview"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":2704
            Key             =   "print"
            Object.Tag             =   "&print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":2C48
            Key             =   "redo"
            Object.Tag             =   "&redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":318C
            Key             =   "save"
            Object.Tag             =   "&save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":36D0
            Key             =   "undo"
            Object.Tag             =   "&undo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":3C14
            Key             =   "font"
            Object.Tag             =   "&font"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":4A68
            Key             =   "wizard"
            Object.Tag             =   "&wizard"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":5344
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":5458
            Key             =   "left"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":556C
            Key             =   "right"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ColorCodeTest.frx":5680
            Key             =   "center"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   3768
      Width           =   7512
      _ExtentX        =   13250
      _ExtentY        =   445
      SimpleText      =   "WildFire Page Editor"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   5715
            MinWidth        =   5715
            Text            =   "WildFire Page Editor"
            TextSave        =   "WildFire Page Editor"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   5785
            MinWidth        =   2187
            TextSave        =   "10:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            Object.Width           =   1305
            MinWidth        =   1305
            TextSave        =   "8/28/00"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5772
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   9456
      _ExtentX        =   16679
      _ExtentY        =   10181
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "Source"
      TabPicture(0)   =   "ColorCodeTest.frx":5794
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "ColorCodeTest.frx":57B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WebBrowser1"
      Tab(1).ControlCount=   1
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   5292
         Left            =   -74880
         TabIndex        =   4
         ToolTipText     =   "This Is Your Webpage"
         Top             =   360
         Width           =   9192
         ExtentX         =   16214
         ExtentY         =   9334
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
      End
      Begin RichTextLib.RichTextBox Text1 
         Height          =   5292
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Type Your Webpage Source In Here Or Use Some Of Our Editor's Of Wizard's"
         Top             =   360
         Width           =   9228
         _ExtentX        =   16277
         _ExtentY        =   9335
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"ColorCodeTest.frx":57CC
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                  Tags"
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   1692
   End
   Begin VB.Image ImgUndo 
      Height          =   180
      Left            =   1800
      Picture         =   "ColorCodeTest.frx":5895
      Top             =   120
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgRedo 
      Height          =   180
      Left            =   1200
      Picture         =   "ColorCodeTest.frx":5DC7
      Top             =   120
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgHelp 
      Height          =   180
      Left            =   3840
      Picture         =   "ColorCodeTest.frx":62F9
      Top             =   240
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgPaste 
      Height          =   180
      Left            =   3480
      Picture         =   "ColorCodeTest.frx":682B
      Top             =   240
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgCut 
      Height          =   180
      Left            =   3000
      Picture         =   "ColorCodeTest.frx":6D5D
      Top             =   240
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgCopy 
      Height          =   180
      Left            =   2400
      Picture         =   "ColorCodeTest.frx":728F
      Top             =   240
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgPrint 
      Height          =   180
      Left            =   840
      Picture         =   "ColorCodeTest.frx":77C1
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgSave 
      Height          =   180
      Left            =   1440
      Picture         =   "ColorCodeTest.frx":7CF3
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgOPEN 
      Height          =   180
      Left            =   1920
      Picture         =   "ColorCodeTest.frx":8225
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImgNEW 
      Height          =   180
      Left            =   2040
      Picture         =   "ColorCodeTest.frx":8757
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   "&Printer Setup"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu bye 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Un&do"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "R&edo"
         Enabled         =   0   'False
         Shortcut        =   ^Q
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Cop&y"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "To&ols"
      Begin VB.Menu mnuToolsForm 
         Caption         =   "Form Wizard"
      End
      Begin VB.Menu mnuToolsFrame 
         Caption         =   "Frame Wizard"
      End
      Begin VB.Menu mnuToolsChatRoom 
         Caption         =   "Chat Room Wizard"
      End
      Begin VB.Menu mnuToolsTable 
         Caption         =   "Table Wizard"
      End
      Begin VB.Menu mnuWizardCSS 
         Caption         =   "CSS Wizard"
      End
      Begin VB.Menu jkgfdjfdjlkjfdkljfjfdklfjdkl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsImageMap 
         Caption         =   "Image Map Wizard"
      End
      Begin VB.Menu mnutoolsImageWizard 
         Caption         =   "Image Wizard"
      End
      Begin VB.Menu toolsdahs1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnufasttag 
      Caption         =   "Fast Tags"
      Begin VB.Menu mnuFastTagsComment 
         Caption         =   "Comment"
      End
      Begin VB.Menu mnuFastTagsBreak 
         Caption         =   "Break"
      End
      Begin VB.Menu mnuFastTagsParagraph 
         Caption         =   "Paragraph"
      End
      Begin VB.Menu mnuFastTagsHr 
         Caption         =   "Horazontal Break"
      End
      Begin VB.Menu dash111111 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStrartStop 
         Caption         =   "<></>"
      End
      Begin VB.Menu mnuFastTagsHtmlTag 
         Caption         =   "<html>"
      End
      Begin VB.Menu mnuFastTagsHTMLEndTag 
         Caption         =   "</html>"
      End
      Begin VB.Menu mnufasttagstitle 
         Caption         =   "<title></title>"
      End
      Begin VB.Menu mnufasttagshead1 
         Caption         =   "<head>"
      End
      Begin VB.Menu mnufasttagshead2 
         Caption         =   "</head>"
      End
      Begin VB.Menu mnufasttagstable 
         Caption         =   "<table></table>"
      End
      Begin VB.Menu mnufasttagsNewRow 
         Caption         =   "<tr></tr>"
      End
      Begin VB.Menu mnufasttagsnewcol 
         Caption         =   "<td></td>"
      End
      Begin VB.Menu mnubold 
         Caption         =   "<b></b>"
      End
      Begin VB.Menu mnuItalics 
         Caption         =   "<i></i>"
      End
      Begin VB.Menu mnuunderline 
         Caption         =   "<u></u>"
      End
      Begin VB.Menu mnufonty 
         Caption         =   "<font></font>"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'These are the variables for Undo and Redo
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
' API stuff for putting bitmaps in menus.
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10
Dim OldString
Dim OldLetter
Dim NewLetter
Dim NewString
Dim OldString1
Dim OldLetter1
Dim NewLetter1
Dim NewString1

' Put a bitmap in a menu item.
Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

    ' Get the menu handle.
    menu_handle = GetMenu(frm.hWnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

    ' Initialize the menu information.
    With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With

    ' Assign the picture.
    SetMenuItemInfo menu_handle, _
        item_numbers(UBound(item_numbers)), _
        True, menu_info
End Sub
Public Function Replace(OldString, NewLetter, OldLetter) As String
    Dim i As Integer
    i = 1


    Do While InStr(i, OldString, OldLetter, vbTextCompare) <> 0
        Replace = Replace & Mid(OldString, i, InStr(i, OldString, OldLetter, vbTextCompare) - i) & NewLetter
        i = InStr(i, OldString, OldLetter, vbTextCompare) + Len(OldLetter)
    Loop
    Replace = Replace & Right(OldString, Len(OldString) - i + 1)
End Function

Private Sub Command1_Click()

End Sub

Private Sub bye_Click()
Dim MeMe, YouYou
MeMe = MsgBox("Are You Sure You Wish To Quit?", vbYesNoCancel + vbQuestion, "Quit?")
If MeMe = vbYes Then YouYou = MsgBox("Save Current Webpage?", vbYesNo + vbQuestion, "Save?")
If YouYou = vbYes Then SavePage
If YouYou = vbNo Then End
If MeMe = vbNo Then Exit Sub
If MeMe = vbCancel Then Exit Sub
End Sub

Private Sub Form_Initialize()
On Error Resume Next
 Form2.Show
End Sub

Private Sub Form_Load()
'NewMe
AddTag "Me", "Not-Seld", "Seld"
AddTag "Me", "Not-Seld", "Seld"
SetMenuBitmap Me, Array(0, 0), ImgNEW.Picture
SetMenuBitmap Me, Array(0, 1), ImgOPEN.Picture
SetMenuBitmap Me, Array(0, 3), ImgSave.Picture
SetMenuBitmap Me, Array(0, 6), ImgPrint.Picture
SetMenuBitmap Me, Array(1, 0), ImgUndo.Picture
SetMenuBitmap Me, Array(1, 1), ImgRedo.Picture
SetMenuBitmap Me, Array(1, 3), ImgCut.Picture
SetMenuBitmap Me, Array(1, 4), ImgCopy.Picture
SetMenuBitmap Me, Array(1, 5), ImgPaste.Picture
    Text1_Change
End Sub

Private Sub mnuASPAdd_Click()
AddScript "ASP Script|*.asp|", "Add ASP Script"
End Sub

Private Sub mnuCGIAdd_Click()
AddScript "CGI Script|*.cgi|", "Add CGI"
End Sub


Private Sub Form_Resize()
On Error Resume Next
SSTab1.Width = Form1.ScaleWidth - 150
Text1.Width = Form1.ScaleWidth - 400
Text1.Height = Form1.ScaleHeight - 1190
WebBrowser1.Height = Form1.ScaleHeight - 1190
WebBrowser1.Width = Form1.ScaleWidth - 400
SSTab1.Height = Form1.ScaleHeight - 700
End Sub

Private Sub mnuCopy_Click()
mnuPaste.Enabled = True
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from Text1 onto the Clipboard
    Clipboard.SetText Text1.SelText
    'Sets the Focus to Text1
    Text1.SetFocus
End Sub

Private Sub mnuCut_Click()
mnuCopy_Click
Text1.SelText = ""
End Sub

Private Sub mnuEditRedo_Click()
'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub mnuEditUndo_Click()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub mnuJAAdd_Click()
AddScript "Java Applet|*.java|", "Add Java Applet"
End Sub

Private Sub mnuJSAdd_Click()
AddScript "JavaScript Files|*.js;*.jscript|", "Add JavaScript"
End Sub

Private Sub mnunew_Click()
NewPage
End Sub

Private Sub mnuOpen_Click()
OpenPage
End Sub

Private Sub mnuPaste_Click()
    'Puts the Text from the clipboard into Text1
    Text1.SelText = Clipboard.GetText
    'Sets the Focus to Text1
    Text1.SetFocus
End Sub

Private Sub mnuPrint_Click()
Form1.Print Text1.Text
Printer.EndDoc
End Sub

Private Sub mnuPrinterSetup_Click()
On Error Resume Next
CMD1.DialogTitle = "Printer Setup"
CMD1.CancelError = True
CMD1.ShowPrinter
End Sub

Private Sub mnuSave_Click()
SavePage
End Sub

Private Sub mnuSaveAs_Click()
SavePage
End Sub

Private Sub mnuSelectAll_Click()
    'Sets the cursors position to zero
    Text1.SelStart = 0
    'Selects the full length of Text1
    Text1.SelLength = Len(Text1.Text)
    'Sets the Focus to Text1
    Text1.SetFocus
End Sub

Private Sub mnuToolsOptions_Click()

End Sub

Private Sub mnuToolsForm_Click()
Form5.Show vbModal, Form1
End Sub

Private Sub mnuToolsFrame_Click()
Form6.Show vbModal, Form1
End Sub

Private Sub Text1_Change()
' ColorCode
 'Basically this updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = Form1.Text1.TextRTF
    End If
    If Text1.Text = "" Then
    mnunew.Enabled = False
    Form1.mnuSave.Enabled = False
    mnuSaveAs.Enabled = False
      Form1.mnuPrint.Enabled = False
      mnuSelectAll.Enabled = False
    Else
    mnunew.Enabled = True
    Form1.mnuSave.Enabled = True
    mnuSaveAs.Enabled = True
    Form1.mnuPrint.Enabled = True
      mnuSelectAll.Enabled = True
End If
mnuEditRedo.Enabled = True
mnuEditUndo.Enabled = True


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SaveFile "c:\windows\system\test.htm", rtfText
WebBrowser1.Navigate "c:\windows\system\test.htm"









End If
End Sub

Private Sub Text1_SelChange()
Text1.ToolTipText = Text1.SelText
End Sub

Private Sub Timer1_Timer()
If Clipboard.GetText = "" Then
mnuPaste.Enabled = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
Select Case Button.Key
Dim Stuff
Case "new"
NewPage
Case "open"
OpenPage
Case "save"
SavePage
Case "font"
CMD1.Flags = 10
CMD1.ShowFont
If CMD1.FontBold = True Then
Text1.SelText = "<b>"
If CMD1.FontItalic = True Then
Text1.SelText = "<i>"
If CMD1.FontUnderline = True Then
Text1.SelText = "<u>"
If CMD1.FontBold & CMD1.FontItalic = True Then
Text1.SelText = "<b><i>"
End If
End If
End If
End If
Text1.SelText = "<font face=" & CMD1.FontName & ">"
Text1.SelText = "<font size=" & CMD1.FontSize & ">"
CMD1.ShowColor
Text1.SelText = "<font color=#" + CMD1.Color + ">"
Case "cut"
mnuPaste.Enabled = True
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from Text1 onto the Clipboard
    Clipboard.SetText Text1.SelText
    'Deletes the Selected Text on Text1
    Text1.SelText = ""
    'Sets the Focus to Text1
    Text1.SetFocus
Case "copy"
mnuPaste.Enabled = True
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from Text1 onto the Clipboard
    Clipboard.SetText Text1.SelText
    'Sets the Focus to Text1
    Text1.SetFocus
Case "paste"
    'Puts the Text from the clipboard into Text1
    Text1.SelText = Clipboard.GetText
    'Sets the Focus to Text1
    Text1.SetFocus
Case "left"
Text1.SelText = "<p align=left>"
Case "center"
Text1.SelText = "<center>"
Case "right"
Text1.SelText = "<p align=right>"
End Select
End Sub
