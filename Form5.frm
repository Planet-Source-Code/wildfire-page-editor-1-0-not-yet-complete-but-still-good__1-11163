VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form5 
   Caption         =   "Form Wizard"
   ClientHeight    =   3864
   ClientLeft      =   2064
   ClientTop       =   1320
   ClientWidth     =   5640
   LinkTopic       =   "Form5"
   ScaleHeight     =   3864
   ScaleWidth      =   5640
   Begin TabDlg.SSTab SSTab1 
      Height          =   3852
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5652
      _ExtentX        =   9970
      _ExtentY        =   6795
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   8
      TabHeight       =   420
      TabCaption(0)   =   "Submit"
      TabPicture(0)   =   "Form5.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SSCommand1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSCommand2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Reset"
      TabPicture(1)   =   "Form5.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text3"
      Tab(1).Control(1)=   "Text4"
      Tab(1).Control(2)=   "SSCommand4"
      Tab(1).Control(3)=   "SSCommand3"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Button"
      TabPicture(2)   =   "Form5.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text5"
      Tab(2).Control(1)=   "Text6"
      Tab(2).Control(2)=   "SSCommand6"
      Tab(2).Control(3)=   "SSCommand5"
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(5)=   "Label6"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Check Box"
      TabPicture(3)   =   "Form5.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text7"
      Tab(3).Control(1)=   "Text8"
      Tab(3).Control(2)=   "SSCommand8"
      Tab(3).Control(3)=   "SSCommand7"
      Tab(3).Control(4)=   "Label7"
      Tab(3).Control(5)=   "Label8"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Radio Button"
      TabPicture(4)   =   "Form5.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text9"
      Tab(4).Control(1)=   "Text10"
      Tab(4).Control(2)=   "SSCommand10"
      Tab(4).Control(3)=   "SSCommand9"
      Tab(4).Control(4)=   "Label9"
      Tab(4).Control(5)=   "Label10"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Text Box"
      TabPicture(5)   =   "Form5.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text11"
      Tab(5).Control(1)=   "Text12"
      Tab(5).Control(2)=   "SSCommand12"
      Tab(5).Control(3)=   "SSCommand11"
      Tab(5).Control(4)=   "Label11"
      Tab(5).Control(5)=   "Label12"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Text Area"
      TabPicture(6)   =   "Form5.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text13"
      Tab(6).Control(1)=   "Text14"
      Tab(6).Control(2)=   "Text15"
      Tab(6).Control(3)=   "Text16"
      Tab(6).Control(4)=   "SSCommand14"
      Tab(6).Control(5)=   "SSCommand13"
      Tab(6).Control(6)=   "Label13"
      Tab(6).Control(7)=   "Label14"
      Tab(6).Control(8)=   "Label15"
      Tab(6).Control(9)=   "Label16"
      Tab(6).ControlCount=   10
      TabCaption(7)   =   "Password"
      TabPicture(7)   =   "Form5.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Text17"
      Tab(7).Control(1)=   "SSCommand16"
      Tab(7).Control(2)=   "SSCommand15"
      Tab(7).Control(3)=   "Label17"
      Tab(7).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   2040
         TabIndex        =   33
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   2040
         TabIndex        =   32
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   288
         Left            =   -72960
         TabIndex        =   29
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text4 
         Height          =   288
         Left            =   -72960
         TabIndex        =   28
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text5 
         Height          =   288
         Left            =   -72960
         TabIndex        =   25
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text6 
         Height          =   288
         Left            =   -72960
         TabIndex        =   24
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text7 
         Height          =   288
         Left            =   -72960
         TabIndex        =   21
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text8 
         Height          =   288
         Left            =   -72960
         TabIndex        =   20
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text9 
         Height          =   288
         Left            =   -72960
         TabIndex        =   17
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text10 
         Height          =   288
         Left            =   -72960
         TabIndex        =   16
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text11 
         Height          =   288
         Left            =   -72960
         TabIndex        =   13
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text12 
         Height          =   288
         Left            =   -72960
         TabIndex        =   12
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text13 
         Height          =   288
         Left            =   -73080
         TabIndex        =   9
         Top             =   600
         Width           =   1932
      End
      Begin VB.TextBox Text14 
         Height          =   288
         Left            =   -73080
         TabIndex        =   8
         Top             =   960
         Width           =   1932
      End
      Begin VB.TextBox Text15 
         Height          =   372
         Left            =   -73080
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1320
         Width           =   492
      End
      Begin VB.TextBox Text16 
         Height          =   372
         Left            =   -71640
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1320
         Width           =   492
      End
      Begin VB.TextBox Text17 
         Height          =   288
         Left            =   -73080
         TabIndex        =   3
         Top             =   600
         Width           =   1932
      End
      Begin Threed.SSCommand SSCommand16 
         Height          =   372
         Left            =   -73080
         TabIndex        =   1
         Top             =   960
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand15 
         Height          =   372
         Left            =   -72000
         TabIndex        =   2
         Top             =   960
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand14 
         Height          =   372
         Left            =   -73080
         TabIndex        =   4
         Top             =   1800
         Width           =   732
         _Version        =   65536
         _ExtentX        =   1291
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand13 
         Height          =   372
         Left            =   -71880
         TabIndex        =   5
         Top             =   1800
         Width           =   732
         _Version        =   65536
         _ExtentX        =   1291
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand12 
         Height          =   372
         Left            =   -72960
         TabIndex        =   10
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand11 
         Height          =   372
         Left            =   -71880
         TabIndex        =   11
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand10 
         Height          =   372
         Left            =   -72960
         TabIndex        =   14
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand9 
         Height          =   372
         Left            =   -71880
         TabIndex        =   15
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand8 
         Height          =   372
         Left            =   -72960
         TabIndex        =   18
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   372
         Left            =   -71880
         TabIndex        =   19
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   372
         Left            =   -72960
         TabIndex        =   22
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   372
         Left            =   -72000
         TabIndex        =   23
         Top             =   1320
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1714
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   372
         Left            =   -72960
         TabIndex        =   26
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   372
         Left            =   -72000
         TabIndex        =   27
         Top             =   1320
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1714
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   372
         Left            =   2040
         TabIndex        =   30
         Top             =   1320
         Width           =   852
         _Version        =   65536
         _ExtentX        =   1503
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "Can&cle"
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   372
         Left            =   3000
         TabIndex        =   31
         Top             =   1320
         Width           =   972
         _Version        =   65536
         _ExtentX        =   1714
         _ExtentY        =   656
         _StockProps     =   78
         Caption         =   "O&k"
      End
      Begin VB.Label Label1 
         Caption         =   "Caption:"
         Height          =   252
         Left            =   1200
         TabIndex        =   50
         Top             =   600
         Width           =   612
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   252
         Left            =   1200
         TabIndex        =   49
         Top             =   960
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Caption:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   48
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   372
         Left            =   -73800
         TabIndex        =   47
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label5 
         Caption         =   "Caption:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   46
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label6 
         Caption         =   "Name:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   45
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label7 
         Caption         =   "Caption:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   44
         Top             =   600
         Width           =   732
      End
      Begin VB.Label Label8 
         Caption         =   "Name:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   43
         Top             =   960
         Width           =   732
      End
      Begin VB.Label Label9 
         Caption         =   "Caption:"
         Height          =   372
         Left            =   -73800
         TabIndex        =   42
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label10 
         Caption         =   "Name:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   41
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label11 
         Caption         =   "Caption:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   40
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label12 
         Caption         =   "Name:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   39
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label13 
         Caption         =   "Caption:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   38
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label14 
         Caption         =   "Name:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   37
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label15 
         Caption         =   "Rows:"
         Height          =   252
         Left            =   -73800
         TabIndex        =   36
         Top             =   1320
         Width           =   732
      End
      Begin VB.Label Label16 
         Caption         =   "Cols:"
         Height          =   372
         Left            =   -72360
         TabIndex        =   35
         Top             =   1320
         Width           =   492
      End
      Begin VB.Label Label17 
         Caption         =   "Name:"
         Height          =   372
         Left            =   -73800
         TabIndex        =   34
         Top             =   600
         Width           =   852
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NowColor As Double
Dim RGBValues(4) As Long
Function RGBPicker()
RGBValues(3) = CLng(NowColor)
RGBValues(0) = RGBValues(3) And 255
RGBValues(1) = (RGBValues(3) And 65280) \ 256&
RGBValues(2) = (RGBValues(3) And 16711680) \ 65535

RGBValues(0) = 255
RGBValues(1) = 255
RGBValues(2) = 255

Picture1.DrawWidth = 2
P = 0
For i = 1 To 254
P = P + 13
Picture1.ForeColor = RGB(RGBValues(0), RGBValues(1), i)
Picture1.Line (0, P)-(245, P)
Next i
End Function


Private Sub Picture3_Click()
Text18.Text = NowColor
RGBPicker
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3_Click
Picture1.BackColor = Picture3.Point(X, Y)
End Sub

Private Sub SSCommand1_Click()
Form1.Text1.SelText = "<input type=submit name=" + Text2.Text + " value=" + Text1.Text + ">"
Unload Me
End Sub

Private Sub SSCommand10_Click()
Unload Me
End Sub

Private Sub SSCommand11_Click()
Form1.Text1.SelText = "<input type=text name=" + Text12.Text + " value=" + Text11.Text + ">"
Unload Me
End Sub

Private Sub SSCommand12_Click()
Unload Me
End Sub

Private Sub SSCommand13_Click()
Form1.Text1.SelText = "<text area cols=" + Text16.Text + " rows=" + Text15.Text + "name=" + Text14.Text + ">" + Text13.Text + "</textarea>"
Unload Me
End Sub

Private Sub SSCommand14_Click()
Unload Me
End Sub

Private Sub SSCommand15_Click()
Form1.Text1.SelText = "<input type=password name=" + Text17.Text + ">"
Unload Me
End Sub

Private Sub SSCommand16_Click()
Unload Me
End Sub

Private Sub SSCommand18_Click()
Form1.Text1.SelText = "<font color=" + Text18.Text + ">"
Unload Me
End Sub

Private Sub SSCommand19_Click()
Unload Me
End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub

Private Sub SSCommand3_Click()
Form1.Text1.SelText = "<input type=reset name=" + Text2.Text + " value=" + Text1.Text + ">"
Unload Me
End Sub

Private Sub SSCommand4_Click()
Unload Me
End Sub

Private Sub SSCommand5_Click()
Form1.Text1.SelText = "<input type=button name=" + Text6.Text + " value=" + Text5.Text + ">"
Unload Me
End Sub

Private Sub SSCommand6_Click()
Unload Me
End Sub

Private Sub SSCommand7_Click()
Form1.Text1.SelText = "<input type=checkbox name=" + Text8.Text + ">" + Text7.Text + ""
Unload Me
End Sub

Private Sub SSCommand8_Click()
Unload Me
End Sub

Private Sub SSCommand9_Click()
Form1.Text1.SelText = "<input type=radio name=" + Text10.Text + ">" + Text9.Text + ""
Unload Me
End Sub

