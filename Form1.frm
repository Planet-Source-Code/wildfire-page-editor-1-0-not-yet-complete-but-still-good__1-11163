VERSION 5.00
Object = "{6AAF67B6-786B-11D4-8AC4-FC5FAB6C1248}#1.0#0"; "WINSOCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1092
   ClientLeft      =   2652
   ClientTop       =   1728
   ClientWidth     =   2448
   LinkTopic       =   "Form1"
   ScaleHeight     =   1092
   ScaleWidth      =   2448
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   312
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1812
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   1812
   End
   Begin Project1.NovaIceWINSOCK NovaIceWINSOCK1 
      Left            =   360
      Top             =   1080
      _ExtentX        =   1715
      _ExtentY        =   974
   End
   Begin VB.Label Label3 
      Height          =   372
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
NovaIceWINSOCK1.StartWinsock "Label3.Caption" = True
NovaIceWINSOCK1.SendData cool, Text2.Text
End Sub

