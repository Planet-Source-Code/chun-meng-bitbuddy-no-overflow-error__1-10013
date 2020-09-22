VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About BitBuddy V1.0"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   945
      TabIndex        =   8
      Top             =   3915
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "PaulHews"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   180
      TabIndex        =   15
      ToolTipText     =   "For helps on solving the overflow problem."
      Top             =   2790
      Width           =   3030
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "and others..."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   180
      TabIndex        =   14
      ToolTipText     =   "www.experts-exchange.com"
      Top             =   3330
      Width           =   3030
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Twister of Twisted Media"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   180
      TabIndex        =   13
      ToolTipText     =   "http://www.twistedmedia.f2s.com"
      Top             =   3060
      Width           =   3030
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "> For Their Codes and Helps.  <"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   180
      TabIndex        =   12
      Top             =   3600
      Width           =   3030
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   3240
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "GivenRandy"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   180
      TabIndex        =   11
      ToolTipText     =   "For helps on solving the overflow problem."
      Top             =   2520
      Width           =   3030
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ameba"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   180
      TabIndex        =   10
      ToolTipText     =   "For helps on solving the overflow problem."
      Top             =   2250
      Width           =   3030
   End
   Begin VB.Label lAuthor 
      Alignment       =   2  'Center
      Caption         =   "by Heng Chun Meng"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   1530
      Width           =   3030
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   690
      Left            =   135
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   315
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Released :"
      Height          =   240
      Left            =   1215
      TabIndex        =   7
      Top             =   855
      Width           =   780
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "23/07/2000"
      Height          =   240
      Left            =   2070
      TabIndex        =   6
      Top             =   855
      Width           =   1140
   End
   Begin VB.Label lThanks 
      Alignment       =   2  'Center
      Caption         =   ">>   Special Thanks   <<"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   1935
      Width           =   3030
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "License :"
      Height          =   240
      Left            =   1215
      TabIndex        =   4
      Top             =   1170
      Width           =   780
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Free"
      Height          =   240
      Left            =   2070
      TabIndex        =   3
      Top             =   1170
      Width           =   1140
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.0"
      Height          =   240
      Left            =   2070
      TabIndex        =   2
      Top             =   540
      Width           =   1140
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   3195
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version :"
      Height          =   240
      Left            =   1215
      TabIndex        =   1
      Top             =   540
      Width           =   780
   End
   Begin VB.Label lTitle 
      Alignment       =   2  'Center
      Caption         =   "Bit Buddy "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1215
      TabIndex        =   0
      Top             =   135
      Width           =   1995
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

