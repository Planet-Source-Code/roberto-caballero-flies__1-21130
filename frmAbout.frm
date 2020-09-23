VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Flies"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Email 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "robcaballero@mail.scbbs-bo.com"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox HTTP 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "http://www.vulliamy.demon.co.uk/alex.html"
      Top             =   3600
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   540
      Left            =   3630
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   540
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   1740
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   540
      Left            =   510
      Picture         =   "frmAbout.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label LabelDueCredits 
      Alignment       =   2  'Center
      Caption         =   "DueCredits"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label LabelContact 
      Alignment       =   2  'Center
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label LabelInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label LabelCopyright 
      Alignment       =   2  'Center
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label LabelFlies 
      Alignment       =   2  'Center
      Caption         =   "Flies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRLF As String
Private Sub Form_Load()
   CRLF = Chr(13) & Chr(10)
   LabelFlies.Caption = "Flies, Version " & App.Major & "." & App.Minor & "." & App.Revision
   LabelCopyright.Caption = "Copyright (C) 2001 Roberto Caballero"
   LabelInfo.Caption = "Programmed using Microsoft's Visual Basic 5.0," & CRLF & _
   "an excelent programming language!"
   LabelContact.Caption = "Send me a message if you like this program" & CRLF & _
   "If asked, I may even send you the source code!" & CRLF & _
   "My email is:"
   LabelDueCredits.Caption = "Based on " + Chr(34) + "Java Flies," + Chr(34) + " an idea by" & CRLF & _
   "Alex Vulliamy and Jeff Cragg" & CRLF & "Visit their web site at"
End Sub
Private Sub Ok_Click()
   Unload Me
End Sub
