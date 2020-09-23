VERSION 5.00
Begin VB.Form frminformacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frminformacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblcreditos 
      BackStyle       =   0  'Transparent
      Caption         =   "To report any bug, to talk if you're bored or you like Guns N' Roses, Avril Lavigne mail me at matias_gnr@yahoo.com.ar"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5490
   End
   Begin VB.Label lblversion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0.0"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual Disks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frminformacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblversion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
