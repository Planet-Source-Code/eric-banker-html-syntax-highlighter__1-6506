VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About EasyASP"
   ClientHeight    =   2940
   ClientLeft      =   4260
   ClientTop       =   3645
   ClientWidth     =   5670
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About EasyASP"
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2200
      Left            =   120
      ScaleHeight     =   2205
      ScaleWidth      =   5460
      TabIndex        =   1
      Top             =   120
      Width           =   5460
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         FillColor       =   &H00008000&
         Height          =   1455
         Left            =   80
         ScaleHeight     =   1455
         ScaleWidth      =   5295
         TabIndex        =   4
         Top             =   680
         Width           =   5295
         Begin VB.Label lblDisclaimer 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "This program is Copyright 2000 Eric Banker. All rights reserved"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   10
            Tag             =   "Warning: ..."
            Top             =   120
            Width           =   5055
         End
         Begin VB.Line Line2 
            X1              =   480
            X2              =   4800
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "For product information and help please email: "
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "ebanker@bigfoot.com"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   3480
            TabIndex        =   8
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Web Sites: "
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.optweb.net/ebanker/easyasp"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   960
            TabIndex        =   6
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.codearchive.com/home/eric"
            ForeColor       =   &H80000002&
            Height          =   255
            Left            =   960
            TabIndex        =   5
            Top             =   960
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   80
         ScaleHeight     =   585
         ScaleWidth      =   5265
         TabIndex        =   2
         Top             =   80
         Width           =   5295
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmAbout.frx":0442
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EasyColorCode Version 1.5"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   0
            TabIndex        =   3
            Tag             =   "Application Title"
            Top             =   120
            Width           =   5205
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub
