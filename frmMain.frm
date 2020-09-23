VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   1695
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<Move>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblLVL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         Height          =   135
         Left            =   0
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         Height          =   6975
         Left            =   1560
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         Height          =   6975
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "by Alex Donavon"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Death Star v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox ShipLeft 
      Height          =   375
      Left            =   8040
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ShipDown 
      Height          =   375
      Left            =   7680
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ShipRight 
      Height          =   375
      Left            =   7320
      Picture         =   "frmMain.frx":0614
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ShipUp 
      Height          =   375
      Left            =   6960
      Picture         =   "frmMain.frx":091E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image ball 
      Height          =   375
      Left            =   8040
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   25
      Left            =   7320
      Picture         =   "frmMain.frx":0C28
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   24
      Left            =   7320
      Picture         =   "frmMain.frx":0F32
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   23
      Left            =   7320
      Picture         =   "frmMain.frx":123C
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   22
      Left            =   7680
      Picture         =   "frmMain.frx":1546
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   21
      Left            =   7680
      Picture         =   "frmMain.frx":1850
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   20
      Left            =   7680
      Picture         =   "frmMain.frx":1B5A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   19
      Left            =   7680
      Picture         =   "frmMain.frx":1E64
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   18
      Left            =   7680
      Picture         =   "frmMain.frx":216E
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   17
      Left            =   7680
      Picture         =   "frmMain.frx":2478
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   16
      Left            =   7680
      Picture         =   "frmMain.frx":2782
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   15
      Left            =   7320
      Picture         =   "frmMain.frx":2A8C
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   14
      Left            =   7680
      Picture         =   "frmMain.frx":2D96
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   13
      Left            =   7320
      Picture         =   "frmMain.frx":30A0
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   12
      Left            =   7320
      Picture         =   "frmMain.frx":33AA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   11
      Left            =   7320
      Picture         =   "frmMain.frx":36B4
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   10
      Left            =   7680
      Picture         =   "frmMain.frx":39BE
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   9
      Left            =   7320
      Picture         =   "frmMain.frx":3CC8
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   8
      Left            =   7320
      Picture         =   "frmMain.frx":3FD2
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   7
      Left            =   7320
      Picture         =   "frmMain.frx":42DC
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   6
      Left            =   7680
      Picture         =   "frmMain.frx":45E6
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   5
      Left            =   7680
      Picture         =   "frmMain.frx":48F0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   4
      Left            =   7320
      Picture         =   "frmMain.frx":4BFA
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   3
      Left            =   7320
      Picture         =   "frmMain.frx":4F04
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   2
      Left            =   7320
      Picture         =   "frmMain.frx":520E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   1
      Left            =   7680
      Picture         =   "frmMain.frx":5518
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Shape1 
      Height          =   375
      Index           =   0
      Left            =   7680
      Picture         =   "frmMain.frx":5822
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "getdata"
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   5040
      Width           =   975
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   9
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   8
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   7
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   6
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   5
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   4
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   3
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   2
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape Food 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   135
      Index           =   0
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   600
      Width           =   135
   End
   Begin VB.Line Line2 
      X1              =   -360
      X2              =   6840
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6840
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   6840
      Y1              =   0
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line14 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line15 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line16 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line17 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line23 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line24 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line25 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line26 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line27 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6960
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line18 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line19 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line20 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line21 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line22 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line28 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line29 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line30 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line31 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line32 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line33 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line34 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line35 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line36 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line37 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line38 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line39 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line40 
      Visible         =   0   'False
      X1              =   0
      X2              =   120
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line41 
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line42 
      Visible         =   0   'False
      X1              =   1800
      X2              =   1800
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line43 
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line44 
      Visible         =   0   'False
      X1              =   1080
      X2              =   1080
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line45 
      Visible         =   0   'False
      X1              =   720
      X2              =   720
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line46 
      Visible         =   0   'False
      X1              =   2160
      X2              =   2160
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line47 
      Visible         =   0   'False
      X1              =   3600
      X2              =   3600
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line48 
      Visible         =   0   'False
      X1              =   3240
      X2              =   3240
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line49 
      Visible         =   0   'False
      X1              =   2880
      X2              =   2880
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line50 
      Visible         =   0   'False
      X1              =   2520
      X2              =   2520
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line51 
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line52 
      Visible         =   0   'False
      X1              =   5400
      X2              =   5400
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line53 
      Visible         =   0   'False
      X1              =   5040
      X2              =   5040
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line54 
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line55 
      Visible         =   0   'False
      X1              =   4320
      X2              =   4320
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line56 
      Visible         =   0   'False
      X1              =   5760
      X2              =   5760
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line58 
      Visible         =   0   'False
      X1              =   6840
      X2              =   6840
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line59 
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line60 
      Visible         =   0   'False
      X1              =   6120
      X2              =   6120
      Y1              =   6840
      Y2              =   6960
   End
   Begin VB.Line Line57 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line61 
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line62 
      Visible         =   0   'False
      X1              =   1080
      X2              =   1080
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line63 
      Visible         =   0   'False
      X1              =   720
      X2              =   720
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line64 
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line65 
      Visible         =   0   'False
      X1              =   1800
      X2              =   1800
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line66 
      Visible         =   0   'False
      X1              =   3240
      X2              =   3240
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line67 
      Visible         =   0   'False
      X1              =   2880
      X2              =   2880
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line68 
      Visible         =   0   'False
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line69 
      Visible         =   0   'False
      X1              =   2160
      X2              =   2160
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line70 
      Visible         =   0   'False
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line71 
      Visible         =   0   'False
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line72 
      Visible         =   0   'False
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line73 
      Visible         =   0   'False
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line74 
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line75 
      Visible         =   0   'False
      X1              =   5400
      X2              =   5400
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line76 
      Visible         =   0   'False
      X1              =   6480
      X2              =   6480
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line77 
      Visible         =   0   'False
      X1              =   6120
      X2              =   6120
      Y1              =   0
      Y2              =   120
   End
   Begin VB.Line Line78 
      Visible         =   0   'False
      X1              =   5760
      X2              =   5760
      Y1              =   0
      Y2              =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DSTAR v1.0 by Alex Donavon (based off of DeathStar for
'Ti calcs by Andrew Von Dollen.)
'
'This project took me a few days to make.  The game engine
'was not too difficult but there were several bugs that I had
'to iron out before deployment.  This is basically a demo
'because I didn't clone Andrew Von Dollen's game (which had more
'levels and also you could be either a box or a diamond) If I were
'to add this code it wouldn't be too hard so I may.
'
'IF YOU LIKE THIS GAME AND WOULD LIKE TO SEE A POSSIBLE
'UPDATE INCLUDING A LEVEL EDITOR AND MORE LEVELS, ETC. THEN PLEASE
'EITHER VOTE GENEROUSLY OR MAIL ME! I'd greatly appreciate
'good votes or comments in my mailbox:
'(Aedseed@msn.com or Aedseed@aol.com)
'
'I really hope everyone likes this.  It's not too difficult as it
'is (or maybe it's just easy for me 'cause I made all the levels!)
'And also my brother had trouble beating it --so I guess it can be
'mind boggling!  (O:
'If i were to add the additional movable box (like in the
'original from the ti calcs) then It would be posible to create
'more challenging levels.
'
'there is only one (1) bug that i can think of.  Sometimes when
'you move and get a point while zooming thru the air it stops
'you and you have to re-press the button.  This is not at all
'serious but i dont know how to fix it.  If u find any other
'bugs PLZ mail me!
'
'THANKS FOR D/ling this!  Have Fun and Good Luck!
'
Dim release As Boolean
Dim Hit As Boolean
Dim NoRight As Boolean
Dim NoLeft As Boolean
Dim NoDown As Boolean
Dim NoUp As Boolean
Dim Direc As String
Dim Stopped As Boolean
Dim fCount As Integer
Dim CurrLevel As Integer
Function load1()
ball.Picture = ShipUp.Picture
Shape1(0).Left = 7680
Shape1(0).Top = 480
Shape1(1).Left = 7680
Shape1(1).Top = 1560
Shape1(2).Left = 7680
Shape1(2).Top = 1200
Shape1(3).Left = 7680
Shape1(3).Top = 840
Shape1(4).Left = 1080
Shape1(4).Top = 1800
Shape1(5).Left = 7320
Shape1(5).Top = 480
Shape1(6).Left = 7320
Shape1(6).Top = 840
Shape1(7).Left = 6480
Shape1(7).Top = 6120
Shape1(8).Left = 7320
Shape1(8).Top = 1920
Shape1(9).Left = 7320
Shape1(9).Top = 2640
Shape1(10).Left = 0
Shape1(10).Top = 4680
Shape1(11).Left = 3240
Shape1(11).Top = 0
Shape1(12).Left = 5760
Shape1(12).Top = 6480
Shape1(13).Left = 7320
Shape1(13).Top = 2280
Shape1(14).Left = 7320
Shape1(14).Top = 1560
Shape1(15).Left = 7320
Shape1(15).Top = 1200
Shape1(16).Left = 7320
Shape1(16).Top = 3000
Shape1(17).Left = 7680
Shape1(17).Top = 2640
Shape1(18).Left = 2160
Shape1(18).Top = 3600
Shape1(19).Left = 2520
Shape1(19).Top = 2160
Shape1(20).Left = 7680
Shape1(20).Top = 3360
Shape1(21).Left = 7680
Shape1(21).Top = 3000
Shape1(22).Left = 7320
Shape1(22).Top = 3360
Shape1(23).Left = 1440
Shape1(23).Top = 5040
Shape1(24).Left = 7680
Shape1(24).Top = 1920
Shape1(25).Left = 7680
Shape1(25).Top = 2280
Food(0).Left = 5880
Food(0).Top = 3360
Food(1).Left = 5520
Food(1).Top = 2280
Food(2).Left = 120
Food(2).Top = 4440
Food(3).Left = 6240
Food(3).Top = 5880
Food(4).Left = 6600
Food(4).Top = 480
Food(5).Left = 4800
Food(5).Top = 6600
Food(6).Left = 1560
Food(6).Top = 2280
Food(7).Left = 120
Food(7).Top = 5160
Food(8).Left = 2280
Food(8).Top = 3000
Food(9).Left = 1200
Food(9).Top = 4800
ball.Left = 3240
ball.Top = 6480
End Function
Function load2()
ball.Picture = ShipUp.Picture
Shape1(0).Left = 7200
Shape1(0).Top = 1800
Shape1(1).Left = 7200
Shape1(1).Top = 1080
Shape1(2).Left = 7200
Shape1(2).Top = 360
Shape1(3).Left = 7200
Shape1(3).Top = 1440
Shape1(4).Left = 0
Shape1(4).Top = 5040
Shape1(5).Left = 720
Shape1(5).Top = 4680
Shape1(6).Left = 7200
Shape1(6).Top = 0
Shape1(7).Left = 7200
Shape1(7).Top = 720
Shape1(8).Left = 1080
Shape1(8).Top = 5400
Shape1(9).Left = 720
Shape1(9).Top = 5760
Shape1(10).Left = 360
Shape1(10).Top = 2520
Shape1(11).Left = 7560
Shape1(11).Top = 720
Shape1(12).Left = 7560
Shape1(12).Top = 360
Shape1(13).Left = 7560
Shape1(13).Top = 0
Shape1(14).Left = 6120
Shape1(14).Top = 0
Shape1(15).Left = 2880
Shape1(15).Top = 1800
Shape1(16).Left = 6480
Shape1(16).Top = 2160
Shape1(17).Left = 0
Shape1(17).Top = 0
Shape1(18).Left = 0
Shape1(18).Top = 1800
Shape1(19).Left = 6120
Shape1(19).Top = 2880
Shape1(20).Left = 5760
Shape1(20).Top = 6480
Shape1(21).Left = 5400
Shape1(21).Top = 4320
Shape1(22).Left = 6480
Shape1(22).Top = 5040
Shape1(23).Left = 1440
Shape1(23).Top = 4320
Shape1(24).Left = 6480
Shape1(24).Top = 4320
Shape1(25).Left = 7560
Shape1(25).Top = 1080
Food(0).Left = 5520
Food(0).Top = 3000
Food(1).Left = 120
Food(1).Top = 6600
Food(2).Left = 6240
Food(2).Top = 5160
Food(3).Left = 5880
Food(3).Top = 120
Food(4).Left = 120
Food(4).Top = 480
Food(5).Left = 6600
Food(5).Top = 4800
Food(6).Left = 5880
Food(6).Top = 6240
Food(7).Left = 3000
Food(7).Top = 1560
Food(8).Left = 1920
Food(8).Top = 3360
Food(9).Left = 840
Food(9).Top = 5520
ball.Left = 2880
ball.Top = 6480
End Function
Function load3()
ball.Picture = ShipRight.Picture
Shape1(0).Left = 2880
Shape1(0).Top = 0
Shape1(1).Left = 2520
Shape1(1).Top = 1800
Shape1(2).Left = 6480
Shape1(2).Top = 1440
Shape1(3).Left = 3240
Shape1(3).Top = 6480
Shape1(4).Left = 360
Shape1(4).Top = 6120
Shape1(5).Left = 360
Shape1(5).Top = 1800
Shape1(6).Left = 6480
Shape1(6).Top = 2160
Shape1(7).Left = 6120
Shape1(7).Top = 5400
Shape1(8).Left = 5040
Shape1(8).Top = 5040
Shape1(9).Left = 5400
Shape1(9).Top = 2880
Shape1(10).Left = 0
Shape1(10).Top = 3240
Shape1(11).Left = 10000
Shape1(11).Top = 10000
Shape1(12).Left = 10000
Shape1(12).Top = 10000
Shape1(13).Left = 10000
Shape1(13).Top = 10000
Shape1(14).Left = 10000
Shape1(14).Top = 10000
Shape1(15).Left = 10000
Shape1(15).Top = 10000
Shape1(16).Left = 10000
Shape1(16).Top = 10000
Shape1(17).Left = 10000
Shape1(17).Top = 10000
Shape1(18).Left = 10000
Shape1(18).Top = 10000
Shape1(19).Left = 10000
Shape1(19).Top = 10000
Shape1(20).Left = 10000
Shape1(20).Top = 10000
Shape1(21).Left = 10000
Shape1(21).Top = 10000
Shape1(22).Left = 10000
Shape1(22).Top = 10000
Shape1(23).Left = 10000
Shape1(23).Top = 10000
Shape1(24).Left = 10000
Shape1(24).Top = 10000
Shape1(25).Left = 10000
Shape1(25).Top = 10000
Food(0).Left = 2640
Food(0).Top = 120
Food(1).Left = 5520
Food(1).Top = 3360
Food(2).Left = 5160
Food(2).Top = 4800
Food(3).Left = 3360
Food(3).Top = 5880
Food(4).Left = 6600
Food(4).Top = 6600
Food(5).Left = 3360
Food(5).Top = 120
Food(6).Left = 120
Food(6).Top = 3720
Food(7).Left = 2640
Food(7).Top = 1560
Food(8).Left = 480
Food(8).Top = 2280
Food(9).Left = 120
Food(9).Top = 3000
ball.Left = 0
ball.Top = 0
ball.Visible = True
ball.Picture = ShipRight.Picture
End Function
Function load4()
ball.Picture = ShipUp.Picture
Shape1(0).Left = 7200
Shape1(0).Top = 1080
Shape1(1).Left = 6480
Shape1(1).Top = 6480
Shape1(2).Left = 3600
Shape1(2).Top = 0
Shape1(3).Left = 7200
Shape1(3).Top = 720
Shape1(4).Left = 0
Shape1(4).Top = 3240
Shape1(5).Left = 6480
Shape1(5).Top = 0
Shape1(6).Left = 1080
Shape1(6).Top = 5760
Shape1(7).Left = 6120
Shape1(7).Top = 6480
Shape1(8).Left = 7200
Shape1(8).Top = 360
Shape1(9).Left = 6480
Shape1(9).Top = 3240
Shape1(10).Left = 1800
Shape1(10).Top = 6120
Shape1(11).Left = 2520
Shape1(11).Top = 720
Shape1(12).Left = 6120
Shape1(12).Top = 1800
Shape1(13).Left = 3960
Shape1(13).Top = 6120
Shape1(14).Left = 7560
Shape1(14).Top = 360
Shape1(15).Left = 1440
Shape1(15).Top = 4320
Shape1(16).Left = 7560
Shape1(16).Top = 720
Shape1(17).Left = 7560
Shape1(17).Top = 0
Shape1(18).Left = 4320
Shape1(18).Top = 360
Shape1(19).Left = 2880
Shape1(19).Top = 5400
Shape1(20).Left = 720
Shape1(20).Top = 0
Shape1(21).Left = 7560
Shape1(21).Top = 1080
Shape1(22).Left = 5400
Shape1(22).Top = 2880
Shape1(23).Left = 7200
Shape1(23).Top = 0
Shape1(24).Left = 5760
Shape1(24).Top = 2160
Shape1(25).Left = 1080
Shape1(25).Top = 6480
Food(0).Left = 5160
Food(0).Top = 3000
Food(1).Left = 5880
Food(1).Top = 2640
Food(2).Left = 4440
Food(2).Top = 5880
Food(3).Left = 1560
Food(3).Top = 6600
Food(4).Left = 3360
Food(4).Top = 120
Food(5).Left = 6600
Food(5).Top = 6240
Food(6).Left = 4440
Food(6).Top = 840
Food(7).Left = 840
Food(7).Top = 480
Food(8).Left = 6600
Food(8).Top = 480
Food(9).Left = 120
Food(9).Top = 6600
ball.Left = 1800
ball.Top = 3240
End Function
Function load5()
ball.Picture = ShipRight.Picture
Shape1(0).Left = 7200
Shape1(0).Top = 1080
Shape1(1).Left = 1080
Shape1(1).Top = 3240
Shape1(2).Left = 1440
Shape1(2).Top = 4320
Shape1(3).Left = 7200
Shape1(3).Top = 720
Shape1(4).Left = 6480
Shape1(4).Top = 0
Shape1(5).Left = 7200
Shape1(5).Top = 2160
Shape1(6).Left = 360
Shape1(6).Top = 360
Shape1(7).Left = 2880
Shape1(7).Top = 3600
Shape1(8).Left = 7200
Shape1(8).Top = 360
Shape1(9).Left = 7560
Shape1(9).Top = 1800
Shape1(10).Left = 5400
Shape1(10).Top = 5760
Shape1(11).Left = 1080
Shape1(11).Top = 2520
Shape1(12).Left = 7200
Shape1(12).Top = 1800
Shape1(13).Left = 5040
Shape1(13).Top = 2160
Shape1(14).Left = 7560
Shape1(14).Top = 360
Shape1(15).Left = 5760
Shape1(15).Top = 720
Shape1(16).Left = 7560
Shape1(16).Top = 720
Shape1(17).Left = 7560
Shape1(17).Top = 0
Shape1(18).Left = 4320
Shape1(18).Top = 3960
Shape1(19).Left = 720
Shape1(19).Top = 5400
Shape1(20).Left = 6120
Shape1(20).Top = 6480
Shape1(21).Left = 7560
Shape1(21).Top = 1080
Shape1(22).Left = 7200
Shape1(22).Top = 1440
Shape1(23).Left = 7200
Shape1(23).Top = 0
Shape1(24).Left = 7560
Shape1(24).Top = 1440
Shape1(25).Left = 0
Shape1(25).Top = 6120
Food(0).Left = 4080
Food(0).Top = 3000
Food(1).Left = 2640
Food(1).Top = 5880
Food(2).Left = 3000
Food(2).Top = 5160
Food(3).Left = 5520
Food(3).Top = 1200
Food(4).Left = 2640
Food(4).Top = 480
Food(5).Left = 3000
Food(5).Top = 2640
Food(6).Left = 840
Food(6).Top = 840
Food(7).Left = 120
Food(7).Top = 6600
Food(8).Left = 3000
Food(8).Top = 4080
Food(9).Left = 6600
Food(9).Top = 6600
ball.Left = 0
ball.Top = 0
End Function
Function load6()
ball.Picture = ShipRight.Picture
Shape1(0).Left = 3600
Shape1(0).Top = 0
Shape1(1).Left = 6480
Shape1(1).Top = 360
Shape1(2).Left = 5040
Shape1(2).Top = 3960
Shape1(3).Left = 5400
Shape1(3).Top = 6480
Shape1(4).Left = 7200
Shape1(4).Top = 4320
Shape1(5).Left = 7200
Shape1(5).Top = 3960
Shape1(6).Left = 3240
Shape1(6).Top = 4320
Shape1(7).Left = 3600
Shape1(7).Top = 3240
Shape1(8).Left = 720
Shape1(8).Top = 6480
Shape1(9).Left = 1440
Shape1(9).Top = 5760
Shape1(10).Left = 1080
Shape1(10).Top = 4680
Shape1(11).Left = 5400
Shape1(11).Top = 720
Shape1(12).Left = 3240
Shape1(12).Top = 2520
Shape1(13).Left = 6480
Shape1(13).Top = 2880
Shape1(14).Left = 0
Shape1(14).Top = 6480
Shape1(15).Left = 7560
Shape1(15).Top = 2880
Shape1(16).Left = 1800
Shape1(16).Top = 6120
Shape1(17).Left = 0
Shape1(17).Top = 5040
Shape1(18).Left = 7200
Shape1(18).Top = 3600
Shape1(19).Left = 5760
Shape1(19).Top = 0
Shape1(20).Left = 7200
Shape1(20).Top = 2880
Shape1(21).Left = 0
Shape1(21).Top = 1800
Shape1(22).Left = 7200
Shape1(22).Top = 2520
Shape1(23).Left = 360
Shape1(23).Top = 5760
Shape1(24).Left = 7560
Shape1(24).Top = 2160
Shape1(25).Left = 6120
Shape1(25).Top = 2160
Food(0).Left = 6600
Food(0).Top = 3360
Food(1).Left = 6600
Food(1).Top = 2640
Food(2).Left = 3720
Food(2).Top = 840
Food(3).Left = 480
Food(3).Top = 480
Food(4).Left = 6240
Food(4).Top = 6600
Food(5).Left = 120
Food(5).Top = 5520
Food(6).Left = 6240
Food(6).Top = 1200
Food(7).Left = 120
Food(7).Top = 2280
Food(8).Left = 6240
Food(8).Top = 5160
Food(9).Left = 4800
Food(9).Top = 480
ball.Left = 360
ball.Top = 6120
End Function
Function load7()
ball.Picture = ShipRight.Picture
Shape1(0).Left = 7680
Shape1(0).Top = 480
Shape1(1).Left = 0
Shape1(1).Top = 6120
Shape1(2).Left = 7680
Shape1(2).Top = 1200
Shape1(3).Left = 7680
Shape1(3).Top = 840
Shape1(4).Left = 360
Shape1(4).Top = 5760
Shape1(5).Left = 7320
Shape1(5).Top = 480
Shape1(6).Left = 7320
Shape1(6).Top = 840
Shape1(7).Left = 5400
Shape1(7).Top = 0
Shape1(8).Left = 360
Shape1(8).Top = 3960
Shape1(9).Left = 6120
Shape1(9).Top = 6120
Shape1(10).Left = 2160
Shape1(10).Top = 0
Shape1(11).Left = 5400
Shape1(11).Top = 2880
Shape1(12).Left = 5760
Shape1(12).Top = 360
Shape1(13).Left = 6480
Shape1(13).Top = 5400
Shape1(14).Left = 3600
Shape1(14).Top = 5760
Shape1(15).Left = 7320
Shape1(15).Top = 1200
Shape1(16).Left = 720
Shape1(16).Top = 0
Shape1(17).Left = 3240
Shape1(17).Top = 1800
Shape1(18).Left = 4320
Shape1(18).Top = 5400
Shape1(19).Left = 0
Shape1(19).Top = 2520
Shape1(20).Left = 5760
Shape1(20).Top = 2520
Shape1(21).Left = 6480
Shape1(21).Top = 2160
Shape1(22).Left = 1080
Shape1(22).Top = 360
Shape1(23).Left = 2520
Shape1(23).Top = 6480
Shape1(24).Left = 6480
Shape1(24).Top = 1440
Shape1(25).Left = 3960
Shape1(25).Top = 4680
Food(0).Left = 1200
Food(0).Top = 120
Food(1).Left = 6600
Food(1).Top = 6600
Food(2).Left = 5880
Food(2).Top = 120
Food(3).Left = 480
Food(3).Top = 5160
Food(4).Left = 5520
Food(4).Top = 3720
Food(5).Left = 1920
Food(5).Top = 840
Food(6).Left = 120
Food(6).Top = 4080
Food(7).Left = 120
Food(7).Top = 2280
Food(8).Left = 2640
Food(8).Top = 120
Food(9).Left = 5520
Food(9).Top = 480
ball.Left = 0
ball.Top = 6480
End Function
Function load8()
ball.Picture = ShipRight.Picture
Shape1(0).Left = 7680
Shape1(0).Top = 480
Shape1(1).Left = 720
Shape1(1).Top = 5040
Shape1(2).Left = 7680
Shape1(2).Top = 1200
Shape1(3).Left = 7680
Shape1(3).Top = 840
Shape1(4).Left = 6120
Shape1(4).Top = 6480
Shape1(5).Left = 7320
Shape1(5).Top = 480
Shape1(6).Left = 7320
Shape1(6).Top = 840
Shape1(7).Left = 5760
Shape1(7).Top = 6120
Shape1(8).Left = 2880
Shape1(8).Top = 5040
Shape1(9).Left = 1080
Shape1(9).Top = 0
Shape1(10).Left = 1440
Shape1(10).Top = 6480
Shape1(11).Left = 6120
Shape1(11).Top = 3240
Shape1(12).Left = 6480
Shape1(12).Top = 3600
Shape1(13).Left = 720
Shape1(13).Top = 3600
Shape1(14).Left = 1080
Shape1(14).Top = 3960
Shape1(15).Left = 3960
Shape1(15).Top = 1080
Shape1(16).Left = 1080
Shape1(16).Top = 4680
Shape1(17).Left = 2880
Shape1(17).Top = 0
Shape1(18).Left = 2520
Shape1(18).Top = 3240
Shape1(19).Left = 2520
Shape1(19).Top = 5400
Shape1(20).Left = 4680
Shape1(20).Top = 4680
Shape1(21).Left = 4320
Shape1(21).Top = 6120
Shape1(22).Left = 6480
Shape1(22).Top = 360
Shape1(23).Left = 2880
Shape1(23).Top = 3600
Shape1(24).Left = 4680
Shape1(24).Top = 6480
Shape1(25).Left = 4320
Shape1(25).Top = 2880
Food(0).Left = 1560
Food(0).Top = 5880
Food(1).Left = 1920
Food(1).Top = 840
Food(2).Left = 1200
Food(2).Top = 3720
Food(3).Left = 480
Food(3).Top = 6240
Food(4).Left = 480
Food(4).Top = 1560
Food(5).Left = 6600
Food(5).Top = 120
Food(6).Left = 4080
Food(6).Top = 3000
Food(7).Left = 6600
Food(7).Top = 6600
Food(8).Left = 6600
Food(8).Top = 3360
Food(9).Left = 1200
Food(9).Top = 4440
ball.Left = 0
ball.Top = 6480
End Function
Function load9()
ball.Picture = ShipRight.Picture
Shape1(0).Left = 7680
Shape1(0).Top = 480
Shape1(1).Left = 6120
Shape1(1).Top = 1440
Shape1(2).Left = 7680
Shape1(2).Top = 1200
Shape1(3).Left = 7680
Shape1(3).Top = 840
Shape1(4).Left = 0
Shape1(4).Top = 3600
Shape1(5).Left = 7320
Shape1(5).Top = 480
Shape1(6).Left = 7320
Shape1(6).Top = 840
Shape1(7).Left = 7680
Shape1(7).Top = 1920
Shape1(8).Left = 6480
Shape1(8).Top = 2160
Shape1(9).Left = 4320
Shape1(9).Top = 6480
Shape1(10).Left = 1440
Shape1(10).Top = 5040
Shape1(11).Left = 7680
Shape1(11).Top = 1560
Shape1(12).Left = 1080
Shape1(12).Top = 4680
Shape1(13).Left = 3600
Shape1(13).Top = 6120
Shape1(14).Left = 360
Shape1(14).Top = 0
Shape1(15).Left = 7320
Shape1(15).Top = 1200
Shape1(16).Left = 6480
Shape1(16).Top = 3960
Shape1(17).Left = 4680
Shape1(17).Top = 5760
Shape1(18).Left = 720
Shape1(18).Top = 1800
Shape1(19).Left = 0
Shape1(19).Top = 4320
Shape1(20).Left = 7320
Shape1(20).Top = 1920
Shape1(21).Left = 6120
Shape1(21).Top = 4320
Shape1(22).Left = 7320
Shape1(22).Top = 1560
Shape1(23).Left = 720
Shape1(23).Top = 2520
Shape1(24).Left = 1080
Shape1(24).Top = 1800
Shape1(25).Left = 6480
Shape1(25).Top = 0
Food(0).Left = 4080
Food(0).Top = 120
Food(1).Left = 6240
Food(1).Top = 6240
Food(2).Left = 6600
Food(2).Top = 2640
Food(3).Left = 120
Food(3).Top = 3360
Food(4).Left = 120
Food(4).Top = 4800
Food(5).Left = 1560
Food(5).Top = 4800
Food(6).Left = 1200
Food(6).Top = 4440
Food(7).Left = 5880
Food(7).Top = 1560
Food(8).Left = 6240
Food(8).Top = 1200
Food(9).Left = 6600
Food(9).Top = 4440
ball.Left = 360
ball.Top = 3960
ball.Visible = True
ball.Picture = ShipRight.Picture
End Function
Function load10()
ball.Picture = ShipUp.Picture
Shape1(0).Left = 2880
Shape1(0).Top = 6480
Shape1(1).Left = 2160
Shape1(1).Top = 4320
Shape1(2).Left = 2160
Shape1(2).Top = 3240
Shape1(3).Left = 3960
Shape1(3).Top = 3240
Shape1(4).Left = 1440
Shape1(4).Top = 5040
Shape1(5).Left = 2880
Shape1(5).Top = 2160
Shape1(6).Left = 4320
Shape1(6).Top = 3960
Shape1(7).Left = 4320
Shape1(7).Top = 360
Shape1(8).Left = 0
Shape1(8).Top = 5040
Shape1(9).Left = 6480
Shape1(9).Top = 4320
Shape1(10).Left = 2160
Shape1(10).Top = 0
Shape1(11).Left = 0
Shape1(11).Top = 3600
Shape1(12).Left = 5040
Shape1(12).Top = 0
Shape1(13).Left = 0
Shape1(13).Top = 360
Shape1(14).Left = 7680
Shape1(14).Top = 120
Shape1(15).Left = 1800
Shape1(15).Top = 3960
Shape1(16).Left = 1800
Shape1(16).Top = 6480
Shape1(17).Left = 2520
Shape1(17).Top = 3600
Shape1(18).Left = 2520
Shape1(18).Top = 5400
Shape1(19).Left = 2880
Shape1(19).Top = 4680
Shape1(20).Left = 5400
Shape1(20).Top = 720
Shape1(21).Left = 5040
Shape1(21).Top = 6480
Shape1(22).Left = 4680
Shape1(22).Top = 720
Shape1(23).Left = 1800
Shape1(23).Top = 5760
Shape1(24).Left = 3240
Shape1(24).Top = 5760
Shape1(25).Left = 7320
Shape1(25).Top = 120
Food(0).Left = 6240
Food(0).Top = 2640
Food(1).Left = 1560
Food(1).Top = 6600
Food(2).Left = 5520
Food(2).Top = 1200
Food(3).Left = 2640
Food(3).Top = 120
Food(4).Left = 5520
Food(4).Top = 6600
Food(5).Left = 120
Food(5).Top = 4800
Food(6).Left = 120
Food(6).Top = 3360
Food(7).Left = 3720
Food(7).Top = 840
Food(8).Left = 4800
Food(8).Top = 6600
Food(9).Left = 120
Food(9).Top = 120
ball.Left = 2520
ball.Top = 6480
End Function
Function ScanForHit()
'control array and for next
For i = 0 To 25 '25 max squares
If ball.Top - Shape1(i).Top < 360 And Shape1(i).Top - ball.Top < 360 And Shape1(i).Left - ball.Left = 360 Then NoRight = True 'scan if immediate collison to the right would occurr
If ball.Top - Shape1(i).Top < 360 And Shape1(i).Top - ball.Top < 360 And ball.Left - Shape1(i).Left = 360 Then NoLeft = True ' ''   '   ''        ' '   '  ' '  left ' ' ' '' '  '
If ball.Top - Shape1(i).Top = 360 And ball.Left - Shape1(i).Left < 360 And Shape1(i).Left - ball.Left < 360 Then NoUp = True
If Shape1(i).Top - ball.Top = 360 And ball.Left - Shape1(i).Left < 360 And Shape1(i).Left - ball.Left < 360 Then NoDown = True
Next i
'
If NoRight = True Or NoLeft = True Or NoUp = True Or NoDown = True Then Stopped = True
End Function
Function ScanFood()
On Error Resume Next 'if for some reason the .wav file can't be found then it just won't play a sound.
'
For i = 0 To 9
If Food(i).Top - ball.Top = 120 And Food(i).Left - ball.Left = 120 Then
'Food(i).Visible = False
Food(i).Top = -1000 'get it outta the way so unreal beeps wont occur
Food(i).Left = -1000
fCount = fCount + 1
PlayWav "getpoint.wav"
End If
Next i
'
'
If fCount = 10 Then 'if all dots collected:
'response = MsgBox("You beat level " & CurrLevel & "!  Load next level?", vbYesNo + vbInformation, "Congratulations!")
'    If response = vbYes Then
    If CurrLevel = 10 Then
    MsgBox "You Win!  Congrats!  If you liked this game then please tell me!  (Aedseed@aol.com)  I'm thinking of adding more to this game (not only more levels, but a whole new concept - being able to control not only the ship but also another thing so you could use one to bank another off and vise versa.  So if you'd like to see more then tell me!  Thanks!  Also I'm thinking of adding a level editor (which would be very simple to make!)", vbOKOnly + vbInformation, "Yay!"
    Unload Me
    Unload frm2
    End
    End If
    LoadNextLvl
'    Else
'    Unload Me
'    End
'    End If
End If
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Ball.Visible = False
release = False 'duh, its being pressed!
'
NoRight = False 'setting them all
'to false lets u move nomatter what
NoLeft = False 'but still prevents going
NoDown = False 'thru walls and squares!
NoUp = False
'
If KeyCode = vbKeyRight Then 'if right arrow key is pressed
    If Stopped = False Then
        If Direc = "Vert" Then Exit Sub 'cant cheat!
    End If
    Direc = "Horiz" 'sets the line you're on as horizontal
    ball.Picture = ShipRight.Picture
Do Until release = True 'move right until right arrow key is released
    If ball.Left = 6480 Then
    Stopped = True
    ScanFood
    Exit Sub 'if on right border exit(cant leave form)
    End If
    ScanForHit 'main collison detection (see ScanForHit above)
    ScanFood
        If NoRight = True Then Exit Sub 'if a collision would occur, exit sub
ball.Left = ball.Left + 360 'if everythings OK, then finally move right
Stopped = False
DoEvents 'prevents freezing and nonstop loops!
Loop
End If
'
If KeyCode = vbKeyLeft Then
    If Stopped = False Then
        If Direc = "Vert" Then Exit Sub 'cant cheat!
    End If
    Direc = "Horiz"
    ball.Picture = ShipLeft.Picture
Do Until release = True
    If ball.Left = 0 Then
    Stopped = True
    ScanFood
    Exit Sub
    End If
    ScanForHit
    ScanFood
        If NoLeft = True Then Exit Sub
ball.Left = ball.Left - 360
Stopped = False
DoEvents
Loop
End If
'
If KeyCode = vbKeyUp Then
    If Stopped = False Then
        If Direc = "Horiz" Then Exit Sub 'cant cheat!
    End If
    Direc = "Vert"
    ball.Picture = ShipUp.Picture
Do Until release = True
    If ball.Top = 0 Then
    Stopped = True
    ScanFood
    Exit Sub
    End If
    ScanForHit
    ScanFood
        If NoUp = True Then Exit Sub
ball.Top = ball.Top - 360
Stopped = False
DoEvents
Loop
End If
'
If KeyCode = vbKeyDown Then
    If Stopped = False Then
        If Direc = "Horiz" Then Exit Sub 'cant cheat!
    End If
    Direc = "Vert"
    ball.Picture = ShipDown.Picture
Do Until release = True
    If ball.Top = 6480 Then
    Stopped = True
    ScanFood
    Exit Sub
    End If
    ScanForHit
    ScanFood
        If NoDown = True Then Exit Sub
ball.Top = ball.Top + 360
Stopped = False
DoEvents
Loop
End If
'
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'ball.Visible = True
release = True 'stops movement immediately
End Sub

Private Sub Form_Load()
Me.Height = 6870
'Me.Width = 6870
Me.BackColor = vbBlack
Frame1.Left = 6840
Frame1.Top = 0
'
fCount = 0
CurrLevel = 0
LoadNextLvl
ball.Picture = ShipUp.Picture
End Sub
Sub LoadNextLvl()
CurrLevel = CurrLevel + 1
lblLVL.Caption = CurrLevel
'ball.Visible = True
fCount = 0
'level loading presets:
If CurrLevel = 1 Then load1
If CurrLevel = 2 Then load2
If CurrLevel = 3 Then
ball.Picture = ShipRight.Picture ' for some reason it never would show the ship at the start of lvl3 and this fixed it...
load3
End If
If CurrLevel = 4 Then load4
If CurrLevel = 5 Then load5
If CurrLevel = 6 Then load6
If CurrLevel = 7 Then load7
If CurrLevel = 8 Then load8
If CurrLevel = 9 Then load9
If CurrLevel = 10 Then load10
End Sub



'all this is just interface crap that deals w/ the 3 buttons:

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
Label6.ForeColor = vbBlack
Label7.ForeColor = vbBlack
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
Label6.ForeColor = vbBlack
Label7.ForeColor = vbBlack
End Sub

Private Sub Label1_Click()
frm2.Show
End Sub

Private Sub Label5_Click()
MsgBox "The object of the game is to collect all of the green dots by moving around with the ship using the Up, Down, Left and Right keys.  Each level has ten points you need to get to advance to the next level.  There are currently ten levels which progressively get more difficult." & vbCrLf & vbCrLf & "This game is based off of a Ti-89 game called DeathStar.  I think it's a pretty fun game and I hope you do too.  If you like it or think it's fun then please Vote!", vbInformation + vbOKOnly, "About:"
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF00&
Label6.ForeColor = vbBlack
Label7.ForeColor = vbBlack
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF00&
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlack
'
FormDrag Me
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF00&
Label5.ForeColor = vbBlack
Label7.ForeColor = vbBlack
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF00&
End Sub

Private Sub Label7_Click()
Unload Me
Unload frm2
End
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF00&
Label5.ForeColor = vbBlack
Label6.ForeColor = vbBlack
End Sub

