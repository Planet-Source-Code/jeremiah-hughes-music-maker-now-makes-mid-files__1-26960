VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Music Maker"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   -1335
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "MusicMaker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton RewindBut 
      Height          =   375
      Left            =   840
      Picture         =   "MusicMaker.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Return to Start"
      Top             =   960
      Width           =   375
   End
   Begin VB.Frame FunctionsFrame 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   240
      Width           =   3015
      Begin VB.CommandButton InsertColumnBut 
         Height          =   375
         Left            =   2520
         Picture         =   "MusicMaker.frx":0BF2
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Insert Space"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton PasteIntoBut 
         Height          =   375
         Left            =   2160
         Picture         =   "MusicMaker.frx":0D44
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Paste (Insert)"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton PasteBut 
         Height          =   375
         Left            =   1800
         Picture         =   "MusicMaker.frx":14F4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Paste (Overwrite)"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CutBut 
         Height          =   375
         Left            =   1440
         Picture         =   "MusicMaker.frx":1CA4
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Cut Selection"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton CopyBut 
         Height          =   375
         Left            =   1080
         Picture         =   "MusicMaker.frx":2454
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Copy Selection"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton FunctionBut 
         Height          =   375
         Index           =   2
         Left            =   720
         Picture         =   "MusicMaker.frx":2C04
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Select Area"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton FunctionBut 
         Height          =   375
         Index           =   1
         Left            =   360
         Picture         =   "MusicMaker.frx":33B2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Start/Insertion Position"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton FunctionBut 
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "MusicMaker.frx":3B60
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Draw Mode"
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.Frame TransposeFrame 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9120
      TabIndex        =   34
      Top             =   240
      Width           =   1215
      Begin VB.CommandButton TransposeBut 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "MusicMaker.frx":4310
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Move All Notes Up One"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton TransposeBut 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "MusicMaker.frx":4AC0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Move All Notes Down One"
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mia"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   31
      Top             =   1080
      Width           =   2895
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   3120
      Max             =   300
      Min             =   20
      TabIndex        =   28
      Top             =   240
      Value           =   120
      Width           =   2895
   End
   Begin VB.OptionButton PlayAndStop 
      Height          =   375
      Index           =   1
      Left            =   1800
      Picture         =   "MusicMaker.frx":5270
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Stop"
      Top             =   960
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton PlayAndStop 
      Height          =   375
      Index           =   0
      Left            =   1320
      Picture         =   "MusicMaker.frx":5A20
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Play"
      Top             =   960
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      LargeChange     =   25
      Left            =   10260
      Max             =   115
      Min             =   12
      TabIndex        =   24
      Top             =   2040
      Value           =   12
      Width           =   255
   End
   Begin VB.Frame MuteTrackFrame 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      Top             =   1080
      Width           =   2895
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H00446699&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H00FF00FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H00FF8080&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H00FFFF00&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H0000FF00&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H0000C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H000080FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox TrackHide 
         BackColor       =   &H000000FF&
         Caption         =   "1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox MusicBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   2040
      Width           =   9615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   40
      Left            =   600
      Max             =   980
      Min             =   20
      TabIndex        =   1
      Top             =   8100
      Value           =   20
      Width           =   9645
   End
   Begin VB.PictureBox SwapScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.Frame EditTrackFrame 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00446699&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00FF00FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00FF8080&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H00FFFF00&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H0000FF00&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H0000C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H000080FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton TrackSel 
         BackColor       =   &H000000FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Serpentine"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.PictureBox BlueBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   70
      Top             =   2040
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.PictureBox GreenBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   69
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox RedBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   600
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   68
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label TopBarOverlay 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   58
      Top             =   1800
      Width           =   9600
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   9720
      TabIndex        =   67
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   9720
      TabIndex        =   66
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   9720
      TabIndex        =   65
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   9720
      TabIndex        =   64
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   9720
      TabIndex        =   63
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   9720
      TabIndex        =   62
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   9720
      TabIndex        =   61
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   9720
      TabIndex        =   60
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   9720
      TabIndex        =   59
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   57
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   56
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   55
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   54
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   53
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   52
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   51
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   50
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   49
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label TopBar 
      BackColor       =   &H00000000&
      Caption         =   "9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   9240
      TabIndex        =   48
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Functions"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Transpose"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   33
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label InstLabel 
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Instruments"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label TempoLabel 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3120
      TabIndex        =   29
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Tempo"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Hide Tracks"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Edit Track"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   0
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   6045
      Left            =   120
      Top             =   2040
      Width           =   10125
   End
   Begin VB.Label SideBar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6315
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   600
      Top             =   1800
      Width           =   9915
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mExporttoMidi 
         Caption         =   "Export to Midi"
         Shortcut        =   ^E
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu midi_devices 
      Caption         =   "Midi Device"
      Begin VB.Menu Device 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mOptions 
      Caption         =   "Options"
      Begin VB.Menu mUseAlphaBlending 
         Caption         =   "Use Alpha Blending"
      End
   End
   Begin VB.Menu mExtra 
      Caption         =   "Extra"
      Begin VB.Menu mReverseNotes 
         Caption         =   "Reverse Notes"
      End
      Begin VB.Menu mFlipNotes 
         Caption         =   "Flip Notes"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim Grid%(7, 1000) '(Tracks, Value)
Dim InstGrid%(7, 1000)
Dim CopyGrid%(7, 1000)
Dim CopyInstGrid%(7, 1000)
Dim CopyGridLen%
Dim A%, B%, C%, D%, E%, F%, G%, Z%
Dim XSize%, YSize%
Dim XSizeMid%, YSizeMid%
Dim Temp$, TempInt%
Dim NoteNames$(127), Sharps(100) As Boolean
Dim Colr As Long, BGColr As Long
Dim GridX%, GridY%
Dim OldGridX%, OldGridY%
Dim MovingGridX%, OldMovingGridX%
Dim HoverX%, HoverY%
Dim OldHoverX%, OldHoverY%
Dim StartX%, StartY%
Dim EndX%, EndY%
Dim NotePlayX%
Dim Note%
Dim CurrentTrack%, TrackColr As Long
Dim TrackColrs(7) As Long, DimmedColrs(7) As Long
Dim TheNotes As Variant
Dim Offset%
Dim Flip%
Dim TopBarOffset%
Dim Tick As Long
Dim PlayingSong As Boolean, PlayX%
Dim OldNote%(7)
Dim NoteType%, OldNoteType%
Dim TempoWait%
Dim CurrentInst%(7), OldInst%(7)
Dim SongLength%, TrackLength%
Dim FilePath$, ProgramPath$
Dim FF%
Dim CurrentlyOpenFile
Dim CursorType As Integer, ColumnX As Integer
Dim PlayBarX As Integer, HasStarted As Boolean
Dim JustScrolled As Boolean, TurningPage As Boolean
Dim MouseIsDown As Boolean
Dim SelectStartX%, SelectEndX%
Dim VisSelStartX%, VisSelEndX%, VisSelLengthX%
Dim TempColumnCursor As Boolean
Dim Resizing As Boolean, FormLoading As Boolean
Dim AlphaBlendOn As Boolean
Dim DeltaTime As Long, LastDeltaTime As Long
Dim DeltaTimeArray(3) As Byte
Dim DeltaBinArray(3) As String
Dim CurrentByte%
Dim TotalBytes As Long
Dim Bytes(10000) As Byte
Dim TrackDataStartByte%
Dim IsDrum(7) As Boolean, DrumNum%(7)
Dim MidiIsDrum As Boolean, MidiDrumNum%
Dim MidiInst%
Dim TempChannel%
Dim TempZ%
Dim DeltaTicks As Integer

Const DrawCursor = 0
Const ColumnCursor = 1
Const SelectCursor = 2

Const BlankNote = 0
Const StartingNote = 1
Const ContinuingNote = 2

'Alpha blending
Private Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal x As Long, ByVal y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean 'only Windows 98 or Later
Dim nBlend&


'Midi Variables (from midi piano program)

'for piano play
Dim numDevices As Long      ' number of midi output devices
Dim curDevice As Long       ' current midi device
Dim hmidi As Long           ' midi output handle
Dim rc As Long              ' return code
Dim midimsg As Long         ' midi output message buffer
Dim channel As Integer      ' midi output channel
Dim volume As Integer       ' midi volume
Dim incra As Integer        ' increment the note
Dim Tempo As Integer        ' set playing speed
Dim incraup As Integer      ' incra-1

Private Sub Combo1_Click()
Dim Instrument%

Instrument = Combo1.ListIndex

'if anything is selected then change all selected notes on current track
'to selected instrument, otherwise just change current instrument
If SelectStartX > -1 Then
    For A = SelectStartX To SelectEndX
        If Grid(CurrentTrack, A) > -1 Then
            InstGrid(CurrentTrack, A) = Instrument
        End If
    Next A
End If

CurrentInst(CurrentTrack) = Instrument
channel = CurrentTrack
ChangeInstrument CurrentInst(CurrentTrack)

'play short sample of instrument
If Not FormLoading Then
    Tick = GetTickCount
    StartNote 67
    Do While Tick + 250 > GetTickCount
        DoEvents
    Loop
    StopNote 67
End If

End Sub

'Note bending test
'Private Sub Command1_Click()
'Dim W As Integer, Abc As Integer, Steps As Integer
'
'Steps = Val(Text1.Text)
'
'If Steps < 1 Then Steps = 1
'
'W = 250 \ Steps
'
'SetPitchBend 0, 64, 0
'
'Tick = GetTickCount
'StartNote 67
'Do While Tick + W > GetTickCount
'    DoEvents
'Loop
'
'
'For A = 1 To Steps
'
'Abc = 64 + A * (64 \ Steps)
'
'If A = Steps Then Abc = 127
'
'SetPitchBend 0, Abc, 0
'
'Tick = GetTickCount
'Do While Tick + W > GetTickCount
'   DoEvents
'Loop
'
'Next A
'
'StopNote 67
'
'SetPitchBend 0, 64, 0
'
'End Sub
'
'Public Function SetPitchBend(ByVal nLSB As Integer, ByVal nMSB As Integer, ByVal nChannel As Integer)
'    midimsg = &HE0 + nLSB * &H100 + nMSB * &H10000 + nChannel
'    midiOutShortMsg hmidi, midimsg
'End Function

Private Sub CopyBut_Click()
MusicBox.SetFocus

If SelectStartX > -1 Then 'something is selected
    
    CopySelectedArea
    
    SelectStartX = -1 'remove blue highlight
    SelectEndX = -1

    DrawGrid
    
    FunctionBut(1).Value = True 'change to red column mode
    
End If

End Sub

Private Sub CutBut_Click()
Dim Jump%

MusicBox.SetFocus

If SelectStartX > -1 Then 'something is selected
    
    CopySelectedArea
        
    'move all notes that are after selected area to start of selected area
    Jump = CopyGridLen + 1
    For A = SelectStartX To 1000 - Jump
        For B = 0 To 7
            If TrackHide(B) = 0 Then
                Grid(B, A) = Grid(B, A + Jump)
                InstGrid(B, A) = InstGrid(B, A + Jump)
            End If
        Next B
    Next A
    
    'change start of severed notes from continuing notes to starting notes
    For A = 0 To 7
        If Grid(A, SelectStartX) >= 1000 And TrackHide(A) = 0 Then
            Grid(A, SelectStartX) = Grid(A, SelectStartX) - 1000
        End If
    Next A

    'delete extra space at end of Grid and InstGrid
    For A = 1001 - Jump To 1000
        For B = 0 To 7
            If TrackHide(B) = 0 Then
                Grid(B, A) = -1
                InstGrid(B, A) = 0
            End If
        Next B
    Next A
    
    SelectStartX = -1 'remove blue highlight
    SelectEndX = -1
    
    FunctionBut(1).Value = True 'change to red column mode

    DrawGrid
    
End If

End Sub

Private Sub CopySelectedArea()

CopyGridLen = SelectEndX - SelectStartX 'record length of selected area
Erase CopyGrid, CopyInstGrid

'copy selected area's data to CopyGrid and CopyInstGrid
For A = SelectStartX To SelectEndX
    For B = 0 To 7
        If TrackHide(B) = 0 Then
            CopyGrid(B, A - SelectStartX) = Grid(B, A)
            CopyInstGrid(B, A - SelectStartX) = InstGrid(B, A)
        End If
    Next B
Next A

'change start of severed notes from continuing notes to starting notes
For A = 0 To 7
    If CopyGrid(A, 0) >= 1000 Then
        CopyGrid(A, 0) = CopyGrid(A, 0) - 1000
    End If
Next A

End Sub


Private Sub device_Click(Index As Integer)

Device(curDevice + 1).Checked = False
Device(Index).Checked = True
curDevice = Index - 1
rc = midiOutClose(hmidi)
rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
   
If (rc <> 0) Then
      MsgBox "Couldn't open midi out, rc = " & rc
End If

End Sub

Private Sub Form_Load()

FormLoading = True

ProgramPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")

'Midi Stuff
   Dim I As Long
   Dim caps As MIDIOUTCAPS
   
   ' Set the first device as midi mapper
   Device(0).Caption = "MIDI Mapper"
   Device(0).Visible = True
   Device(0).Enabled = True
   
   ' Get the rest of the midi devices
   numDevices = midiOutGetNumDevs()
   
   For I = 0 To (numDevices - 1)
      midiOutGetDevCaps I, caps, Len(caps)
        Device(I + 1).Caption = caps.szPname
        Device(I + 1).Visible = True
      Device(I + 1).Enabled = True
   Next
   
   'Select the MIDI Mapper as the default device
   device_Click (0)
   
   ' Set the default channel
   channel = 0
   
   ' Set volume range
   volume = 127



'Initialize
ClearGrids

For A = 100 To 227 'load up instrument names
    Combo1.AddItem A - 100 & " " & LoadResString(A)
Next A

For A = 35 To 81 'load up drum names
    Combo1.AddItem A + 93 & " " & LoadResString(A)
Next A

Combo1.ListIndex = 0

XSize = 39 'width of musicbox (columns)
YSize = 24 'height of musicbox (rows)
XSizeMid = (XSize + 1) \ 2
YSizeMid = (YSize + 1) \ 2

CursorType = 0 'draw mode
ColumnX = 0 'red play start position bar starts at 0
StartX = 0 'leftmost column displayed
EndX = StartX + XSize 'rightmost column displayed
StartY = 55 'topmost row displayed
EndY = StartY + YSize 'bottommost row displayed
SelectStartX = -1 'no area is selected
SelectEndX = -1 'no area is selected

TopBarOffset = 40 'for printing topbar (column #'s) 40 pixles from left edge of form

Tempo = 120
HScroll2.Value = Tempo

CurrentTrack = 0
For A = 0 To 7
    TrackColrs(A) = TrackSel(A).BackColor
Next A
TrackColr = TrackColrs(0)

DimmedColrs(0) = &HC0C0FF
DimmedColrs(1) = &HC0E0FF
DimmedColrs(2) = &HC0FFFF
DimmedColrs(3) = &HC0FFC0
DimmedColrs(4) = &HFFFFC0
DimmedColrs(5) = &HFFC0C0
DimmedColrs(6) = &HFFC0FF
DimmedColrs(7) = &HAACCEE
BGColr = MusicBox.BackColor
nBlend = vbBlue - CLng(155) * (vbYellow + 1) 'used for alphablending

If Dir(ProgramPath & "MusicMaker.dat") <> "" Then
    FF = FreeFile
    Open ProgramPath & "MusicMaker.dat" For Input As #FF
    Input #FF, TempInt
    Close #FF
    
    AlphaBlendOn = TempInt
    mUseAlphaBlending.Checked = AlphaBlendOn
    
End If

'Assign names to all 128 notes
'notes go from high 10G (note 0) to low 0C (note 128)
'to better suit screen's top to bottom Y coordinate system
'when playing note, number is flipped (127 - #)
TheNotes = Split("C,C#,D,D#,E,F,F#,G,G#,A,A#,B", ",")
For A = 0 To 11
    For B = 0 To 11
        C = A * 12 + B
        If C < 128 Then
            NoteNames(127 - C) = A & TheNotes(B)
        End If
    Next B
Next A

PrintSideBar

DrawGrid

'check to see if file should be loaded on startup
If Len(Command) > 2 Then
    FilePath = Mid(Command, 2, Len(Command) - 2)
    If Dir(FilePath) <> "" Then
        CurrentlyOpenFile = FilePath
        OpenFile
    End If
End If

FormLoading = False

End Sub

Private Sub Form_Resize()

Resizing = True

If Me.WindowState = 1 Then Exit Sub

'if not fullscreen and form size is smaller than minimum allowed then change
If WindowState <> 2 Then
    If Me.ScaleWidth < 718 Then
        Me.Width = ScaleX(718, vbPixels, vbTwips)
    End If

    If Me.ScaleHeight < 325 Then
        Me.Height = ScaleY(325, vbPixels, vbTwips)
    End If
End If

'figure out new width and height of grid
XSize = (Me.ScaleWidth - 70) \ 16 - 1
YSize = (Me.ScaleHeight - 165) \ 16 - 1

'resize shapes and controls
MusicBox.Width = (XSize + 1) * 16 + 1 'width of musicbox (pixels)
MusicBox.Height = (YSize + 1) * 16 + 1 'height of musicbox (pixels)
SwapScreen.Width = MusicBox.Width
SwapScreen.Height = MusicBox.Height
BlueBox.Width = MusicBox.Width
BlueBox.Height = MusicBox.Height
GreenBox.Height = MusicBox.Height
RedBox.Height = MusicBox.Height
TopBarOverlay.Width = MusicBox.Width - 1 'clickable overlay for topbar
Shape1.Width = MusicBox.Width + 34 'thin black borer around musicbox
Shape1.Height = MusicBox.Height + 2
Shape2.Width = MusicBox.Width + 20 'topbar black background box
SideBar.Height = MusicBox.Height + 20 'sidebar
HScroll1.Width = MusicBox.Width + 2 'horizontal scrollbar
HScroll1.Top = MusicBox.Height + 139
VScroll1.Height = MusicBox.Height 'vertical scrollbar
VScroll1.Left = MusicBox.Width + 43

XSizeMid = (XSize + 1) \ 2
YSizeMid = (YSize + 1) \ 2

EndX = StartX + XSize 'rightmost column displayed
'if endx was at the limit and the form was expanded
If EndX > 999 Then
    EndX = 999
    StartX = EndX - XSize
End If

EndY = StartY + YSize 'bottommost row displayed
'if endy was at the limit and the form was expanded
If EndY > 127 Then
    EndY = 127
    StartY = EndY - YSize
End If

'reset scrollbar parameters
HScroll1.Min = XSizeMid
HScroll1.Max = 999 - (XSize - XSizeMid)
HScroll1.LargeChange = XSize + 1
HScroll1.Value = StartX + XSizeMid 'value of horizontal scrollbar = middle column # of displayed screen
VScroll1.Min = YSizeMid
VScroll1.Max = 127 - (YSize - YSizeMid)
VScroll1.LargeChange = YSize + 1
VScroll1.Value = StartY + YSizeMid 'value of vertical scrollbar = middle row# of displayed screen

'trim off extra form border (if not fullscreen)
If Me.WindowState <> 2 Then
    Me.Width = ScaleX(MusicBox.Width + 77, vbPixels, vbTwips)
    Me.Height = ScaleX(MusicBox.Height + 212, vbPixels, vbTwips)
End If

'print side and top bars
PrintSideBar
PrintTopBar

JustScrolled = True

DrawGrid

Resizing = False

MusicBox.SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
' Close current midi device
rc = midiOutClose(hmidi)

End

End Sub


Private Sub FunctionBut_Click(Index As Integer)
MusicBox.SetFocus

CursorType = Index

Select Case CursorType

Case Is = 0
    MusicBox.MousePointer = 0

Case Is = 1
    MusicBox.MousePointer = 2

Case Is = 2
    MusicBox.MousePointer = 2

End Select

End Sub

Private Sub FunctionBut_GotFocus(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub HScroll1_Change()

If Resizing Then Exit Sub

StartX = HScroll1.Value - XSizeMid
EndX = StartX + XSize
PrintTopBar
If Not TurningPage Then DrawGrid
JustScrolled = True

End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub HScroll2_Change()

Tempo = HScroll2.Value
TempoWait = 1000 \ (Tempo / 15)
TempoLabel.Caption = Tempo

End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

Private Sub mExit_Click()
' Close current midi device
rc = midiOutClose(hmidi)
End
End Sub

Private Sub mExporttoMidi_Click()
'file DialogBox settings
With CommonDialog1
    .FileName = "*.mid"
    .Filter = "*.mid"
    .DialogTitle = "Export to Midi File"
    .ShowSave
    FilePath = .FileName
End With

'cancel button pressed
If Right(FilePath, 5) = "*.mid" Or FilePath = "" Then
    Exit Sub
End If

'if file doesn't end with ".mia" then add it
If Right(FilePath, 4) <> ".mid" Then FilePath = FilePath & ".mid"

MidiFile

End Sub

Private Sub mFlipNotes_Click()
Dim FlipStartX%, FlipEndX%, After%

MusicBox.SetFocus

'if something is selected, flip only those notes
'otherwise flip all notes
If SelectStartX > -1 Then
    FlipStartX = SelectStartX
    FlipEndX = SelectEndX
Else
    FlipStartX = 0
    FlipEndX = 1000
End If

'flip notes
For A = 0 To 7
    If TrackHide(A) = 0 Then
        For B = FlipStartX To FlipEndX
            D = Grid(A, B)
            If D > -1 Then
                If D < 128 Then
                    D = 127 - Grid(A, B)
                Else
                    D = 1000 + (1127 - Grid(A, B))
                End If
                Grid(A, B) = D
            End If
        Next B
    End If
Next A

'change cut off notes after flipped area to starting notes
If SelectStartX > -1 Then
    If FlipEndX < 1000 Then
        After = FlipEndX + 1
        For B = 0 To 7
            If TrackHide(B) = 0 Then
                If Grid(B, After) >= 1000 Then
                    Grid(B, After) = Grid(B, After) - 1000
                End If
            End If
        Next B
    End If
End If

DrawGrid
End Sub

Private Sub mNew_Click()
Dim M As VbMsgBoxResult
        
M = MsgBox("Are you sure you want to clear the board and start a new song?", vbOKCancel)

If M = vbOK Then
    FormLoading = True
    CurrentlyOpenFile = ""
    ClearGrids
    For G = 0 To 7
        TrackSel(G).Enabled = True
        TrackHide(G).Enabled = True
        TrackHide(G).Value = False
        CurrentInst(G) = 0
    Next G
    TrackSel(0).Value = True
    TrackHide(0).Enabled = False
    HScroll2.Value = 120
    channel = 0
    ChangeInstrument 0
    Combo1.ListIndex = 0
    TrackColr = TrackColrs(0)
    Me.Caption = "Music Maker"
    ColumnX = 0
    TurningPage = True
    HScroll1.Value = HScroll1.Min
    TurningPage = False
    DrawGrid
    MusicBox.SetFocus
    FormLoading = False
End If
        
End Sub

Private Sub mOpen_Click()

With CommonDialog1
    .FileName = "*.mia"
    .Filter = "*.mia"
    .DialogTitle = "Open Music File"
    .ShowOpen
    FilePath = .FileName
End With

If Right(FilePath, 5) = "*.mia" Or FilePath = "" Then 'cancel button pressed
    Exit Sub
End If

If Dir(FilePath) <> "" Then
    CurrentlyOpenFile = FilePath
    OpenFile
End If
    
End Sub

Private Sub mReverseNotes_Click()
Dim TempGrid%(1000), TempInstGrid%(1000)
Dim RevStartX%, RevEndX%, After%

MusicBox.SetFocus

FindEndOfSong

'if something is selected, reverse only those notes
'otherwise flip all notes
If SelectStartX > -1 Then
    RevStartX = SelectStartX
    RevEndX = SelectEndX
Else
    RevStartX = 0
    RevEndX = SongLength
End If


'Reverse song
For Z = 0 To 7
        
    If TrackHide(Z) = 0 Then
    
        'reset oldnote and oldinst variables
        OldNote(Z) = -1
        OldInst(Z) = -1
        
        'note loop
        For A = RevEndX To RevStartX Step -1
        
            'store instrument in TempInstGrid
            TempInstGrid(A) = InstGrid(Z, A)
                        
            'find note types
            If OldNote(Z) = -1 Then OldNoteType = BlankNote
            If OldNote(Z) >= 0 And OldNote(Z) < 1000 Then OldNoteType = StartingNote
            If OldNote(Z) >= 1000 Then OldNoteType = ContinuingNote
            If Grid(Z, A) = -1 Then NoteType = BlankNote
            If Grid(Z, A) >= 0 And Grid(Z, A) < 1000 Then NoteType = StartingNote
            If Grid(Z, A) >= 1000 Then NoteType = ContinuingNote
            
            'Blank note
            If NoteType = BlankNote Then TempGrid(A) = -1
            
            'Starting note
            If NoteType <> BlankNote And OldNoteType <> ContinuingNote Then
                TempGrid(A) = Grid(Z, A) Mod 1000
            End If
            
            'Continuing note
            If NoteType <> BlankNote And OldNoteType = ContinuingNote Then
                TempGrid(A) = Grid(Z, A) Mod 1000 + 1000
            End If
            
            OldNote(Z) = Grid(Z, A)
            
        Next A
        
        'Copy TempGrid to Grid and reverse instruments
        For A = RevStartX To RevEndX
            B = RevStartX + (RevEndX - A)
            Grid(Z, A) = TempGrid(B)
            InstGrid(Z, A) = TempInstGrid(B)
        Next A
        
    End If
    
Next Z

'change cut off notes after flipped area to starting notes
If SelectStartX > -1 Then
    If RevEndX < 1000 Then
        After = RevEndX + 1
        For B = 0 To 7
            If TrackHide(B) = 0 Then
                If Grid(B, After) >= 1000 Then
                    Grid(B, After) = Grid(B, After) - 1000
                End If
            End If
        Next B
    End If
End If

DrawGrid

End Sub

Private Sub mSave_Click()

If CurrentlyOpenFile = "" Then
    mSaveAs_Click
    Exit Sub
End If

FilePath = CurrentlyOpenFile
SaveFile

End Sub

Private Sub mSaveAs_Click()

'file DialogBox settings
With CommonDialog1
    .FileName = "*.mia"
    .Filter = "*.mia"
    .DialogTitle = "Save Music File as"
    .ShowSave
    FilePath = .FileName
End With

'cancel button pressed
If Right(FilePath, 5) = "*.mia" Or FilePath = "" Then
    Exit Sub
End If

'if file doesn't end with ".mia" then add it
If Right(FilePath, 4) <> ".mia" Then FilePath = FilePath & ".mia"

SaveFile

End Sub

Private Sub SaveFile()

FF = FreeFile

Open FilePath For Output As #FF

For A = 0 To 7
    
    TrackLength = 0
    For B = 1000 To 0 Step -1 'find length of track
        If Grid(A, B) > -1 Then
            TrackLength = B
            Exit For
        End If
    Next B
    
    Temp = ""
    For B = 0 To TrackLength
        Temp = Temp & Grid(A, B) & ","
    Next B
    
    Print #FF, Left(Temp, Len(Temp) - 1)
    
    Temp = ""
    For B = 0 To TrackLength
        Temp = Temp & InstGrid(A, B) & ","
    Next B
    
    Print #FF, Left(Temp, Len(Temp) - 1)
    
Next A

Print #FF, Tempo

Temp = ""
For A = 0 To 7
    Temp = Temp & CurrentInst(A) & ","
Next A

Print #FF, Left(Temp, Len(Temp) - 1)

Close #FF

CurrentlyOpenFile = FilePath
'MsgBox "Saved " & FilePath

Me.Caption = "Music Maker - " & FilePath

End Sub

Private Sub OpenFile()
Dim V As Variant, W As Variant

ClearGrids

'load file
FF = FreeFile

Open FilePath For Input As #FF

For A = 0 To 7
    Line Input #FF, Temp
    V = Split(Temp, ",")
    
    For B = 0 To UBound(V)
        Grid(A, B) = Val(V(B))
    Next B
    
    Erase V
    
    Line Input #FF, Temp
    V = Split(Temp, ",")
    
    For B = 0 To UBound(V)
        InstGrid(A, B) = Val(V(B))
    Next B
Next A

If Not EOF(FF) Then
    Input #FF, Tempo
    HScroll2.Value = Tempo
End If

If Not EOF(FF) Then
    Line Input #FF, Temp
    W = Split(Temp, ",")
    For A = 0 To 7
        CurrentInst(A) = Val(W(A))
    Next A
Else
    For A = 0 To 7
        CurrentInst(A) = InstGrid(A, 0)
    Next A
End If

FormLoading = True
Combo1.ListIndex = CurrentInst(CurrentTrack)
FormLoading = False

Close #FF

StartX = 0
EndX = XSize
ColumnX = 0
DrawGrid

Me.Caption = "Music Maker - " & FilePath

End Sub

Private Sub mUseAlphaBlending_Click()

If mUseAlphaBlending.Checked Then
    mUseAlphaBlending.Checked = False
    AlphaBlendOn = False
    TempInt = 0
Else
    mUseAlphaBlending.Checked = True
    AlphaBlendOn = True
    TempInt = 1
End If

FF = FreeFile
Open ProgramPath & "MusicMaker.dat" For Output As #FF
Print #FF, TempInt
Close #FF

DrawGrid

End Sub

Private Sub MusicBox_KeyDown(KeyCode As Integer, Shift As Integer)

'if in drawmode and Ctrl key is held down
'If CursorType = DrawCursor And Shift = 2 And TempColumnCursor = False Then
'    TempColumnCursor = True
'    CursorType = ColumnCursor
'    MusicBox.MousePointer = 2
'End If

End Sub


Private Sub MusicBox_KeyUp(KeyCode As Integer, Shift As Integer)

'if in TempColumnCursor mode and Ctrl key is released
'If TempColumnCursor = True And Shift = 0 Then
'    TempColumnCursor = False
'    CursorType = DrawCursor
'    MusicBox.MousePointer = 0
'End If

End Sub

Private Sub MusicBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If PlayingSong Then Exit Sub

'prevents note from being drawn in MouseMove and MouseUp when a file was
'opened by a double-click and cursor was over MusicBox when button was
'released (MouseUp activated with no MouseDown)
MouseIsDown = True

GridX = x \ 16
GridY = y \ 16

Select Case CursorType

'*** Draw Mode ***
Case Is = DrawCursor

    'New note
    If Button = 1 Then

        'play notes
        channel = CurrentTrack
        If TrackHide(CurrentTrack) = 0 Then StartNote StartY + GridY 'play note outside startnotes sub because note hasn't been set yet
        NotePlayX = GridX 'keep current x coordinate
        StartNotes
    
        'draw square
        GetRow (GridY)
        MusicBox.Line (GridX * 16 + 1, GridY * 16 + 1)-(GridX * 16 + 15, GridY * 16 + 15), BGColr, BF 'draw white erasing square
        MusicBox.Line (GridX * 16 + 2, GridY * 16 + 2)-(GridX * 16 + 14, GridY * 16 + 14), TrackColr, B 'draw colored hollow square
        OldGridX = GridX 'record starting point for new note

    End If

    'Erase note
    If Button = 2 Then
        If Grid(CurrentTrack, StartX + GridX) = -1 Then Exit Sub 'if trying to delete blank spot
        Do While Grid(CurrentTrack, StartX + GridX) >= 1000 'find start of note
            GridX = GridX - 1
        Loop
        Grid(CurrentTrack, GridX + StartX) = -1 'erase start of note
        InstGrid(CurrentTrack, GridX + StartX) = 0 'erase instrument
    
        GridX = GridX + 1
    
        Do While Grid(CurrentTrack, GridX + StartX) >= 1000 'find all continuations of note
            Grid(CurrentTrack, GridX + StartX) = -1 'erase continuation of note
            InstGrid(CurrentTrack, GridX + StartX) = 0 'erase instrument
            GridX = GridX + 1
        Loop
    
        DrawGrid
    End If


'*** Red Column ****
Case Is = ColumnCursor

    ColumnX = StartX + GridX
    DrawGrid


'*** Select Mode ***
Case Is = SelectCursor
    
    If Button = 1 Then
    
        If SelectStartX = -1 Or StartX + GridX < SelectStartX Then 'if no area selected or clicked behind selected area
        
            'start selecting new area
            
            'reset SelectStartX and SelectEndX
            SelectStartX = -1
            SelectEndX = -1
            
            'reset MusicBox window and get background
            DrawGrid
            GetBox
        
            'draw rectangle
            MusicBox.Line (GridX * 16, 0)-(GridX * 16 + 16, MusicBox.Height - 1), vbBlue, B
            OldGridX = GridX 'record starting point for selected area

            'record selected area start position
            SelectStartX = StartX + GridX
        
        Else
        
            'keep SelectStartX but get new SelectEndX
            
            'reset MusicBox window and get background
            DrawGrid
            GetBox
        
            'draw rectangle
            MusicBox.Line (GridX * 16, 0)-(GridX * 16 + 16, MusicBox.Height - 1), vbBlue, B
            OldGridX = GridX 'record starting point for selected area

        End If
        
        
    End If
        
    'unselect
    If Button = 2 Then
    
        'erase any previously selected area
        SelectStartX = -1
        SelectEndX = -1
        
        DrawGrid

    End If

End Select

End Sub

Private Sub MusicBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If PlayingSong Then
    If Not (Button = 0 And CursorType = DrawCursor) Then Exit Sub
End If

'prevents note from being drawn in MouseMove and MouseUp when a file was
'opened by a double-click and cursor was over MusicBox when button was
'released (MouseUp activated with no MouseDown)
If MouseIsDown = False And Button > 0 Then Exit Sub

Select Case CursorType

'*** Draw Mode ***
Case Is = DrawCursor

    'display instrument of current track under cursor
    If Button = 0 Then

        HoverX = x \ 16
    
        If HoverX = OldHoverX Then Exit Sub
    
        If Grid(CurrentTrack, StartX + HoverX) > -1 Then
            InstLabel.Caption = Combo1.List(InstGrid(CurrentTrack, StartX + HoverX))
        Else
            InstLabel.Caption = ""
        End If
        
        OldHoverX = HoverX
    End If
    
    'continue drawing note
    If Button = 1 Then

        MovingGridX = x \ 16

        If MovingGridX = OldMovingGridX Then Exit Sub 'if cursor hasn't moved to another square
        OldMovingGridX = MovingGridX 'remember cursor position
    
        If MovingGridX > XSize Then MovingGridX = XSize 'prevent from going past edge of screen

        If MovingGridX < GridX Then MovingGridX = GridX 'if cursor is behind starting point

        PutRow (GridY)

        MusicBox.Line (GridX * 16 + 1, GridY * 16 + 1)-(MovingGridX * 16 + 14, GridY * 16 + 14), BGColr, BF
        MusicBox.Line (GridX * 16 + 2, GridY * 16 + 2)-(MovingGridX * 16 + 14, GridY * 16 + 14), TrackColr, B
        
    End If
    
    
'*** Select Mode ***
Case Is = SelectCursor

    'continue selecting area
    If Button = 1 Then

        MovingGridX = x \ 16

        If MovingGridX = OldMovingGridX Then Exit Sub 'if cursor hasn't moved to another column
        OldMovingGridX = MovingGridX 'remember cursor position
    
        If MovingGridX > XSize Then MovingGridX = XSize 'prevent from going past edge of screen

        If MovingGridX < GridX Then MovingGridX = GridX 'if cursor is behind starting point

        PutBox 'redraw entire MusicBox window from SwapScreen
        
        MusicBox.Line (GridX * 16, 0)-(MovingGridX * 16 + 16, MusicBox.Height - 1), vbBlue, B
    
    End If

End Select


End Sub

Private Sub MusicBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If PlayingSong Then Exit Sub

'prevents note from being drawn in MouseMove and MouseUp when a file was
'opened by a double-click and cursor was over MusicBox when button was
'released (MouseUp activated with no MouseDown)
If MouseIsDown Then
    MouseIsDown = False
Else
    Exit Sub
End If

Select Case CursorType

'*** Draw Mode ***
Case Is = DrawCursor

    GridX = x \ 16
    
    'finish drawing note
    If Button = 1 Then

        If GridX > XSize Then GridX = XSize 'prevent from going past edge of screen

        channel = CurrentTrack
        If TrackHide(CurrentTrack) = 0 Then StopNote StartY + GridY 'stop playing note that's being placed
        StopNotes

        If GridX < OldGridX Then GridX = OldGridX 'if cursor is behind note's starting point

        'Record to Grid
        Grid(CurrentTrack, StartX + OldGridX) = StartY + GridY 'record starting note
        InstGrid(CurrentTrack, StartX + OldGridX) = CurrentInst(CurrentTrack) 'record starting note's instrument
    
        If GridX > OldGridX Then
            For A = OldGridX + 1 To GridX
                Grid(CurrentTrack, StartX + A) = 1000 + StartY + GridY 'record continuing notes
                InstGrid(CurrentTrack, StartX + A) = CurrentInst(CurrentTrack) 'record continuing notes' instruments
            Next A
        End If
    
        'Check for partial note after new note
        If Grid(CurrentTrack, StartX + GridX + 1) >= 1000 Then
            Grid(CurrentTrack, StartX + GridX + 1) = Grid(CurrentTrack, StartX + GridX + 1) - 1000 'change it to starting note
        End If
    
        DrawGrid

    End If


'*** Select Mode ***
Case Is = SelectCursor

    'finish selecting area
    If Button = 1 Then
    
        GridX = x \ 16
        
        If GridX > XSize Then GridX = XSize 'prevent from going past edge of screen

        If GridX < OldGridX Then GridX = OldGridX 'if cursor is behind note's starting point
    
        SelectEndX = StartX + GridX
        DrawGrid

    End If
    
    
End Select

End Sub


Private Sub DrawGrid()
Dim CX%, CY%, CX2%
Dim NoteNum%, NoteLength%
Dim BB%, CC%

MusicBox.Cls

'Draw sharp (#) rows different color
For A = 0 To YSize
    If Sharps(A) Then
        MusicBox.Line (0, A * 16)-(MusicBox.Width - 1, A * 16 + 15), &HE0E0E0, BF
    End If
Next A

'Draw columns on MusicBox
For A = 0 To MusicBox.Width - 1 Step 16
    If (StartX + A \ 16) Mod 4 = 0 Then
        Colr = &HFFCC44
    'Else
    '    Colr = &HD0D0D0
    'End If
        MusicBox.Line (A, 0)-(A, MusicBox.Height), Colr
    End If
Next A

'Draw notes
For A = 0 To 7
    E = A
    If A >= CurrentTrack Then E = A + 1 'make sure to draw all other tracks before CurrentTrack
    If A = 7 Then E = CurrentTrack 'CurrentTrack is drawn last so it's on the top
    
    If TrackHide(E) = 0 Then 'if track is hidden then don't draw it
        If A < 7 Then 'if not drawing top layer then make it dimmed
            Colr = DimmedColrs(E)
        Else
            Colr = TrackColrs(E)
        End If
        
'new note drawing method
        
        B = StartX - 1
        
        'StartX to EndX loop
        Do
            B = B + 1

            'see if there's a note
            C = Grid(E, B)
            If C > -1 Then
                                        
                NoteNum = C Mod 1000
    
                NoteLength = B
                'Loop to find length of note
                Do
                    NoteLength = NoteLength + 1
                    
                    If Grid(E, NoteLength) < 1000 Or NoteLength > EndX Then
                        NoteLength = NoteLength - 1
                        CX = (B - StartX) * 16
                        CY = (NoteNum - StartY) * 16
                        
                        If NoteLength = B Then
                            CX2 = CX
                        Else
                             CX2 = CX + (NoteLength - B) * 16
                        End If
                        
                        MusicBox.Line (CX + 2, CY + 2)-(CX2 + 14, CY + 14), Colr, BF
                        B = NoteLength
                        Exit Do
                    End If

                Loop While NoteLength <= EndX

            End If

        Loop While B <= EndX
        
'old note drawing method
'
'        For B = StartX To StartX + XSize
'            C = Grid(E, B)
'            If C >= StartY And C <= EndY Then 'draw only notes in viewable area (Y range)
'                BB = (B - StartX) * 16
'                CC = (C - StartY) * 16
'                MusicBox.Line (BB + 2, CC + 2)-(BB + 14, CC + 14), Colr, BF
'            End If
'
'            If C >= 1000 Then
'                D = C - 1000
'                If D >= StartY And D <= EndY Then
'                    F = (B - StartX) * 16
'                    G = (D - StartY) * 16
'                    MusicBox.Line (F - 2, G + 2)-(F + 14, G + 14), Colr, BF
'                End If
'            End If
'        Next B

    End If
Next A

'draw selected area if visible
If SelectStartX > -1 Then 'there is a selected area
    If Not (SelectStartX > EndX Or SelectEndX < StartX) Then 'at least partially visible
    
        If SelectStartX < StartX Then 'start is off screen
            VisSelStartX = 0
        Else
            VisSelStartX = (SelectStartX - StartX) * 16
        End If
        
        If SelectEndX > EndX Then 'end is off screen
            VisSelEndX = MusicBox.Width - 1
        Else
            VisSelEndX = (SelectEndX - StartX) * 16 + 16
        End If
        
        'draw blue box
        If AlphaBlendOn Then
            VisSelLengthX = VisSelEndX - VisSelStartX + 1
            AlphaBlend MusicBox.hDC, VisSelStartX, 0, VisSelLengthX, MusicBox.Height, BlueBox.hDC, VisSelStartX, 0, VisSelLengthX, MusicBox.Height, nBlend
        Else
            MusicBox.FillStyle = 7
            MusicBox.FillColor = vbBlue
            MusicBox.Line (VisSelStartX, 0)-(VisSelEndX, MusicBox.Height - 1), vbBlue, B
            MusicBox.FillStyle = 1
        End If
        
    End If
End If

'draw start position column if visible
If ColumnX >= StartX And ColumnX <= StartX + XSize Then
    A = (ColumnX - StartX) * 16
    
    If AlphaBlendOn Then
        AlphaBlend MusicBox.hDC, A, 0, 16, MusicBox.Height, RedBox.hDC, 0, 0, 16, MusicBox.Height, nBlend
    Else
        MusicBox.FillStyle = 7
        MusicBox.FillColor = vbRed
        MusicBox.Line (A, 0)-(A + 16, MusicBox.Height - 1), vbRed, B
    End If
    
End If

MusicBox.FillStyle = 1

End Sub

Sub GetRow(y%)
Dim RowY As Long

RowY = y * 16

BitBlt SwapScreen.hDC, 0, 0, MusicBox.Width, 15, MusicBox.hDC, 0, RowY, vbSrcCopy

End Sub

Sub PutRow(y%)
Dim RowY As Long

RowY = y * 16

BitBlt MusicBox.hDC, 0, RowY, MusicBox.Width, 15, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub

Sub GetColumn(x%)
Dim ColX As Long

ColX = x * 16

BitBlt SwapScreen.hDC, 0, 0, 17, MusicBox.Height, MusicBox.hDC, ColX, 0, vbSrcCopy

End Sub

Sub PutColumn(x%)
Dim ColX As Long

ColX = x * 16

BitBlt MusicBox.hDC, ColX, 0, 17, MusicBox.Height, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub

Sub GetBox()

BitBlt SwapScreen.hDC, 0, 0, MusicBox.Width, MusicBox.Height, MusicBox.hDC, 0, 0, vbSrcCopy

End Sub

Sub PutBox()

BitBlt MusicBox.hDC, 0, 0, MusicBox.Width, MusicBox.Height, SwapScreen.hDC, 0, 0, vbSrcCopy

End Sub

Private Sub PasteBut_Click()
Dim PasteStartX%, PasteEndX%, After%
MusicBox.SetFocus

PasteStartX = ColumnX
PasteEndX = PasteStartX + CopyGridLen

If PasteEndX > 1000 Then PasteEndX = 1000

For A = PasteStartX To PasteEndX
    For B = 0 To 7
        If TrackHide(B) = 0 Then
            Grid(B, A) = CopyGrid(B, A - PasteStartX)
            InstGrid(B, A) = CopyInstGrid(B, A - PasteStartX)
        End If
    Next B
Next A

'change cut off notes after pasted area to starting notes
If PasteEndX < 1000 Then
    After = PasteEndX + 1
    For B = 0 To 7
        If TrackHide(B) = 0 Then
            If Grid(B, After) >= 1000 Then
                Grid(B, After) = Grid(B, After) - 1000
            End If
        End If
    Next B
End If
            
DrawGrid

End Sub

Private Sub PasteIntoBut_Click()
MusicBox.SetFocus

InsertSpaces CopyGridLen + 1

PasteBut_Click

End Sub

Private Sub InsertColumnBut_Click()
MusicBox.SetFocus

InsertSpaces 1
DrawGrid

End Sub

Private Sub InsertSpaces(NumOfSpaces%)
Dim DestStartX%

DestStartX = ColumnX + NumOfSpaces

'move everything from ColumnX to 1000 over NumOfSpaces spaces
For A = 1000 To DestStartX Step -1
    For B = 0 To 7
        If TrackHide(B) = 0 Then
            Grid(B, A) = Grid(B, A - NumOfSpaces)
            InstGrid(B, A) = InstGrid(B, A - NumOfSpaces)
        End If
    Next B
Next A

'erase grids where the spaces were put
For A = ColumnX To ColumnX + NumOfSpaces - 1
    For B = 0 To 7
        If TrackHide(B) = 0 Then
            Grid(B, A) = -1
            InstGrid(B, A) = 0
        End If
    Next B
Next A

'change start of severed notes from continuing notes to starting notes
For A = 0 To 7
    If Grid(A, DestStartX) >= 1000 And TrackHide(A) = 0 Then
        Grid(A, DestStartX) = Grid(A, DestStartX) - 1000
    End If
Next A

End Sub

Private Sub PlayAndStop_Click(Index As Integer)
MusicBox.SetFocus

If Index = 0 Then
    PlayingSong = True
    DisableControls
    PlaySong
Else
    EnableControls
    PlayingSong = False
End If

End Sub

Private Sub PlayAndStop_GotFocus(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub RewindBut_Click()

MusicBox.SetFocus

StartX = 0
EndX = XSize
ColumnX = 0
DrawGrid

HScroll1.Value = XSizeMid

End Sub


Private Sub TopBarOverlay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If PlayingSong Then Exit Sub

Dim TopBarX%

TopBarX = ScaleX(x, vbTwips, vbPixels) \ 16

ColumnX = StartX + TopBarX
DrawGrid


End Sub

Private Sub TrackHide_Click(Index As Integer)
MusicBox.SetFocus

If TrackHide(Index).Value = 1 Then
    TrackSel(Index).Enabled = False
Else
    TrackSel(Index).Enabled = True
End If

DrawGrid

End Sub

Private Sub TrackSel_Click(Index As Integer)
MusicBox.SetFocus

'enable all Hide Track buttons
For A = 0 To 7
    TrackHide(A).Enabled = True
Next A

'disable Hide Track button that corresponds to selected track
TrackHide(Index).Enabled = False

CurrentTrack = Index

channel = CurrentTrack
ChangeInstrument CurrentInst(CurrentTrack)
Combo1.ListIndex = CurrentInst(CurrentTrack)

TrackColr = TrackColrs(Index)

DrawGrid
End Sub

Private Sub StartNote(Index As Integer)

If IsDrum(channel) Then
    Flip = DrumNum(channel)
    TempChannel = 9
Else
    Flip = 127 - Index 'notes recorded on grid are 127 - midi number
    TempChannel = channel
End If

midimsg = &H90 + ((Flip) * &H100) + (volume * &H10000) + TempChannel
midiOutShortMsg hmidi, midimsg

End Sub

Private Sub StopNote(Index As Integer)

If IsDrum(channel) Then
    Flip = DrumNum(channel)
    TempChannel = 9
Else
    Flip = 127 - Index 'notes recorded on grid are 127 - midi number
    TempChannel = channel
End If
   
midimsg = &H80 + ((Flip) * &H100) + TempChannel
midiOutShortMsg hmidi, midimsg
   
End Sub

Private Sub StartNotes()

For A = 0 To 7
    If TrackHide(A) = 0 And CurrentTrack <> A Then 'if not muted and not current track
        
        B = (Grid(A, (StartX + NotePlayX))) Mod 1000
        
        If B > 0 Then
            channel = A
            ChangeInstrument InstGrid(A, StartX + NotePlayX) 'change channel's instrument for note being played
            StartNote B
        End If
    End If
Next A

End Sub
Private Sub StopNotes()

For A = 0 To 7
    If TrackHide(A) = 0 And CurrentTrack <> A Then 'if not muted and not current track
    
        B = (Grid(A, (StartX + NotePlayX))) Mod 1000
        
        If B > 0 Then
            channel = A
            ChangeInstrument InstGrid(A, StartX + NotePlayX) 'change channel's instrument for note being played
            StopNote B
        End If
    End If
Next A

End Sub

Private Sub TrackSel_GotFocus(Index As Integer)
MusicBox.SetFocus

End Sub

Private Sub TransposeBut_Click(Index As Integer)
Dim TransStartX%, TransEndX%

MusicBox.SetFocus

'if something is selected, transpose only those notes
'otherwise transpose all notes
If SelectStartX > -1 Then
    TransStartX = SelectStartX
    TransEndX = SelectEndX
Else
    TransStartX = 0
    TransEndX = 1000
End If

'cliked + or - ?
If Index = 0 Then
    C = 1
Else
    C = -1
End If

'transpose notes
For A = 0 To 7
    If TrackHide(A) = 0 Then
        For B = TransStartX To TransEndX
            If Grid(A, B) > -1 Then
                D = (Grid(A, B) + C)
                
                If D = -1 Then D = 127
                If D = 128 Then D = 0
                If D = 999 Then D = 1127
                If D = 1128 Then D = 1000
                
                Grid(A, B) = D
            End If
        Next B
    End If
Next A

DrawGrid

End Sub


Private Sub VScroll1_Change()

If Resizing Then Exit Sub

StartY = VScroll1.Value - YSizeMid
EndY = StartY + YSize
PrintSideBar
DrawGrid
JustScrolled = True 'for when user scrolls window while song is playing

End Sub

Private Sub PrintSideBar()
'Print note names on sidebar
Temp = ""
B = StartY
For A = 0 To YSize
    Temp = Temp & NoteNames(B) & vbCrLf
    If Right(NoteNames(B), 1) = "#" Then 'Find all sharps
        Sharps(A) = True
    Else
        Sharps(A) = False
    End If
    B = B + 1
Next A
SideBar.Caption = Temp
End Sub

Private Sub PrintTopBar()

TopBarOffset = 40

'Print column numbers on topbar
B = 0
For A = StartX To StartX + XSize
    If A Mod 4 = 0 Then
        TopBar(B).Visible = True
        TopBar(B).Left = (A - StartX) * 16 + TopBarOffset
        TopBar(B).Caption = A \ 4
        B = B + 1
    End If
Next A

'make unused labels invisible
For A = B To 18
    TopBar(A).Visible = False
Next A

End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub PlaySong()

'if red starting column is not on the screen then move screen
If ColumnX < StartX Or ColumnX > StartX + XSize Then
    StartX = ColumnX
    EndX = StartX + XSize
    If StartX > 999 - XSize Then
        StartX = 999 - XSize 'startx can't be higher than 999 - xsize
        EndX = 999
    End If
End If

PlayX = ColumnX 'start playing from red column
GetColumn StartX
HScroll1.Value = StartX + XSizeMid
HasStarted = False
JustScrolled = False

DrawGrid

'reset oldnote and oldinst variables
For A = 0 To 7
    OldNote(A) = -1
    OldInst(A) = -1
Next A

FindEndOfSong 'so we know when to stop playing

Tick = GetTickCount 'get curent time in milliseconds (I hate the built in VB Timer control!)


'*** Start of song playing loop ***

Do While PlayingSong

    'test how much time left before next note after all routines are done
    'If PlayBarX = 0 Then
    '   InstLabel.Caption = "Extra time = " & (Tick + TempoWait) - GetTickCount & " / " & TempoWait
    'End If

    'loop until 16th note pause has elapsed
    Do While GetTickCount < Tick + TempoWait
        DoEvents
    Loop
    
    Tick = GetTickCount

    PlayNotes
    
    DrawGraphics
    
Loop

'wait another 16th to finish playing last note
Do While GetTickCount < Tick + TempoWait
    DoEvents
Loop

'stop all notes
For A = 0 To 7
    If TrackHide(A) = 0 Then
        channel = A
        If OldNote(A) > -1 Then
            StopNote OldNote(A) Mod 1000
        End If
    End If
Next A

DrawGrid

End Sub

Private Sub PlayNotes()


'Play notes
For Z = 0 To 7
    
    If TrackHide(Z) = 0 Then
    
        channel = Z
        
        'find note types
        If OldNote(Z) = -1 Then OldNoteType = BlankNote
        If OldNote(Z) >= 0 And OldNote(Z) < 1000 Then OldNoteType = StartingNote
        If OldNote(Z) >= 1000 Then OldNoteType = ContinuingNote
        If Grid(Z, PlayX) = -1 Then NoteType = BlankNote
        If Grid(Z, PlayX) >= 0 And Grid(Z, PlayX) < 1000 Then NoteType = StartingNote
        If Grid(Z, PlayX) >= 1000 Then NoteType = ContinuingNote
            
        'change instrument if needed
        'if new note is being played and instrument differs from old instrument
        If NoteType = StartingNote And InstGrid(Z, PlayX) <> OldInst(Z) Then
            ChangeInstrument InstGrid(Z, PlayX)
        End If
        
        'stop old note and start new one
        If OldNoteType <> BlankNote And NoteType = StartingNote Then
            StopNote OldNote(Z) Mod 1000
            StartNote Grid(Z, PlayX) Mod 1000
        End If
        
        'stop old note but don't start new one
        If OldNoteType <> BlankNote And NoteType = BlankNote Then
            StopNote OldNote(Z) Mod 1000
        End If
        
        'start new note but don't stop old
        If OldNoteType = BlankNote And NoteType = StartingNote Then
            StartNote Grid(Z, PlayX) Mod 1000
        End If
        
        'set oldnote to current note's value
        OldNote(Z) = Grid(Z, PlayX)
            
        'set oldinst to current instrument's value if note was started
        If NoteType = StartingNote Then
            OldInst(Z) = InstGrid(Z, PlayX)
        End If
        
    End If
        
Next Z

End Sub

Private Sub DrawGraphics()

'replace background that position bar was drawn over
If PlayBarX < XSize And HasStarted And (Not JustScrolled) Then
    PutColumn PlayBarX
Else
    HasStarted = True
End If
    
JustScrolled = False
    
'show playing position bar
PlayBarX = PlayX - StartX
    
If PlayBarX > XSize Then 'if play bar scrolled off screen move screen over one page
    StartX = PlayX 'move over one page
    EndX = StartX + XSize
    If StartX > 999 - XSize Then
        StartX = 999 - XSize 'prevent showing past end of grid
        EndX = 999
    End If
    PlayBarX = 0 'playbar starts at the left again
    TurningPage = True 'prevents hscroll1_change from calling drawgrid sub
    HScroll1.Value = StartX + XSizeMid 'update horizontal scroll bar
    TurningPage = False 'allow hscroll1_change to call drawgrid sub again
    DrawGrid
    JustScrolled = False 'prevent skipping redraw of leftmost column on next loop
End If
    
GetColumn PlayBarX 'get background which will be behind playbar
        
'draw green bar
E = PlayBarX * 16

If AlphaBlendOn Then
    AlphaBlend MusicBox.hDC, E, 0, 16, MusicBox.Height, GreenBox.hDC, 0, 0, 16, MusicBox.Height, nBlend
    MusicBox.Refresh
Else
    MusicBox.FillStyle = 7
    MusicBox.FillColor = vbGreen
    MusicBox.Line (E, 0)-(E + 16, MusicBox.Height - 1), vbGreen, B
    MusicBox.FillStyle = 1
End If

'increase playx (move forward one note)
PlayX = PlayX + 1
    
If PlayX > SongLength Then
    PlayAndStop(1).Value = True
    PlayingSong = False
End If

End Sub

Private Sub ChangeInstrument(Inst As Integer)

If Inst < 128 Then
    'melody instrument
    midiOutShortMsg hmidi, &HB0 + channel
    midiOutShortMsg hmidi, 32 * &H100 + &HB0 + channel
    midiOutShortMsg hmidi, Inst * &H100 + &HC0 + channel
    IsDrum(channel) = False
Else
    'percussion instrument
    IsDrum(channel) = True
    DrumNum(channel) = Inst - 93
End If



End Sub

Private Sub FindEndOfSong()

SongLength = 0

For A = 0 To 7
    If TrackHide(A) = 0 Then
        For B = 1000 To 0 Step -1
            If Grid(A, B) > -1 Then
                If B > SongLength Then SongLength = B
                Exit For
            End If
        Next B
    End If
Next A
                  
End Sub

Private Sub ClearGrids()

'clear grids
For A = 0 To 7
    For B = 0 To 1000
        Grid(A, B) = -1
        InstGrid(A, B) = 0
    Next B
Next A

End Sub

Private Sub DisableControls()

FunctionsFrame.Enabled = False
EditTrackFrame.Enabled = False
MuteTrackFrame.Enabled = False
TransposeFrame.Enabled = False
Combo1.Enabled = False
RewindBut.Enabled = False

mNew.Enabled = False
mOpen.Enabled = False
mSave.Enabled = False
mSaveAs.Enabled = False
mExporttoMidi.Enabled = False

mFlipNotes.Enabled = False

End Sub


Private Sub EnableControls()

FunctionsFrame.Enabled = True
EditTrackFrame.Enabled = True
MuteTrackFrame.Enabled = True
TransposeFrame.Enabled = True
Combo1.Enabled = True
RewindBut.Enabled = True

mNew.Enabled = True
mOpen.Enabled = True
mSave.Enabled = True
mSaveAs.Enabled = True
mExporttoMidi.Enabled = True

mFlipNotes.Enabled = True

End Sub

Private Sub MidiFile()

'*** MIDI File Header Chunk (14 bytes) ***

'MThd
Bytes(1) = Asc("M")
Bytes(2) = Asc("T")
Bytes(3) = Asc("h")
Bytes(4) = Asc("d")

'Length of header to follow (always 6 bytes)
Bytes(5) = 0
Bytes(6) = 0
Bytes(7) = 0
Bytes(8) = 6

'6 byte header
Bytes(9) = 0
Bytes(10) = 1   '0 - single-track
                '1 - multiple tracks, synchronous
                '2 - multiple tracks, asynchronous

Bytes(11) = 0
Bytes(12) = 9   'number of tracks

Bytes(13) = 0
Bytes(14) = 48  'number of delta-time ticks per quarter note
                '12 ticks per 16th note


'*** Track 1 ***
'MTrk
Bytes(15) = Asc("M")
Bytes(16) = Asc("T")
Bytes(17) = Asc("r")
Bytes(18) = Asc("k")

'Number of bytes in track data to follow
Bytes(19) = 0
Bytes(20) = 0
Bytes(21) = 0
Bytes(22) = 19

'Set Tempo
Bytes(23) = 0 ' delta time
Bytes(24) = 255 'FF (Meta command)
Bytes(25) = 81 'tempo command
Bytes(26) = 3 '3 bytes to describe tempo

ConvertTempo 'this writes the 3 tempo bytes

'Set Time Signature
Bytes(30) = 0 'delta time
Bytes(31) = 255 'FF (Meta command)
Bytes(32) = 88 'Time Signature command
Bytes(33) = 4 '4 bytes to follow
Bytes(34) = 4 'numerator
Bytes(35) = 2 'denominator (2 = 1/4, 3 = 1/8, 4 = 1/16, etc.)
Bytes(36) = 24 '# of midi clocks in metronome tick
Bytes(37) = 8 '# of 32nd notes in a quarter note

'end of track
Bytes(38) = 0 'delta time
Bytes(39) = 255 'End track comand (3 bytes)
Bytes(40) = 47
Bytes(41) = 0


'***** Start Writing Tracks *****
CurrentByte = 41
FindEndOfSong
SongLength = SongLength + 1

'track loop
For Z = 0 To 7

IncByte

OldNote(Z) = -1
OldInst(Z) = -1
LastDeltaTime = 0
TotalBytes = 0

'MTrk
Bytes(CurrentByte) = Asc("M"): IncByte
Bytes(CurrentByte) = Asc("T"): IncByte
Bytes(CurrentByte) = Asc("r"): IncByte
Bytes(CurrentByte) = Asc("k"): IncByte

'Increase CurrentByte by 4 for track size to be filled in after
'track is written
CurrentByte = CurrentByte + 4

'This will be subtracted from A after writing the track to
'determine the number of bytes of data the track consists of
TrackDataStartByte = CurrentByte


If TrackHide(Z) = 0 Then 'skip note counting loop if track is hidden

    'Note Counting Loop
    For A = 0 To SongLength
        
        'find note types
        If OldNote(Z) = -1 Then OldNoteType = BlankNote
        If OldNote(Z) >= 0 And OldNote(Z) < 1000 Then OldNoteType = StartingNote
        If OldNote(Z) >= 1000 Then OldNoteType = ContinuingNote
        If A = SongLength Then
            NoteType = BlankNote
        Else
            If Grid(Z, A) = -1 Then NoteType = BlankNote
            If Grid(Z, A) >= 0 And Grid(Z, A) < 1000 Then NoteType = StartingNote
            If Grid(Z, A) >= 1000 Then NoteType = ContinuingNote
        End If
        
        
        'change instrument if needed
        'if new note is being played and instrument differs from old instrument
        MidiInst = (InstGrid(Z, A))
        If NoteType = StartingNote And MidiInst <> OldInst(Z) Then
            'Choose instrument
            If MidiInst < 128 Then
                'melody instrument
                DeltaTime = A - LastDeltaTime
                ConvertDeltaTime
                Bytes(CurrentByte) = 192 + Z 'choose instrument command + channel #
                    IncByte
                Bytes(CurrentByte) = MidiInst 'instrument #
                    IncByte
                LastDeltaTime = A
                MidiIsDrum = False
            Else
                'precussion instrument
                MidiIsDrum = True
                MidiDrumNum = MidiInst - 93
            End If
        End If
    
        'stop old note and start new one
        If OldNoteType <> BlankNote And NoteType = StartingNote Then
            'stop old note
            DeltaTime = A - LastDeltaTime
            ConvertDeltaTime
            MidiFileStopNote
            'start new note
            Bytes(CurrentByte) = 0 'delta time
                IncByte
            MidiFileStartNote
            LastDeltaTime = A
        End If
            
        'stop old note but don't start new one
        If OldNoteType <> BlankNote And NoteType = BlankNote Then
            DeltaTime = A - LastDeltaTime
            ConvertDeltaTime
            MidiFileStopNote
            LastDeltaTime = A
        End If
            
        'start new note but don't stop old
        If OldNoteType = BlankNote And NoteType = StartingNote Then
            DeltaTime = A - LastDeltaTime
            ConvertDeltaTime
            MidiFileStartNote
            LastDeltaTime = A
        End If
            
        'set oldnote to current note's value
        OldNote(Z) = Grid(Z, A)
                
        'set oldinst to current instrument's value if note was started
        If NoteType = StartingNote Then
            OldInst(Z) = InstGrid(Z, A)
        End If
    
    Next A

End If


'end of track
DeltaTime = SongLength - LastDeltaTime 'delta time
ConvertDeltaTime
Bytes(CurrentByte) = 255
    IncByte
Bytes(CurrentByte) = 47
    IncByte
Bytes(CurrentByte) = 0

TotalBytes = CurrentByte - TrackDataStartByte + 1
ConvertTrackLength


Next Z

'write file
If Dir(FilePath) <> "" Then Kill FilePath

FF = FreeFile

Open FilePath For Binary Access Write As #FF

For A = 1 To CurrentByte
    Put #FF, A, Bytes(A)
Next A

Close #FF

End Sub


Private Sub ConvertDeltaTime()
Dim Power%, DB As Long
Dim DA%, DC%, DD%

'convert DeltaTime to binary
DB = DeltaTime * 12 '12 delta ticks per 16th note
Temp = ConvertToBinary(DB)
    
'make length of Temp be a multiple of 7
DA = Len(Temp) Mod 7
If DA <> 0 Then
    Temp = String(7 - DA, "0") & Temp
End If

'Break into 7-bit bytes and store in DeltaBinArray
DD = Len(Temp) \ 7 - 1
For DA = 0 To DD
    DeltaBinArray(DA) = Mid(Temp, DA * 7 + 1, 7)
    'add most significant bit to 7-bit bytes
    If DA = DD Then
        DeltaBinArray(DA) = "0" & DeltaBinArray(DA)
    Else
        DeltaBinArray(DA) = "1" & DeltaBinArray(DA)
    End If
Next DA

'convert binary strings in DeltaBinArray to bytes in DeltaTimeArray
For DA = 0 To DD
    DB = 0
    For DC = 1 To 8
        If Mid(DeltaBinArray(DA), DC, 1) = "1" Then
            DB = DB + 2 ^ (8 - DC)
        End If
    Next DC
    DeltaTimeArray(DA) = DB
Next DA

'write delta time to Bytes and increase CurrentByte
For DA = 0 To DD
    Bytes(CurrentByte) = DeltaTimeArray(DA)
        IncByte
Next DA

End Sub

Private Sub ConvertTrackLength()

Dim Power%, DB As Long
Dim DA%, DC%
Dim TotalBytesArray(3) As String

'convert to binary
DB = TotalBytes
Temp = ConvertToBinary(DB)
    
'add necessary 0's to make temp 32 characters long
Temp = String(32 - Len(Temp), "0") & Temp

'Break into 8-bit bytes and store in TotalBytesArray
For DA = 0 To 3
    TotalBytesArray(DA) = Mid(Temp, DA * 8 + 1, 8)
Next DA

'convert binary strings in TotalBytesArray to bytes and
'write to Bytes
For DA = 0 To 3
    DB = 0
    For DC = 1 To 8
        If Mid(TotalBytesArray(DA), DC, 1) = "1" Then
            DB = DB + 2 ^ (8 - DC)
        End If
    Next DC
    Bytes(TrackDataStartByte - (4 - DA)) = DB
Next DA

End Sub

Private Sub ConvertTempo()

Dim Power%, DB As Long
Dim DA%, DC%
Dim BytesArray(2) As String
Dim MPQN As Long 'microseconds per quarter note

MPQN = 60000000 \ Tempo

DB = MPQN
Temp = ConvertToBinary(DB)
    
'add necessary 0's to make temp 24 characters long
Temp = String(24 - Len(Temp), "0") & Temp

'Break into 8-bit bytes and store in BytesArray
For DA = 0 To 2
    BytesArray(DA) = Mid(Temp, DA * 8 + 1, 8)
Next DA

'convert binary strings in BytesArray to bytes and
'write to Bytes
For DA = 0 To 2
    DB = 0
    For DC = 1 To 8
        If Mid(BytesArray(DA), DC, 1) = "1" Then
            DB = DB + 2 ^ (8 - DC)
        End If
    Next DC
    Bytes(27 + DA) = DB
Next DA

End Sub

Private Sub IncByte()

CurrentByte = CurrentByte + 1

End Sub


Private Function ConvertToBinary(BInput As Long) As String
Dim Power%, BTemp As String
Dim BNum As Long
Dim BA%, BB As Long

BNum = BInput

'convert to binary
Power = 0
Do
    If BNum < 2 ^ Power Then Exit Do
    Power = Power + 1
Loop

Power = Power - 1

If Power < 0 Then Power = 0

BTemp = ""

For BA = Power To 0 Step -1
    BB = 2 ^ BA
    If BNum >= BB Then
        BTemp = BTemp & "1"
        BNum = BNum - BB
    Else
        BTemp = BTemp & "0"
    End If
Next BA

ConvertToBinary = BTemp

End Function

Private Sub MidiFileStartNote()

If MidiIsDrum Then
    Flip = MidiDrumNum
    TempZ = 9
Else
    Flip = 127 - Grid(Z, A)
   TempZ = Z
End If

Bytes(CurrentByte) = 144 + TempZ 'start note command + channel #
    IncByte
Bytes(CurrentByte) = Flip 'note #
    IncByte
Bytes(CurrentByte) = 127 ' velocity
    IncByte

End Sub

Private Sub MidiFileStopNote()

If MidiIsDrum Then
    Flip = MidiDrumNum
    TempZ = 9
Else
    Flip = 127 - (OldNote(Z) Mod 1000)
    TempZ = Z
End If

Bytes(CurrentByte) = 128 + TempZ 'stop note command + channel #
    IncByte
Bytes(CurrentByte) = Flip 'note #
    IncByte
Bytes(CurrentByte) = 0 ' velocity
    IncByte
            
End Sub
