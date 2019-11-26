VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5028
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5028
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne elsTop 
      Height          =   4248
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7980
      _cx             =   14076
      _cy             =   7493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.Frame fraOne 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   4332
         Left            =   3264
         TabIndex        =   1
         Top             =   0
         Width           =   4692
         Begin VB.Label lblIntroduction 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizard.frx":0000
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   10.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1380
            Left            =   360
            TabIndex        =   7
            Top             =   744
            Width           =   4140
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Renaissance SMS Shooter Import Wizard"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   10.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   408
            TabIndex        =   6
            Top             =   192
            Width           =   3996
         End
      End
      Begin VB.Image imgIntro 
         Height          =   4332
         Left            =   24
         Top             =   24
         Width           =   3204
      End
   End
   Begin SizerOneLibCtl.ElasticOne elsBottom 
      Height          =   780
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4248
      Width           =   7980
      _cx             =   14076
      _cy             =   1376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmWizard.frx":00C4
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   348
         Left            =   5484
         TabIndex        =   4
         Top             =   216
         Width           =   1188
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Enabled         =   0   'False
         Height          =   348
         Left            =   4248
         TabIndex        =   3
         Top             =   216
         Width           =   1188
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   348
         Left            =   6720
         TabIndex        =   5
         Top             =   216
         Width           =   1188
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdNext_Click()
Unload Me
frmSource.Show
End Sub
