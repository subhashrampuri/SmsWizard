VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODL7.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSource 
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
      Begin MSComDlg.CommonDialog DlgFile 
         Left            =   216
         Top             =   1704
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
      Begin VB.Frame fraTwo 
         Height          =   2460
         Left            =   672
         TabIndex        =   10
         Top             =   1584
         Width           =   6924
         Begin VB.TextBox txtsPassword 
            Appearance      =   0  'Flat
            Height          =   324
            IMEMode         =   3  'DISABLE
            Left            =   1584
            PasswordChar    =   "*"
            TabIndex        =   17
            Top             =   1776
            Visible         =   0   'False
            Width           =   2436
         End
         Begin VB.TextBox txtsUserName 
            Appearance      =   0  'Flat
            Height          =   324
            Left            =   1584
            TabIndex        =   15
            Top             =   1272
            Visible         =   0   'False
            Width           =   2436
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "...."
            Height          =   348
            Left            =   5808
            TabIndex        =   14
            Top             =   1248
            Width           =   948
         End
         Begin VB.TextBox txtFileName 
            Appearance      =   0  'Flat
            Height          =   324
            Left            =   1584
            TabIndex        =   13
            Top             =   744
            Width           =   5196
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            Caption         =   "&Password :"
            Height          =   192
            Index           =   5
            Left            =   312
            TabIndex        =   18
            Top             =   1824
            Visible         =   0   'False
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            Caption         =   "&User Name :"
            Height          =   192
            Index           =   4
            Left            =   288
            TabIndex        =   16
            Top             =   1344
            Visible         =   0   'False
            Width           =   936
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            Caption         =   "File Name :"
            Height          =   192
            Index           =   3
            Left            =   288
            TabIndex        =   12
            Top             =   816
            Width           =   828
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            Height          =   192
            Index           =   2
            Left            =   288
            TabIndex        =   11
            Top             =   288
            Width           =   6084
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmOne 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   996
         Left            =   0
         TabIndex        =   5
         Top             =   -24
         Width           =   7980
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "From where do you want to import data?  You can import data from on the following source"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   144
            TabIndex        =   7
            Top             =   576
            Width           =   7488
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDataSource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Choose a Data Source"
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
            Left            =   144
            TabIndex        =   6
            Top             =   168
            Width           =   2172
         End
      End
      Begin TrueOleDBList70.TDBCombo cboSource_lID 
         Height          =   312
         Left            =   2328
         TabIndex        =   9
         Top             =   1200
         Width           =   5256
         _ExtentX        =   9271
         _ExtentY        =   550
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2731"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   0
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   -1  'True
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   5
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   311.811
         AutoSize        =   -1  'True
         GapHeight       =   36.283
         ListField       =   ""
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   -1  'True
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         AddItemSeparator=   ";"
         _PropDict       =   $"frmSource.frx":0000
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=28,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=43"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblSource 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Source Name :"
         Height          =   192
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   1248
         Width           =   1452
      End
   End
   Begin SizerOneLibCtl.ElasticOne elsBottom 
      Height          =   780
      Left            =   0
      TabIndex        =   1
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
      _GridInfo       =   $"frmSource.frx":0087
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   348
         Left            =   5484
         TabIndex        =   3
         Top             =   216
         Width           =   1188
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Enabled         =   0   'False
         Height          =   348
         Left            =   4248
         TabIndex        =   2
         Top             =   216
         Width           =   1188
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   348
         Left            =   6720
         TabIndex        =   4
         Top             =   216
         Width           =   1188
      End
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sType As String
Dim sFileTitle As String
Private ofrmFormat As frmFormat
Attribute ofrmFormat.VB_VarHelpID = -1
Private ofrmExcel As frmExcel
Private Sub cboSource_lID_ItemChange()
Dim sou_lID As Long
    If cboSource_lID.SelectedItem <> "" Then
        sou_lID = cboSource_lID.SelectedItem
        If sou_lID = 0 Then
            sType = "*.xls|*.xls"
            lblSource(2).Caption = "To Connect to Microsoft Excel, you must first choose an excel file"
            fraTwo.Visible = True
            pDisable
        ElseIf sou_lID = 1 Then
            sType = "*.txt|*.txt"
            lblSource(2).Caption = "Text file can be delimited or fixed field. To connect, you must select a file"
            fraTwo.Visible = True
            pDisable
        ElseIf sou_lID = 2 Then
            sType = "*.csv|*.csv"
            lblSource(2).Caption = "CSV file can be delimited or fixed field. To connect, you must select a file"
            fraTwo.Visible = True
            pDisable
        ElseIf sou_lID = 3 Then
            sType = "*.mdb|*.mdb"
            lblSource(2).Caption = "To connect select a database and provide username and password."
            fraTwo.Visible = True
            lblSource(4).Visible = True
            lblSource(5).Visible = True
            txtsUserName.Visible = True
            txtsPassword.Visible = True
        End If
     Else
        sou_lID = ""
        sType = ""
        fraTwo.Visible = False
     End If
End Sub
Private Sub cmdBrowse_Click()
 On Error GoTo LocalErr
    With DlgFile
        .CancelError = True
        .Filter = sType
        .ShowOpen
        txtFileName.Text = .FileName
        sFileTitle = .FileTitle
    End With
LocalErr:
End Sub
Private Sub cmdCancel_Click()
End
End Sub
Private Sub cmdNext_Click()
    If cboSource_lID.SelectedItem = 0 Then
        Set ofrmExcel = New frmExcel
        With ofrmExcel
            .sSource = txtFileName.Text
            .sTitle = sFileTitle
            .Show
        End With
    Else
        With ofrmFormat
            .sSource = txtFileName.Text
            .sTitle = sFileTitle
            .Show
        End With
    End If
End Sub
Private Sub Form_Load()
    pInitializeComponent
End Sub
Private Sub pInitializeComponent()
    Me.Caption = "Choose Data Source"
    Set ofrmFormat = New frmFormat
    cboSource_lID.AddItem "Microsoft Excel 97 - 2000", 0
    cboSource_lID.AddItem "Text File", 1
    cboSource_lID.AddItem "Comma Separated Value (CSV)", 2
    cboSource_lID.AddItem "Microsoft Access", 3
    fraTwo.Visible = False
End Sub
Private Sub pDisable()
    lblSource(4).Visible = False
    lblSource(5).Visible = False
    txtsUserName.Visible = False
    txtsPassword.Visible = False
End Sub
