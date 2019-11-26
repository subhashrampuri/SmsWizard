VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "TODG8.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODL7.OCX"
Begin VB.Form frmSab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5628
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9012
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5628
   ScaleWidth      =   9012
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne elsRight 
      Height          =   5256
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2736
      _cx             =   4826
      _cy             =   9271
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
      BackColor       =   -2147483634
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   3
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
      TagWidth        =   1400
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
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H8000000E&
         Cancel          =   -1  'True
         Caption         =   "Display"
         Height          =   324
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2088
         Width           =   1536
      End
      Begin VB.TextBox txtsab_tDate 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   120
         TabIndex        =   6
         Top             =   1584
         Width           =   2448
      End
      Begin TrueOleDBList70.TDBCombo cbosec_lID 
         Height          =   312
         Left            =   120
         TabIndex        =   4
         Top             =   792
         Width           =   2448
         _ExtentX        =   4318
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
         _PropDict       =   $"frmSab.frx":0000
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Date :"
         Height          =   192
         Left            =   120
         TabIndex        =   5
         Top             =   1248
         Width           =   420
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Grade - Section :"
         Height          =   192
         Left            =   120
         TabIndex        =   3
         Top             =   456
         Width           =   1188
      End
      Begin VB.Label lblHeader 
         BackColor       =   &H8000000D&
         Height          =   252
         Left            =   24
         TabIndex        =   10
         Top             =   0
         Width           =   2676
      End
   End
   Begin SizerOneLibCtl.ElasticOne elsLeft 
      Height          =   5256
      Left            =   2736
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6276
      _cx             =   11070
      _cy             =   9271
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
      BackColor       =   -2147483634
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   4
      AutoSizeChildren=   1
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
      Begin TrueOleDBGrid80.TDBGrid grdList 
         Height          =   5112
         Left            =   72
         TabIndex        =   8
         Top             =   72
         Width           =   6132
         _ExtentX        =   10816
         _ExtentY        =   9017
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "General Number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   508
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3048"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2963"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         AllowArrows     =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=780,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=780,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
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
   End
   Begin SizerOneLibCtl.ElasticOne elsBottom 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5256
      Width           =   9012
      _cx             =   15896
      _cy             =   656
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
      BackColor       =   -2147483634
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      GridRows        =   1
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmSab.frx":0087
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000E&
         Caption         =   "E&xit"
         Height          =   324
         Left            =   8436
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   24
         Width           =   552
      End
      Begin VB.Image imgSmartAnalyzer 
         Height          =   324
         Left            =   24
         Top             =   24
         Width           =   3852
      End
   End
End
Attribute VB_Name = "frmSab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===================================================
' frmSab.frm
' Satish
' Renaissance Technologies Pvt. Ltd.
' Copyright © 2004  (services@renaissanceind.com)
' August 18 2004
' Hungarian Notation
'===================================================
Public Event Change()
Public Event CloseUp(ByVal bChanged As Boolean)
Public Event Loaded()
Private WithEvents m_oSab As cSab
Attribute m_oSab.VB_VarHelpID = -1
Private WithEvents m_oList As cList
Attribute m_oList.VB_VarHelpID = -1
Dim m_bSave As Boolean
Dim m_sec_lID As Variant
Dim m_tProperties As tProperties
Private m_oRes As New cResource
Dim m_xdbAbsentee As XArrayDB
Public Property Set Res(ByRef cRes As cResource)
    Set m_oRes = cRes
End Property
Public Property Get Res() As cResource
    Set Res = m_oRes
End Property
Private Sub pInitalizeComponent()
     Me.Caption = "Attendance"
     Me.Icon = Res.foGetResourcePicture(19, "Custom")
     imgSmartAnalyzer.Picture = Res.foGetResourcePicture(103, "Custom")
    
     Set m_oSab = New cSab
     Set m_oSab.Connection = g_Connection
       
     Set m_oList = New cList
     Set m_oList.Connection = g_Connection
     
     Set m_xdbAbsentee = New XArrayDB
     pClearFields
End Sub
Private Sub pDestroyComponent()
     Set m_oSab = Nothing
     Set m_oList = Nothing
     Set m_xdbAbsentee = Nothing
End Sub
Private Sub cbosec_lID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pEnterAsTab KeyCode
    End If
End Sub
Private Sub cbosec_lID_Validate(Cancel As Boolean)
 On Error Resume Next
    If fsCheckNumberAsString(cbosec_lID.SelectedItem) <> "" Then
        m_sec_lID = cbosec_lID.Columns(1).CellText(cbosec_lID.SelectedItem)
        
     Else
        m_sec_lID = ""
     End If
End Sub
Private Sub cmdDisplay_Click()
    'validates here
    If Not fbCheckFields Then
        'Exit the procedure if fields are not validated.
        Exit Sub
    End If
    pPopulateStudentGrid
End Sub
Private Sub cmdExit_Click()
  Unload Me
End Sub
Private Sub Form_Activate()
    Dim tdbCol As TrueOleDBGrid80.Column
    For Each tdbCol In grdList.Columns
        tdbCol.AutoSize
    Next
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    pEnterAsTab KeyAscii
End Sub
Private Sub Form_Load()
    pInitalizeProperties
    pInitalizeComponent
    pCreateGridColumns
    pInitalizeGridProperties
    pLoadSection
    RaiseEvent Loaded
End Sub
Private Sub pDelete()
'======================================================
' Assigns the field values and updates to the database
' using Datalayer
'======================================================
    Dim vAL As Variant
    Dim vResponse As Variant
    
    'Set the MousePointer to Hourglass for busy.
    Screen.MousePointer = vbHourglass
  
    'Assign values for updation.
    vAL = AssignValues()
    g_Connection.BeginTrans
    'Call datalayers update function.
    If m_oSab.fbDelete(vAL) Then
       'Commit transaction on successfull insertion to
        'the database.
        g_Connection.CommitTrans
        m_bSave = True
        'Set the MousePointer to normal.
        Screen.MousePointer = vbNormal
    Else
         'Rollsback Transaction.
        'Check ErrorGenerated Event of Datalayer object
        'for errors.
        g_Connection.RollbackTrans
        Screen.MousePointer = vbNormal
        MsgBox Res.fsGetResourceString(10272), vbExclamation, App.Title
    End If
    
End Sub
Private Sub pAdd()
'======================================================
' Assigns the field values and inserts to the database
' using Datalayer.
'======================================================
    Dim lPrimarykey As Long
    Dim vResponse As Variant
    Dim vAL As Variant
    
    'Set the MousePointer to Hourglass for busy.
    Screen.MousePointer = vbHourglass
    
    'Assign values for inserting.
    vAL = AssignValues()
        
    'Begin transaction since multiple table gets affected
    'during insertion.
    g_Connection.BeginTrans
    'Call Datalayers update function.
    If m_oSab.fbAddNew(vAL, lPrimarykey) Then
        'Item ID is set to Tag property for immediate
        'updation of data.
        'Commit transaction on successfull insertion to
        'the database.
        m_bSave = True
        g_Connection.CommitTrans
        'Set the MousePointer to normal.
        Screen.MousePointer = vbNormal
    Else
        'Rollsback Transaction.
        'Check ErrorGenerated Event of Datalayer object
        'for errors.
        g_Connection.RollbackTrans
        'Set the MousePointer to normal.
        Screen.MousePointer = vbNormal
        MsgBox Res.fsGetResourceString(10272), vbExclamation, App.Title
    End If
End Sub
Private Function AssignValues() As Variant
'======================================================
' Assigns the field values to variant used during
' insertion & updation.
' Return : Variant
'======================================================
    Dim vAL As Variant
    ReDim vAL(2)
    vAL(0) = grdList.Columns(0).Text
    vAL(1) = Trim(txtsab_tDate.Text)
    vAL(2) = grdList.Columns(1).Text
    
    AssignValues = vAL
End Function
Private Sub pClearFields()
 On Error Resume Next
    txtsab_tDate.Text = ""
    m_sec_lID = ""
    cbosec_lID.Text = ""
End Sub
Private Sub pRefreshGridData()
    pPopulateStudentGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent CloseUp(m_bSave)
End Sub
Private Sub grdList_AfterUpdate()
    pRefreshGridData
End Sub
Private Sub grdList_BeforeUpdate(Cancel As Integer)
  If Len(grdList.Columns.Item(0)) = 0 Then
    pAdd
  Else
    pDelete
  End If
End Sub
Private Sub grdList_Error(ByVal DataError As Integer, Response As Integer)
     Response = 0
End Sub
Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pEnterAsTab KeyCode
    End If
End Sub
Private Function fbCheckFields() As Boolean
    Dim bCheck As Boolean
    bCheck = True
        
    If bCheck Then
        If Len(m_sec_lID) = 0 Then
            MsgBox Res.fsGetResourceString(10270), vbExclamation, App.Title
            bCheck = False
            cbosec_lID.SetFocus
        End If
    End If

    If bCheck Then
        If Len(Trim(txtsab_tDate.Text)) = 0 Then
            MsgBox Res.fsGetResourceString(10271), vbExclamation, App.Title
            bCheck = False
            txtsab_tDate.SetFocus
        End If
    End If

    If bCheck Then
        If Not fbValidateDate(Trim(txtsab_tDate.Text)) Then
            MsgBox Res.fsGetResourceString(10271), vbExclamation, App.Title
            bCheck = False
            txtsab_tDate.SetFocus
        End If
    End If
    
  fbCheckFields = bCheck
End Function
Private Sub pInitalizeProperties()
  'Call global default properties subroutine
   pDefaultProperties m_tProperties
End Sub
Private Sub grdList_Validate(Cancel As Boolean)
On Error Resume Next
grdList.Update
    
End Sub
Private Sub m_oSab_ErrorGenerated(ByVal enuErrorFlag As prjChristelDL.enuErrFlag, ByVal lErrNumber As Long, ByVal sErrString As String)
 Select Case enuErrorFlag
   
        Case enuErrAdd
        
        Select Case lErrNumber
        
            Case 2627 'Duplicate Data
            MsgBox Res.fsGetResourceString(2007) & Chr(10) _
            , vbExclamation, App.Title
               
            Case Else
            MsgBox Res.fsGetResourceString(2001) & Chr(10) _
            & Res.fsGetResourceString(2000) & Space(1) & lErrNumber & Space(1) & "to" & Chr(10) _
            & App.CompanyName _
            , vbCritical, App.Title
        
        End Select
      
        Case enuErrUpdate
      
        Select Case lErrNumber
        
            Case 2627  'Duplicate Data
            MsgBox Res.fsGetResourceString(2008) & Chr(10) _
            , vbExclamation, App.Title

               
        Case Else
            MsgBox Res.fsGetResourceString(2002) & Chr(10) _
            & Res.fsGetResourceString(2000) & Space(1) & lErrNumber & Space(1) & "to" & Chr(10) _
            & App.CompanyName _
            , vbCritical, App.Title
        End Select
            
        Case enuErrDelete
        
        Select Case lErrNumber
    
            Case 547
            MsgBox Res.fsGetResourceString(2018) _
            , vbExclamation, App.Title
       
            Case Else
            MsgBox Res.fsGetResourceString(2003) & Chr(10) _
            & Res.fsGetResourceString(2000) & Space(1) & lErrNumber & Space(1) & "to" & Chr(10) _
            & App.CompanyName _
            , vbCritical, App.Title
             
        End Select
             
             
        Case enuErrEdit
            MsgBox Res.fsGetResourceString(2004) & Chr(10) _
            & Res.fsGetResourceString(2000) & Space(1) & lErrNumber & Space(1) & "to" & Chr(10) _
            & App.CompanyName _
            , vbCritical, App.Title
  
        Case enuErrLoading
            MsgBox Res.fsGetResourceString(2005) & Chr(10) _
            & Res.fsGetResourceString(2000) & Space(1) & lErrNumber & Space(1) & "to" & Chr(10) _
            & App.CompanyName _
            , vbCritical, App.Title
                                
        Case enuErrUnKnown
            MsgBox Res.fsGetResourceString(2006) & Chr(10) _
            & Res.fsGetResourceString(2000) & Space(1) & lErrNumber & Space(1) & "to" & Chr(10) _
            & App.CompanyName _
            , vbCritical, App.Title
                
        End Select
End Sub
Private Sub pLoadSection()
   On Error Resume Next
    Dim spGetSql As String
    Dim adors As ADODB.Recordset
    Dim i As Integer
    On Error Resume Next
    
    cbosec_lID.Clear
    spGetSql = "SELECT tblSection.sec_lID, tblSection.sec_sFullName" _
        & " FROM  tblSection ORDER BY tblSection.sec_sFullName"
    If m_oList.fbGetlist(adors, spGetSql) Then
        Do While Not adors.EOF
        cbosec_lID.AddItem adors(1) & ";" & adors(0), i
        i = i + 1
        adors.MoveNext
        DoEvents
        Loop
    End If
    i = 0
    adors.Close
    Set adors = Nothing
  End Sub
Private Sub txtsab_tDate_Validate(Cancel As Boolean)
  On Error Resume Next
    If Len(Trim(txtsab_tDate.Text)) <> 0 And Not fbValidateDate(Trim(txtsab_tDate.Text)) Then
            MsgBox Res.fsGetResourceString(2009), vbExclamation, App.Title
            Cancel = True
            txtsab_tDate.SetFocus
    ElseIf Len(Trim(txtsab_tDate.Text)) <> 0 Then
           txtsab_tDate.Text = Format(Trim(txtsab_tDate.Text), "Short Date")
           
    End If
End Sub

Private Sub pCreateGridColumns()
 Dim tdbCol As TrueOleDBGrid80.Column
 
  'Clear Default Columns
  For Each tdbCol In grdList.Columns
    grdList.Columns.Remove tdbCol.ColIndex
  Next
  
  'sab_lID
  Set tdbCol = grdList.Columns.Add(0)
  tdbCol.Caption = "sab_lID"
  
  'sta_lID
  Set tdbCol = grdList.Columns.Add(1)
  tdbCol.Caption = "sta_lID"
  
  'Student No
  Set tdbCol = grdList.Columns.Add(2)
  tdbCol.Caption = "Student ID"
  tdbCol.Width = "200"
  tdbCol.Locked = True
  tdbCol.AllowFocus = False
  
  'Student Name
  Set tdbCol = grdList.Columns.Add(3)
  tdbCol.Caption = "Student Name"
  tdbCol.Width = "200"
  tdbCol.Locked = True
  tdbCol.AllowFocus = False
  
   
  'Mark absentee
  Set tdbCol = grdList.Columns.Add(4)
  tdbCol.Caption = "Mark Absentee"
  tdbCol.ValueItems.Presentation = dbgCheckBox
 
   
        
  Set tdbCol = Nothing
        
  'Set unique properties for the columns
  For Each tdbCol In grdList.Columns
    If tdbCol.ColIndex = 0 Or tdbCol.ColIndex = 1 Then
       tdbCol.Visible = False
       tdbCol.AllowSizing = False
    Else
      tdbCol.Visible = True
      
    End If
 
  Next
 
  'Rebind the grid for changes
  grdList.ReBind
  Set tdbCol = Nothing
End Sub
Private Sub pInitalizeGridProperties()
  'Call global grid properties subroutine
  pSetGridProperties grdList, m_tProperties
    
  grdList.ExtendRightColumn = True
  grdList.Appearance = dbgTrack3D
  grdList.AllowColSelect = True
  grdList.AllowUpdate = True
End Sub
Private Sub pPopulateStudentGrid()
  Dim vArray As Variant
  Dim adors As New ADODB.Recordset
  Dim vAL As Variant
 On Error Resume Next
  ReDim vAL(1)
  vAL(0) = txtsab_tDate.Text
  vAL(1) = m_sec_lID
  m_xdbAbsentee.ReDim 0, -1, 0, 5
     
  If m_oSab.fbGetAbsentee(adors, vAL) = True Then
      vArray = adors.GetRows
      m_xdbAbsentee.LoadRows (vArray)
  End If
  
  m_xdbAbsentee.DefaultColumnType(0) = XTYPE_LONG  'sab_lID
  m_xdbAbsentee.DefaultColumnType(1) = XTYPE_LONG 'student no
  m_xdbAbsentee.DefaultColumnType(2) = XTYPE_STRING 'student name
  m_xdbAbsentee.DefaultColumnType(3) = XTYPE_STRING
  m_xdbAbsentee.DefaultColumnType(4) = XTYPE_BOOLEAN 'check box
 
  Set grdList.Array = m_xdbAbsentee
  grdList.ReBind
End Sub

