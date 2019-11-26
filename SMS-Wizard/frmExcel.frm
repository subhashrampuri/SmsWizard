VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "TODG8.OCX"
Begin VB.Form frmExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6048
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7896
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6048
   ScaleWidth      =   7896
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne elsTop 
      Height          =   5268
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7896
      _cx             =   13928
      _cy             =   9292
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
      Begin VB.Frame fraFour 
         Height          =   780
         Left            =   120
         TabIndex        =   8
         Top             =   984
         Width           =   7644
         Begin VB.CheckBox chkHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "First row has column names"
            ForeColor       =   &H80000008&
            Height          =   324
            Left            =   144
            TabIndex        =   9
            Top             =   288
            Width           =   2220
         End
      End
      Begin VB.Frame frmOne 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   948
         Left            =   0
         TabIndex        =   5
         Top             =   -24
         Width           =   7980
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "You can choose one or more sheets to import"
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
            Top             =   528
            Width           =   3684
         End
         Begin VB.Label lblDataSource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Source Sheet"
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
            Width           =   1968
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetails 
         Height          =   3144
         Left            =   144
         TabIndex        =   10
         Top             =   1944
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   5546
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
      Height          =   780
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5268
      Width           =   7896
      _cx             =   13928
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
      _GridInfo       =   $"frmExcel.frx":0000
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   348
         Left            =   5424
         TabIndex        =   3
         Top             =   216
         Width           =   1176
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Enabled         =   0   'False
         Height          =   348
         Left            =   4200
         TabIndex        =   2
         Top             =   216
         Width           =   1176
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   348
         Left            =   6648
         TabIndex        =   4
         Top             =   216
         Width           =   1176
      End
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sFileName As String
Private sFileTitle As String
Dim m_xdbList As XArrayDB
Dim iNoOfRows As Long
Dim sPath As String
Dim sSheets As String
Public Property Let sSource(ByVal sStr As String)
    sFileName = sStr
End Property
Public Property Let sTitle(ByVal sFile As String)
    sFileTitle = sFile
End Property
Private Sub cmdCancel_Click()
  End
End Sub
Private Sub cmdNext_Click()
Dim i As Integer
sSheets = ""
    grdDetails.MoveFirst
    For i = 0 To m_xdbList.UpperBound(1)
        If grdDetails.Columns(0).Value = 1 Then
            If sSheets = "" Then
                sSheets = grdDetails.Columns(1).Text
            Else
                sSheets = sSheets & "|" & grdDetails.Columns(1).Text
            End If
        End If
        grdDetails.MoveNext
    Next
End Sub
Private Sub Form_Load()
    pInitializeComponent
    pCreateGridColumns
    pInitalizeGridProperties
    pPopulateData
End Sub
Private Sub pInitializeComponent()
Dim sConstr As String
    
    Set m_xdbList = New XArrayDB
 '   Set objCon = New ADODB.Connection
    sPath = Mid(sFileName, 1, Len(sFileName) - Len(sFileTitle))
    
    ' DriverId=790: Excel 97/2000
    ' DriverId=22: Excel 5/95
    ' DriverId=278: Excel 4
    ' DriverId=534: Excel 3z
 
  '  sConstr = "DRIVER={Microsoft Excel Driver (*.xls)};DriverId=790;ReadOnly=True;" & _
   '     "DBQ=" & sPath & sFileTitle & ";"
                      
    'objCon.Open sConstr

End Sub
Private Sub pDestroyComponent()
 Set m_xdbList = Nothing
 'Set objCon = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    pDestroyComponent
End Sub
Private Sub pCreateGridColumns()
 Dim tdbCol As TrueOleDBGrid80.Column
 Dim iLoop As Integer
 Dim objRs As New ADODB.Recordset
 Dim sSource As String
 Dim iDel As Long
 Dim iCap As Long

 For Each tdbCol In grdDetails.Columns
    grdDetails.Columns.Remove tdbCol.ColIndex
 Next
 
    Set tdbCol = grdDetails.Columns.Add(0)
    tdbCol.Width = "2000"
    tdbCol.Locked = True
    tdbCol.AllowFocus = True
    tdbCol.Caption = "Select"
    tdbCol.DropDownList = False
    tdbCol.ValueItems.Presentation = dbgCheckBox
    tdbCol.Visible = True
    
    Set tdbCol = grdDetails.Columns.Add(1)
    tdbCol.Width = "2000"
    tdbCol.Locked = True
    tdbCol.AllowFocus = True
    tdbCol.Caption = "Source Sheets"
    tdbCol.Visible = True
        
  Set tdbCol = Nothing
 
  'Rebind the grid for changes
   grdDetails.ReBind
   Set tdbCol = Nothing
End Sub
Private Sub pInitalizeGridProperties()
    grdDetails.FetchRowStyle = True
    grdDetails.AllowArrows = True
    grdDetails.AllowDelete = False
    grdDetails.AllowAddNew = False
    grdDetails.EmptyRows = True
    grdDetails.RecordSelectors = False
    grdDetails.TabAction = dbgControlNavigation
    grdDetails.MultiSelect = dbgMultiSelectExtended
    grdDetails.ExtendRightColumn = True
    grdDetails.Appearance = dbgTrack3D
    grdDetails.AllowColSelect = True
    grdDetails.AllowUpdate = True
    grdDetails.EditActive = True
End Sub
Private Sub pPopulateData()
Dim wbook As Excel.Workbook
Dim vArray As Variant
Dim iLoop As Long
   
Set wbook = Excel.Application.Workbooks.Open(sPath & sFileTitle)

ReDim vArray(wbook.Sheets.Count)
m_xdbList.ReDim 0, -1, 0, 2
iNoOfRows = wbook.Sheets.Count
For iLoop = 1 To wbook.Sheets.Count
    m_xdbList.AppendRows (1)
    m_xdbList(m_xdbList.UpperBound(1), 0) = 0
    m_xdbList(m_xdbList.UpperBound(1), 1) = wbook.Sheets(iLoop).Name
Next
     
  Set grdDetails.Array = m_xdbList
  grdDetails.ReBind
  Excel.Application.Quit
End Sub
Private Sub grdDetails_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Select Case grdDetails.Col
 
  Case 0
       If grdDetails.Columns(0).Value = 0 Then
            grdDetails.Columns(0).Value = 1
       Else
            grdDetails.Columns(0).Value = 0
       End If
  Case Else
  
  End Select
End Sub

