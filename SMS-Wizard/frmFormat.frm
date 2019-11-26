VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "TODG8.OCX"
Begin VB.Form frmFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6528
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6528
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne elsTop 
      Height          =   5748
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7980
      _cx             =   14076
      _cy             =   10139
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
         Left            =   4248
         TabIndex        =   18
         Top             =   2160
         Width           =   3540
         Begin VB.CheckBox chkHeading 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "First row has column names"
            ForeColor       =   &H80000008&
            Height          =   324
            Left            =   144
            TabIndex        =   19
            Top             =   288
            Width           =   2220
         End
      End
      Begin VB.Frame fraThree 
         Caption         =   "Specify Column Delimiter"
         Height          =   780
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   4044
         Begin VB.TextBox txtOther 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   2568
            TabIndex        =   17
            Top             =   336
            Width           =   1188
         End
         Begin VB.OptionButton optOther 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Other"
            ForeColor       =   &H80000008&
            Height          =   204
            Left            =   1776
            TabIndex        =   16
            Top             =   360
            Width           =   660
         End
         Begin VB.OptionButton optComma 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Comma"
            ForeColor       =   &H80000008&
            Height          =   204
            Left            =   816
            TabIndex        =   15
            Top             =   360
            Width           =   852
         End
         Begin VB.OptionButton optTab 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Tab"
            ForeColor       =   &H80000008&
            Height          =   204
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   588
         End
      End
      Begin VB.Frame fraTwo 
         Height          =   924
         Left            =   120
         TabIndex        =   8
         Top             =   1152
         Width           =   7716
         Begin VB.OptionButton optFormat 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fixed Width"
            ForeColor       =   &H80000008&
            Height          =   324
            Left            =   1632
            TabIndex        =   11
            Top             =   552
            Width           =   1116
         End
         Begin VB.OptionButton optDelimited 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delimited."
            ForeColor       =   &H80000008&
            Height          =   324
            Left            =   1632
            TabIndex        =   9
            Top             =   240
            Width           =   972
         End
         Begin VB.Label lblFormat 
            AutoSize        =   -1  'True
            Caption         =   "Information is aligned into columns of equal width."
            Height          =   192
            Index           =   2
            Left            =   2856
            TabIndex        =   12
            Top             =   600
            Width           =   3480
         End
         Begin VB.Label lblFormat 
            AutoSize        =   -1  'True
            Caption         =   "The columns are separated by any character(s)."
            Height          =   192
            Index           =   0
            Left            =   2832
            TabIndex        =   10
            Top             =   288
            Width           =   3420
         End
      End
      Begin VB.Frame frmOne 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   0
         TabIndex        =   5
         Top             =   -24
         Width           =   7980
         Begin VB.Label lblSource 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "To Import the data, confirm the source file format. Confirm that the file properties are correctly detected before proceeding."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   144
            TabIndex        =   7
            Top             =   528
            Width           =   7488
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDataSource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select File Format"
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
            Width           =   1824
         End
      End
      Begin TrueOleDBGrid80.TDBGrid grdDetails 
         Height          =   2088
         Left            =   144
         TabIndex        =   20
         Top             =   3384
         Width           =   7596
         _ExtentX        =   13399
         _ExtentY        =   3683
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
      Top             =   5748
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
      _GridInfo       =   $"frmFormat.frx":0000
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
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sFileName As String
Private sFileTitle As String
Dim m_xdbList As XArrayDB
Dim objCon As ADODB.Connection
Dim iNoOfColumns As Long
Dim sPath As String
Public Property Let sSource(ByVal sStr As String)
    sFileName = sStr
End Property
Public Property Let sTitle(ByVal sFile As String)
    sFileTitle = sFile
End Property
Private Sub cmdCancel_Click()
    End
End Sub
Private Sub Form_Load()

Open sPath & "\schema.ini" For Output As #1
    Print #1, "[" & sFileTitle & "]"
    Print #1, "Format=Delimited( )"
    Print #1, "ColNameHeader = true"
    Print #1, "MaxScanRows = 0"
Close #1
    pInitializeComponent
    pReadData
    pCreateGridColumns
    pInitalizeGridProperties
    pPopulateData
    
'Kill sPath & "schema.ini"
End Sub
Private Sub pInitializeComponent()
Dim sConstr As String
   
    Set m_xdbList = New XArrayDB
    Set objCon = New ADODB.Connection
    sPath = Mid(sFileName, 1, Len(sFileName) - Len(sFileTitle))
    sConstr = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
            "DefaultDir=" & sPath & ";"
                      
    objCon.Open sConstr

End Sub
Private Sub pDestroyComponent()
 Set m_xdbList = Nothing
 Set objCon = Nothing
End Sub
Private Sub pReadData()
Dim objRs As New ADODB.Recordset
Dim sSource As String
Dim iLoop As Long

    sSource = "SELECT * FROM [" & sFileTitle & "]"
    objRs.Open sSource, objCon, adOpenKeyset
    iNoOfColumns = objRs.Fields.Count
   
    objRs.Close
    Set objRs = Nothing
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
 
 'Delete columns
 For iDel = 0 To grdDetails.Columns.Count - 1
    If grdDetails.Columns.Count <> 1 Then
        grdDetails.Columns.Remove iDel
    End If
 Next
 
 If chkHeading.Value = 1 Then
    pOpenConnection
    sSource = "SELECT * FROM [" & sFileTitle & "]"
    objRs.Open sSource, objCon, adOpenKeyset
 End If
 
 If iNoOfColumns <> 1 Then
 
 For iLoop = 1 To iNoOfColumns - 1
  Set tdbCol = grdDetails.Columns.Add(iLoop)
    tdbCol.Width = "2000"
    tdbCol.Locked = True
    tdbCol.AllowFocus = True
 Next
 End If
    
  Set tdbCol = Nothing
        
  'Set caption for the columns
  For iCap = 0 To iNoOfColumns - 1
     If chkHeading.Value = 1 Then
        grdDetails.Columns(iCap).Caption = objRs.Fields(iCap).Name
        objRs.MoveNext
     Else
        grdDetails.Columns(iCap).Caption = "Col" & iLoop + 1
     End If
       ' tdbCol.Visible = True
        grdDetails.Columns(iCap).Visible = True
  Next
 
  'Rebind the grid for changes
   grdDetails.ReBind
   Set tdbCol = Nothing
   Set objRs = Nothing
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
End Sub
Private Sub pPopulateData()
Dim vArray As Variant
Dim adoRs As New ADODB.Recordset
Dim iLoop As Long
m_xdbList.ReDim 0, -1, 0, iNoOfColumns - 1
  
  If fvReturnData(adoRs) Then
   vArray = adoRs.GetRows
   m_xdbList.LoadRows vArray
  End If
  
 For iLoop = 0 To iNoOfColumns - 1
     m_xdbList.DefaultColumnType(iLoop) = XTYPE_STRING
 Next
  Set grdDetails.Array = m_xdbList
  grdDetails.ReBind
  adoRs.Close
  Set adoRs = Nothing
End Sub
Private Function fvReturnData(adoRs As Recordset) As Boolean
Dim sSource As String
    sSource = "SELECT * FROM [" & sFileTitle & "]"
    adoRs.Open sSource, objCon, adOpenKeyset
    
fvReturnData = True
End Function
Private Sub optComma_Click()
Dim objRs As New ADODB.Recordset
Dim sSource As String
Dim iLoop As Long

Open sPath & "\schema.ini" For Output As #1
    Print #1, "[" & sFileTitle & "]"
    Print #1, "Format = CSVDelimited"
    If chkHeading.Value = 1 Then
        Print #1, "ColNameHeader = True"
    Else
        Print #1, "ColNameHeader = false"
    End If
    Print #1, "MaxScanRows = 0"
 
    Close #1
    
    pOpenConnection
    sSource = "SELECT * FROM [" & sFileTitle & "]"
    objRs.Open sSource, objCon, adOpenKeyset
    iNoOfColumns = objRs.Fields.Count
    objRs.Close
    Set objRs = Nothing
    
    pCreateGridColumns
    pInitalizeGridProperties
    pPopulateData
    
    Kill (sPath & "schema.ini")
End Sub
Private Sub optFormat_Click()
Dim objRs As New ADODB.Recordset
Dim sSource As String
Dim iLoop As Long

Open sPath & "\schema.ini" For Output As #1
    Print #1, "[" & sFileTitle & "]"
    Print #1, "Format=FixedLength"
    If chkHeading.Value = 1 Then
        Print #1, "ColNameHeader = True"
    Else
        Print #1, "ColNameHeader = false"
    End If
    Print #1, "MaxScanRows = 0"
 
    Close #1
    
    pOpenConnection
    sSource = "SELECT * FROM [" & sFileTitle & "]"
    objRs.Open sSource, objCon, adOpenForwardOnly
    iNoOfColumns = objRs.Fields.Count
    objRs.Close
    Set objRs = Nothing
    
    pCreateGridColumns
    pInitalizeGridProperties
    pPopulateData
    
    Kill (sPath & "schema.ini")
End Sub
Private Sub optTab_Click()
Dim objRs As New ADODB.Recordset
Dim sSource As String
Dim iLoop As Long

Open sPath & "\schema.ini" For Output As #1
    Print #1, "[" & sFileTitle & "]"
    Print #1, "Format = TabDelimited"
    If chkHeading.Value = 1 Then
        Print #1, "ColNameHeader = True"
    Else
        Print #1, "ColNameHeader = false"
    End If
    Print #1, "MaxScanRows = 0"
 
    Close #1
    
    pOpenConnection
    sSource = "SELECT * FROM [" & sFileTitle & "]"
    objRs.Open sSource, objCon, adOpenKeyset
    iNoOfColumns = objRs.Fields.Count
    objRs.Close
    Set objRs = Nothing
    
    pCreateGridColumns
    pInitalizeGridProperties
    pPopulateData
    
    Kill (sPath & "schema.ini")
End Sub
Private Sub pOpenConnection()
Dim sConstr As String

   sConstr = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
            "DefaultDir=" & sPath & ";"
                 
   If objCon.State = 1 Then
    objCon.Close
   End If
    objCon.Open sConstr
End Sub
Private Sub txtOther_Validate(Cancel As Boolean)
Dim objRs As New ADODB.Recordset
Dim sSource As String
Dim iLoop As Long

Open sPath & "\schema.ini" For Output As #1
    Print #1, "[" & sFileTitle & "]"
    Print #1, "Format = Delimited(" & txtOther.Text & ")"
    If chkHeading.Value = 1 Then
        Print #1, "ColNameHeader = True"
    Else
        Print #1, "ColNameHeader = false"
    End If
    Print #1, "MaxScanRows = 0"
 
    Close #1
    
    pOpenConnection
    sSource = "SELECT * FROM [" & sFileTitle & "]"
    objRs.Open sSource, objCon, adOpenKeyset
    iNoOfColumns = objRs.Fields.Count
    objRs.Close
    Set objRs = Nothing
    
    pCreateGridColumns
    pInitalizeGridProperties
    pPopulateData
    
    Kill (sPath & "schema.ini")
End Sub
