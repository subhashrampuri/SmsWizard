VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "TODG8.OCX"
Begin VB.Form frmTables 
   Caption         =   "Form1"
   ClientHeight    =   7332
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7548
   LinkTopic       =   "Form1"
   ScaleHeight     =   7332
   ScaleWidth      =   7548
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   372
      Left            =   6216
      TabIndex        =   11
      Top             =   1896
      Width           =   852
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   468
      Left            =   6648
      TabIndex        =   10
      Top             =   3144
      Width           =   876
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Extract"
      Height          =   372
      Left            =   96
      TabIndex        =   8
      Top             =   3576
      Width           =   2004
   End
   Begin VB.ListBox lstFields 
      Height          =   2352
      Left            =   3048
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   1080
      Width           =   2748
   End
   Begin VB.ListBox lstTable 
      Height          =   2352
      Left            =   96
      TabIndex        =   6
      Top             =   1080
      Width           =   2868
   End
   Begin VB.CommandButton Command2 
      Caption         =   "...."
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   804
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3096
      Top             =   504
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   1056
      TabIndex        =   3
      Top             =   144
      Width           =   1956
   End
   Begin VB.TextBox txtsPassword 
      Height          =   324
      Left            =   1056
      TabIndex        =   1
      Top             =   528
      Width           =   1956
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tables"
      Height          =   348
      Left            =   3552
      TabIndex        =   0
      Top             =   528
      Width           =   804
   End
   Begin TrueOleDBGrid80.TDBGrid grdAccess 
      Height          =   2112
      Left            =   288
      TabIndex        =   9
      Top             =   4200
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   3725
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
      Splits(0).AllowRowSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3048"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2963"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
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
      MultiSelect     =   2
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   144
      Width           =   1044
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   204
      Left            =   144
      TabIndex        =   2
      Top             =   504
      Width           =   780
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCon As New ADODB.Connection
Dim sSelTable As String
Dim sQueryString As String
Dim m_iNoOfColumns As Integer
Dim m_iIntPreviousPos As Integer
Private Sub Command1_Click()
 On Error GoTo LocalErr
    Dim objRs As New ADODB.Recordset
    
    If Len(Trim(txtsPassword.Text)) <> 0 Then
       objCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
        txtFileName & "; Jet OLEDB:Database Password=" & txtsPassword.Text & ";"
    Else
        objCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(txtFileName.Text)
    End If
    
    objCon.Open
    
    Set objRs = objCon.OpenSchema(adSchemaTables)
    
    lstTable.Clear
    While objRs.EOF <> True
        If Left(objRs.Fields("Table_Name").Value, 4) <> "MSys" Then
            lstTable.AddItem objRs.Fields("Table_Name")
        End If
        objRs.MoveNext
    Wend
    
    If objRs.State = 1 Then
        objRs.Close
    End If
    Set objRs = Nothing
    Exit Sub
LocalErr:
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Command2_Click()
 On Error GoTo LocalErr
    With cdlg
        .CancelError = True
        .Filter = "mdb|*.mdb"
        .ShowOpen
        txtFileName.Text = .FileName
    End With
    Exit Sub
LocalErr:
    txtFileName.Text = ""
End Sub
Private Sub Command3_Click()
    pCreateAccessGridColumns
    pInitalizeAccessGridProperties
    pPopulateAccessRecord
End Sub
Private Sub Command4_Click()
    Dim i As Integer
    Dim iCnt As Integer
    Dim sSel_1 As String
    Dim sSel_2 As String
    
    For i = 0 To lstFields.ListCount - 1
        If lstFields.Selected(i) = True Then
            iCnt = iCnt + 1
            sSel_1 = lstFields.List(i)
        End If
    Next
    
    If iCnt > 2 Then
        MsgBox "Please select less than or equal to 2 columns"
    End If
    
    MsgBox grdAccess.Columns(sSel_1).Text
End Sub

Private Sub Command5_Click()
    MsgBox grdAccess.Columns("Email").Order
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If objCon.State = 1 Then
        objCon.Close
    End If
    Set objCon = Nothing
End Sub
Private Sub pCreateAccessGridColumns()
    Dim tdbCol As TrueOleDBGrid80.Column
    Dim objRs As New ADODB.Recordset
    Dim iDel As Long
    Dim iLoop As Integer
    Dim i As Integer
    
 On Error GoTo LocalErr
 
    'Delete columns
    For iDel = grdAccess.Columns.Count To 0 Step -1
        If grdAccess.Columns.Count >= 1 Then
          grdAccess.Columns.Remove iDel - 1
        End If
    Next
 
    For i = 0 To lstTable.ListCount - 1
        If lstTable.Selected(i) = True Then
            sSelTable = lstTable.List(i)
            If objRs.State = 1 Then
                objRs.Close
            End If
            objRs.Open sSelTable, objCon, adOpenKeyset, adLockReadOnly, adCmdTable
            m_iNoOfColumns = objRs.Fields.Count
        End If
    Next
    
    If m_iNoOfColumns <> 1 Then
        For iLoop = 0 To m_iNoOfColumns - 1
            Set tdbCol = grdAccess.Columns.Add(iLoop)
            tdbCol.Width = "2000"
            tdbCol.Locked = True
            tdbCol.Caption = objRs.Fields(iLoop).Name
            tdbCol.AllowFocus = True
            tdbCol.Visible = True
            lstFields.AddItem objRs.Fields(iLoop).Name
        Next
    End If
    
    Set tdbCol = Nothing
    
    'Rebind the grid for changes
    grdAccess.ReBind
    Set tdbCol = Nothing
    Set objRs = Nothing
    Exit Sub
LocalErr:
    If Err.Number = -2147217900 Then
        MsgBox "Table name should be one word!", vbExclamation, App.Title
    Else
        MsgBox Err.Description, vbExclamation, App.Title
    End If
End Sub
Private Sub pInitalizeAccessGridProperties()
    grdAccess.FetchRowStyle = True
    grdAccess.AllowArrows = True
    grdAccess.AllowDelete = False
    grdAccess.AllowAddNew = False
    grdAccess.EmptyRows = True
    grdAccess.RecordSelectors = False
    grdAccess.TabAction = dbgControlNavigation
    grdAccess.MultiSelect = dbgMultiSelectExtended
    grdAccess.ExtendRightColumn = True
    grdAccess.Appearance = dbgTrack3D
    grdAccess.AllowColSelect = True
    grdAccess.AllowUpdate = True
    grdAccess.EditActive = True
    grdAccess.AllowColSelect = True
    grdAccess.AllowColMove = True
    
End Sub
Private Sub pPopulateAccessRecord()

On Error GoTo LocalErr
 
    Dim vArray As Variant
    Dim i As Long
    Dim objRs As New ADODB.Recordset
    Dim objXDBList As New XArrayDB
      
    ReDim vArray(m_iNoOfColumns)
    
    objXDBList.ReDim 0, -1, 0, m_iNoOfColumns - 1
    
    objRs.Open sSelTable, objCon, adOpenKeyset, adLockReadOnly, adCmdTable
     
    Do While objRs.EOF = False
        For i = 0 To m_iNoOfColumns - 1
            vArray(i) = objRs(i)
        Next
        objXDBList.AppendRows (1)
        For i = 0 To m_iNoOfColumns - 1
            objXDBList(objXDBList.UpperBound(1), i) = vArray(i)
        Next
        objRs.MoveNext
   Loop
    
    Set grdAccess.Array = objXDBList
    grdAccess.ReBind
    Exit Sub
LocalErr:
    MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub pReadSelectedColumns()
    Dim i As Integer
    Dim j As Integer
    
    grdAccess.MoveFirst
    For i = grdAccess.SelStartCol To grdAccess.SelEndCol
       ' MsgBox grdAccess.Columns(grdAccess.Columns(grdAccess.SelStartCol).Order).Caption
        
       MsgBox grdAccess.Columns(grdAccess.SelStartCol).Order
        
        'MsgBox grdAccess.Columns(grdAccess.Columns(grdAccess.SelEndCol).Order).Caption
        'grdAccess.MoveNext
    Next
End Sub
'Private Sub grdAccess_ColMove(ByVal Position As Integer, Cancel As Integer)
'    grdAccess.Columns(m_iIntPreviousPos).Order = Position
'End Sub
'Private Sub grdAccess_HeadClick(ByVal ColIndex As Integer)
'    m_iIntPreviousPos = ColIndex
'End Sub
