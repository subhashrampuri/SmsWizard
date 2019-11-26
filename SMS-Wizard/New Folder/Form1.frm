VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   324
      Left            =   384
      TabIndex        =   1
      Top             =   480
      Width           =   2316
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   324
      Left            =   600
      TabIndex        =   0
      Top             =   1056
      Width           =   1596
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCon As ADODB.Connection
Private Sub Command1_Click()
Dim sConstr As String
Dim objRs As New ADODB.Recordset

 On Error GoTo LocalErr
 
   Set objCon = New ADODB.Connection
  
'     DriverId = 790: Excel 97 / 2000
'     DriverId = 22: Excel 5 / 95
'     DriverId = 278: Excel 4
'     DriverId = 534: Excel 3
 

    sConstr = "DRIVER={Microsoft Excel Driver (*.xls)};DriverId=22;ReadOnly=True;" & _
        "DefaultDir=" & App.Path & ";DBQ=demo1.xls;FirstRowHasNames=0" & ";"
                      
'   With objCon
'        .Provider = "Microsoft.Jet.OLEDB.4.0"
'        .ConnectionString = "Data Source=" & App.Path & "\Automobile-Dealers.xls;" & _
'            "Extended Properties=Excel 9.0;"
'        .Open
'   End With
    
    objCon.Open sConstr
    
    objRs.Open "SELECT * FROM [Bangalore$]", objCon, adOpenKeyset, adLockReadOnly, adCmdText
    
    MsgBox objRs.Fields.Count
    
    MsgBox objRs.Fields(0).Name
    
    MsgBox objRs.RecordCount
    
    objRs.MoveFirst
       
   While objRs.EOF = False
    MsgBox objRs.Fields(0) & "-" & objRs.Fields(1)
    
    objRs.MoveNext
   Wend
   objRs.Close
   
   Set objRs = Nothing
   Set objCon = Nothing
   Exit Sub
LocalErr:
    MsgBox Err.Description

    Resume Next
End Sub
Private Sub Command2_Click()
    Dim objApp As Object
    Dim objBook As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim i As Integer
    
 On Error GoTo LocalErr
 
    Set objApp = Nothing
    Set objBook = Nothing
    Set objSheet = Nothing
    
    Set objApp = CreateObject("Excel.Application")
    Set objBook = objApp.Workbooks.Open(App.Path & "\demo1.xls")
    
    Set objSheet = objBook.Worksheets("Bangalore")
    
    For i = 1 To 50
        If objSheet.Cells(i, 1) <> 0 Then
             MsgBox objSheet.Cells(i, 1)
        End If
    Next
    
    Set objApp = Nothing
    Set objBook = Nothing
    Set objSheet = Nothing
    
    Exit Sub
LocalErr:
    MsgBox Err.Description
End Sub
