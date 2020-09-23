VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   1095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   420
      Width           =   1245
   End
   Begin VB.ComboBox cmbOp 
      Height          =   315
      ItemData        =   "Filter.frx":0000
      Left            =   2400
      List            =   "Filter.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   675
   End
   Begin VB.ComboBox cmbField 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   1
      Top             =   600
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4860
      TabIndex        =   0
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   195
      Left            =   3240
      TabIndex        =   6
      Top             =   210
      Width           =   975
   End
   Begin VB.Label lblField 
      Caption         =   "Field"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   210
      Width           =   1185
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub DisplayFields(ByVal rstData As ADODB.Recordset)
   Dim lngCount   As Long
   Dim lngStart   As Long
   Dim lngEnd     As Long
   
   On Error GoTo EHError
   
   lngStart = 0
   lngEnd = rstData.Fields.Count - 1
   
   For lngCount = lngStart To lngEnd
      cmbField.AddItem rstData.Fields(lngCount).Name
   Next lngCount
   
   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strFilter As String
   
   strFilter = cmbField.Text & " " & cmbOp.Text & " " & txtValue.Text
   Call frmGrid.FilterData(strFilter)
   Unload Me
End Sub
