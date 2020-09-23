VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Categories"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "Grid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2730
      TabIndex        =   4
      Top             =   90
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dgdData 
      Height          =   3825
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6747
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      Height          =   375
      Left            =   1830
      TabIndex        =   2
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   930
      TabIndex        =   1
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   855
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------
'Project       :  A simple template for 'Master Data'
'File Name     :  frmGrid.frm
'Description   :  This form contains the functionality for
'                 Viewing, sorting, filtering data
'Created By    :  Gerard D'Mello
'Created Date  :  01/11/2001
'-------------------------------------------------------------------

Option Explicit

Dim mrstData      As ADODB.Recordset
Dim mstrQuery     As String

Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub ShowData(ByVal strQuery As String)
   On Error GoTo EHError
   
   mstrQuery = strQuery
   Call InitData

   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub InitData()
   Set mrstData = New ADODB.Recordset
   mrstData.CursorLocation = adUseClient
   mrstData.Open mstrQuery, gconDatabase, adOpenStatic
   Set dgdData.DataSource = mrstData
   dgdData.ClearFields
   dgdData.ReBind
End Sub

Private Sub cmdFilter_Click()
   Call frmFilter.DisplayFields(mrstData)
   frmFilter.Show vbModal
End Sub

Private Sub cmdRefresh_Click()
   On Error GoTo EHError
   
   mrstData.Filter = ""
   mrstData.Sort = ""
   mrstData.Requery
   dgdData.ClearFields
   dgdData.ReBind
   
   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdSort_Click()
   On Error GoTo EHError
   
   If dgdData.SelStartCol >= 0 Then
      mrstData.Sort = mrstData.Fields(dgdData.SelStartCol).Name
   End If
   
   Exit Sub
   
EHError:
   Select Case Err.Number
      Case -2147217824
         MsgBox "Sort cannot be performed on this column!", vbInformation + vbOKOnly, "Sort"
      
      Case Else
         MsgBox "Error : " & Err.Number & vbCrLf & _
            "Description : " & Err.Description & vbCrLf & _
            "Source : " & Err.Source & vbCrLf & vbCrLf & _
            "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
   End Select
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   'This will resize the grid when the form is resized
   dgdData.Height = Me.ScaleHeight - dgdData.Top - 60
   dgdData.Width = Me.ScaleWidth - 120
End Sub

Public Sub FilterData(ByVal strFilter As String)
   On Error GoTo EHError
   
   If strFilter <> "" Then
      mrstData.Filter = strFilter
   End If
   
   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub
