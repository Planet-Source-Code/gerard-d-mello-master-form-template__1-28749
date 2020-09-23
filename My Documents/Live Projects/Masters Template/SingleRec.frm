VERSION 5.00
Begin VB.Form frmCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categories"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "SingleRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   795
      Left            =   4950
      Picture         =   "SingleRec.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2070
      Width           =   795
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   5940
      TabIndex        =   12
      Top             =   3060
      Width           =   5940
      Begin VB.CommandButton cmdFirst 
         Height          =   330
         Left            =   0
         Picture         =   "SingleRec.frx":00CE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   330
         Left            =   345
         Picture         =   "SingleRec.frx":0410
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   330
         Left            =   5190
         Picture         =   "SingleRec.frx":0752
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Height          =   330
         Left            =   5535
         Picture         =   "SingleRec.frx":0A94
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   690
         TabIndex        =   17
         Top             =   30
         Width           =   4500
      End
   End
   Begin VB.TextBox txtId 
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   300
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   4140
      Picture         =   "SingleRec.frx":0DD6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   3330
      Picture         =   "SingleRec.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "&Grid"
      Height          =   795
      Left            =   2535
      Picture         =   "SingleRec.frx":1762
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   795
      Left            =   1725
      Picture         =   "SingleRec.frx":1C18
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   795
      Left            =   930
      Picture         =   "SingleRec.frx":205A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2070
      Width           =   795
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   795
      Left            =   120
      Picture         =   "SingleRec.frx":249C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2070
      Width           =   795
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   1560
      Width           =   5265
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   930
      Width           =   1935
   End
   Begin VB.Label lblLabel 
      Caption         =   "Category Id"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   11
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label lblLabel 
      Caption         =   "Description"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   1350
      Width           =   1365
   End
   Begin VB.Label lblLabel 
      Caption         =   "Category Name"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------
'Project       :  A simple template for 'Master Data'
'File Name     :  Categories.frm
'Description   :  This form contains the functionality for
'                 Adding/Modifying/Deleteing Records.
'                 You can also move through the records
'                 ie. MoveFirst, MoveLast, MovePrevious, MoveNext.
'Created By    :  Gerard D'Mello
'Created Date  :  01/11/2001
'-------------------------------------------------------------------

Option Explicit

'User defined datatype to indicate whether user is Adding or Editing
'a single record.
Private Type Mode
   Add As Boolean
   Edit As Boolean
End Type

Private mudtMode As Mode
Private WithEvents mrstData As ADODB.Recordset  'Event Sink for the recordset
Attribute mrstData.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
   On Error GoTo EHError
   
   mudtMode.Add = True
   Call EnableEditing
   mrstData.AddNew
   lblStatus.Caption = "Adding new record"
   
   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdCancel_Click()
   On Error GoTo EHError
   
   mudtMode.Add = False
   mudtMode.Edit = False
   Call DisableEditing
   mrstData.CancelUpdate
   mrstData.MoveFirst

   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdDelete_Click()
   Dim intResponse As Integer
      
   On Error GoTo EH_Error
   
   'First check if any records exist
   If Not mrstData.EOF Then
      intResponse = MsgBox("Are you sure you want to delete the selected Category?", vbYesNoCancel + vbQuestion, "Delete Category")
      'If user is sure he/she wants to delete the record
      If intResponse = vbYes Then
         With mrstData
            .Delete adAffectCurrent
            .MoveNext                     'Move to the next record
            If .EOF Then                  'If EOF is reached
               .MovePrevious
            End If
         End With
      End If
   Else
      MsgBox "No records exist. Cannot delete!", vbInformation + vbOKOnly, "Categories"
   End If
   
   Exit Sub
   
EH_Error:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdEdit_Click()
   'First check if any records exist
   If Not mrstData.EOF Then
      mudtMode.Edit = True
      Call EnableEditing
      lblStatus.Caption = "Editing Record"
   Else
      MsgBox "No records exist. Cannot Edit!", vbInformation + vbOKOnly, "Categories"
   End If
End Sub

Private Sub cmdFirst_Click()
   On Error GoTo EH_Error
   
   'First check if any records exist
   If Not mrstData.BOF Then
      mrstData.MoveFirst
   End If
   
   Exit Sub
   
EH_Error:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdGrid_Click()
   frmGrid.Show
   Call frmGrid.ShowData("SELECT CategoryID, CategoryName, Description FROM Categories")
End Sub

Private Sub cmdLast_Click()
   On Error GoTo EH_Error
   
   'First check if any records exist
   If Not mrstData.EOF Then
      mrstData.MoveLast
   End If
   
   Exit Sub
   
EH_Error:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdNext_Click()
   On Error GoTo EH_Error
   
   'First check if any records exist
   If Not mrstData.EOF Then
      mrstData.MoveNext
      If mrstData.EOF Then mrstData.MoveLast
   End If
   
   Exit Sub
   
EH_Error:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdPrevious_Click()
   On Error GoTo EH_Error
   
   If Not mrstData.BOF Then
      mrstData.MovePrevious
      If mrstData.BOF Then mrstData.MoveFirst
   End If
   
   Exit Sub
   
EH_Error:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdRefresh_Click()
   On Error GoTo EH_Error
   
   mrstData.Requery
   
   Exit Sub
   
EH_Error:
   MsgBox "Error : " & Err.Number & vbCrLf & _
      "Description : " & Err.Description & vbCrLf & _
      "Source : " & Err.Source & vbCrLf & vbCrLf & _
      "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub

Private Sub cmdSave_Click()
   Dim varBookMark   As Variant
   
   On Error GoTo EH_Error
   
   If fValidData Then      'Check if data entered is valid or not
      'if user has added a new record
      If mudtMode.Add Then
         mrstData.Update
         mudtMode.Add = False
         'Set a bookmark, so that after requerying data (recd. pointer
         'goes to 1st record), the newly added record can be displayed
         varBookMark = mrstData.Bookmark
         mrstData.Requery
      End If
      
      'if user edited an existing record
      If mudtMode.Edit Then
         mrstData.Update
         mudtMode.Edit = False
         'Used bookmark here, since the Record counter label does not get refreshed.
         varBookMark = mrstData.Bookmark
      End If
      
      Call DisableEditing
      mrstData.Bookmark = varBookMark
   End If

   Exit Sub
   
EH_Error:
   Select Case Err.Number
      Case -2147467259
         MsgBox "A category with the same name already exists. Cannot save!", _
               vbInformation + vbOKOnly, "Categories"
      Case Else
      MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
   End Select
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   'Initialize Form
   txtId.Enabled = False    'Disable this control, bcos it's an Identity field
   Call InitDataControl
   Call DisableEditing
   Screen.MousePointer = vbDefault
End Sub

'----------------------------------------------------------------
'Description   :  Used to Initialize the recordset object and to bind
'                 controls to the recordset
'Parameters    :  None
'Returns       :  None
'----------------------------------------------------------------
Private Sub InitDataControl()
   'Create a new recordset and populate the recordset
   Set mrstData = New ADODB.Recordset
   mrstData.CursorLocation = adUseClient
   mrstData.Open "SELECT CategoryID, CategoryName, Description FROM Categories", _
         gconDatabase, adOpenStatic, adLockPessimistic
   
   'Bind controls to the recordset
   Set txtId.DataSource = mrstData
   txtId.DataField = "CategoryId"
   Set txtName.DataSource = mrstData
   txtName.DataField = "CategoryName"
   Set txtDescription.DataSource = mrstData
   txtDescription.DataField = "Description"
End Sub

'----------------------------------------------------------------
'Description   :  Dont allow user to edit data, ie when user is
'                 browsing thru the records.
'Parameters    :  None
'Returns       :  None
'----------------------------------------------------------------
Private Sub DisableEditing()
   txtName.Locked = True
   txtDescription.Locked = True
      
   cmdAdd.Enabled = True
   cmdEdit.Enabled = True
   cmdDelete.Enabled = True
   cmdGrid.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   cmdRefresh.Enabled = True
      
   cmdFirst.Enabled = True
   cmdLast.Enabled = True
   cmdPrevious.Enabled = True
   cmdNext.Enabled = True
End Sub

'----------------------------------------------------------------
'Description   :  Allow user to edit data, ie when user is adding
'                 or editing a record.
'Parameters    :  None
'Returns       :  None
'----------------------------------------------------------------
Private Sub EnableEditing()
   txtName.Locked = False
   txtDescription.Locked = False
   
   cmdAdd.Enabled = False
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdGrid.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   cmdRefresh.Enabled = False

   cmdFirst.Enabled = False
   cmdLast.Enabled = False
   cmdPrevious.Enabled = False
   cmdNext.Enabled = False
End Sub

Private Sub mrstData_MoveComplete( _
               ByVal adReason As ADODB.EventReasonEnum, _
               ByVal pError As ADODB.Error, _
               ByRef adStatus As ADODB.EventStatusEnum, _
               ByVal pRecordset As ADODB.Recordset)
   If Not pRecordset.EOF Then
      lblStatus.Caption = "Record " & pRecordset.AbsolutePosition & " of " & pRecordset.RecordCount
   Else
      lblStatus.Caption = "No Records"
   End If
End Sub

Private Function fValidData() As Boolean
   fValidData = True
   
   If Trim(txtName.Text) = "" Then
      MsgBox "Category Name cannot be blank.  Cannot Save!", _
            vbInformation + vbOKOnly, "Categories"
      txtName.SetFocus
      fValidData = False
   End If
   If Trim(txtDescription.Text) = "" Then
      MsgBox "Category Description cannot be blank.  Cannot Save!", _
            vbInformation + vbOKOnly, "Categories"
      txtDescription.SetFocus
      fValidData = False
   End If
End Function
