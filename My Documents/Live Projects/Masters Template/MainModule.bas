Attribute VB_Name = "Module1"
Option Explicit

Public gconDatabase        As ADODB.Connection

Public Sub Main()
   On Error GoTo EHError
   
   Set gconDatabase = New ADODB.Connection
   gconDatabase.CursorLocation = adUseClient
   gconDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Nwind.mdb"
   
   frmCategories.Show
   
   Exit Sub
   
EHError:
   MsgBox "Error : " & Err.Number & vbCrLf & _
         "Description : " & Err.Description & vbCrLf & _
         "Source : " & Err.Source & vbCrLf & vbCrLf & _
         "If the problem persists contact your software vendor!", vbCritical + vbOKOnly, "Error"
End Sub
