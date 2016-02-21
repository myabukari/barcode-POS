Attribute VB_Name = "Module1"

Public Function OpenConnection(cnn As ADODB.Connection, CnnMsg As String) As Boolean
On Error GoTo OkError

If cnn.State = 0 Then
   cnn.CursorLocation = adUseClient
   cnn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ZuksData;Data Source=nxomen", "sa", "abu"
   'cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\database\Product.mdb;Persist Security Info=False;Jet OLEDB:Database Password="
End If

OpenConnection = True
CnnMsg = "Connection to Data Server Successful."

Exit Function
OkError:
    OpenConnection = False
     If cnn.State = 1 Then
        cnn.Close
     End If
        CnnMsg = "There is a Problem Connecting to the Database! - Call the System Administrator for Assistance."
     Exit Function
End Function
Sub CenterForm(CurForm As Form)
    
    
    On Error GoTo ERRHANDLER
    Dim cMouseStatus As clsMousePointer
    Set cMouseStatus = New clsMousePointer
    
    Dim xPos As Integer
    Dim yPos As Integer
    Dim wForm As Integer
    Dim hForm As Integer
    Dim nTop As Integer
    Dim nLeft As Integer
    
    xPos = Screen.Width
    yPos = Screen.Height
    wForm = CurForm.Width
    hForm = CurForm.Height
    nTop = (yPos - hForm) / 2
    nLeft = (xPos - wForm) / 2
    CurForm.Top = nTop
    CurForm.Left = nLeft
    Exit Sub
    
ERRHANDLER:
    MsgBox "Unexpected error in procedure: CenterForm" & vbCrLf & _
        "Error #" & Err.Number & ": " & Err.Description, _
        vbCritical + vbOKOnly, App.Title
End Sub


