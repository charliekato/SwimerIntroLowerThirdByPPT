VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formServerSelect 
   Caption         =   "�T�[�o�[�I��"
   ClientHeight    =   2364
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   3144
   OleObjectBlob   =   "formServerSelect.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formServerSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnOK_Click
    End If
End Sub

Private Sub txtBoxServerName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnOK_Click
    End If
End Sub


Private Sub btnOK_Click()
    Dim serverName As String
    serverName = txtBoxServerName.Text

    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String
    Unload Me
'    On Error GoTo MyError
    Set LaneOrder.MyCon = New ADODB.Connection
    LaneOrder.MyCon.connectionString = "Provider=SQLOLEDB;Data Source=" & serverName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    LaneOrder.MyCon.Open
    Dim eventPick As formEventNoPick
    Set eventPick = New formEventNoPick
    
    myQuery = "SELECT ���ԍ�, ��1 FROM ���ݒ�"
    myRecordset.Open myQuery, LaneOrder.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        formEventNoPick.listEvent.AddItem Right("   " & myRecordset!���ԍ�, 3) & "   " & if_not_null_string(myRecordset!��1)
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
    formEventNoPick.show vbModeless
    Exit Sub
MyError:
    MsgBox ("cannot access server " & serverName)
End Sub


