VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEventNoPick 
   Caption         =   "���I��"
   ClientHeight    =   6900
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   6984
   OleObjectBlob   =   "formEventNoPick.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formEventNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'
'  formEventNoPick
'
Public classBasedRace As Boolean
'Public ClassExist As Boolean


Function class_exist(dummy As String) As Boolean
    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String

    myQuery = "select * from �N���X where ���ԍ� = " & EventNo
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    If myRecordset.EOF Then
        class_exist = False
    Else
        class_exist = True
    End If
    myRecordset.Close
    Set myRecordset = Nothing


End Function

Public Function class_based_race(dummy As String) As Boolean
    Dim myQuery As String
    Dim myRecordset As New ADODB.Recordset
    myQuery = "SELECT COUNT(1) as NUM from �v���O���� where ���ԍ�=" & LaneOrder.EventNo & " and �N���X�ԍ� > 0 "
    myRecordset.Open myQuery, LaneOrder.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    If myRecordset!NUM = 0 Then
       class_based_race = False
    Else
      class_based_race = True
    End If
    myRecordset.Close
    Set myRecordset = Nothing
End Function

Private Sub btnClose_Click()

    LaneOrder.MyCon.Close
    Set LaneOrder.MyCon = Nothing

    Unload Me
End Sub

Private Sub listEvent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' �G���^�[�L�[�������ꂽ�Ƃ��A CommandButton1 ���N���b�N
        Call btnOK_Click
    End If
End Sub






Sub CreateTableIfNotExists()
    Dim cmd As Object 'ADODB.Command
    Dim sql As String


    sql = "IF NOT EXISTS (" & _
          "SELECT 1 " & _
          "FROM INFORMATION_SCHEMA.TABLES " & _
          "WHERE TABLE_NAME = '�����'" & _
          ") " & _
          "BEGIN " & _
          "CREATE TABLE ����� (" & _
          "���ԍ� smallINT NOT NULL, " & _
          "���Z�ԍ� smallINT NOT NULL, " & _
          "����� smallint NOT NULL " & _
           " CONSTRAINT PK_����� PRIMARY KEY (���ԍ�, ���Z�ԍ�)" & _
          "); " & _
          "END;"
    
    On Error GoTo ErrorHandler
    
    
    ' ADODB.Command �I�u�W�F�N�g���쐬����SQL�����s
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = LaneOrder.MyCon
    cmd.CommandText = sql
    cmd.Execute
    
    ' ���\�[�X�����

    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    ' �G���[����
    Debug.Print "�G���[���������܂���: " & Err.Description

    Set cmd = Nothing

End Sub


Sub CopyToPrintStatusIfNotExists(target���ԍ� As Integer)
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim connectionString As String
    Dim checkSql As String
    Dim insertSql As String
    
   
    
    ' ���݊m�F�pSQL��
    checkSql = "SELECT 1 FROM ����� WHERE ���ԍ� = " & target���ԍ� & ";"
    
    ' �}���pSQL��
    insertSql = "INSERT INTO ����� (���ԍ�, ���Z�ԍ�, �����) " & _
                "SELECT ���ԍ�, ���Z�ԍ�, 0 " & _
                "FROM �v���O���� " & _
                "WHERE ���ԍ� = " & target���ԍ� & ";"
    
    On Error GoTo ErrorHandler
    
    ' ADODB.Connection �I�u�W�F�N�g���쐬

    
    ' ���݊m�F�pSQL�����s
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = LaneOrder.MyCon
    cmd.CommandText = checkSql
    
    Set rs = cmd.Execute
    If rs.EOF Then
        ' ���R�[�h�����݂��Ȃ��ꍇ�̂ݑ}��SQL�����s
        cmd.CommandText = insertSql
        cmd.Execute

    End If
    
    ' ���\�[�X�����
    rs.Close

    Set rs = Nothing
    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    ' �G���[����
    Debug.Print "�G���[���������܂���: " & Err.Description
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing

End Sub
Sub add_list_item(row As Integer, item1 As String, item2 As String, item3 As String, item4 As String, item5 As String)
    formPrgNoPick.listPrg.AddItem ("")
    formPrgNoPick.listPrg.List(row, 0) = item1
    formPrgNoPick.listPrg.List(row, 1) = item2
    formPrgNoPick.listPrg.List(row, 2) = item3
    formPrgNoPick.listPrg.List(row, 3) = item4
    formPrgNoPick.listPrg.List(row, 4) = item5

End Sub


Private Sub DispProgram()
   Dim gender(5) As String
    gender(1) = "�j�q"
    gender(2) = "���q"
    gender(3) = "����"
    gender(4) = "����"
    Dim Yk(10) As String
    Yk(3) = "�^�C������"
    Yk(5) = "A����"
    Yk(6) = "����"
    Dim selectedItem As String
    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String
    Dim row As Integer

    selectedItem = listEvent.Value
    LaneOrder.EventNo = CInt(Left(selectedItem, 3))
    Call CreateTableIfNotExists
    CopyToPrintStatusIfNotExists (LaneOrder.EventNo)
    classBasedRace = class_based_race("")
    ClassExist = class_exist("")
    formPrgNoPick.listPrg.Width = 340
    formPrgNoPick.listPrg.ColumnCount = 5
    formPrgNoPick.Caption = selectedItem + "  �@���Z�I��"

    If ClassExist And classBasedRace Then
        formPrgNoPick.listPrg.ColumnWidths = "30pt;90pt;105pt;70pt;25pt"
        Call add_list_item(0, "#", "�N���X", "���", "�\/��", "st")
        row = 1
        myQuery = "SELECT �v���O����.�\���p���Z�ԍ� as ���Z�ԍ�, �N���X.�N���X���� as �N���X, " & _
              "�v���O����.���ʃR�[�h as ����, �v���O����.�\���R�[�h," & _
              "����.���� as ����, ���.��� as ��� FROM �v���O����" + _
              " INNER JOIN ��� ON ���.��ڃR�[�h = �v���O����.��ڃR�[�h " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & LaneOrder.EventNo & " AND " + _
              " �N���X.���ԍ� = " & LaneOrder.EventNo & _
              " and (�v���O����.�\���R�[�h=3 or �v���O����.�\���R�[�h=5 or �v���O����.�\���R�[�h=6) " & _
              " order by �v���O����.�\���p���Z�ԍ� asc;"
              
            myRecordset.Open myQuery, LaneOrder.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordset.EOF

                Call add_list_item(row, Right("   " & myRecordset!���Z�ԍ�, 3), myRecordset!�N���X, _
                    gender(myRecordset!����) + myRecordset!���� + myRecordset!���, Yk(myRecordset!�\���R�[�h), "")
                row = row + 1
                myRecordset.MoveNext
            Loop
    Else
        formPrgNoPick.listPrg.ColumnWidths = "30pt;20pt;120pt;130pt;20pt"
        Call add_list_item(0, "No.", "", "���", "�\/��", "st")
        row = 1
        myQuery = "SELECT �v���O����.�\���p���Z�ԍ� as ���Z�ԍ�,  " & _
              "�v���O����.���ʃR�[�h as ����, �v���O����.�\���R�[�h," & _
              "����.���� as ����, ���.��� as ��� FROM �v���O����" + _
              " INNER JOIN ��� ON ���.��ڃR�[�h = �v���O����.��ڃR�[�h " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & LaneOrder.EventNo & _
              " and (�v���O����.�\���R�[�h=3 or �v���O����.�\���R�[�h=5 or �v���O����.�\���R�[�h=6) " & _
              " order by �v���O����.�\���p���Z�ԍ� asc;"
            myRecordset.Open myQuery, LaneOrder.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordset.EOF
                Call add_list_item(row, Right("   " & myRecordset!���Z�ԍ�, 3), "", _
                    gender(myRecordset!����) + myRecordset!���� + myRecordset!���, Yk(myRecordset!�\���R�[�h), "")

                row = row + 1
                myRecordset.MoveNext
            Loop
    End If
    formPrgNoPick.LastRow = row - 1
    

    myRecordset.Close
    Set myRecordset = Nothing
   ' Call LaneOrder.init_senshu("")
    If ClassExist Then
'        Call LaneOrder.init_class("")
    End If
End Sub

Private Sub btnOK_Click()
    Call DispProgram

    
    Unload Me
    formPrgNoPick.show vbModeless
End Sub


