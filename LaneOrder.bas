Attribute VB_Name = "LaneOrder"
' slide-1 �` 12 �l��ڗp
' slide-13�` 24 �����[��ڗp
'
' slide-1�@�@�@�S���[���̑I��Љ�p
'�@�@�@�l��ڂɑΉ�
'
'
'
' slide-2�@�@�@�V�L�^�ꗗ
'
'
' slide-3�`12 ���[�����̑I��Љ�
'
' slide-3 : 0���[���@�@�@slide-4 : 1���[��      slide-5 : 2���[�� .....   slide-12 : 9���[��
'
'
' slide-13    �S���[���̑I��Љ�
'     �����[��ڂɑΉ�
'
'
'
' slide-14�@�@�@�V�L�^�ꗗ
'
'
' slide-15�`24 ���[�����̑I��Љ�
'
' slide-15 : 0���[���@�@�@slide-16 : 1���[��      slide-17 : 2���[�� .....   slide-24 : 9���[��
'
'
'
' �g�p����textbox�̖��O
'
' �Sslide���ʂ�textbox (���o���ȂǂɎg�p)
'
'    �@PN(���Z�ԍ�)�A���ʁA�N���X�A�����A���
'
'   ����!! �������O�� textBox���������ꍇ�@�G���[�ɂ͂Ȃ炸�A��ɂ݂�����textBox�ɍ������݂���܂��B
'
'
' slide-1
'
' ���H0 , ����0 , ����0 , �N���X0 , �\�I�^�C��0 ,�@�@<--�����̐����̓��[���ԍ�
' ���H1 , ����1 , ����1 , �N���X1 , �\�I�^�C��1 ,
' ���H2 , ����2 , ����2 , �N���X2 , �\�I�^�C��2 ,
' ...
' ���H9 , ����9 , ����9 , �N���X9 , �\�I�^�C��9 ,
'
' ���H0, ���H1, ���H2, ..., ���H9  �́A ���[���ԍ��� ���H0�̒��g��0�ɂȂ�B
' �I��̂��Ȃ��Ƃ���͕\������Ȃ�
'
'slide-2.
'
'
' ���Elabel , ���E�L�^ , ���E������, ���E����  , ���E���t ,
' ���{label , ���{�L�^ , ���{������, ���{����  , ���{���t ,
' �w��label , �w���L�^ , �w��������, �w������  , �w�����t ,
' ���Zlabel , ���Z�L�^ , ���Z������, ���Z����  , ���Z���t ,
' ���wlabel , ���w�L�^ , ���w������, ���w����  , ���w���t,
' �w��label , �w���L�^ , �w��������, �w������  , �w�����t,
' ��label , ���L�^ , �������� , ������ , �����t ,
' �����Zlabel , �����Z�L�^ , �����Z������ , �����Z���� , �����Z���t ,
' �����wlabel , �����w�L�^ , �����w������ , �����w���� , �����w���t ,
' ���w��label , ���w���L�^ , ���w�������� , ���w������ , ���w�����t ,
'
'
'�X���C�h3�`12
'
'���H�A�����A�����A�\�I�^�C��
'
'slide-13   slide-13�ȍ~�̓����[�p
'
' ���H0 , ����10 , ����20, ����30, ����40, �`�[����0 , �N���X0 , �\�I�^�C��0 ,
' ���H1 , ����11 , ����21, ����31, ����41, �`�[����1 , �N���X1 , �\�I�^�C��1 ,
' ���H2 , ����12 , ����21, ����32, ����42, �`�[����2 , �N���X2 , �\�I�^�C��2 ,
' ...
' ���H9 , ����19 , ����29 ,����39 ,����49, �`�[����9 , �N���X9 , �\�I�^�C��9 ,
'
'    �����̐����̓��[���ԍ��B���j�� �́@����1
'
'
'slide-14.
' slide-2 �Ɠ��������ɂ̓`�[�������͂���A�����҂ɂ�4�l�̖��O������
'
' ���Elabel , ���E�L�^ , ���E������, ���E����  , ���E���t ,
' ���{label , ���{�L�^ , ���{������, ���{����  , ���{���t ,
' �w��label , �w���L�^ , �w��������, �w������  , �w�����t ,
' ���Zlabel , ���Z�L�^ , ���Z������, ���Z����  , ���Z���t ,
' ���wlabel , ���w�L�^ , ���w������, ���w����  , ���w���t,
' �w��label , �w���L�^ , �w��������, �w������  , �w�����t,
' ��label , ���L�^ , �������� , ������ , �����t ,
' �����Zlabel , �����Z�L�^ , �����Z������ , �����Z���� , �����Z���t ,
' �����wlabel , �����w�L�^ , �����w������ , �����w���� , �����w���t ,
' ���w��label , ���w���L�^ , ���w�������� , ���w������ , ���w�����t ,

'
'�X���C�h15�`24
'
'���H�A����1, ����2, ����3, ����4, �����A�\�I�^�C��
'
'  slide 3�`12�Ɠ��������A�I�薼�� �����ł͂Ȃ��A ����1(���j��)�` ����4(��l�j��)�ɂȂ�B








Option Explicit
Option Base 0
Const �I��Љ� As Boolean = False
Const DefaultServerName = "localhost"
Public ZeroUse As Boolean

Public TOPSLIDE As Integer   ' relay�� 13, �l��� 1
Dim DispItems As Variant
Dim DispItems2 As Variant
Dim DispItems3 As Variant ' �V�L�^�p
Dim DispItems4 As Variant ' �V�L�^�p
    Public MyCon As ADODB.Connection
    Public EventNo As Integer


    Public MaxClassNo As Integer

'    Public ClassTable() As String
Sub setDispItems(dummy As String)
    DispItems = Array("PN", "����", "���", "����", "�N���X")
End Sub
Sub setDispItems2(dummy As String)
    DispItems2 = Array("�`�[����", "����1", "����2", "����3", "����4", "�N���X", "����", "�\�I�^�C��", _
    "����", "����", "���H", "����", "���")
End Sub
Sub setDispItems3(dummy As String)
    DispItems3 = Array("���E", "���{", "�w��", "���Z", "���w", "�w��", "��", "�����Z", "�����w", "���w��", "���")
End Sub
Sub setDispItems4(dummy As String)
    DispItems4 = Array("label", "�L�^", "�ێ���", "����", "���t")
End Sub
Sub ClearAllTextBoxes(dummy As Integer)
    Dim field As Variant
    Dim i As Integer
    For Each field In DispItems2
        For i = 0 To 9
            Call show2(field, "", i)
        Next i
    Next
  
End Sub

Sub ClearAllRecordTextBoxes(dummy As String)
    Dim field As Variant
    Dim field2 As Variant
    Dim boxName As String
    For Each field In DispItems
        Call show(field, "", TOPSLIDE + 1)
    Next
    For Each field In DispItems3
        For Each field2 In DispItems4
            boxName = CStr(field) + CStr(field2)
            Call show(boxName, "", TOPSLIDE + 1)
        Next
    Next
End Sub

Sub ClearAllTextBoxesx(dummy As String)
    Dim sld As slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Text = ""
                End If
            End If
        Next shp
    Next sld


End Sub

Sub ���[�����쐬()
    Dim ss As formServerSelect

   

    Set ss = New formServerSelect
    ss.txtBoxServerName = DefaultServerName
    ss.show
End Sub




Function if_not_null(obj As Variant) As Integer
    If IsNull(obj) Then
        if_not_null = 0
    Else
        if_not_null = obj
    End If
End Function

Function if_not_null_string(obj As Variant) As String
    If IsNull(obj) Then
        if_not_null_string = ""
    Else
        if_not_null_string = obj
    End If
End Function






Sub get_race_title(ByVal prgNo As Integer, ByRef Class As String, _
            ByRef genderStr As String, ByRef distance As String, ByRef styleNo As Integer)
    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String

     
    
    If formEventNoPick.class_based_race("") Then
        myQuery = "SELECT �N���X.�N���X���� as �N���X, �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �N���X.���ԍ� = " & EventNo & " AND " & _
              " �v���O����.���Z�ԍ� = " & prgNo & ";"
    Else
        myQuery = "SELECT  �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �v���O����.���Z�ԍ� = " & prgNo & ";"
    End If
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        If classBasedRace Then
            Class = myRecordset!�N���X
        Else
            Class = ""
        End If
        genderStr = gender(myRecordset!����)
        distance = myRecordset!����
        styleNo = myRecordset!���
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
                
              
    
End Sub





Function RelayDistance(distance As String) As String
    If distance = " 200m" Then
        RelayDistance = " 4�~50m"
        Exit Function
    End If
    If distance = " 400m" Then
        RelayDistance = " 4�~100m"
        Exit Function
    End If
    If distance = " 800m" Then
        RelayDistance = " 4�~200m"
        Exit Function
    End If
End Function





Sub ShowSlide(slideNo As Integer)
    Dim ssw As SlideShowWindow
    Dim pres As Presentation

    Set pres = ActivePresentation

    ' ���łɃX���C�h�V���[�����s�����ǂ����m�F
    On Error Resume Next
    Set ssw = pres.SlideShowWindow
    On Error GoTo 0

    ' �X���C�h�V���[���N�����Ă��Ȃ��ꍇ�͊J�n
    If ssw Is Nothing Then
        With pres.SlideShowSettings
            .StartingSlide = 1
            .EndingSlide = pres.Slides.Count
            .Run
        End With
        Set ssw = pres.SlideShowWindow
    End If

    ' �w��̃X���C�h�ԍ��ֈړ�
    ssw.View.GotoSlide slideNo
End Sub
'���[����
Sub FillNewRecords(prgNo As Integer, classNo As Integer, gender As Integer, _
                   distance As Integer, style As Integer)
    Dim myQuery As String

    Dim myRecordset As New ADODB.Recordset

    myQuery = _
"    with base as ( " & _
"    select   " & _
"      �L�^����, " & _
"      case �V�L�^.�L�^�敪�ԍ� " & _
"        when 0 then '' " & _
"        else �N���X���� " & _
"      end as �V�L�^�敪, "
myQuery = myQuery & _
"      �����R�[�h, " & _
"      ��ڃR�[�h, " & _
"      �L�^, " & _
"      ���t, " & _
"      �L�^�ێ���, " & _
"      ���� " & _
"     from �V�L�^ "
myQuery = myQuery & _
"    inner join �V�L�^���� on �V�L�^����.�L�^�敪�ԍ�=�V�L�^.�L�^�敪�ԍ�" & _
"                        and  �V�L�^����.�L�^���̔ԍ�=�V�L�^.�L�^���̔ԍ�" & _
"                and  �V�L�^����.���ԍ�=�V�L�^.���ԍ�" & _
"    LEFT join �N���X on �N���X.�N���X�ԍ�=�V�L�^.�L�^�敪�ԍ�" & _
"             and �N���X.�N���X�ԍ�=" & classNo & _
"             and �N���X.���ԍ�=�V�L�^.���ԍ�"
myQuery = myQuery & _
"     where �V�L�^.���ԍ�=" & EventNo & _
"       and �����R�[�h=" & distance & _
"       and ��ڃR�[�h=" & style & _
"       and ���ʃR�[�h= " & gender & " )" & _
"       select  * from base where base.�V�L�^�敪 is not null"


    Dim slideNo As Integer
    Debug.Print ("" + myQuery)
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    Dim dateStr As String
 
    Do Until myRecordset.EOF
        dateStr = Left(myRecordset("���t"), 4)
        slideNo = TOPSLIDE + 1
            If myRecordset("�L�^����") = "���{" Then
                Call show("���{label", "���{�L�^", slideNo)
                Call show("���{�L�^", myRecordset("�L�^"), slideNo)
                Call show("���{���t", dateStr, slideNo)
                Call show("���{�ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("���{����", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "�w��" Then
                Call show("�w��label", "�w���L�^", slideNo)
                Call show("�w���L�^", myRecordset("�L�^"), slideNo)
                Call show("�w�����t", dateStr, slideNo)
                Call show("�w���ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("�w������", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "���Z" Then
                Call show("���Zlabel", "���Z�L�^", slideNo)
                Call show("���Z�L�^", myRecordset("�L�^"), slideNo)
                Call show("���Z���t", dateStr, slideNo)
                Call show("���Z�ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("���Z����", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "���w" Then
                Call show("���wlabel", "���w�L�^", slideNo)
                Call show("���w�L�^", myRecordset("�L�^"), slideNo)
                Call show("���w���t", dateStr, slideNo)
                Call show("���w�ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("���w����", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "�w��" Then
                Call show("�w��label", "�w���L�^", slideNo)
                Call show("�w���L�^", myRecordset("�L�^"), slideNo)
                Call show("�w�����t", dateStr, slideNo)
                Call show("�w���ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("�w������", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "���ꌧ" Then
                Call show("��label", "���L�^", slideNo)
                Call show("���L�^", myRecordset("�L�^"), slideNo)
                Call show("�����t", dateStr, slideNo)
                Call show("���ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("������", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "�����Z" Then
                Call show("�����Zlabel", "�����Z�L�^", slideNo)
                Call show("�����Z�L�^", myRecordset("�L�^"), slideNo)
                Call show("�����Z���t", dateStr, slideNo)
                Call show("�����Z�ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("�����Z����", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "�����w" Then
                Call show("�����wlabel", "�����w�L�^", slideNo)
                Call show("�����w�L�^", myRecordset("�L�^"), slideNo)
                Call show("�����w���t", dateStr, slideNo)
                Call show("�����w�ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("�����w����", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "���w��" Then
                Call show("���w��label", "���w���L�^", slideNo)
                Call show("���w���L�^", myRecordset("�L�^"), slideNo)
                Call show("���w�����t", dateStr, slideNo)
                Call show("���w���ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("���w������", myRecordset("����"), slideNo)
            End If
            If myRecordset("�L�^����") = "���" Then
                Call show("���label", "���L�^", slideNo)
                Call show("���L�^", myRecordset("�L�^"), slideNo)
                Call show("�����t", dateStr, slideNo)
                Call show("���ێ���", myRecordset("�L�^�ێ���"), slideNo)
                Call show("����", myRecordset("����"), slideNo)
            End If

        myRecordset.MoveNext
    Loop
    myRecordset.Close
    'MyCon.Close
    Set myRecordset = Nothing

    'Set MyCon = Nothing


End Sub
Function GetZeroUse(prgNo As Integer) As Boolean
    Dim myQuery As String

    Dim myRecordset As New ADODB.Recordset

    myQuery = "select ���H  from �L�^ " & _
              " inner join �v���O���� on �v���O����.���Z�ԍ�=�L�^.���Z�ԍ� " & _
              "        and �v���O����.���ԍ�=�L�^.���ԍ� " & _
              "  where �\���p���Z�ԍ� = " & prgNo & " And �L�^.���ԍ� = " & EventNo
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    ZeroUse = False
    Do Until myRecordset.EOF
        If myRecordset("���H") = 10 Then
            GetZeroUse = True
        End If
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    'MyCon.Close
    Set myRecordset = Nothing
End Function

Sub FillAll(prgNo As Integer)

    Call setDispItems("")
    Call setDispItems2("")
    Call setDispItems3("")
    Call setDispItems4("")
    Call ClearAllTextBoxes(0)

    Call ClearAllRecordTextBoxes("")
    Call FillOutLaneInfo(prgNo)
End Sub
'���[����
Sub FillOutLaneInfo(prgNo As Integer)
    Dim myQuery As String

    Dim myRecordset As New ADODB.Recordset
    ZeroUse = GetZeroUse(prgNo)

    myQuery = " SELECT �\���p���Z�ԍ�," & _
    "    case �v���O����.���ʃR�[�h " & _
    "        when 1 then '�j�q'" & _
    "        when 2 then '���q'" & _
    "        when 3 then '����'" & _
    "        when 4 then '����'" & _
    "      end as ����, " & _
    "    ����," & _
    "    ���," & _
    "    �g, " & _
    "    ���H, " & _
    "    �`�[����, " & _
    "    �I��.���� as ����, " & _
    "    �I��.��������1 as ����, "
    myQuery = myQuery & _
    "    ���.��ڃR�[�h, " & _
    "    ����.�����R�[�h, " & _
    "    �v���O����.���ʃR�[�h, " & _
    "    �L�^.�V�L�^����N���X, " & _
    "    �I��1.���� as ����1, " & _
    "    �I��2.���� as ����2, " & _
    "    �I��3.���� as ����3, " & _
    "    �I��4.���� as ����4, " & _
    "    �N���X���� as �N���X, " & _
    "    �\�I�^�C��"
    myQuery = myQuery & _
    "  FROM �L�^" & _
    "  INNER JOIN �v���O���� " & _
    "           ON �v���O����.���ԍ� = �L�^.���ԍ�" & _
    "          AND �v���O����.���Z�ԍ� = �L�^.���Z�ԍ�" & _
    "  LEFT JOIN �I�� on �I��.���ԍ�=�L�^.���ԍ� " & _
    "    AND �I��.�I��ԍ�=�L�^.�I��ԍ�" & _
    "  INNER JOIN ���� on ����.�����R�[�h=�v���O����.�����R�[�h" & _
    "  INNER JOIN ��� on ���.��ڃR�[�h=�v���O����.��ڃR�[�h LeFt join �����[�`�[�� " & _
    "            on �����[�`�[��.���ԍ� = �L�^.���ԍ� " & _
    "           and �����[�`�[��.�`�[���ԍ�=�L�^.�I��ԍ� " & _
    "  LEFT JOIN �I�� as �I��1" & _
    "            ON �I��1.���ԍ� = �L�^.���ԍ�" & _
    "           AND �I��1.�I��ԍ� = �L�^.��P�j��" & _
    "  LEFT JOIN �I�� as �I��2" & _
    "            ON �I��2.���ԍ� = �L�^.���ԍ�" & _
    "           AND �I��2.�I��ԍ� = �L�^.��Q�j��"
    
    myQuery = myQuery & _
    "  left JOIN �I�� as �I��3" & _
    "         ON �I��3.���ԍ� = �L�^.���ԍ�" & _
    "        AND �I��3.�I��ԍ� = �L�^.��R�j��" & _
    "  left JOIN �I�� as �I��4" & _
    "         ON �I��4.���ԍ� = �L�^.���ԍ�" & _
    "        AND �I��4.�I��ԍ� = �L�^.��S�j��" & _
    "  LEFT JOIN �N���X " & _
    "         ON �N���X.���ԍ� = �L�^.���ԍ� " & _
    "       AND �N���X.�N���X�ԍ� = �L�^.�V�L�^����N���X" & _
    "  WHERE �L�^.���ԍ� = " & EventNo & " and " & _
    "         �v���O����.�\���p���Z�ԍ�=" & prgNo & _
    "  ORDER BY  ���H "

    Dim field As Variant
    Dim laneNo As Integer
    Dim first As Boolean
    first = True
    Dim classNo As Integer
    Dim distanceCode As Integer
    Dim genderCode As Integer
    Dim shumokuCode As Integer


    Dim relayMember As String
    Dim slideNo As Integer
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        If myRecordset("�V�L�^����N���X") > 0 Then
            classNo = myRecordset("�V�L�^����N���X")
        End If
        If first Then
            first = False
            
            distanceCode = myRecordset("�����R�[�h")
            genderCode = myRecordset("���ʃR�[�h")
            shumokuCode = myRecordset("��ڃR�[�h")
            If shumokuCode > 5 Then
                TOPSLIDE = 13  ' relay
            Else
                TOPSLIDE = 1 ' individual
            End If
            For slideNo = TOPSLIDE To TOPSLIDE + 1
                Call show("PN", myRecordset("�\���p���Z�ԍ�"), slideNo)
                Call show("����", myRecordset("����"), slideNo)
                Call show("�N���X", if_not_null_string(myRecordset("�N���X")), slideNo)
                If shumokuCode > 5 Then
                    Call show("����", RelayDistance(myRecordset("����")), slideNo)
                Else
                    
                    Call show("����", myRecordset("����"), slideNo)
                End If
                Call show("���", myRecordset("���"), slideNo)
            Next slideNo

        End If


        '-------------  slide 1 �� slide 3�`12�ɍ�������-----------------------
        laneNo = myRecordset("���H")
        If ZeroUse Then
            laneNo = laneNo - 1
        End If
        If laneNo < 12 Then
            For Each field In DispItems2
                If CStr(field) = "���H" Then
                    Call show2(field, laneNo, laneNo)
                ElseIf CStr(field) = "����" Then
                    If shumokuCode > 5 Then
                        Call show2("����", RelayDistance(myRecordset("����")), laneNo)   '<--relaydistance
                    Else
                        Call show2("����", myRecordset("����"), laneNo)
                    
                    End If
                Else
                    Call show2(field, if_not_null_string(myRecordset(field)), laneNo)
                End If
            Next
        End If

        
        
        myRecordset.MoveNext
    Loop
                    ' �N���[�Y�Ɖ��
    myRecordset.Close
    'MyCon.Close
    Set myRecordset = Nothing
    'Set MyCon = Nothing
    Call FillNewRecords(prgNo, classNo, genderCode, distanceCode, shumokuCode)
End Sub




Sub name_text_box(boxNo As Integer, myName As String)
    Dim slide As slide
    Set slide = ActivePresentation.Slides(1)
    slide.Shapes(boxNo).Name = myName
End Sub
Sub show(ByVal txtBoxName As String, dispText As String, slideNo As Integer)

    ' �X���C�h�̎擾
    Dim slide As slide
    Dim shp As Shape
    Dim shapeExists As Boolean
    If slideNo = TOPSLIDE And txtBoxName = "����" Then
        Debug.Print (">> " + dispText)
    End If
    On Error Resume Next
    Set slide = ActivePresentation.Slides(slideNo) ' was slideIndex

     Set shp = slide.Shapes(txtBoxName)
     shapeExists = Not shp Is Nothing
    On Error GoTo 0
    If shapeExists Then
        slide.Shapes(txtBoxName).TextFrame.TextRange = dispText
    End If
End Sub
' show2 ---
'  topSlide (1 or 13) �́@textbox�� 3�`12 �������́A 15�`24��slide��textbox�ɍ������݂���
'
'  args :
'     txtBoxName :  textBox �̖��O
'     dispText  :   ����textBox �ɓ���镶��
'     laneNo  :  ���[��No.
' Global Variable that is used:
'     TOPSLIDE
Sub show2(ByVal txtBoxName As String, ByVal dispText As String, ByVal laneNo As Integer)

    ' �X���C�h�̎擾
    Dim slide As slide
    Dim shp As Shape
    Dim shapeExists As Boolean
    Dim myTextBoxName As String
   

    myTextBoxName = txtBoxName & laneNo
    On Error Resume Next
    Set slide = ActivePresentation.Slides(TOPSLIDE) ' top slide

    Set shp = slide.Shapes(myTextBoxName)
    shapeExists = Not shp Is Nothing
'    On Error GoTo 0
    If shapeExists Then
        slide.Shapes(myTextBoxName).TextFrame.TextRange = dispText
    End If
    '--- �elane
    Dim laneSlide As Integer

    laneSlide = laneNo + 2 + TOPSLIDE


    Set slide = ActivePresentation.Slides(laneSlide)
    Set shp = slide.Shapes(txtBoxName)
    shapeExists = Not shp Is Nothing

    If shapeExists Then
        slide.Shapes(txtBoxName).TextFrame.TextRange = dispText
    End If
    On Error GoTo 0
End Sub


Sub InitTextBox()
    Dim sld As slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                shp.TextFrame.TextRange.Text = shp.Name
            End If
        Next shp
    Next sld

End Sub
Sub DisplayTextBoxName(ByVal txtBoxName As String)
    Dim sld As slide
    Dim shp As Shape
    Dim shapeExists As Boolean

    ' ���ׂẴX���C�h�����ɏ���
    For Each sld In ActivePresentation.Slides
        On Error Resume Next
        Set shp = sld.Shapes(txtBoxName)
        shapeExists = Not shp Is Nothing
        On Error GoTo 0
        
        If shapeExists Then
            ' TextBox�̖��O��TextRange�ɐݒ�
            shp.TextFrame.TextRange.Text = txtBoxName
        End If
        
        ' ���̃X���C�h��
        Set shp = Nothing
    Next sld
End Sub






