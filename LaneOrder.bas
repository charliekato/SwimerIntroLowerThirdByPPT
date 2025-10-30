Attribute VB_Name = "LaneOrder"
' slide-1 ～ 12 個人種目用
' slide-13～ 24 リレー種目用
'
' slide-1　　　全レーンの選手紹介用
'　　　個人種目に対応
'
'
'
' slide-2　　　新記録一覧
'
'
' slide-3～12 レーン毎の選手紹介
'
' slide-3 : 0レーン　　　slide-4 : 1レーン      slide-5 : 2レーン .....   slide-12 : 9レーン
'
'
' slide-13    全レーンの選手紹介
'     リレー種目に対応
'
'
'
' slide-14　　　新記録一覧
'
'
' slide-15～24 レーン毎の選手紹介
'
' slide-15 : 0レーン　　　slide-16 : 1レーン      slide-17 : 2レーン .....   slide-24 : 9レーン
'
'
'
' 使用するtextboxの名前
'
' 全slide共通のtextbox (見出しなどに使用)
'
'    　PN(競技番号)、性別、クラス、距離、種目
'
'   注意!! 同じ名前の textBoxがあった場合　エラーにはならず、先にみつかったtextBoxに差し込みされます。
'
'
' slide-1
'
' 水路0 , 氏名0 , 所属0 , クラス0 , 予選タイム0 ,　　<--末尾の数字はレーン番号
' 水路1 , 氏名1 , 所属1 , クラス1 , 予選タイム1 ,
' 水路2 , 氏名2 , 所属2 , クラス2 , 予選タイム2 ,
' ...
' 水路9 , 氏名9 , 所属9 , クラス9 , 予選タイム9 ,
'
' 水路0, 水路1, 水路2, ..., 水路9  は、 レーン番号で 水路0の中身は0になる。
' 選手のいないところは表示されない
'
'slide-2.
'
'
' 世界label , 世界記録 , 世界所持者, 世界所属  , 世界日付 ,
' 日本label , 日本記録 , 日本所持者, 日本所属  , 日本日付 ,
' 学生label , 学生記録 , 学生所持者, 学生所属  , 学生日付 ,
' 高校label , 高校記録 , 高校所持者, 高校所属  , 高校日付 ,
' 中学label , 中学記録 , 中学所持者, 中学所属  , 中学日付,
' 学童label , 学童記録 , 学童所持者, 学童所属  , 学童日付,
' 県label , 県記録 , 県所持者 , 県所属 , 県日付 ,
' 県高校label , 県高校記録 , 県高校所持者 , 県高校所属 , 県高校日付 ,
' 県中学label , 県中学記録 , 県中学所持者 , 県中学所属 , 県中学日付 ,
' 県学童label , 県学童記録 , 県学童所持者 , 県学童所属 , 県学童日付 ,
'
'
'スライド3～12
'
'水路、氏名、所属、予選タイム
'
'slide-13   slide-13以降はリレー用
'
' 水路0 , 氏名10 , 氏名20, 氏名30, 氏名40, チーム名0 , クラス0 , 予選タイム0 ,
' 水路1 , 氏名11 , 氏名21, 氏名31, 氏名41, チーム名1 , クラス1 , 予選タイム1 ,
' 水路2 , 氏名12 , 氏名21, 氏名32, 氏名42, チーム名2 , クラス2 , 予選タイム2 ,
' ...
' 水路9 , 氏名19 , 氏名29 ,氏名39 ,氏名49, チーム名9 , クラス9 , 予選タイム9 ,
'
'    末尾の数字はレーン番号。第一泳者 は　氏名1
'
'
'slide-14.
' slide-2 と同じ所属にはチーム名がはいり、所持者には4人の名前が入る
'
' 世界label , 世界記録 , 世界所持者, 世界所属  , 世界日付 ,
' 日本label , 日本記録 , 日本所持者, 日本所属  , 日本日付 ,
' 学生label , 学生記録 , 学生所持者, 学生所属  , 学生日付 ,
' 高校label , 高校記録 , 高校所持者, 高校所属  , 高校日付 ,
' 中学label , 中学記録 , 中学所持者, 中学所属  , 中学日付,
' 学童label , 学童記録 , 学童所持者, 学童所属  , 学童日付,
' 県label , 県記録 , 県所持者 , 県所属 , 県日付 ,
' 県高校label , 県高校記録 , 県高校所持者 , 県高校所属 , 県高校日付 ,
' 県中学label , 県中学記録 , 県中学所持者 , 県中学所属 , 県中学日付 ,
' 県学童label , 県学童記録 , 県学童所持者 , 県学童所属 , 県学童日付 ,

'
'スライド15～24
'
'水路、氏名1, 氏名2, 氏名3, 氏名4, 所属、予選タイム
'
'  slide 3～12と同じだが、選手名は 氏名ではなく、 氏名1(第一泳者)～ 氏名4(第四泳者)になる。








Option Explicit
Option Base 0
Const 選手紹介 As Boolean = False
Const DefaultServerName = "localhost"
Public ZeroUse As Boolean

Public TOPSLIDE As Integer   ' relay時 13, 個人種目 1
Dim DispItems As Variant
Dim DispItems2 As Variant
Dim DispItems3 As Variant ' 新記録用
Dim DispItems4 As Variant ' 新記録用
    Public MyCon As ADODB.Connection
    Public EventNo As Integer


    Public MaxClassNo As Integer

'    Public ClassTable() As String
Sub setDispItems(dummy As String)
    DispItems = Array("PN", "距離", "種目", "性別", "クラス")
End Sub
Sub setDispItems2(dummy As String)
    DispItems2 = Array("チーム名", "氏名1", "氏名2", "氏名3", "氏名4", "クラス", "性別", "予選タイム", _
    "氏名", "所属", "水路", "距離", "種目")
End Sub
Sub setDispItems3(dummy As String)
    DispItems3 = Array("世界", "日本", "学生", "高校", "中学", "学童", "県", "県高校", "県中学", "県学童", "大会")
End Sub
Sub setDispItems4(dummy As String)
    DispItems4 = Array("label", "記録", "保持者", "所属", "日付")
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

Sub レーン順作成()
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
        myQuery = "SELECT クラス.クラス名称 as クラス, プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN クラス ON クラス.クラス番号=プログラム.クラス番号 " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " クラス.大会番号 = " & EventNo & " AND " & _
              " プログラム.競技番号 = " & prgNo & ";"
    Else
        myQuery = "SELECT  プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " プログラム.競技番号 = " & prgNo & ";"
    End If
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        If classBasedRace Then
            Class = myRecordset!クラス
        Else
            Class = ""
        End If
        genderStr = gender(myRecordset!性別)
        distance = myRecordset!距離
        styleNo = myRecordset!種目
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
                
              
    
End Sub





Function RelayDistance(distance As String) As String
    If distance = " 200m" Then
        RelayDistance = " 4×50m"
        Exit Function
    End If
    If distance = " 400m" Then
        RelayDistance = " 4×100m"
        Exit Function
    End If
    If distance = " 800m" Then
        RelayDistance = " 4×200m"
        Exit Function
    End If
End Function





Sub ShowSlide(slideNo As Integer)
    Dim ssw As SlideShowWindow
    Dim pres As Presentation

    Set pres = ActivePresentation

    ' すでにスライドショーが実行中かどうか確認
    On Error Resume Next
    Set ssw = pres.SlideShowWindow
    On Error GoTo 0

    ' スライドショーが起動していない場合は開始
    If ssw Is Nothing Then
        With pres.SlideShowSettings
            .StartingSlide = 1
            .EndingSlide = pres.Slides.Count
            .Run
        End With
        Set ssw = pres.SlideShowWindow
    End If

    ' 指定のスライド番号へ移動
    ssw.View.GotoSlide slideNo
End Sub
'レーン順
Sub FillNewRecords(prgNo As Integer, classNo As Integer, gender As Integer, _
                   distance As Integer, style As Integer)
    Dim myQuery As String

    Dim myRecordset As New ADODB.Recordset

    myQuery = _
"    with base as ( " & _
"    select   " & _
"      記録名称, " & _
"      case 新記録.記録区分番号 " & _
"        when 0 then '' " & _
"        else クラス名称 " & _
"      end as 新記録区分, "
myQuery = myQuery & _
"      距離コード, " & _
"      種目コード, " & _
"      記録, " & _
"      日付, " & _
"      記録保持者, " & _
"      所属 " & _
"     from 新記録 "
myQuery = myQuery & _
"    inner join 新記録名称 on 新記録名称.記録区分番号=新記録.記録区分番号" & _
"                        and  新記録名称.記録名称番号=新記録.記録名称番号" & _
"                and  新記録名称.大会番号=新記録.大会番号" & _
"    LEFT join クラス on クラス.クラス番号=新記録.記録区分番号" & _
"             and クラス.クラス番号=" & classNo & _
"             and クラス.大会番号=新記録.大会番号"
myQuery = myQuery & _
"     where 新記録.大会番号=" & EventNo & _
"       and 距離コード=" & distance & _
"       and 種目コード=" & style & _
"       and 性別コード= " & gender & " )" & _
"       select  * from base where base.新記録区分 is not null"


    Dim slideNo As Integer
    Debug.Print ("" + myQuery)
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    Dim dateStr As String
 
    Do Until myRecordset.EOF
        dateStr = Left(myRecordset("日付"), 4)
        slideNo = TOPSLIDE + 1
            If myRecordset("記録名称") = "日本" Then
                Call show("日本label", "日本記録", slideNo)
                Call show("日本記録", myRecordset("記録"), slideNo)
                Call show("日本日付", dateStr, slideNo)
                Call show("日本保持者", myRecordset("記録保持者"), slideNo)
                Call show("日本所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "学生" Then
                Call show("学生label", "学生記録", slideNo)
                Call show("学生記録", myRecordset("記録"), slideNo)
                Call show("学生日付", dateStr, slideNo)
                Call show("学生保持者", myRecordset("記録保持者"), slideNo)
                Call show("学生所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "高校" Then
                Call show("高校label", "高校記録", slideNo)
                Call show("高校記録", myRecordset("記録"), slideNo)
                Call show("高校日付", dateStr, slideNo)
                Call show("高校保持者", myRecordset("記録保持者"), slideNo)
                Call show("高校所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "中学" Then
                Call show("中学label", "中学記録", slideNo)
                Call show("中学記録", myRecordset("記録"), slideNo)
                Call show("中学日付", dateStr, slideNo)
                Call show("中学保持者", myRecordset("記録保持者"), slideNo)
                Call show("中学所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "学童" Then
                Call show("学童label", "学童記録", slideNo)
                Call show("学童記録", myRecordset("記録"), slideNo)
                Call show("学童日付", dateStr, slideNo)
                Call show("学童保持者", myRecordset("記録保持者"), slideNo)
                Call show("学童所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "滋賀県" Then
                Call show("県label", "県記録", slideNo)
                Call show("県記録", myRecordset("記録"), slideNo)
                Call show("県日付", dateStr, slideNo)
                Call show("県保持者", myRecordset("記録保持者"), slideNo)
                Call show("県所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "県高校" Then
                Call show("県高校label", "県高校記録", slideNo)
                Call show("県高校記録", myRecordset("記録"), slideNo)
                Call show("県高校日付", dateStr, slideNo)
                Call show("県高校保持者", myRecordset("記録保持者"), slideNo)
                Call show("県高校所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "県中学" Then
                Call show("県中学label", "県中学記録", slideNo)
                Call show("県中学記録", myRecordset("記録"), slideNo)
                Call show("県中学日付", dateStr, slideNo)
                Call show("県中学保持者", myRecordset("記録保持者"), slideNo)
                Call show("県中学所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "県学童" Then
                Call show("県学童label", "県学童記録", slideNo)
                Call show("県学童記録", myRecordset("記録"), slideNo)
                Call show("県学童日付", dateStr, slideNo)
                Call show("県学童保持者", myRecordset("記録保持者"), slideNo)
                Call show("県学童所属", myRecordset("所属"), slideNo)
            End If
            If myRecordset("記録名称") = "大会" Then
                Call show("大会label", "大会記録", slideNo)
                Call show("大会記録", myRecordset("記録"), slideNo)
                Call show("大会日付", dateStr, slideNo)
                Call show("大会保持者", myRecordset("記録保持者"), slideNo)
                Call show("大会所属", myRecordset("所属"), slideNo)
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

    myQuery = "select 水路  from 記録 " & _
              " inner join プログラム on プログラム.競技番号=記録.競技番号 " & _
              "        and プログラム.大会番号=記録.大会番号 " & _
              "  where 表示用競技番号 = " & prgNo & " And 記録.大会番号 = " & EventNo
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    ZeroUse = False
    Do Until myRecordset.EOF
        If myRecordset("水路") = 10 Then
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
'レーン順
Sub FillOutLaneInfo(prgNo As Integer)
    Dim myQuery As String

    Dim myRecordset As New ADODB.Recordset
    ZeroUse = GetZeroUse(prgNo)

    myQuery = " SELECT 表示用競技番号," & _
    "    case プログラム.性別コード " & _
    "        when 1 then '男子'" & _
    "        when 2 then '女子'" & _
    "        when 3 then '混成'" & _
    "        when 4 then '混合'" & _
    "      end as 性別, " & _
    "    距離," & _
    "    種目," & _
    "    組, " & _
    "    水路, " & _
    "    チーム名, " & _
    "    選手.氏名 as 氏名, " & _
    "    選手.所属名称1 as 所属, "
    myQuery = myQuery & _
    "    種目.種目コード, " & _
    "    距離.距離コード, " & _
    "    プログラム.性別コード, " & _
    "    記録.新記録判定クラス, " & _
    "    選手1.氏名 as 氏名1, " & _
    "    選手2.氏名 as 氏名2, " & _
    "    選手3.氏名 as 氏名3, " & _
    "    選手4.氏名 as 氏名4, " & _
    "    クラス名称 as クラス, " & _
    "    予選タイム"
    myQuery = myQuery & _
    "  FROM 記録" & _
    "  INNER JOIN プログラム " & _
    "           ON プログラム.大会番号 = 記録.大会番号" & _
    "          AND プログラム.競技番号 = 記録.競技番号" & _
    "  LEFT JOIN 選手 on 選手.大会番号=記録.大会番号 " & _
    "    AND 選手.選手番号=記録.選手番号" & _
    "  INNER JOIN 距離 on 距離.距離コード=プログラム.距離コード" & _
    "  INNER JOIN 種目 on 種目.種目コード=プログラム.種目コード LeFt join リレーチーム " & _
    "            on リレーチーム.大会番号 = 記録.大会番号 " & _
    "           and リレーチーム.チーム番号=記録.選手番号 " & _
    "  LEFT JOIN 選手 as 選手1" & _
    "            ON 選手1.大会番号 = 記録.大会番号" & _
    "           AND 選手1.選手番号 = 記録.第１泳者" & _
    "  LEFT JOIN 選手 as 選手2" & _
    "            ON 選手2.大会番号 = 記録.大会番号" & _
    "           AND 選手2.選手番号 = 記録.第２泳者"
    
    myQuery = myQuery & _
    "  left JOIN 選手 as 選手3" & _
    "         ON 選手3.大会番号 = 記録.大会番号" & _
    "        AND 選手3.選手番号 = 記録.第３泳者" & _
    "  left JOIN 選手 as 選手4" & _
    "         ON 選手4.大会番号 = 記録.大会番号" & _
    "        AND 選手4.選手番号 = 記録.第４泳者" & _
    "  LEFT JOIN クラス " & _
    "         ON クラス.大会番号 = 記録.大会番号 " & _
    "       AND クラス.クラス番号 = 記録.新記録判定クラス" & _
    "  WHERE 記録.大会番号 = " & EventNo & " and " & _
    "         プログラム.表示用競技番号=" & prgNo & _
    "  ORDER BY  水路 "

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
        If myRecordset("新記録判定クラス") > 0 Then
            classNo = myRecordset("新記録判定クラス")
        End If
        If first Then
            first = False
            
            distanceCode = myRecordset("距離コード")
            genderCode = myRecordset("性別コード")
            shumokuCode = myRecordset("種目コード")
            If shumokuCode > 5 Then
                TOPSLIDE = 13  ' relay
            Else
                TOPSLIDE = 1 ' individual
            End If
            For slideNo = TOPSLIDE To TOPSLIDE + 1
                Call show("PN", myRecordset("表示用競技番号"), slideNo)
                Call show("性別", myRecordset("性別"), slideNo)
                Call show("クラス", if_not_null_string(myRecordset("クラス")), slideNo)
                If shumokuCode > 5 Then
                    Call show("距離", RelayDistance(myRecordset("距離")), slideNo)
                Else
                    
                    Call show("距離", myRecordset("距離"), slideNo)
                End If
                Call show("種目", myRecordset("種目"), slideNo)
            Next slideNo

        End If


        '-------------  slide 1 と slide 3～12に差し込み-----------------------
        laneNo = myRecordset("水路")
        If ZeroUse Then
            laneNo = laneNo - 1
        End If
        If laneNo < 12 Then
            For Each field In DispItems2
                If CStr(field) = "水路" Then
                    Call show2(field, laneNo, laneNo)
                ElseIf CStr(field) = "距離" Then
                    If shumokuCode > 5 Then
                        Call show2("距離", RelayDistance(myRecordset("距離")), laneNo)   '<--relaydistance
                    Else
                        Call show2("距離", myRecordset("距離"), laneNo)
                    
                    End If
                Else
                    Call show2(field, if_not_null_string(myRecordset(field)), laneNo)
                End If
            Next
        End If

        
        
        myRecordset.MoveNext
    Loop
                    ' クローズと解放
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

    ' スライドの取得
    Dim slide As slide
    Dim shp As Shape
    Dim shapeExists As Boolean
    If slideNo = TOPSLIDE And txtBoxName = "距離" Then
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
'  topSlide (1 or 13) の　textboxと 3～12 もしくは、 15～24のslideのtextboxに差し込みする
'
'  args :
'     txtBoxName :  textBox の名前
'     dispText  :   そのtextBox に入れる文字
'     laneNo  :  レーンNo.
' Global Variable that is used:
'     TOPSLIDE
Sub show2(ByVal txtBoxName As String, ByVal dispText As String, ByVal laneNo As Integer)

    ' スライドの取得
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
    '--- 各lane
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

    ' すべてのスライドを順に処理
    For Each sld In ActivePresentation.Slides
        On Error Resume Next
        Set shp = sld.Shapes(txtBoxName)
        shapeExists = Not shp Is Nothing
        On Error GoTo 0
        
        If shapeExists Then
            ' TextBoxの名前をTextRangeに設定
            shp.TextFrame.TextRange.Text = txtBoxName
        End If
        
        ' 次のスライドへ
        Set shp = Nothing
    Next sld
End Sub






