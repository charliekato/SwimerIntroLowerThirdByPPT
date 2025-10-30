VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrgNoPick 
   Caption         =   "競技選択"
   ClientHeight    =   7344
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   8856
   OleObjectBlob   =   "formPrgNoPick.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formPrgNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'
'  formPrgNoPick
'
Public LastRow As Integer





Private Sub btnAllLane_Click()
     LaneOrder.ShowSlide (LaneOrder.TOPSLIDE)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnRecord_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 1)
End Sub
Private Sub btmLane0_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 2)
End Sub
Private Sub btnLane1_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 3)
End Sub

Private Sub btnLane2_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 4)
End Sub

Private Sub btnLane3_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 5)
End Sub
Private Sub btnLane4_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 6)
End Sub
Private Sub btnLane5_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 7)
End Sub
Private Sub btnLane6_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 8)
End Sub
Private Sub btnLane7_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 9)
End Sub
Private Sub btnLane8_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 10)
End Sub
Private Sub btnLane9_Click()
    LaneOrder.ShowSlide (LaneOrder.TOPSLIDE + 11)
End Sub


Private Sub btnPreView_Click()
    Dim printPrgNo As Integer
 '   On Error GoTo subEnd

    printPrgNo = CInt(Left(listPrg.Value, 3))
    Call LaneOrder.FillAll(printPrgNo)

subEnd:
End Sub




Private Sub listPrg_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnPreView_Click
    End If
End Sub


