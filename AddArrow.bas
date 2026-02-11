Attribute VB_Name = "Module1"
Option Explicit

' 選択中セル(ActiveCell)の位置に、左右両矢印 + 文字 を作ってグループ化して挿入
Public Sub InsertArrowTextGroup_AtActiveCell()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ActiveCell Is Nothing Then Exit Sub

    Dim labelText As String
    labelText = InputBox("両矢印の中央に入れる文字を入力してください", "文字入力", "テキスト")

    If StrPtr(labelText) = 0 Then Exit Sub ' キャンセル
    ' 空文字も許容（不要なら If labelText = "" Then Exit Sub など）

    Dim anchor As Range
    Set anchor = ActiveCell

    ' サイズ（セルに合わせたいならここを調整）
    Dim w As Single, h As Single
    w = 260
    h = 42

    Dim leftPos As Single, topPos As Single
    leftPos = anchor.Left
    topPos = anchor.Top

    Dim shpArrow As Shape, shpText As Shape, grp As Shape
    Dim namesToGroup(1 To 2) As Variant
    Dim uid As String
    uid = Format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 100000))

    ' 1) 左右両矢印
    Set shpArrow = ws.Shapes.AddShape(msoShapeLeftRightArrow, leftPos, topPos, w, h)
    With shpArrow
        .Name = "LRArrow_" & uid
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(230, 230, 230) ' 塗り（好みで変更）
    End With

    ' 2) 文字（透明テキストボックス）
    Set shpText = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, w, h)
    With shpText
        .Name = "Label_" & uid
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse

        ' フォント調整（必要なら）
        ' .TextFrame2.TextRange.Font.Name = "Meiryo UI"
        ' .TextFrame2.TextRange.Font.Size = 11
    End With

    ' 3) グループ化
    namesToGroup(1) = shpArrow.Name
    namesToGroup(2) = shpText.Name
    Set grp = ws.Shapes.Range(namesToGroup).Group

    ' 4) 仕上げ
    With grp
        .Name = "ArrowTextGroup_" & uid
        .Left = leftPos
        .Top = topPos
        .Placement = xlMoveAndSize ' セル移動/サイズ変更に追従
    End With
End Sub

