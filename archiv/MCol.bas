Attribute VB_Name = "MCol"
Option Explicit


Public Sub Col_SwapItems(col As Collection, ByVal i1 As Long, i2 As Long)
    Dim c As Long: c = col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or col.Count <= i1 Then Exit Sub
    If i2 <= 0 Or col.Count < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim obj1, obj2
    If IsObject(col.Item(i1)) Then Set obj1 = col.Item(i1) Else obj1 = col.Item(i1)
    If IsObject(col.Item(i2)) Then Set obj2 = col.Item(i2) Else obj2 = col.Item(i2)
    col.Remove i1: col.Add obj2, , i1:     col.Remove i2
    If i2 < c Then col.Add obj1, , i2 Else col.Add obj1
End Sub

Public Sub Col_MoveUp(col As Collection, ByVal i As Long)
    Col_SwapItems col, i, i - 1
End Sub

Public Sub Col_MoveDown(col As Collection, ByVal i As Long)
    Col_SwapItems col, i, i + 1
End Sub

