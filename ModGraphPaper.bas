Attribute VB_Name = "ModGraphPaper"
Option Explicit

'MakeGraphPaper・・・元場所：FukamiAddins3.ModGraphPaper
'PxToWidth     ・・・元場所：FukamiAddins3.ModGraphPaper



Sub MakeGraphPaper(TargetSheet As Worksheet, InputPx As Long, Optional MessageIrunaraTrue As Boolean = False)
'引数
'TargetSheet            ・・・対象のシート
'InputPx                ・・・マス幅（＝高さ）のピクセル値
'[MessageIrunaraTrue]   ・・・方眼紙作成後にメッセージを表示するかどうか

    Dim SetHeight As Double
    Dim SetWidth  As Double
    
    SetHeight = 0.6 * InputPx
    
    SetWidth = PxToWidth(InputPx)
    
    With TargetSheet.Cells
        .ColumnWidth = SetWidth
        .RowHeight = SetHeight
    End With
     
    If MessageIrunaraTrue Then
        MsgBox ("幅高さ共に" & InputPx & "ピクセル" & vbLf & _
                "幅(ColumnWidth):" & SetWidth & "ポイント" & vbLf & _
                "幅(Width):" & Range("A1").Width & vbLf & _
                "高さ(RowHeight):" & SetHeight & "ポイント" & vbLf & _
                "高さ(Height):" & Range("A1").Height)
    End If
    
End Sub

Private Function PxToWidth(Px As Long)
'ピクセル値を幅に変換する

    Dim Output As Double
    If Px <= 4 Then
        Output = 0.06 * Px
    ElseIf Px = 5 Then
        Output = 0.29
    ElseIf Px <= 12 Then
        Output = 0.06 * (Px - 6) + 0.35
    ElseIf Px = 13 Then
        Output = 0.76
    ElseIf Px <= 17 Then
        Output = 0.06 * (Px - 14) + 0.82
    Else
        Output = 0.1 * (Px - 18) + 1.1
    End If
    
    PxToWidth = Output
    
End Function


