Attribute VB_Name = "ModGraphPaper"
Option Explicit

'MakeGraphPaper�E�E�E���ꏊ�FFukamiAddins3.ModGraphPaper
'PxToWidth     �E�E�E���ꏊ�FFukamiAddins3.ModGraphPaper



Sub MakeGraphPaper(TargetSheet As Worksheet, InputPx As Long, Optional MessageIrunaraTrue As Boolean = False)
'����
'TargetSheet            �E�E�E�Ώۂ̃V�[�g
'InputPx                �E�E�E�}�X���i�������j�̃s�N�Z���l
'[MessageIrunaraTrue]   �E�E�E���ᎆ�쐬��Ƀ��b�Z�[�W��\�����邩�ǂ���

    Dim SetHeight As Double
    Dim SetWidth  As Double
    
    SetHeight = 0.6 * InputPx
    
    SetWidth = PxToWidth(InputPx)
    
    With TargetSheet.Cells
        .ColumnWidth = SetWidth
        .RowHeight = SetHeight
    End With
     
    If MessageIrunaraTrue Then
        MsgBox ("����������" & InputPx & "�s�N�Z��" & vbLf & _
                "��(ColumnWidth):" & SetWidth & "�|�C���g" & vbLf & _
                "��(Width):" & Range("A1").Width & vbLf & _
                "����(RowHeight):" & SetHeight & "�|�C���g" & vbLf & _
                "����(Height):" & Range("A1").Height)
    End If
    
End Sub

Private Function PxToWidth(Px As Long)
'�s�N�Z���l�𕝂ɕϊ�����

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


