Attribute VB_Name = "Module13"
Sub x�`�F�b�N�{�b�N�X�̒l����Âo�͂����������Ńe�L�X�g�{�b�N�X�ɏ���������Ȃ��ꍇ�̓{�b�N�X���폜()
    

Dim ppApp As New PowerPoint.Application
ppApp.Visible = True

Dim ppPrs As PowerPoint.Presentation
Set ppPrs = ppApp.Presentations.Open(ThisWorkbook.Path & "\���q���C�A�E�g�x�[�X�i�����Łj12.17-3.pptx")

Dim countSld As Long '�X���C�h��

Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("�f�[�^")

Dim i As Long
i = 2

Do While ws.Cells(i, 1).Value <> ""

    countSld = ppPrs.Slides.Count '���݂̃X���C�h�����J�E���g
    If countSld Mod 2 = 0 Then
        ppPrs.Slides(2).Duplicate.MoveTo toPos:=countSld + 1
    Else
        ppPrs.Slides(1).Duplicate.MoveTo toPos:=countSld + 1
    End If
    ppPrs.Slides(countSld + 1).Shapes("�y�[�W").TextFrame.TextRange.Text = ws.Cells(i, 1).Value '���Ə���
    ppPrs.Slides(countSld + 1).Shapes("���Ə���").TextFrame.TextRange.Text = ws.Cells(i, 2).Value '���Ə���
    ppPrs.Slides(countSld + 1).Shapes("�ꌾ���b�Z�[�W").TextFrame.TextRange.Text = ws.Cells(i, 4).Value '���b�Z�[�W
    ppPrs.Slides(countSld + 1).Shapes("�����^�C�g��").TextFrame.TextRange.Text = ws.Cells(i, 5).Value '�����^�C�g��
    ppPrs.Slides(countSld + 1).Shapes("�������e").TextFrame.TextRange.Text = ws.Cells(i, 6).Value '�������e
    ppPrs.Slides(countSld + 1).Shapes("���Ə���2").TextFrame.TextRange.Text = ws.Cells(i, 2).Value '���Ə���2
    ppPrs.Slides(countSld + 1).Shapes("�X�֔ԍ�").TextFrame.TextRange.Text = ws.Cells(i, 7).Value '�X�֔ԍ�
    ppPrs.Slides(countSld + 1).Shapes("�Z��").TextFrame.TextRange.Text = ws.Cells(i, 8).Value '�Z��
    ppPrs.Slides(countSld + 1).Shapes("������").TextFrame.TextRange.Text = ws.Cells(i, 9).Value '������
    ppPrs.Slides(countSld + 1).Shapes("�d�b�ԍ�").TextFrame.TextRange.Text = ws.Cells(i, 10).Value '�d�b�ԍ�
    If IsEmpty(ws.Cells(i, 11).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("���[���A�h���X�A�C�R��").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���[���A�h���X").Visible = msoFalse
    Else
        ppPrs.Slides(countSld + 1).Shapes("���[���A�h���X").TextFrame.TextRange.Text = ws.Cells(i, 11).Value '���[���A�h���X
    End If
    
    ppPrs.Slides(countSld + 1).Shapes("�Ŋ��w").TextFrame.TextRange.Text = ws.Cells(i, 13).Value '�Ŋ�w�P
    ppPrs.Slides(countSld + 1).Shapes("�Ŋ��w2").TextFrame.TextRange.Text = ws.Cells(i, 14).Value '�Ŋ�w2
    ppPrs.Slides(countSld + 1).Shapes("�J�n����").TextFrame.TextRange.Text = Format(ws.Cells(i, 15).Value, "hh:mm") '�J�n����
    ppPrs.Slides(countSld + 1).Shapes("�I������").TextFrame.TextRange.Text = Format(ws.Cells(i, 16).Value, "hh:mm") '�I������
    ppPrs.Slides(countSld + 1).Shapes("�J���j��").TextFrame.TextRange.Text = ws.Cells(i, 18).Value '�J����
    
    Dim tmp As Variant '���Ə����
    tmp = Split(ws.Cells(i, 3).Value, ", ")
    If IsEmpty(tmp) Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 0 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 1 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").TextFrame.TextRange.Text = tmp(3)
    End If
    
    If UBound(tmp) = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 5 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 6 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���Ə����9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 7 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����8").TextFrame.TextRange.Text = tmp(7)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 8 Then
        ppPrs.Slides(countSld + 1).Shapes("���Ə����1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����8").TextFrame.TextRange.Text = tmp(7)
        ppPrs.Slides(countSld + 1).Shapes("���Ə����9").TextFrame.TextRange.Text = tmp(8)
    End If
    
    Dim tmp2 As Variant '��Q�Ҏ��
    tmp2 = Split(ws.Cells(i, 22).Value, ", ")
    
    If IsEmpty(tmp2) Then
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 0 Then
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 1 Then
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��4").TextFrame.TextRange.Text = tmp2(3)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��4").TextFrame.TextRange.Text = tmp2(3)
        ppPrs.Slides(countSld + 1).Shapes("��Q�Ҏ��5").TextFrame.TextRange.Text = tmp2(4)
    End If
    
    '���}�͈͏���
    If IsEmpty(ws.Cells(i, 19).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("���}�͈̓A�C�R��").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���}�͈́F(���x��)").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���}�͈�").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���}�͈̓A�C�R��(���x��)").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���}�͈̓A�C�R��(�g)").Visible = msoFalse
    Else
        ppPrs.Slides(countSld + 1).Shapes("���}�͈�").TextFrame.TextRange.Text = ws.Cells(i, 19).Value '���}�͈�
    End If
    
    '��ÃP�A����
    If IsEmpty(ws.Cells(i, 20).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("��ÃA�C�R��").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("��ÃA�C�R��(�g)").Visible = msoFalse
    End If
    
    '���H����
    If IsEmpty(ws.Cells(i, 21).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("���H�A�C�R��").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("���H�A�C�R��(�g)").Visible = msoFalse
    End If
    
    '��O��������
    If IsEmpty(ws.Cells(i, 17).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("��O����").Visible = msoFalse
    End If
    
    '�w�\��
    If ws.Cells(i, 27).Value = 1 Then
        'ppPrs.Slides(countSld + 1).Shapes("���}�͈�").TextFrame.TextRange.Text = ws.Cells(i, 19).Value
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 5 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 6 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 7 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 8 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 9 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 10 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 11 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 12 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 13 Then
        ppPrs.Slides(countSld + 1).Shapes("�w�\��1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("�w�\��12").Visible = msoFalse
    End If
    
    
    i = i + 1
Loop

'ppApp.Quit
'Set ppApp = Nothing
End Sub



