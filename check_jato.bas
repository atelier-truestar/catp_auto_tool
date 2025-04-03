Attribute VB_Name = "check_jato"
Option Explicit

Sub check_jato_list(ByVal button_no as Integer)

    Dim current_sht_name As String
    Dim result_body, result_wheel, result_seat, result_battery As String
    Dim check_filter_on As Boolean
    Dim check_filtermode as Boolean
    Dim insert_order As Integer
    
    '�Ō�A�e���ʂ�JATO�V�[�g�ɖ߂邽�߁A���݂̃V�[�g�����擾
    current_sht_name = ActiveSheet.Name
    '���݊J���Ă���V�[�g��ActiveSheet�ɐݒ�
    Worksheets(current_sht_name).Activate
    '�����̃{�^���ԍ����擾
    insert_order = button_no
    '�t�B���^�[���������Ă��邩�ǂ������m�F
    check_filter_on = ActiveSheet.AutoFilterMode
    check_filtermode = ActiveSheet.FilterMode
    '�ȉ��A�t�B���^�[�̏�Ԃŏ����𕪊�
    If check_filter_on = False or (check_filter_on=True and check_filtermode=False) Then
        '�t�B���^�[����������Ă���ꍇ�A�܂��̓I�[�g�t�B���^�[�����I���̏ꍇ�Alist_all()��Sub���R�[������
        Call list_all(current_sht_name, insert_order)
    Else
        '�t�B���^�[���������Ă���ꍇ�Alist_filter()��Sub���R�[������
        Call list_filter(current_sht_name, insert_order)
    End If
        

End Sub

Sub list_all(ByVal sht_tab_name as string,ByVal insert_order as integer)
    Dim last_row As Long
    Dim startRange, endRange As Range
    Dim targetRange As Range
    Dim filterEndRow As Long
    Dim UniqueList(3) As Variant
    Dim UniqueList_a(3) As Variant
    Dim RangeList(3) As String
    Dim ws As Worksheet
    Dim i as integer

    'RangeList��A2����D2�܂ł͈̔͂��i�[
    RangeList(0) = "A2"
    RangeList(1) = "B2"
    RangeList(2) = "C2"
    RangeList(3) = "D2"
    'Dim temp_sht_name As String
    '�e���|�����̃V�[�g����ɂ���
    ThisWorkbook.Sheets("JATO_WorkArea").Cells.Clear
    '���݊J���Ă���V�[�g��ActiveSheet�ɐݒ�
    Worksheets(sht_tab_name).Activate
    'A��̍ŏI�s���擾
    last_row = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    'E15����H��last_row�܂ł͈̔͂�I�����A�R�s�[
    ActiveSheet.Range("E15:H" & last_row).Select
    Selection.Copy
    '�e���|�����̃V�[�g�Ƀy�[�X�g
    ThisWorkbook.Sheets("JATO_WorkArea").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Set ws = ThisWorkbook.Sheets("JATO_WorkArea")
    ws.Activate
    '4���ڕ������[�v�ŏ���
    For i = 0 To 3
        '���X�g�̊J�n�Z�����擾
        Set startRange = ws.Range(RangeList(i))
        '���X�g�̍ŏI�s���擾
        filterEndRow = startRange.End(xlDown).Row
        '���X�g�̍ŏI�Z�����擾
        Set endRange = ws.Cells(filterEndRow, startRange.Column)
        '���X�g�͈̔͂��擾
        Set targetRange = ws.Range(startRange, endRange)
        '�d�������������X�g���쐬
        UniqueList(i) = WorksheetFunction.Unique(targetRange)
        'Uniquelist�̓��e�`�F�b�N
        If UBound(UniqueList(i)) = 1 Then
            If IsEmpty(UniqueList(i)(1)) Then
                UniqueList_a(i) = "None"
            Else: UniqueList_a(i) = UniqueList(i)
            End If
        Else: UniqueList_a(i) = "All"
        End If
        '���ʂ�cells(1, i+6)�ɑ��
        ws.Cells(1, i + 6) = UniqueList_a(i)
    Next i
    '�y�[�X�g�����͈͂��Z���N�g
    ws.Range("F1:I1").Select
    Selection.Copy
    Worksheets(sht_tab_name).Activate
    'ActiveSheet��E��Ainsert_order+8�s�ڂɃy�[�X�g
    ActiveSheet.Cells(insert_order + 8, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'MSRP���ʃZ���Ɉړ�
    ActiveSheet.Range("O10").Select
    Selection.Copy
    'AcrtiveSheet��I��Ainsert_order+8�s�ڂɃy�[�X�g
    ActiveSheet.Cells(insert_order + 8, 9).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '�R�s�[���w�����
    Application.CutCopyMode = False
end Sub

Sub list_filter(ByVal sht_tab_name As String, ByVal insert_order As Integer)
    Dim last_row As Long
    Dim startRange, endRange As Range
    Dim targetRange As Range
    Dim filterEndRow As Long
    Dim UniqueList(3) As Variant
    Dim UniqueList_a(3) As Variant
    Dim RangeList(3) As String
    Dim ws As Worksheet
    Dim i as integer

    'RangeList��A2����D2�܂ł͈̔͂��i�[
    RangeList(0) = "A2"
    RangeList(1) = "B2"
    RangeList(2) = "C2"
    RangeList(3) = "D2"
    'Dim temp_sht_name As String
    '�e���|�����̃V�[�g����ɂ���
    ThisWorkbook.Sheets("JATO_WorkArea").Cells.Clear
    '���݊J���Ă���V�[�g��ActiveSheet�ɐݒ�
    Worksheets(sht_tab_name).Activate
    'A��̍ŏI�s���擾
    last_row = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    'E15����H��last_row�܂ł͈̔͂�I�����A�R�s�[
    ActiveSheet.Range("E15:H" & last_row).Select
    Selection.Copy
    '�e���|�����̃V�[�g�Ƀy�[�X�g
    ThisWorkbook.Sheets("JATO_WorkArea").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Set ws = ThisWorkbook.Sheets("JATO_WorkArea")
    ws.Activate
    '4���ڕ������[�v�ŏ���
    For i = 0 To 3
        '���X�g�̊J�n�Z�����擾
        Set startRange = ws.Range(RangeList(i))
        '���X�g�̍ŏI�s���擾
        filterEndRow = startRange.End(xlDown).Row
        '���X�g�̍ŏI�Z�����擾
        Set endRange = ws.Cells(filterEndRow, startRange.Column)
        '���X�g�͈̔͂��擾
        Set targetRange = ws.Range(startRange, endRange)
        '�d�������������X�g���쐬
        UniqueList(i) = WorksheetFunction.Unique(targetRange)
        'Uniquelist�̓��e�`�F�b�N
        If UBound(UniqueList(i)) = 1 Then
            If IsEmpty(UniqueList(i)(1)) Then
                UniqueList_a(i) = "None"
            Else: UniqueList_a(i) = UniqueList(i)
            End If
        Else
            UniqueList_a(i) = Join(WorksheetFunction.Transpose(UniqueList(i)), ",")
        End If
        '���ʂ�cells(1, i+6)�ɑ��
        ws.Cells(1, i + 6) = UniqueList_a(i)
    Next i
    '�y�[�X�g�����͈͂��Z���N�g
    ws.Range("F1:I1").Select
    Selection.Copy
    Worksheets(sht_tab_name).Activate
    'ActiveSheet��E��Ainsert_order+8�s�ڂɃy�[�X�g
    ActiveSheet.Cells(insert_order + 8, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'MSRP���ʃZ���Ɉړ�
    ActiveSheet.Range("O10").Select
    Selection.Copy
    'AcrtiveSheet��I��Ainsert_order+8�s�ڂɃy�[�X�g
    ActiveSheet.Cells(insert_order + 8, 9).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '�R�s�[���w�����
    Application.CutCopyMode = False

End Sub

