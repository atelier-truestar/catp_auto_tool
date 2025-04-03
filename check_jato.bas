Attribute VB_Name = "check_jato"
Option Explicit

Sub check_jato_list(ByVal button_no as Integer)

    Dim current_sht_name As String
    Dim result_body, result_wheel, result_seat, result_battery As String
    Dim check_filter_on As Boolean
    Dim check_filtermode as Boolean
    Dim insert_order As Integer
    
    '最後、各国別のJATOシートに戻るため、現在のシート名を取得
    current_sht_name = ActiveSheet.Name
    '現在開いているシートをActiveSheetに設定
    Worksheets(current_sht_name).Activate
    '引数のボタン番号を取得
    insert_order = button_no
    'フィルターがかかっているかどうかを確認
    check_filter_on = ActiveSheet.AutoFilterMode
    check_filtermode = ActiveSheet.FilterMode
    '以下、フィルターの状態で処理を分岐
    If check_filter_on = False or (check_filter_on=True and check_filtermode=False) Then
        'フィルターが解除されている場合、またはオートフィルターが無選択の場合、list_all()のSubをコールする
        Call list_all(current_sht_name, insert_order)
    Else
        'フィルターがかかっている場合、list_filter()のSubをコールする
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

    'RangeListにA2からD2までの範囲を格納
    RangeList(0) = "A2"
    RangeList(1) = "B2"
    RangeList(2) = "C2"
    RangeList(3) = "D2"
    'Dim temp_sht_name As String
    'テンポラリのシートを空にする
    ThisWorkbook.Sheets("JATO_WorkArea").Cells.Clear
    '現在開いているシートをActiveSheetに設定
    Worksheets(sht_tab_name).Activate
    'A列の最終行を取得
    last_row = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    'E15からHのlast_rowまでの範囲を選択し、コピー
    ActiveSheet.Range("E15:H" & last_row).Select
    Selection.Copy
    'テンポラリのシートにペースト
    ThisWorkbook.Sheets("JATO_WorkArea").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Set ws = ThisWorkbook.Sheets("JATO_WorkArea")
    ws.Activate
    '4項目分をループで処理
    For i = 0 To 3
        'リストの開始セルを取得
        Set startRange = ws.Range(RangeList(i))
        'リストの最終行を取得
        filterEndRow = startRange.End(xlDown).Row
        'リストの最終セルを取得
        Set endRange = ws.Cells(filterEndRow, startRange.Column)
        'リストの範囲を取得
        Set targetRange = ws.Range(startRange, endRange)
        '重複を除いたリストを作成
        UniqueList(i) = WorksheetFunction.Unique(targetRange)
        'Uniquelistの内容チェック
        If UBound(UniqueList(i)) = 1 Then
            If IsEmpty(UniqueList(i)(1)) Then
                UniqueList_a(i) = "None"
            Else: UniqueList_a(i) = UniqueList(i)
            End If
        Else: UniqueList_a(i) = "All"
        End If
        '結果をcells(1, i+6)に代入
        ws.Cells(1, i + 6) = UniqueList_a(i)
    Next i
    'ペーストした範囲をセレクト
    ws.Range("F1:I1").Select
    Selection.Copy
    Worksheets(sht_tab_name).Activate
    'ActiveSheetのE列、insert_order+8行目にペースト
    ActiveSheet.Cells(insert_order + 8, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'MSRP結果セルに移動
    ActiveSheet.Range("O10").Select
    Selection.Copy
    'AcrtiveSheetのI列、insert_order+8行目にペースト
    ActiveSheet.Cells(insert_order + 8, 9).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'コピー元指定解除
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

    'RangeListにA2からD2までの範囲を格納
    RangeList(0) = "A2"
    RangeList(1) = "B2"
    RangeList(2) = "C2"
    RangeList(3) = "D2"
    'Dim temp_sht_name As String
    'テンポラリのシートを空にする
    ThisWorkbook.Sheets("JATO_WorkArea").Cells.Clear
    '現在開いているシートをActiveSheetに設定
    Worksheets(sht_tab_name).Activate
    'A列の最終行を取得
    last_row = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    'E15からHのlast_rowまでの範囲を選択し、コピー
    ActiveSheet.Range("E15:H" & last_row).Select
    Selection.Copy
    'テンポラリのシートにペースト
    ThisWorkbook.Sheets("JATO_WorkArea").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Set ws = ThisWorkbook.Sheets("JATO_WorkArea")
    ws.Activate
    '4項目分をループで処理
    For i = 0 To 3
        'リストの開始セルを取得
        Set startRange = ws.Range(RangeList(i))
        'リストの最終行を取得
        filterEndRow = startRange.End(xlDown).Row
        'リストの最終セルを取得
        Set endRange = ws.Cells(filterEndRow, startRange.Column)
        'リストの範囲を取得
        Set targetRange = ws.Range(startRange, endRange)
        '重複を除いたリストを作成
        UniqueList(i) = WorksheetFunction.Unique(targetRange)
        'Uniquelistの内容チェック
        If UBound(UniqueList(i)) = 1 Then
            If IsEmpty(UniqueList(i)(1)) Then
                UniqueList_a(i) = "None"
            Else: UniqueList_a(i) = UniqueList(i)
            End If
        Else
            UniqueList_a(i) = Join(WorksheetFunction.Transpose(UniqueList(i)), ",")
        End If
        '結果をcells(1, i+6)に代入
        ws.Cells(1, i + 6) = UniqueList_a(i)
    Next i
    'ペーストした範囲をセレクト
    ws.Range("F1:I1").Select
    Selection.Copy
    Worksheets(sht_tab_name).Activate
    'ActiveSheetのE列、insert_order+8行目にペースト
    ActiveSheet.Cells(insert_order + 8, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'MSRP結果セルに移動
    ActiveSheet.Range("O10").Select
    Selection.Copy
    'AcrtiveSheetのI列、insert_order+8行目にペースト
    ActiveSheet.Cells(insert_order + 8, 9).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'コピー元指定解除
    Application.CutCopyMode = False

End Sub

