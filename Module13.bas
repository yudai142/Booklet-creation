Attribute VB_Name = "Module13"
Sub xチェックボックスの値を一つづつ出力し､条件分岐でテキストボックスに書き換え､ない場合はボックスを削除()
    

Dim ppApp As New PowerPoint.Application
ppApp.Visible = True

Dim ppPrs As PowerPoint.Presentation
Set ppPrs = ppApp.Presentations.Open(ThisWorkbook.Path & "\冊子レイアウトベース（完成版）12.17-3.pptx")

Dim countSld As Long 'スライド数

Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("データ")

Dim i As Long
i = 2

Do While ws.Cells(i, 1).Value <> ""

    countSld = ppPrs.Slides.Count '現在のスライド数をカウント
    If countSld Mod 2 = 0 Then
        ppPrs.Slides(2).Duplicate.MoveTo toPos:=countSld + 1
    Else
        ppPrs.Slides(1).Duplicate.MoveTo toPos:=countSld + 1
    End If
    ppPrs.Slides(countSld + 1).Shapes("ページ").TextFrame.TextRange.Text = ws.Cells(i, 1).Value '事業所名
    ppPrs.Slides(countSld + 1).Shapes("事業所名").TextFrame.TextRange.Text = ws.Cells(i, 2).Value '事業所名
    ppPrs.Slides(countSld + 1).Shapes("一言メッセージ").TextFrame.TextRange.Text = ws.Cells(i, 4).Value 'メッセージ
    ppPrs.Slides(countSld + 1).Shapes("活動タイトル").TextFrame.TextRange.Text = ws.Cells(i, 5).Value '活動タイトル
    ppPrs.Slides(countSld + 1).Shapes("活動内容").TextFrame.TextRange.Text = ws.Cells(i, 6).Value '活動内容
    ppPrs.Slides(countSld + 1).Shapes("事業所名2").TextFrame.TextRange.Text = ws.Cells(i, 2).Value '事業所名2
    ppPrs.Slides(countSld + 1).Shapes("郵便番号").TextFrame.TextRange.Text = ws.Cells(i, 7).Value '郵便番号
    ppPrs.Slides(countSld + 1).Shapes("住所").TextFrame.TextRange.Text = ws.Cells(i, 8).Value '住所
    ppPrs.Slides(countSld + 1).Shapes("建物名").TextFrame.TextRange.Text = ws.Cells(i, 9).Value '建物名
    ppPrs.Slides(countSld + 1).Shapes("電話番号").TextFrame.TextRange.Text = ws.Cells(i, 10).Value '電話番号
    If IsEmpty(ws.Cells(i, 11).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("メールアドレスアイコン").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("メールアドレス").Visible = msoFalse
    Else
        ppPrs.Slides(countSld + 1).Shapes("メールアドレス").TextFrame.TextRange.Text = ws.Cells(i, 11).Value 'メールアドレス
    End If
    
    ppPrs.Slides(countSld + 1).Shapes("最寄り駅").TextFrame.TextRange.Text = ws.Cells(i, 13).Value '最寄駅１
    ppPrs.Slides(countSld + 1).Shapes("最寄り駅2").TextFrame.TextRange.Text = ws.Cells(i, 14).Value '最寄駅2
    ppPrs.Slides(countSld + 1).Shapes("開始時刻").TextFrame.TextRange.Text = Format(ws.Cells(i, 15).Value, "hh:mm") '開始時刻
    ppPrs.Slides(countSld + 1).Shapes("終了時刻").TextFrame.TextRange.Text = Format(ws.Cells(i, 16).Value, "hh:mm") '終了時刻
    ppPrs.Slides(countSld + 1).Shapes("開所曜日").TextFrame.TextRange.Text = ws.Cells(i, 18).Value '開所日
    
    Dim tmp As Variant '事業所種別
    tmp = Split(ws.Cells(i, 3).Value, ", ")
    If IsEmpty(tmp) Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 0 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 1 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").TextFrame.TextRange.Text = tmp(3)
    End If
    
    If UBound(tmp) = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 5 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 6 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("事業所種別9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 7 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別8").TextFrame.TextRange.Text = tmp(7)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 8 Then
        ppPrs.Slides(countSld + 1).Shapes("事業所種別1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別8").TextFrame.TextRange.Text = tmp(7)
        ppPrs.Slides(countSld + 1).Shapes("事業所種別9").TextFrame.TextRange.Text = tmp(8)
    End If
    
    Dim tmp2 As Variant '障害者種別
    tmp2 = Split(ws.Cells(i, 22).Value, ", ")
    
    If IsEmpty(tmp2) Then
        ppPrs.Slides(countSld + 1).Shapes("障害者種別1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 0 Then
        ppPrs.Slides(countSld + 1).Shapes("障害者種別1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 1 Then
        ppPrs.Slides(countSld + 1).Shapes("障害者種別1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("障害者種別1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("障害者種別5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("障害者種別1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別4").TextFrame.TextRange.Text = tmp2(3)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("障害者種別1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別4").TextFrame.TextRange.Text = tmp2(3)
        ppPrs.Slides(countSld + 1).Shapes("障害者種別5").TextFrame.TextRange.Text = tmp2(4)
    End If
    
    '送迎範囲処理
    If IsEmpty(ws.Cells(i, 19).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("送迎範囲アイコン").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("送迎範囲：(ラベル)").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("送迎範囲").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("送迎範囲アイコン(ラベル)").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("送迎範囲アイコン(枠)").Visible = msoFalse
    Else
        ppPrs.Slides(countSld + 1).Shapes("送迎範囲").TextFrame.TextRange.Text = ws.Cells(i, 19).Value '送迎範囲
    End If
    
    '医療ケア処理
    If IsEmpty(ws.Cells(i, 20).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("医療アイコン").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("医療アイコン(枠)").Visible = msoFalse
    End If
    
    '給食処理
    If IsEmpty(ws.Cells(i, 21).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("給食アイコン").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("給食アイコン(枠)").Visible = msoFalse
    End If
    
    '例外時刻処理
    If IsEmpty(ws.Cells(i, 17).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("例外時刻").Visible = msoFalse
    End If
    
    '背表紙
    If ws.Cells(i, 27).Value = 1 Then
        'ppPrs.Slides(countSld + 1).Shapes("送迎範囲").TextFrame.TextRange.Text = ws.Cells(i, 19).Value
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 5 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 6 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 7 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 8 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 9 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 10 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 11 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 12 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 13 Then
        ppPrs.Slides(countSld + 1).Shapes("背表紙1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("背表紙12").Visible = msoFalse
    End If
    
    
    i = i + 1
Loop

'ppApp.Quit
'Set ppApp = Nothing
End Sub



