Attribute VB_Name = "Module13"
Sub xƒ`ƒFƒbƒNƒ{ƒbƒNƒX‚Ì’l‚ğˆê‚Â‚Ã‚Âo—Í‚µ¤ğŒ•ªŠò‚ÅƒeƒLƒXƒgƒ{ƒbƒNƒX‚É‘‚«Š·‚¦¤‚È‚¢ê‡‚Íƒ{ƒbƒNƒX‚ğíœ()
    

Dim ppApp As New PowerPoint.Application
ppApp.Visible = True

Dim ppPrs As PowerPoint.Presentation
Set ppPrs = ppApp.Presentations.Open(ThisWorkbook.Path & "\ûqƒŒƒCƒAƒEƒgƒx[ƒXiŠ®¬”Åj12.17-3.pptx")

Dim countSld As Long 'ƒXƒ‰ƒCƒh”

Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("ƒf[ƒ^")

Dim i As Long
i = 2

Do While ws.Cells(i, 1).Value <> ""

    countSld = ppPrs.Slides.Count 'Œ»İ‚ÌƒXƒ‰ƒCƒh”‚ğƒJƒEƒ“ƒg
    If countSld Mod 2 = 0 Then
        ppPrs.Slides(2).Duplicate.MoveTo toPos:=countSld + 1
    Else
        ppPrs.Slides(1).Duplicate.MoveTo toPos:=countSld + 1
    End If
    ppPrs.Slides(countSld + 1).Shapes("ƒy[ƒW").TextFrame.TextRange.Text = ws.Cells(i, 1).Value '–‹ÆŠ–¼
    ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠ–¼").TextFrame.TextRange.Text = ws.Cells(i, 2).Value '–‹ÆŠ–¼
    ppPrs.Slides(countSld + 1).Shapes("ˆêŒ¾ƒƒbƒZ[ƒW").TextFrame.TextRange.Text = ws.Cells(i, 4).Value 'ƒƒbƒZ[ƒW
    ppPrs.Slides(countSld + 1).Shapes("Šˆ“®ƒ^ƒCƒgƒ‹").TextFrame.TextRange.Text = ws.Cells(i, 5).Value 'Šˆ“®ƒ^ƒCƒgƒ‹
    ppPrs.Slides(countSld + 1).Shapes("Šˆ“®“à—e").TextFrame.TextRange.Text = ws.Cells(i, 6).Value 'Šˆ“®“à—e
    ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠ–¼2").TextFrame.TextRange.Text = ws.Cells(i, 2).Value '–‹ÆŠ–¼2
    ppPrs.Slides(countSld + 1).Shapes("—X•Ö”Ô†").TextFrame.TextRange.Text = ws.Cells(i, 7).Value '—X•Ö”Ô†
    ppPrs.Slides(countSld + 1).Shapes("ZŠ").TextFrame.TextRange.Text = ws.Cells(i, 8).Value 'ZŠ
    ppPrs.Slides(countSld + 1).Shapes("Œš•¨–¼").TextFrame.TextRange.Text = ws.Cells(i, 9).Value 'Œš•¨–¼
    ppPrs.Slides(countSld + 1).Shapes("“d˜b”Ô†").TextFrame.TextRange.Text = ws.Cells(i, 10).Value '“d˜b”Ô†
    If IsEmpty(ws.Cells(i, 11).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("ƒ[ƒ‹ƒAƒhƒŒƒXƒAƒCƒRƒ“").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("ƒ[ƒ‹ƒAƒhƒŒƒX").Visible = msoFalse
    Else
        ppPrs.Slides(countSld + 1).Shapes("ƒ[ƒ‹ƒAƒhƒŒƒX").TextFrame.TextRange.Text = ws.Cells(i, 11).Value 'ƒ[ƒ‹ƒAƒhƒŒƒX
    End If
    
    ppPrs.Slides(countSld + 1).Shapes("ÅŠñ‚è‰w").TextFrame.TextRange.Text = ws.Cells(i, 13).Value 'ÅŠñ‰w‚P
    ppPrs.Slides(countSld + 1).Shapes("ÅŠñ‚è‰w2").TextFrame.TextRange.Text = ws.Cells(i, 14).Value 'ÅŠñ‰w2
    ppPrs.Slides(countSld + 1).Shapes("ŠJn").TextFrame.TextRange.Text = Format(ws.Cells(i, 15).Value, "hh:mm") 'ŠJn
    ppPrs.Slides(countSld + 1).Shapes("I—¹").TextFrame.TextRange.Text = Format(ws.Cells(i, 16).Value, "hh:mm") 'I—¹
    ppPrs.Slides(countSld + 1).Shapes("ŠJŠ—j“ú").TextFrame.TextRange.Text = ws.Cells(i, 18).Value 'ŠJŠ“ú
    
    Dim tmp As Variant '–‹ÆŠí•Ê
    tmp = Split(ws.Cells(i, 3).Value, ", ")
    If IsEmpty(tmp) Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 0 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 1 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").Visible = msoFalse
    End If
    
    If UBound(tmp) = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").TextFrame.TextRange.Text = tmp(3)
    End If
    
    If UBound(tmp) = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 5 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 6 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 7 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê8").TextFrame.TextRange.Text = tmp(7)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê9").Visible = msoFalse
    End If
    
    If UBound(tmp) = 8 Then
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê1").TextFrame.TextRange.Text = tmp(0)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê2").TextFrame.TextRange.Text = tmp(1)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê3").TextFrame.TextRange.Text = tmp(2)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê4").TextFrame.TextRange.Text = tmp(3)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê5").TextFrame.TextRange.Text = tmp(4)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê6").TextFrame.TextRange.Text = tmp(5)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê7").TextFrame.TextRange.Text = tmp(6)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê8").TextFrame.TextRange.Text = tmp(7)
        ppPrs.Slides(countSld + 1).Shapes("–‹ÆŠí•Ê9").TextFrame.TextRange.Text = tmp(8)
    End If
    
    Dim tmp2 As Variant 'áŠQÒí•Ê
    tmp2 = Split(ws.Cells(i, 22).Value, ", ")
    
    If IsEmpty(tmp2) Then
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 0 Then
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 1 Then
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê4").TextFrame.TextRange.Text = tmp2(3)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê5").Visible = msoFalse
    End If
    
    If UBound(tmp2) = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê1").TextFrame.TextRange.Text = tmp2(0)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê2").TextFrame.TextRange.Text = tmp2(1)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê3").TextFrame.TextRange.Text = tmp2(2)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê4").TextFrame.TextRange.Text = tmp2(3)
        ppPrs.Slides(countSld + 1).Shapes("áŠQÒí•Ê5").TextFrame.TextRange.Text = tmp2(4)
    End If
    
    '‘—Œ}”ÍˆÍˆ—
    If IsEmpty(ws.Cells(i, 19).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍƒAƒCƒRƒ“").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍF(ƒ‰ƒxƒ‹)").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍ").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍƒAƒCƒRƒ“(ƒ‰ƒxƒ‹)").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍƒAƒCƒRƒ“(˜g)").Visible = msoFalse
    Else
        ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍ").TextFrame.TextRange.Text = ws.Cells(i, 19).Value '‘—Œ}”ÍˆÍ
    End If
    
    'ˆã—ÃƒPƒAˆ—
    If IsEmpty(ws.Cells(i, 20).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("ˆã—ÃƒAƒCƒRƒ“").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("ˆã—ÃƒAƒCƒRƒ“(˜g)").Visible = msoFalse
    End If
    
    '‹‹Hˆ—
    If IsEmpty(ws.Cells(i, 21).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("‹‹HƒAƒCƒRƒ“").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("‹‹HƒAƒCƒRƒ“(˜g)").Visible = msoFalse
    End If
    
    '—áŠOˆ—
    If IsEmpty(ws.Cells(i, 17).Value) Then
        ppPrs.Slides(countSld + 1).Shapes("—áŠO").Visible = msoFalse
    End If
    
    '”w•\†
    If ws.Cells(i, 27).Value = 1 Then
        'ppPrs.Slides(countSld + 1).Shapes("‘—Œ}”ÍˆÍ").TextFrame.TextRange.Text = ws.Cells(i, 19).Value
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 2 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 3 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 4 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 5 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 6 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 7 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 8 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 9 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 10 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 11 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 12 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†13").Visible = msoFalse
    End If
    
    If ws.Cells(i, 27).Value = 13 Then
        ppPrs.Slides(countSld + 1).Shapes("”w•\†1").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†2").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†3").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†4").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†5").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†6").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†7").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†8").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†9").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†10").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†11").Visible = msoFalse
        ppPrs.Slides(countSld + 1).Shapes("”w•\†12").Visible = msoFalse
    End If
    
    
    i = i + 1
Loop

'ppApp.Quit
'Set ppApp = Nothing
End Sub



