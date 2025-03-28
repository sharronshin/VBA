# VBA
VBA 코드

Sub CopyMatchingValuesWithCondition()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, lastRowDest As Long
    Dim i As Long, j As Long

    ' 시트 설정
    Set wsSrc = ThisWorkbook.Sheets("적요") ' 적요 시트
    Set wsDest = ThisWorkbook.Sheets("지급내역") ' 지급내역 시트

    ' 적요 시트의 A열과 B열 마지막 행 찾기
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    ' 지급내역 시트의 K열과 L열 마지막 행 찾기
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, 12).End(xlUp).Row ' L열 = 12번째 열, K열 = 11번째 열

    ' A열 값과 L열 값이 일치하는 경우 B열 값을 하나씩 확인하여 K열에서 포함 여부 검사
    For i = 2 To lastRow
        For j = 2 To lastRowDest
            ' A열 값이 지급내역 시트 L열 값과 일치하면
            If wsSrc.Cells(i, 1).Value = wsDest.Cells(j, 12).Value Then
                ' B열 값이 K열에 포함되어 있으면
                If InStr(1, wsDest.Cells(j, 11).Value, wsSrc.Cells(i, 2).Value, vbTextCompare) > 0 Then
                    wsDest.Cells(j, 3).Value = wsSrc.Cells(i, 3).Value ' C열 값을 지급내역 시트 D열에 복사
                End If
            End If
        Next j
    Next i

    MsgBox "작업 완료!", vbInformation
End Sub


