Option Explicit


Public Function СТОИМОСТЬДС(startDate As Date, endDate As Date, plata As Double, nds As Double) As Variant
    On Error GoTo ErrorHandler
    ' Переменные для хранения дат периодов с разными % ставками
    Dim startDate1 As Date
    Dim endDate1 As Date
    
    ' Переменные для хранения процентов и сумм
    Dim summ As Double
    Dim precent As Double
    Dim janDays As Long
    Dim i As Integer
    Dim f As Double
    f = WorksheetFunction.CountA(Worksheets("КредитныеПроценты").Range("C:C"))
    
    ' Проверяем наличие НДС
    If nds > 0 Then
        plata = plata + nds
    End If

    ' Вычисление количества дней в месяце находящихся в периоде
    If startDate <> 0 And endDate <> 0 And plata <> 0 Then
        If endDate > startDate Then
            For i = 1 To f
                startDate1 = Worksheets("КредитныеПроценты").Range("A" & i)
                endDate1 = Worksheets("КредитныеПроценты").Range("B" & i)
                If startDate < endDate1 Then
                    If endDate >= startDate1 Then
                        janDays = DateDiff("d", IIf(startDate >= startDate1, startDate, startDate1), IIf(endDate <= endDate1, endDate, endDate1)) + 1
                        precent = Worksheets("КредитныеПроценты").Range("C" & i).Value
                        summ = plata * janDays * precent / 365 + summ
                        
                        precent = 0
                    End If
                End If
            Next i
        End If
    Else
        СТОИМОСТЬДС = ""
        Exit Function
    End If
    
    СТОИМОСТЬДС = summ
    Exit Function
ErrorHandler:
    СТОИМОСТЬДС = CVErr(xlErrValue)
End Function
