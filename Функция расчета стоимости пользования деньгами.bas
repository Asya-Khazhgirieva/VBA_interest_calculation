Attribute VB_Name = "�������"
Option Explicit


Public Function �����������(startDate As Date, endDate As Date, plata As Double, nds As Double) As Double
Attribute �����������.VB_Description = "������ ��������� ����������� ��������"
Attribute �����������.VB_ProcData.VB_Invoke_Func = " \n14"
    ' ���������� ��� �������� ��� �������� � ������� % ��������
    Dim startDate1 As Date
    Dim endDate1 As Date
    
    ' ���������� ��� �������� ��������� � ����
    Dim summ As Double
    Dim precent As Double
    Dim janDays As Long
    Dim i As Integer
    Dim f As Double
    f = WorksheetFunction.CountA(Worksheets("�����������������").Range("C:C"))
    
    ' ��������� ������� ���
    If nds > 0 Then
        plata = plata + nds
    End If
    
    ' ���������� ���������� ���� � ������ ����������� � �������
    If endDate > startDate Then
        For i = 1 To f
            startDate1 = Worksheets("�����������������").Range("A" & i)
            endDate1 = Worksheets("�����������������").Range("B" & i)
            If startDate < endDate1 Then
                If endDate >= startDate1 Then
                    janDays = DateDiff("d", IIf(startDate >= startDate1, startDate, startDate1), IIf(endDate <= endDate1, endDate, endDate1)) + 1
                    precent = Worksheets("�����������������").Range("C" & i).Value
                    summ = plata * janDays * precent / 365 + summ
                    
                    precent = 0
                End If
            End If
        Next i
    End If
    
    ����������� = summ
End Function
