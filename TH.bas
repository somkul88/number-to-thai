Attribute VB_Name = "TH"
Option Explicit
'Main Function
Function thtext(ByVal MyNumber)
    Dim Dollars, Cents, Temp, Chk1, Chk3
    Dim DecimalPlace, Count, Test1, Test2, Test3
    ReDim Place(9) As String
    Place(2) = "�ѹ"
    Place(3) = "��ҹ"
    Place(4) = "�ѹ��ҹ"
    Place(5) = "��ҹ��ҹ"
    Chk1 = "0"
    Chk3 = 2
    Count = 1
    Test3 = 0
    ' String representation of amount.
    MyNumber = WorksheetFunction.Round(MyNumber, 2)
    MyNumber = Trim(Str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
     Test1 = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
    If DecimalPlace > 0 Then
        Cents = GetTens(Test1)
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Do While MyNumber <> ""
    If Count = Chk3 Then
        Temp = GetThousand(Right(MyNumber, 3))
        Test2 = Right(MyNumber, 3)
        If Right(Test2, 1) = "0" Then
        Dollars = Temp & Dollars
        Test3 = 1
        End If
    Else
        Temp = GetHundreds(Right(MyNumber, 3))
    End If
        If Test3 <> 1 Then Dollars = Temp & Place(Count) & Dollars
    If Count = Chk3 Then
        Chk3 = Chk3 + 2
        Test3 = 0
    End If
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Dollars
        Case ""
            Dollars = "No Baht"
            Chk1 = "1"
        Case "˹��"
            Dollars = "˹�觺ҷ"
         Case Else
            Dollars = Dollars & "�ҷ"
    End Select
    Select Case Cents
        Case ""
            Cents = "��ǹ"
            If Chk1 = "1" Then
            Chk1 = "1"
            Else
            Chk1 = "0"
            End If
        Case "˹��"
            Cents = "˹��ʵҧ��"
              Case Else
            Cents = Cents & "ʵҧ��"
    End Select
    If Chk1 = "1" Then
    thtext = ""
    Else
    thtext = "----" & Dollars & Cents & "----"
    End If
End Function
' Converts a number from 1000-9999 into text
Function GetThousand(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & "�ʹ"
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTensx(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetThousand = Result
End Function
' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & "����"
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function
' Converts a number from 10 to 99 into text.
Function GetTensx(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "˹������"
            Case 11: Result = "˹������˹��"
            Case 12: Result = "˹�������ͧ"
            Case 13: Result = "˹���������"
            Case 14: Result = "˹���������"
            Case 15: Result = "˹���������"
            Case 16: Result = "˹������ˡ"
            Case 17: Result = "˹��������"
            Case 18: Result = "˹������Ỵ"
            Case 19: Result = "˹���������"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "�ͧ����"
            Case 3: Result = "�������"
            Case 4: Result = "�������"
            Case 5: Result = "�������"
            Case 6: Result = "ˡ����"
            Case 7: Result = "������"
            Case 8: Result = "Ỵ����"
            Case 9: Result = "�������"
            Case Else
        End Select
        Result = Result & GetDigit(Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTensx = Result
End Function
' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "�Ժ"
            Case 11: Result = "�Ժ���"
            Case 12: Result = "�Ժ�ͧ"
            Case 13: Result = "�Ժ���"
            Case 14: Result = "�Ժ���"
            Case 15: Result = "�Ժ���"
            Case 16: Result = "�Ժˡ"
            Case 17: Result = "�Ժ��"
            Case 18: Result = "�ԺỴ"
            Case 19: Result = "�Ժ���"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "����Ժ"
            Case 3: Result = "����Ժ"
            Case 4: Result = "����Ժ"
            Case 5: Result = "����Ժ"
            Case 6: Result = "ˡ�Ժ"
            Case 7: Result = "���Ժ"
            Case 8: Result = "Ỵ�Ժ"
            Case 9: Result = "����Ժ"
            Case Else
        End Select
        If Right(TensText, 1) = "1" Then
        Result = Result & "���"
        Else
        Result = Result & GetDigit(Right(TensText, 1))  ' Retrieve ones place.
        End If
    End If
    GetTens = Result
End Function
' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "˹��"
        Case 2: GetDigit = "�ͧ"
        Case 3: GetDigit = "���"
        Case 4: GetDigit = "���"
        Case 5: GetDigit = "���"
        Case 6: GetDigit = "ˡ"
        Case 7: GetDigit = "��"
        Case 8: GetDigit = "Ỵ"
        Case 9: GetDigit = "���"
        Case Else: GetDigit = ""
    End Select
End Function
