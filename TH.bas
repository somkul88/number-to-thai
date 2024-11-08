Attribute VB_Name = "TH"
Option Explicit
'Main Function
Function thtext(ByVal MyNumber)
    Dim Dollars, Cents, Temp, Chk1, Chk3
    Dim DecimalPlace, Count, Test1, Test2, Test3
    ReDim Place(9) As String
    Place(2) = "พัน"
    Place(3) = "ล้าน"
    Place(4) = "พันล้าน"
    Place(5) = "ล้านล้าน"
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
        Case "หนึ่ง"
            Dollars = "หนึ่งบาท"
         Case Else
            Dollars = Dollars & "บาท"
    End Select
    Select Case Cents
        Case ""
            Cents = "ถ้วน"
            If Chk1 = "1" Then
            Chk1 = "1"
            Else
            Chk1 = "0"
            End If
        Case "หนึ่ง"
            Cents = "หนึ่งสตางค์"
              Case Else
            Cents = Cents & "สตางค์"
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
        Result = GetDigit(Mid(MyNumber, 1, 1)) & "แสน"
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
        Result = GetDigit(Mid(MyNumber, 1, 1)) & "ร้อย"
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
            Case 10: Result = "หนึ่งหมื่น"
            Case 11: Result = "หนึ่งหมื่นหนึ่ง"
            Case 12: Result = "หนึ่งหมื่นสอง"
            Case 13: Result = "หนึ่งหมื่นสาม"
            Case 14: Result = "หนึ่งหมื่นสี่"
            Case 15: Result = "หนึ่งหมื่นห้า"
            Case 16: Result = "หนึ่งหมื่นหก"
            Case 17: Result = "หนึ่งหมื่นเจ็ด"
            Case 18: Result = "หนึ่งหมื่นแปด"
            Case 19: Result = "หนึ่งหมื่นเก้า"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "สองหมื่น"
            Case 3: Result = "สามหมื่น"
            Case 4: Result = "สี่หมื่น"
            Case 5: Result = "ห้าหมื่น"
            Case 6: Result = "หกหมื่น"
            Case 7: Result = "เจ็ดหมื่น"
            Case 8: Result = "แปดหมื่น"
            Case 9: Result = "เก้าหมื่น"
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
            Case 10: Result = "สิบ"
            Case 11: Result = "สิบเอ็ด"
            Case 12: Result = "สิบสอง"
            Case 13: Result = "สิบสาม"
            Case 14: Result = "สิบสี่"
            Case 15: Result = "สิบห้า"
            Case 16: Result = "สิบหก"
            Case 17: Result = "สิบเจ็ด"
            Case 18: Result = "สิบแปด"
            Case 19: Result = "สิบเก้า"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "ยี่สิบ"
            Case 3: Result = "สามสิบ"
            Case 4: Result = "สี่สิบ"
            Case 5: Result = "ห้าสิบ"
            Case 6: Result = "หกสิบ"
            Case 7: Result = "เจ็ดสิบ"
            Case 8: Result = "แปดสิบ"
            Case 9: Result = "เก้าสิบ"
            Case Else
        End Select
        If Right(TensText, 1) = "1" Then
        Result = Result & "เอ็ด"
        Else
        Result = Result & GetDigit(Right(TensText, 1))  ' Retrieve ones place.
        End If
    End If
    GetTens = Result
End Function
' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "หนึ่ง"
        Case 2: GetDigit = "สอง"
        Case 3: GetDigit = "สาม"
        Case 4: GetDigit = "สี่"
        Case 5: GetDigit = "ห้า"
        Case 6: GetDigit = "หก"
        Case 7: GetDigit = "เจ็ด"
        Case 8: GetDigit = "แปด"
        Case 9: GetDigit = "เก้า"
        Case Else: GetDigit = ""
    End Select
End Function
