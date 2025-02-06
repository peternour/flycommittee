Attribute VB_Name = "AbuHani"

Function NoToTxt(TheNo As Double, MyCur As String, MySubCur As String) As String
Dim MyArry1(0 To 9) As String
Dim MyArry2(0 To 9) As String
Dim MyArry3(0 To 9) As String
Dim Myno As String
Dim GetNo As String
Dim RdNo As String
Dim My100 As String
Dim My10 As String
Dim My1 As String
Dim My11 As String
Dim My12 As String
Dim GetTxt As String
Dim Mybillion As String
Dim MyMillion As String
Dim MyThou As String
Dim MyHun As String
Dim MyFraction As String
Dim MyAnd As String
Dim i As Integer
Dim ReMark As String


If TheNo > 999999999999.99 Then Exit Function

If TheNo < 0 Then
TheNo = TheNo * -1
ReMark = "Ì »ﬁÏ ·ﬂ„ "
Else
ReMark = "›ﬁÿ "
End If

If TheNo = 0 Then
NoToTxt = "’›—"
Exit Function
End If

MyAnd = " Ê"
MyArry1(0) = ""
MyArry1(1) = "„«∆…"
MyArry1(2) = "„«∆ «‰"
MyArry1(3) = "À·«À„«∆…"
MyArry1(4) = "√—»⁄„«∆…"
MyArry1(5) = "Œ„”„«∆…"
MyArry1(6) = "” „«∆…"
MyArry1(7) = "”»⁄„«∆…"
MyArry1(8) = "À„«‰„«∆…"
MyArry1(9) = " ”⁄„«∆…"

MyArry2(0) = ""
MyArry2(1) = " ⁄‘—"
MyArry2(2) = "⁄‘—Ê‰"
MyArry2(3) = "À·«ÀÊ‰"
MyArry2(4) = "√—»⁄Ê‰"
MyArry2(5) = "Œ„”Ê‰"
MyArry2(6) = "” Ê‰"
MyArry2(7) = "”»⁄Ê‰"
MyArry2(8) = "À„«‰Ê‰"
MyArry2(9) = " ”⁄Ê‰"

MyArry3(0) = ""
MyArry3(1) = "Ê«Õœ"
MyArry3(2) = "«À‰«‰"
MyArry3(3) = "À·«À…"
MyArry3(4) = "√—»⁄…"
MyArry3(5) = "Œ„”…"
MyArry3(6) = "” …"
MyArry3(7) = "”»⁄…"
MyArry3(8) = "À„«‰Ì…"
MyArry3(9) = " ”⁄…"
'======================

GetNo = Format(TheNo, "000000000000.00")

i = 0
Do While i < 15

If i < 12 Then
Myno = Mid$(GetNo, i + 1, 3)
Else
Myno = "0" + Mid$(GetNo, i + 2, 2)
End If

If (Mid$(Myno, 1, 3)) > 0 Then

RdNo = Mid$(Myno, 1, 1)
My100 = MyArry1(RdNo)
RdNo = Mid$(Myno, 3, 1)
My1 = MyArry3(RdNo)
RdNo = Mid$(Myno, 2, 1)
My10 = MyArry2(RdNo)

If Mid$(Myno, 2, 2) = 11 Then My11 = "≈ÕœÏ ⁄‘—"
If Mid$(Myno, 2, 2) = 12 Then My12 = "≈À‰Ï ⁄‘—"
If Mid$(Myno, 2, 2) = 10 Then My10 = "⁄‘—…"

If ((Mid$(Myno, 1, 1)) > 0) And ((Mid$(Myno, 2, 2)) > 0) Then My100 = My100 + MyAnd
If ((Mid$(Myno, 3, 1)) > 0) And ((Mid$(Myno, 2, 1)) > 1) Then My1 = My1 + MyAnd

GetTxt = My100 + My1 + My10

If ((Mid$(Myno, 3, 1)) = 1) And ((Mid$(Myno, 2, 1)) = 1) Then
GetTxt = My100 + My11
If ((Mid$(Myno, 1, 1)) = 0) Then GetTxt = My11
End If

If ((Mid$(Myno, 3, 1)) = 2) And ((Mid$(Myno, 2, 1)) = 1) Then
GetTxt = My100 + My12
If ((Mid$(Myno, 1, 1)) = 0) Then GetTxt = My12
End If

If (i = 0) And (GetTxt <> "") Then
If ((Mid$(Myno, 1, 3)) > 10) Then
Mybillion = GetTxt + " „·Ì«—"
Else
Mybillion = GetTxt + " „·Ì«—« "
If ((Mid$(Myno, 1, 3)) = 2) Then Mybillion = " „·Ì«—"
If ((Mid$(Myno, 1, 3)) = 2) Then Mybillion = " „·Ì«—«‰"
End If
End If

If (i = 3) And (GetTxt <> "") Then

If ((Mid$(Myno, 1, 3)) > 10) Then
MyMillion = GetTxt + " „·ÌÊ‰"
Else
MyMillion = GetTxt + " „·«ÌÌ‰"
If ((Mid$(Myno, 1, 3)) = 1) Then MyMillion = " „·ÌÊ‰"
If ((Mid$(Myno, 1, 3)) = 2) Then MyMillion = " „·ÌÊ‰«‰"
End If
End If

If (i = 6) And (GetTxt <> "") Then
If ((Mid$(Myno, 1, 3)) > 10) Then
MyThou = GetTxt + " √·›"
Else
MyThou = GetTxt + " ¬·«›"
If ((Mid$(Myno, 3, 1)) = 1) Then MyThou = " √·›"
If ((Mid$(Myno, 3, 1)) = 2) Then MyThou = " √·›«‰"
End If
End If

If (i = 9) And (GetTxt <> "") Then MyHun = GetTxt
If (i = 12) And (GetTxt <> "") Then MyFraction = GetTxt
End If

i = i + 3
Loop

If (Mybillion <> "") Then
If (MyMillion <> "") Or (MyThou <> "") Or (MyHun <> "") Then Mybillion = Mybillion + MyAnd
End If

If (MyMillion <> "") Then
If (MyThou <> "") Or (MyHun <> "") Then MyMillion = MyMillion + MyAnd
End If

If (MyThou <> "") Then
If (MyHun <> "") Then MyThou = MyThou + MyAnd
End If

If MyFraction <> "" Then
If (Mybillion <> "") Or (MyMillion <> "") Or (MyThou <> "") Or (MyHun <> "") Then
NoToTxt = ReMark + Mybillion + MyMillion + MyThou + MyHun + " " + MyCur + MyAnd + MyFraction + " " + MySubCur
Else
NoToTxt = ReMark + MyFraction + " " + MySubCur
End If
Else
NoToTxt = ReMark + Mybillion + MyMillion + MyThou + MyHun + " " + MyCur
End If

End Function
