Attribute VB_Name = "�Զ�����"
Option Explicit

Sub AutoRun()
Dim i As Integer, j As Integer, r As Integer, k&, k1&, k2&
Dim a As String, b As String, vbStatus As String
Dim wk As Workbook, wk1 As Workbook, wk2 As Workbook, wk3 As Workbook
Dim mDate
Dim arr, brr, crr
Dim test

Set wk = ThisWorkbook
On Error GoTo EH
Application.ScreenUpdating = False
vbStatus = MsgBox("ȷ��Ҫ����״̬��", vbInformation + vbOKCancel, "��ʾ")
If vbStatus = vbCancel Then End

If Sheets("ȫ����").FilterMode Then
    Sheets("ȫ����").ShowAllData
End If
j = Range("B1000").End(xlUp).Row

mDate = Range("B1")
If Day(mDate) > 26 Then
    a = Month(mDate) + 1
Else
    a = Month(mDate)
End If
If Left(a, 1) = "0" Then a = Right(a, 1)

'������;״̬
For i = j To 3 Step -1
    If Range("E" & i) = "���" Then
        Rows(i).Delete
    ElseIf Range("E" & i) = "����1" Then
        Range("E" & i) = "���"
    ElseIf Range("E" & i) = "�ս�2" Then
        Range("E" & i) = "����1"

    ElseIf Range("E" & i) = "����1" Then
        Range("E" & i) = "����2"
    End If

Next

Set wk1 = OpenBooks("\\Hanhd\05--���ϼƻ�\�����ռƻ�.xlsm", True)
Set wk2 = OpenBooks("\\Hanhd\05--���ϼƻ�\����ĥ�����ƻ�.xlsx", True)
Set wk3 = OpenBooks("\\Yt2\ϸ�۽����¼\2018��ϸ�۽����¼\2018." & a & ".xls", True)
wk.Sheets("ȫ����").Activate

'
Sheets("����1").Range("B:B").ClearContents
Sheets("����2").Range("B:B").ClearContents
Sheets("����ĥ").Range("B:B").ClearContents
Sheets("�ս�2").Range("B:B").ClearContents

Application.Calculation = xlCalculationAutomatic
''���µ�����������
ReDim arr(50, 1)
With wk1.Sheets("����")
    If .FilterMode Then
        .ShowAllData
    End If
    r = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = r - 100 To r
        test = mDate
        If .Cells(i, "A") = mDate And InStr(.Cells(i, "D"), "�Ի�") = 0 Then
            arr(k, 0) = .Cells(i, "B")
            arr(k, 1) = .Cells(i, "C")
            k = k + 1
        End If
    Next
    If k > 0 Then
        ReDim brr(k - 1, 3)
        For i = 0 To UBound(brr)
            brr(i, 0) = arr(i, 0)
            brr(i, 1) = arr(i, 1)
            brr(i, 2) = 575
            brr(i, 3) = "����"
        Next
    End If

End With
r = Cells(Rows.Count, 2).End(xlUp).Row + 1
Range(Cells(r, "B"), Cells(r + UBound(brr), "E")) = brr
Erase arr: Erase brr: r = 0: k = 0: i = 0

'������Ƭ�Ի����κ������ռƻ�
a = CStr(Format(mDate, "yymmdd"))

With wk1.Sheets(a)
    For i = 6 To 36
         If Cells(i, "C") <> "" Then
            r = r + 1
         End If
    Next
    ReDim arr(r, 3)
    For i = 6 To 36
        If .Cells(i, "C") <> "" And .Cells(i, "D") <> "" And .Cells(i, "E") <> "" Then
            arr(k, 0) = .Cells(i, "C")
            arr(k, 1) = .Cells(i, "D")
            arr(k, 2) = .Cells(i, "E")
            If InStr(.Cells(i, "F"), "����ĥ��") > 0 Then
                arr(k, 3) = "����2"
                With Sheets("����2")
                    k1 = k1 + 1
                    .Cells(k1, "B") = arr(k, 1)
                End With
            Else
                arr(k, 3) = "����1"
                With Sheets("����1")
                    k2 = k2 + 1
                    .Cells(k2, "B") = arr(k, 1)
                End With
            End If
            k = k + 1
        End If
    Next

    For i = 0 To UBound(arr)
        If Mid(arr(i, 1), 8, 1) >= "A" Then
            r = Cells(500, "C").End(xlUp).Row + 1
            Cells(r, "B") = arr(i, 0)
            Cells(r, "C") = arr(i, 1)
            Cells(r, "D") = arr(i, 2)
            Cells(r, "E") = "����"
        End If
    Next
End With

'��������״̬
r = Cells(Rows.Count, 2).End(xlUp).Row
brr = Range(Cells(3, "A"), Cells(r, "E"))
For i = 1 To UBound(brr)
    For j = 0 To UBound(arr)
        If brr(i, 3) = arr(j, 1) Then
           brr(i, 5) = arr(j, 3)
        End If
    Next
Next



Erase arr: r = 0: k = 0

'��������ĥ״̬
With wk2.Sheets("2018������ĥ�ƻ���")
    For i = 5 To 100
        If .Cells(i, "A") = mDate Then
            r = r + 1
        End If
    Next
    ReDim arr(r)
    For i = 5 To 100
        If .Cells(i, "A") = mDate Then
            arr(k) = .Cells(i, "D")
            k = k + 1
        End If
    Next
    For i = 1 To UBound(brr)
        For j = 0 To UBound(arr)
            If brr(i, 3) = arr(j) Then
               brr(i, 5) = "����ĥ"
            End If
        Next
    Next

End With

For i = 0 To UBound(arr)
    Sheets("����ĥ").Cells(i + 1, 2) = arr(i)
Next
Erase arr: r = 0: k = 0: j = 0
'���ĵ�һ���ս�״̬
With wk3.Sheets("��һ��")
    r = .Cells(3, 2).End(xlDown).Row
    For i = 3 To r
        If .Cells(i, "A") = mDate - 1 Then
            j = j + 1
        End If
    Next
    ReDim arr(j)
    For i = 3 To r
        If .Cells(i, "A") = mDate - 1 Then
            arr(k) = .Cells(i, "C")
            k = k + 1
        End If
    Next
    For i = 1 To UBound(brr)
        For j = 0 To UBound(arr)
            If brr(i, 3) = arr(j) Then
               brr(i, 5) = "�ս�2"
               brr(i, 1) = mDate - 1
            End If
        Next
    Next

End With

For i = 0 To UBound(arr)
    Sheets("�ս�2").Cells(i + 1, 2) = arr(i)
Next
Erase arr: r = 0: k = 0: j = 0

r = Cells(Rows.Count, 2).End(xlUp).Row
Range(Cells(3, "A"), Cells(r, "E")) = brr
Range("F3:G3").Copy
Range("F3:G300").PasteSpecial xlPasteFormulas, xlPasteSpecialOperationNone

'�����ս�1״̬
For i = 3 To r
    If Cells(i, "E") = "����ĥ" And Cells(i, "F") = 0 Then
        Cells(i, "E") = "�ս�1"
    End If
    '����װСͰ����
    If Cells(i, "E") <> "����1" And Cells(i, "E") <> "����2" And Cells(i, "E") <> "����" And Cells(i, "I") <> "ʵ����" And Cells(i, "D") = "" Then
        If Mid(Cells(i, "C"), 8, 1) >= "A" Then
            Cells(i, "D") = 600
        Else
            Cells(i, "D") = 575
        End If
    End If
Next

wk1.Close False: wk2.Close False: wk3.Close False
Set wk1 = Nothing: Set wk2 = Nothing: Set wk3 = Nothing

'��������ƻ�
a = CStr(Format(mDate - 1, "yymmdd"))
b = "2018������ƻ�" & a
For Each wk1 In Workbooks
    If wk1.Name = "2018������ƻ�.xlsm" Then wk1.Close False
Next
FileCopy "\\Hanhd\05--���ϼƻ�\2018������ƻ�.xlsm", "\\Hanhd\05--���ϼƻ�\����ƻ�����\" & b & ".xlsm"
Set wk1 = OpenBooks("\\Hanhd\05--���ϼƻ�\2018������ƻ�.xlsm", False, "12.3")

'
''��������ƻ�
'wk1.Sheets("�����ƻ�").Range("B3:BA10").Copy
'wk1.Sheets("�����ƻ�").Range("B3:BA10").Paste
''ˢ�¼��㰴ť
'Application.Run
''���·��ϱ���
'
''ˢ�·������󱨱�
Application.ScreenUpdating = True
Exit Sub
EH:
MsgBox "���ִ�����رչ��������Ҳ����棬Ȼ���ֶ����±���", vbCritical + vbOKOnly, "Visual Basic for Application"
Application.ScreenUpdating = True
End Sub

Sub test ()
dim i%,j%
dim wk1,wk2

set wk1=Workbooks("2018������ƻ�.xlsm")
Application.ScreenUpdating=False

End Sub

'--------------------------
'     �򿪹�����
'--------------------------
Private Function OpenBooks(a As String, myReadOnly As Boolean, Optional pwd As String) As Workbook
Dim wk As Workbook
Dim myBoolean As Boolean
Dim b() As String, i As Integer
b() = Split(a, "\")
i = UBound(b)
For Each wk In Workbooks
    If wk.Name = b(i) Then
       myBoolean = True
       Exit For
    End If
Next
If myBoolean Then
Set OpenBooks = Workbooks(b(i))
Else
'Set wk2 = Workbooks(b)  '����
If pwd <> "" Then
    Set OpenBooks = Workbooks.Open(a, , myReadOnly, Password:=pwd)
Else
    Set OpenBooks = Workbooks.Open(a, , myReadOnly)
End If
End If
End Function

'************************************
'���ܣ����ҵ�Ԫ��
'
'���� findText��Ҫ���ҵ����� withinRng:���ҵ�Ԫ������
'�����ҵ��ĵ�Ԫ���ַ
'************************************
Function RngFind(findText, withinRng As Range) As String
Dim rng As Range
Dim i As Integer
Dim findAddress As String
Set rng = withinRng.Find(findText, lookat:=xlWhole)
On Error GoTo 1
findAddress = rng.Address
If Not rng Is Nothing Then
    
    Do
        RngFind = RngFind & "," & rng.Address
        Set rng = withinRng.FindNext(rng)
        
    Loop While Not rng Is Nothing And rng.Address <> findAddress
    RngFind = Right(RngFind, Len(RngFind) - 1)

    
End If
Exit Function
1:
RngFind = ""
End Function

