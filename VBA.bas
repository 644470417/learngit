Attribute VB_Name = "自动程序"
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
vbStatus = MsgBox("确定要更新状态吗？", vbInformation + vbOKCancel, "提示")
If vbStatus = vbCancel Then End

If Sheets("全流程").FilterMode Then
    Sheets("全流程").ShowAllData
End If
j = Range("B1000").End(xlUp).Row

mDate = Range("B1")
If Day(mDate) > 26 Then
    a = Month(mDate) + 1
Else
    a = Month(mDate)
End If
If Left(a, 1) = "0" Then a = Right(a, 1)

'更新在途状态
For i = j To 3 Step -1
    If Range("E" & i) = "入库" Then
        Rows(i).Delete
    ElseIf Range("E" & i) = "测试1" Then
        Range("E" & i) = "入库"
    ElseIf Range("E" & i) = "烧结2" Then
        Range("E" & i) = "测试1"

    ElseIf Range("E" & i) = "氢碎1" Then
        Range("E" & i) = "氢碎2"
    End If

Next

Set wk1 = OpenBooks("\\Hanhd\05--粉料计划\氢碎日计划.xlsm", True)
Set wk2 = OpenBooks("\\Hanhd\05--粉料计划\气流磨生产计划.xlsx", True)
Set wk3 = OpenBooks("\\Yt2\细粉搅拌记录\2018年细粉搅拌记录\2018." & a & ".xls", True)
wk.Sheets("全流程").Activate

'
Sheets("氢碎1").Range("B:B").ClearContents
Sheets("氢碎2").Range("B:B").ClearContents
Sheets("气流磨").Range("B:B").ClearContents
Sheets("烧结2").Range("B:B").ClearContents

Application.Calculation = xlCalculationAutomatic
''更新当日熔炼批次
ReDim arr(50, 1)
With wk1.Sheets("汇总")
    If .FilterMode Then
        .ShowAllData
    End If
    r = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = r - 100 To r
        test = mDate
        If .Cells(i, "A") = mDate And InStr(.Cells(i, "D"), "对混") = 0 Then
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
            brr(i, 3) = "熔炼"
        Next
    End If

End With
r = Cells(Rows.Count, 2).End(xlUp).Row + 1
Range(Cells(r, "B"), Cells(r + UBound(brr), "E")) = brr
Erase arr: Erase brr: r = 0: k = 0: i = 0

'更新铸片对混批次和氢碎日计划
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
            If InStr(.Cells(i, "F"), "次日磨粉") > 0 Then
                arr(k, 3) = "氢碎2"
                With Sheets("氢碎2")
                    k1 = k1 + 1
                    .Cells(k1, "B") = arr(k, 1)
                End With
            Else
                arr(k, 3) = "氢碎1"
                With Sheets("氢碎1")
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
            Cells(r, "E") = "熔炼"
        End If
    Next
End With

'更改氢碎状态
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

'更改气流磨状态
With wk2.Sheets("2018年气流磨计划表")
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
               brr(i, 5) = "气流磨"
            End If
        Next
    Next

End With

For i = 0 To UBound(arr)
    Sheets("气流磨").Cells(i + 1, 2) = arr(i)
Next
Erase arr: r = 0: k = 0: j = 0
'更改单一粉烧结状态
With wk3.Sheets("单一粉")
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
               brr(i, 5) = "烧结2"
               brr(i, 1) = mDate - 1
            End If
        Next
    Next

End With

For i = 0 To UBound(arr)
    Sheets("烧结2").Cells(i + 1, 2) = arr(i)
Next
Erase arr: r = 0: k = 0: j = 0

r = Cells(Rows.Count, 2).End(xlUp).Row
Range(Cells(3, "A"), Cells(r, "E")) = brr
Range("F3:G3").Copy
Range("F3:G300").PasteSpecial xlPasteFormulas, xlPasteSpecialOperationNone

'更改烧结1状态
For i = 3 To r
    If Cells(i, "E") = "气流磨" And Cells(i, "F") = 0 Then
        Cells(i, "E") = "烧结1"
    End If
    '更改装小桶重量
    If Cells(i, "E") <> "氢碎1" And Cells(i, "E") <> "氢碎2" And Cells(i, "E") <> "熔炼" And Cells(i, "I") <> "实验料" And Cells(i, "D") = "" Then
        If Mid(Cells(i, "C"), 8, 1) >= "A" Then
            Cells(i, "D") = 600
        Else
            Cells(i, "D") = 575
        End If
    End If
Next

wk1.Close False: wk2.Close False: wk3.Close False
Set wk1 = Nothing: Set wk2 = Nothing: Set wk3 = Nothing

'备份需求计划
a = CStr(Format(mDate - 1, "yymmdd"))
b = "2018年需求计划" & a
For Each wk1 In Workbooks
    If wk1.Name = "2018年需求计划.xlsm" Then wk1.Close False
Next
FileCopy "\\Hanhd\05--粉料计划\2018年需求计划.xlsm", "\\Hanhd\05--粉料计划\需求计划备份\" & b & ".xlsm"
Set wk1 = OpenBooks("\\Hanhd\05--粉料计划\2018年需求计划.xlsm", False, "12.3")

'
''备份需求计划
'wk1.Sheets("熔炼计划").Range("B3:BA10").Copy
'wk1.Sheets("熔炼计划").Range("B3:BA10").Paste
''刷新计算按钮
'Application.Run
''更新粉料报表
'
''刷新粉料需求报表
Application.ScreenUpdating = True
Exit Sub
EH:
MsgBox "出现错误，请关闭工作簿并且不保存，然后手动更新报表。", vbCritical + vbOKOnly, "Visual Basic for Application"
Application.ScreenUpdating = True
End Sub

Sub test ()
dim i%,j%
dim wk1,wk2

set wk1=Workbooks("2018年需求计划.xlsm")
Application.ScreenUpdating=False

End Sub

'--------------------------
'     打开工作簿
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
'Set wk2 = Workbooks(b)  '调试
If pwd <> "" Then
    Set OpenBooks = Workbooks.Open(a, , myReadOnly, Password:=pwd)
Else
    Set OpenBooks = Workbooks.Open(a, , myReadOnly)
End If
End If
End Function

'************************************
'功能：查找单元格
'
'参数 findText：要查找的区域 withinRng:查找单元格区域
'返回找到的单元格地址
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

