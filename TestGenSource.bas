Attribute VB_Name = "Module1"
Sub Button1_Click()
Dim ins(20) As Double
Dim positives() As Double
Dim negatives() As Double
Dim insPrint(20) As String
Dim currValue As Double
Dim I As Byte, ptr As Byte, S As Byte
For I = 0 To 19
currValue = CDbl(Range(addres_helper(I, 4)).Value)
If currValue >= 0 Then
ptr = getNextIndexForArray(positives)
ReDim Preserve positives(ptr)
positives(ptr) = currValue
Else
ptr = getNextIndexForArray(negatives)
ReDim Preserve negatives(ptr)
negatives(ptr) = currValue
End If

Next I
'debug output
For I = 0 To UBound(positives) - LBound(positives)
     insPrint(I) = CStr(positives(I))
 Next I

For I = (UBound(positives) - LBound(positives) + 1) To UBound(positives) - LBound(positives) + 1 + UBound(negatives) - LBound(negatives)
     insPrint(I) = CStr(negatives(I - (UBound(positives) - LBound(positives) + 1)))
 Next I
'MsgBox (Join(insPrint, vbCrLf))
'Sheet output
For S = 0 To 19
If S <= UBound(positives) - LBound(positives) Then
Range(addres_helper(S, 9)).Value = positives(S)
Else
Range(addres_helper(S, 9)).Value = negatives(S - (UBound(positives) - LBound(positives) + 1))
End If
Next S
End Sub

Sub TestGenerate()
'Random testi
Dim filePath As String
    filePath = "C:\Users\Kristine\source\RTU\RisAlg\3MD\dip107\src\test\resources\dip107\positive-tests.csv"
Dim a As Double
Dim Randresults(20)
For I = 0 To 20
'vai programmai j�saprot milzu skaitlu zinaatniskais pieraksts???laikam buus tomeer jaasaprot...
a = Application.WorksheetFunction.RandBetween(-999, 0) + (Application.WorksheetFunction.RandBetween(0, 100001) / 100000)
Range("B1").Value = a
'so have to walk it, cann not simply toString or can I?

Randresults(I) = CStr(a) + ", " + """" + ExcelToTestInput(Range("B3:F12")) + """"

'MsgBox (Randresults(I))
Next I
Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(filePath, True, True)
    Fileout.Write "a, TheLastOneIsExpectedOutputResult" + vbCrLf
    Fileout.Write Join(Randresults, vbCrLf)
    Fileout.Close
End Sub

Function ExcelToTestInput(ByVal myRange As Range) As String
    RangeToString = ""
    If Not myRange Is Nothing Then
        Dim myCell As Range
        For Each myCell In myRange
        If myCell.Value = "" Then

        ElseIf myCell.Address = "$B$3" Or myCell.Address = "$B$8" Then
        ExcelToTestInput = ExcelToTestInput & myCell.Value & vbCrLf
        ElseIf Mid(myCell.Address, 2, 1) = "F" Then
        ExcelToTestInput = ExcelToTestInput & vbTab & Format(myCell, "0.00") & vbCrLf
        ElseIf Mid(myCell.Address, 2, 1) = "B" Then
        ExcelToTestInput = ExcelToTestInput & Format(myCell, "0.00")
        Else
            ExcelToTestInput = ExcelToTestInput & vbTab & Format(myCell, "0.00")
        End If
        Next myCell
    End If
End Function

Function getNextIndexForArray(a() As Double) As Byte
If (Not a) = -1 Then
    ' Array has NOT been initialized
    getNextIndexForArray = 0
Else
getNextIndexForArray = UBound(a) + 1
End If
End Function


Function addres_helper(a As Byte, startRow As Byte) As String
addres_helper = CStr(Chr(66 + (a Mod 5))) + CStr(Application.WorksheetFunction.Floor_Math(a / 5, 1) + startRow)
End Function
