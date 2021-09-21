Attribute VB_Name = "Module1"





'References required:
'Microsoft Scripting Runtime
'UIAutomationClient

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function MessageBoxW Lib "User32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If

Public Function calculate_calculator(inp As String, Multipl As String) As String
    #If VBA7 Then
        Dim CalcHwnd As LongPtr
    #Else
        Dim CalcHwnd As Long
    #End If
    Dim keypadDict As Scripting.Dictionary
    Dim CalculatorResult As String
    Dim CalculatorExpression As String
    
    CalcHwnd = Find_Calculator()
    
    If CalcHwnd <> 0 Then
        Set keypadDict = Build_Keys_Dict(CalcHwnd)
        Click_Keys "|CE|(" + Replace(Replace(UCase(inp), "E-", "E|+-|"), "E+", "E") + ")X(" + Replace(Replace(UCase(Multipl), "E-", "E|+-|"), "E+", "E") + ")=", CalcHwnd, keypadDict
        CalculatorResult = Get_Result(CalcHwnd)
        'remove thousands seperator!
        calculate_calculator = Replace(CalculatorResult, ",", "")
    Else
        MsgBox "Calculator isn't running"
        calculate_calculator = "0.0"
    End If
End Function


Public Sub Test_Automate_Calculator()

    #If VBA7 Then
        Dim CalcHwnd As LongPtr
    #Else
        Dim CalcHwnd As Long
    #End If
    Dim keypadDict As Scripting.Dictionary
    Dim CalculatorResult As String
    Dim CalculatorExpression As String
    
    CalcHwnd = Find_Calculator()
    
    If CalcHwnd <> 0 Then
        Set keypadDict = Build_Keys_Dict(CalcHwnd)
        Click_Keys "|CE|3.6+5=|SQRT||RECIP|=", CalcHwnd, keypadDict
        CalculatorResult = Get_Result(CalcHwnd)
        CalculatorExpression = Get_Expression(CalcHwnd)
        MsgBoxW "Result:  " & CalculatorResult & vbCrLf & _
                "Expression: " & CalculatorExpression
    Else
        MsgBox "Calculator isn't running"
    End If
    
End Sub


Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Microsoft Excel") As VbMsgBoxResult
    Prompt = Prompt & vbNullChar 'Add null terminators
    Title = Title & vbNullChar
    MsgBoxW = MessageBoxW(Application.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function


#If VBA7 Then
Public Function Find_Calculator() As LongPtr
#Else
Public Function Find_Calculator() As Long
#End If
   
    'Find the Calculator window and return its window handle

    Dim UIAuto As IUIAutomation
    Dim Desktop As IUIAutomationElement
    Dim CalcWindow As IUIAutomationElement
    Dim ControlTypeAndNameCond As IUIAutomationCondition
    Dim WindowPattern As IUIAutomationWindowPattern
    
    Find_Calculator = 0

    'Create UIAutomation object
    
    Set UIAuto = New CUIAutomation
    
    'Conditions to find the main Calculator window on the Desktop
    'ControlType:   UIA_WindowControlTypeId (0xC370)
    'Name:          "Calculator"
    
    With UIAuto
        Set Desktop = .GetRootElement
        Set ControlTypeAndNameCond = .CreateAndCondition(.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_WindowControlTypeId), _
                                                         .CreatePropertyCondition(UIA_NamePropertyId, "Calculator"))
    End With
    Set CalcWindow = Desktop.FindFirst(TreeScope_Children, ControlTypeAndNameCond)
    
    If Not CalcWindow Is Nothing Then
    
        'Restore the Calculator window, because it must not be minimised (off screen/iconic) in order to find the keypad keys
        
        If CalcWindow.CurrentIsOffscreen Then
            Set WindowPattern = CalcWindow.GetCurrentPattern(UIA_WindowPatternId)
            WindowPattern.SetWindowVisualState WindowVisualState.WindowVisualState_Normal
            DoEvents
            Sleep 100
        End If
        
        'Return the Calculator's window handle
        
        Find_Calculator = CalcWindow.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
                
    End If

End Function


#If VBA7 Then
Public Function Build_Keys_Dict(CalcHwnd As LongPtr) As Scripting.Dictionary
#Else
Public Function Build_Keys_Dict(CalcHwnd As Long) As Scripting.Dictionary
#End If

    'Create a dictionary which maps each keypad key to its UI automation element via the AutomationId string
    
    Dim keysMapping As Variant
    Dim i As Long
    Dim key As cKey
    
    keysMapping = Split("0,num0Button,1,num1Button,2,num2Button,3,num3Button,4,num4Button,5,num5Button,6,num6Button,7,num7Button,8,num8Button,9,num9Button," & _
                        ".,decimalSeparatorButton,/,divideButton,X,multiplyButton,-,minusButton,+,plusButton,=,equalButton,%,percentButton," & _
                        "(,openParenthesisButton,),closeParenthesisButton,E,expButton," & _
                        "|+-|,negateButton,|RECIP|,invertButton,|SQR|,xpower2Button,|SQRT|,squareRootButton," & _
                        "|CE|,clearEntryButton,|C|,clearButton,|BS|,backSpaceButton," & _
                        "|MC|,ClearMemoryButton,|MR|,MemRecall,|M+|,MemPlus,|M-|,MemMinus,|MS|,memButton", ",")
    
    Set Build_Keys_Dict = New Scripting.Dictionary
   
    For i = 0 To UBound(keysMapping) Step 2
        Set key = New cKey
        key.keypadKey = keysMapping(i)
        Set key.UIelement = Find_Key(CalcHwnd, CStr(keysMapping(i + 1)))
        Build_Keys_Dict.Add keysMapping(i), key
    Next

End Function


#If VBA7 Then
Private Function Find_Key(CalcHwnd As LongPtr, keyAutomationId As String) As IUIAutomationElement
#Else
Private Function Find_Key(CalcHwnd As Long, keyAutomationId As String) As IUIAutomationElement
#End If

    'Find the specified Calculator key by its AutomationId
    
    Dim UIAuto As IUIAutomation
    Dim Calc As IUIAutomationElement
    Dim KeyCond As IUIAutomationCondition
    
    'Get the Calculator automation element from its window handle
    
    Set UIAuto = New CUIAutomation
    Set Calc = UIAuto.ElementFromHandle(ByVal CalcHwnd)
    
    'Condition to find the specified Calculator key, for example
    'AutomationId:   "num3Button"
    
    Set KeyCond = UIAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, keyAutomationId)
    
    'Must use TreeScope_Descendants to find the keypad keys, rather than TreeScope_Children, because the Calculator keys are not immediate children of the Calculator window.
    'TreeScope_Descendants searches the element's descendants, including children.  TreeScope_Children searches only the element's immediate children.
    'Note that the memory keys don't exist if the Calculator is in 'Keep on top' mode
    
    Set Find_Key = Calc.FindFirst(TreeScope_Descendants, KeyCond)
    
End Function


#If VBA7 Then
Public Sub Click_Keys(keys As String, CalcHwnd As LongPtr, Keypad_Dict As Dictionary)
#Else
Public Sub Click_Keys(keys As String, CalcHwnd As Long, Keypad_Dict As Dictionary)
#End If

    'Automate the Calculator by clicking the specified keys

    Dim UIAuto As IUIAutomation
    Dim Calc As IUIAutomationElement
    Dim InvokePattern As IUIAutomationInvokePattern
    Dim i As Long, p As Long
    Dim thisKey As String
    Dim key As cKey
    
    'Get the Calculator automation element from its window handle
    
    Set UIAuto = New CUIAutomation
    Set Calc = UIAuto.ElementFromHandle(ByVal CalcHwnd)
    
    'Parse the keys string, looking up each key in the keypad dictionary and clicking the key via its UIAutomation element
    
    For i = 1 To Len(keys)
    
        thisKey = UCase(Mid(keys, i, 1))
        If thisKey = "|" Then
            'Special key surrounded by "|"
            p = InStr(i + 1, keys, "|")
            thisKey = Mid(keys, i, p + 1 - i)
            i = p
        End If
        
        If Keypad_Dict.Exists(thisKey) Then
            Set key = Keypad_Dict(thisKey)
            If Not (key.UIelement Is Nothing) Then
            Set InvokePattern = key.UIelement.GetCurrentPattern(UIA_InvokePatternId)
            Else
            If thisKey = "|C|" Then
            thisKey = "|CE|"
            Set key = Keypad_Dict(thisKey)
            Set InvokePattern = key.UIelement.GetCurrentPattern(UIA_InvokePatternId)
            Else
            If thisKey = "|CE|" Then
            thisKey = "|C|"
            Set key = Keypad_Dict(thisKey)
            Set InvokePattern = key.UIelement.GetCurrentPattern(UIA_InvokePatternId)
            End If
            End If
            End If
            InvokePattern.Invoke
            DoEvents
            Sleep 100
        Else
            MsgBox "Key '" & thisKey & "' not found in keypad dictionary. Check syntax of keys argument", vbExclamation
        End If
        
    Next
        
End Sub


#If VBA7 Then
Public Function Get_Result(CalcHwnd As LongPtr) As String
#Else
Public Function Get_Result(CalcHwnd As Long) As String
#End If

    'Extract the Calculator result string
    
    Dim UIAuto As IUIAutomation
    Dim Calc As IUIAutomationElement
    Dim ResultCond As IUIAutomationCondition
    Dim Result As IUIAutomationElement
    
    'Get the Calculator automation element from its window handle
    
    Set UIAuto = New CUIAutomation
    Set Calc = UIAuto.ElementFromHandle(ByVal CalcHwnd)
    
    'Condition to find the Calculator results
    'Name:   "Display is 7.82842712474619"
    'AutomationId:   "CalculatorResults"
    
    Set ResultCond = UIAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "CalculatorResults")
    Set Result = Calc.FindFirst(TreeScope_Descendants, ResultCond)
    
    If Result Is Nothing Then
        Set ResultCond = UIAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "CalculatorAlwaysOnTopResults")
        Set Result = Calc.FindFirst(TreeScope_Descendants, ResultCond)
    End If
    
    Get_Result = Mid(Result.CurrentName, InStr(Result.CurrentName, " is ") + Len(" is "))
    
End Function


#If VBA7 Then
Public Function Get_Expression(CalcHwnd As LongPtr) As String
#Else
Public Function Get_Expression(CalcHwnd As Long) As String
#End If

    'Extract the Calculator expression string

    Dim UIAuto As IUIAutomation
    Dim Calc As IUIAutomationElement
    Dim ExpressionCond As IUIAutomationCondition
    Dim Expression As IUIAutomationElement
    
    'Get the IE automation element from its window handle
    
    Set UIAuto = New CUIAutomation
    Set Calc = UIAuto.ElementFromHandle(ByVal CalcHwnd)
    
    'Condition to find the Calculator expression, if it exists
    'Name:   "Expression is ?(8) + 5="
    'AutomationId:   "CalculatorExpression"
    
    Set ExpressionCond = UIAuto.CreatePropertyCondition(UIA_AutomationIdPropertyId, "CalculatorExpression")
    
    Set Expression = Calc.FindFirst(TreeScope_Descendants, ExpressionCond)
    
    If Not Expression Is Nothing Then
        Get_Expression = Mid(Expression.CurrentName, InStr(Expression.CurrentName, " is ") + Len(" is "))
    Else
        Get_Expression = ""
    End If
    
End Function

Sub Button1_Click()
Dim ins(20) As String
Dim positives() As String
Dim negatives() As String
Dim insPrint(20) As String
Dim currValue As Double
Dim i As Byte, ptr As Byte, S As Byte
For i = 0 To 19
currValue = CDbl(Range(addres_helper(i, 4)).Value)
If currValue >= 0 Then
ptr = getNextIndexForArray(positives)
ReDim Preserve positives(ptr)
positives(ptr) = Range(addres_helper(i, 4)).Value
Else
ptr = getNextIndexForArray(negatives)
ReDim Preserve negatives(ptr)
negatives(ptr) = Range(addres_helper(i, 4)).Value
End If

Next i
'debug output
For i = 0 To UBound(positives) - LBound(positives)
     insPrint(i) = CStr(positives(i))
 Next i

For i = (UBound(positives) - LBound(positives) + 1) To UBound(positives) - LBound(positives) + 1 + UBound(negatives) - LBound(negatives)
     insPrint(i) = CStr(negatives(i - (UBound(positives) - LBound(positives) + 1)))
 Next i
'MsgBox (Join(insPrint, vbCrLf))
'Sheet output
For S = 0 To 19
Range(addres_helper(S, 9)).NumberFormat = "@"
If S <= UBound(positives) - LBound(positives) Then
Range(addres_helper(S, 9)).Value = positives(S)
Else
Range(addres_helper(S, 9)).Value = negatives(S - (UBound(positives) - LBound(positives) + 1))
End If
Next S
End Sub
 
Sub TestGenerate()
'Random testi
'fuck = SplitRe("-2.32758848538112E+28", "[eE]")
'format_helper ("-2.32758848538112E-3")
Dim filePath As String
    filePath = "C:\Users\Kristine\source\RTU\RisAlg\3MD\dip107\src\test\resources\dip107\positive-tests.csv"
Dim a As Double
Dim i As Integer
Dim Randresults(20)
Dim fuckyou
Range("B1").Value = -249.04835
fuckyou = RecalculateUsingCalculator(Range("B4:F7"))
'so have to walk it, cann not simply toString or can I?
Button1_Click
Randresults(0) = """" + CStr(-249.04835) + """" + ", " + """" + ExcelToTestInput(Range("B3:F12")) + """"

For i = 1 To 19
'vai programmai jâsaprot milzu skaitlu zinaatniskais pieraksts???laikam buus tomeer jaasaprot...
a = Application.WorksheetFunction.RandBetween(-999, 0) + (Application.WorksheetFunction.RandBetween(0, 100001) / 100000)
Range("B1").Value = a
'Now calcuate using calculator because excel calculates wrongly
'wtf CANNOT PASS THE RETURN VALUE OF FUNCTION, MUST CREATE A VARIABLE?????
'Test_Automate_Calculator
fuckyou = RecalculateUsingCalculator(Range("B4:F7"))
'so have to walk it, cann not simply toString or can I?
Button1_Click
'turns out that jUnit cannot parse negative decimals properly - it adds a space after each of the decimals chars???
'so the first param MUST be in quotes as well...
Randresults(i) = """" + CStr(a) + """" + ", " + """" + ExcelToTestInput(Range("B3:F12")) + """"

'MsgBox (Randresults(I))
Next i
'add incorrect input test
Dim wrongInputs() As String
wrongInputs = Split("-1-2-3, " + Chr(34) + "input-output error" + Chr(34) + ":wasd, " + Chr(34) + "input-output error" + Chr(34) + ":âð, " + Chr(34) + "input-output error" + Chr(34) + ":0.1.23.4.5, " + Chr(34) + "input-output error" + Chr(34) + ":", ":")
Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(filePath, True, True)
    Fileout.Write "a, TheLastOneIsExpectedOutputResult" + vbCrLf
    Fileout.Write Join(Randresults, vbCrLf) + vbCrLf + Join(wrongInputs, vbCrLf)
    Fileout.Close
End Sub

'RecalculateUsingCalculator
Function RecalculateUsingCalculator(ByVal myRange As Range)
    If Not myRange Is Nothing Then
        Dim myCell As Range
        Dim prevCell As Range
        Dim tStr As String
        For Each myCell In myRange
        If myCell.Address = "$B$4" Then
        myCell.Value = 0.1
        tStr = "0.1"
        Else
        tStr = calculate_calculator(CStr(Range("B1").Value2), tStr)
        myCell.NumberFormat = "@"
        myCell.Value = UCase(tStr)
        End If
        Next myCell
    End If
End Function

Function GetPrevAddress(currAddr As String) As String
Dim rrow As Integer
Dim col As String
rrow = CInt(Mid(currAddr, 2, Len(currAddr) - 1))
col = Mid(currAddr, 1, 1)
If col = "B" Then
GetPrevAddress = "F" + CStr(rrow - 1)
Else
GetPrevAddress = CStr(Chr(Asc(col) - 1)) + CStr(rrow)
End If
End Function

Function ExcelToTestInput(ByVal myRange As Range) As String
    If Not myRange Is Nothing Then
        Dim myCell As Range
        Dim tStr As String
        For Each myCell In myRange
        If (Mid(myCell.Address, 4, 2) <> "3" And Mid(myCell.Address, 4, 2) <> "8" And myCell.Value <> "") Then
        'even storing at seperate variable did not help as Excel CALCULATES differently than calc - therefore no system of counting
        'avaliable => impossible to write any tests, because the TRUTH is not known...
        tStr = format_helper(myCell.Value2)
        End If
        If myCell.Value = "" Then

        ElseIf myCell.Address = "$B$3" Or myCell.Address = "$B$8" Then
        ExcelToTestInput = ExcelToTestInput & myCell.Value & vbCrLf
        ElseIf Mid(myCell.Address, 2, 1) = "F" And Not (Mid(myCell.Address, 4, 2) = "12") Then
        ExcelToTestInput = ExcelToTestInput & vbTab & tStr & vbCrLf
        ElseIf Mid(myCell.Address, 2, 1) = "B" Then
        ExcelToTestInput = ExcelToTestInput & tStr
        ElseIf (Mid(myCell.Address, 4, 2) <> "3" And Mid(myCell.Address, 4, 2) <> "8" And myCell.Value <> "") Then
            ExcelToTestInput = ExcelToTestInput & vbTab & tStr
        End If
        Next myCell
    End If
End Function

Function getNextIndexForArray(a() As String) As Byte
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

Public Function h_discount(a As String) As Integer
If a > 15 Then
h_discount = 250
Else
h_discount = 0
End If
End Function

Public Function format_helper(inpt As String) As String
Dim Sign As String
Dim power As Double
Dim deciSp As String
Dim baseRH As String
Dim baseLH As String
Dim num As Integer
Dim str() As String
Dim strr() As String
Dim inp As String
inp = CStr(inpt)

If (Mid(inp, 1, 1) = "-") Then
Sign = "-"
inp = Mid(inp, 2, Len(inp) - 1)
End If
str = SplitRe(inp, "[eE]")
If (UBound(str) - LBound(str)) < 1 Then
format_helper = round_helper(Sign + inp)
Exit Function
End If

power = CDbl(str(1))
deciSp = "."
strr = Split(str(0), deciSp)
If (UBound(strr) - LBound(strr)) < 1 Then
baseRH = ""
Else
baseRH = strr(1)
End If
baseLH = strr(0)

If (power >= 0) Then
If (power > Len(baseRH)) Then
baseRH = baseRH + String(power - Len(baseRH), "0")
End If
'baseRH = baseRH.slice(0,power) + deciSp + baseRH.slice(power);
'should test could it be that power>= Len(baseRH)-1?
baseRH = Mid(baseRH, 1, power) + deciSp + Mid(baseRH, power)
'if (baseRH.charAt(baseRH.length-1) ==deciSp) baseRH =baseRH.slice(0,-1); // If decSep at RH end? => remove it
If (Mid(baseRH, Len(baseRH), 1) = deciSp) Then
baseRH = Mid(baseRH, 1, Len(baseRH) - 2)
End If
Else
'num= Math.abs(power) - baseLH.length;
num = Abs(power) - Len(baseLH)
'if (num>0) baseLH = "0".repeat(num) + baseLH;                       // Pad with "0" at LH
If (num > 0) Then
baseLH = String(num, "0") + baseLH
End If
'baseLH = baseLH.slice(0, power) + deciSp + baseLH.slice(power);     // Insert "." at the correct place into LH base
baseLH = Mid(baseLH, 1, Len(baseLH) + power) + deciSp + Mid(baseLH, 1 + Len(baseLH) + power, Abs(power))
'if (baseLH.charAt(0) == deciSp) baseLH="0" + baseLH;
If (Mid(baseLH, 1, 1) = deciSp) Then
baseLH = "0" + baseLH
End If
End If
'return Sign + (baseLH + baseRH).replace(/^0*(\d+|\d+\.\d+?)\.?0*$/,"$1");
format_helper = round_helper(Sign + (baseLH + baseRH))
End Function

'for now works only with numdigits=2
Function round_helper(inp As String) As String
Dim numdigits As Integer
numdigits = 2
Dim str() As String
str = Split(inp, ".")
If UBound(str) - LBound(str) = 1 Then

If (Len(str(1)) > numdigits) Then
If (CInt(Mid(str(1), numdigits + 1, 1)) >= 5) Then
If (CInt(Mid(str(1), numdigits, 1)) <> 9) Then
round_helper = str(0) + "." + Mid(str(1), 1, numdigits - 1) + CStr(CInt(Mid(str(1), numdigits, 1)) + 1)
Else
'Todo add while cycle for rounding up to before the decimal seperator!
round_helper = str(0) + "." + Mid(str(1), 1, numdigits - 2) + CStr(CInt(Mid(str(1), numdigits - 2 + 1, 1)) + 1) + "0"
End If
Else
round_helper = str(0) + "." + Mid(str(1), 1, 2)
End If
Else
round_helper = str(0) + "." + str(1) + String(numdigits - Len(str(1)), "0")
End If

End If

End Function

'nesaturees vbNullChar taapeec varu drosi izmantot so!
Public Function SplitRe(Text As String, Pattern As String, Optional IgnoreCase As Boolean) As String()
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.MultiLine = True
    End If

    re.IgnoreCase = IgnoreCase
    re.Pattern = Pattern
    SplitRe = Strings.Split(re.Replace(Text, ChrW(-1)), ChrW(-1))
End Function
