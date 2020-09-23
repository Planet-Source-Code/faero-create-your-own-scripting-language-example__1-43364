Attribute VB_Name = "Lang"
'''''''Edit this!'''''
'
'Put each command in the array, seperated by commas as shown.
'After each command, put it's associated function (see below)
'
'For example, the MessageBox command has the function
'that makes it work called MessageboxCmnd
'So you put in MESSAGEBOX,MessageboxCmnd
'
'Your function must only have one argument, type long,
'for the line number you are on
'The function must be put in the class module :)
'Use nothing for a function name to make nothing happen
'
'See the premade functions for an example of what i mean!!!
Const scriptComms As String = "COMMENT,nothing,MESSAGEBOX,MessageboxCmnd,FAERO,FaeroCmnd"

'Dont edit the following declerations

Global scriptLines() As String
Global commRef() As String

Function runScript()
    commRef = Split(scriptComms, ",", -1, vbTextCompare)
    Dim theLine As Long
    For theLine = 0 To UBound(scriptLines)
        DoCommand theLine
    Next theLine
End Function
Function DoCommand(commandLine As Long)
    Dim cName As String
    Dim theComm As Long
    Dim langObject As New langObj
    cName = UCase(GetCommName(scriptLines(commandLine)))
    For theComm = 0 To UBound(commRef) Step 2
        If cName = commRef(theComm) Then
            If LCase(commRef(theComm + 1)) = "nothing" Then
                Exit For
            Else
                CallByName langObject, commRef(theComm + 1), VbMethod, commandLine
            End If
            'ALL FUNCTIONS ARE IN THE LANGOBJ CLASS
            Exit For
        End If
    Next theComm
End Function
Function GetCommName(text$)
    cLine$ = Trim(text$)
    'cline$ now contains the line of code
    'with no leading/trailing spaces
    cLine$ = Replace(cLine$, Chr(9), "") 'cline$ has no tabs
    part$ = Mid$(cLine$, 1, 1)
    If part$ = "'" Then
        GetCommName = "COMMENT"
        Exit Function
    Else
    For i = 1 To Len(cLine$)
        parter$ = Mid$(cLine$, i, 1)
        If parter$ = " " Or parter$ = "(" Or parter$ = "=" Then
            Exit For
        ElseIf parter$ = "%" Or parter$ = "$" Then
            'Its a line of code that edits a variable value.
            GetCommName = "VAR"
            Exit Function
        Else
            theComm$ = theComm$ + parter$
        End If
    Next i
    GetCommName = theComm$
    Exit Function
    End If
End Function

Function GetBrackets(text$) As String
    usage$ = Trim(text$)
    enderBrack = -1
    starterBrack = -1
    
    starterBrack = InStr(1, usage$, "(") + 1
    enderBrack = InStrRev(usage$, ")")
    
    If enderBrack = -1 Then
        enderBrack = Len(usage$) + 1
    End If
    
    If Not starterBrack = -1 Then
        returner$ = Mid$(usage$, starterBrack, _
        enderBrack - starterBrack)
    Else
        returner$ = "error"
    End If
    
    GetBrackets = returner$
End Function
Function GetParam(text$, paramnum) As String
    
    pnum = paramnum
    t = 1
    numofp = 1
    ignored = False
    text$ = Trim(text$)
    text$ = Replace(text$, Chr(9), "")
    For t = 1 To Len(text$)
        part$ = Mid$(text$, t, 1)
        If part$ = "," Then
            If Not ignored Then
                numofp = numofp + 1
            End If
        ElseIf part$ = Chr(34) Then
            ignored = Not ignored
        End If
    Next t
    
    If numofp < pnum Then
        'Debug
    End If
    
    i = 1
    curParam = 1
    ignored = False
    
    For i = 1 To Len(text$)
        thePart$ = Mid$(text$, i, 1)
        If thePart$ = "," Then
            If Not ignored Then
                curParam = curParam + 1
                If curParam = pnum + 1 Then
                    Exit For
                End If
            End If
        ElseIf part$ = Chr(34) Then
            ignored = Not ignored
        Else
            If curParam = pnum Then
                theParam$ = theParam$ + thePart$
            End If
        End If
    Next i
    GetParam = theParam$
End Function

Function CountParams(text$) As Integer
    
    numofp = 1
    ignored = False
    text$ = Trim(text$)
    text$ = Replace(text$, Chr(9), "")
    For t = 1 To Len(text$)
        part$ = Mid$(text$, t, 1)
        If part$ = "," Then
            If Not ignored Then
                numofp = numofp + 1
            End If
        ElseIf part$ = Chr(34) Then
            ignored = Not ignored
        End If
    Next t
    CountParams = numofp

End Function
Function DataType(text$) As Integer
    
    text$ = Trim(text$)
    If Left(text$, 1) = Chr(34) Then
        DataType = 1
        Exit Function
    Else
        DataType = 2
        Exit Function
    End If

End Function
Function GetData(text$, Optional ByRef stringData As String, Optional ByRef intData As Double)
    
    text$ = Trim(text$)
    If Left(text$, 1) = Chr(34) Then
        stringData = Mid(text$, 2, Len(text$) - 2)
    Else
        intData = Val(text$)
    End If

End Function
