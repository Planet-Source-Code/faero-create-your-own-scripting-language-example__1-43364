VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "langObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Sub MessageboxCmnd(commandLine As Long)
    Dim cLine As String
    cLine = scriptLines(commandLine)
    numParams = CountParams(cLine)
    
    If Not numParams = 3 Then
        MsgBox "Wrong number of paramaters --- " & cLine, vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim theParams As String
    theParams = GetBrackets(cLine)
    param1$ = GetParam(theParams, 1): param2$ = GetParam(theParams, 2): param3$ = GetParam(theParams, 3)
    dtype1 = DataType(param1$): dtype2 = DataType(param2$): dtype3 = DataType(param3$)
    
    If Not dtype1 = 1 Then
        MsgBox "First paramater must be a string! --- " & cLine, vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    ElseIf Not dtype2 = 1 Then
        MsgBox "Second paramater must be a string! --- " & cLine, vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    ElseIf Not dtype3 = 2 Then
        MsgBox "Third paramater must be an integer! --- " & cLine, vbApplicationModal + vbExclamation, "Error"
        Exit Sub
    End If
    Dim msgboxText As String
    Dim msgboxTitle As String
    Dim msgboxSpec As Double
    GetData param1, msgboxText
    GetData param2, msgboxTitle
    GetData param3, , msgboxSpec
    MsgBox msgboxText, msgboxSpec, msgboxTitle
End Sub
Public Sub FaeroCmnd(commandLine As Long)
    Dim cLine As String
    cLine = scriptLines(commandLine)
    numParams = CountParams(cLine)
    
    'pssh, who cares about the paramaters for this command
    
    MsgBox "This is a cool example, is it not?", vbYesNo, "Faero()"
End Sub

