Attribute VB_Name = "modMain"
Option Explicit

Const defaultHost = "0.0.0.0"
Const defaultPort = 5555
Const defaultCategory = "Pyxcel"
Const modUDF = "udfs"


'https://stackoverflow.com/questions/35750449/how-can-i-assign-a-variant-to-a-variant-in-vba
Private Sub LetSet(ByRef variable As Variant, ByVal value As Variant)
    If IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If
End Sub


Private Sub auto_open()
    Workbooks.Open
    Application.Caption = ("[PYXCEL]")
    Application.Calculation = xlCalculationAutomatic
    GenerateUDFs
End Sub


Private Function ServerUrl() As String
    Dim sHost As String: sHost = Environ("PYXCEL_HOST")
    Dim sPort As String: sPort = Environ("PYXCEL_PORT")
    
    If sHost = vbNullString Then
        sHost = defaultHost
    End If
    
    If Not IsNumeric(sPort) Then
        sPort = defaultPort
    End If
    
    ServerUrl = "http://" + sHost + ":" + sPort + "/"
End Function


'Asks the pyxcel server to import a given module
Function ImportModule(File As Variant) As Boolean

    Dim dArgs As Dictionary: Set dArgs = New Dictionary
    
    dArgs.Add "dir", CurDir()
    dArgs.Add "file", File
    
    ImportModule = QueryAPI("import", dArgs)

    GenerateUDFs
    
End Function

Function ExecutePython(commands As Variant) As Variant

    Dim sCommands As String
    Dim i As Integer
    
    If TypeName(commands) = "Range" Then
        Dim cell As Range
        For Each cell In commands
            If cell <> "" Then
                sCommands = sCommands & cell & vbCrLf
            End If
        Next cell
    Else
        sCommands = commands
    End If

    Dim dReq As Dictionary: Set dReq = New Dictionary
    dReq.Add "command", sCommands
    ExecutePython = JsonToRange(QueryAPI("execute", dReq))

End Function

Private Function QueryAPI(QueryName As String, Optional arg As Object) As Variant
    Dim httpClient As ServerXMLHTTP60
    Set httpClient = New MSXML2.ServerXMLHTTP60
    Dim sUrl, sMsg As String
    
    If arg Is Nothing Then
        sMsg = "{}"
    Else
        sMsg = Application.Run("JsonConverter.ConvertToJson", arg)
    End If
    
    sUrl = ServerUrl + QueryName
    With httpClient
        .Open "POST", sUrl, False
        .setRequestHeader "Content-type", "application/json"
        .setRequestHeader "Cache-Control", "max-age=0"
        .send sMsg
        sMsg = .responseText
    End With
    Dim oRes As Object
    Set oRes = Application.Run("JsonConverter.ParseJson", sMsg)
    If Not oRes.Item("success") Then
        Dim dError As Dictionary
        Set dError = oRes.Item("error")
        Debug.Print "Response: ", dError.Item("description")
        Err.Raise vbObjectError + dError.Item("code"), "QueryAPI", dError.Item("description")
    End If
    LetSet QueryAPI, oRes.Item("result")

End Function

Private Function JsonToRange(var As Variant) As Variant

    If IsObject(var) Then
        'List or Matrix
    
        Dim i, j As Long
        Dim rng As Variant
        Dim nRows, nCols As Integer
    
        If TypeOf var.Item(1) Is Collection Then
            'Matrix
            nCols = var.Count
            nRows = var.Item(1).Count
            ReDim rng(1 To nCols, 1 To nRows)
            For i = 1 To nCols
                For j = 1 To nRows
                    rng(i, j) = var.Item(i).Item(j)
                Next j
            Next i
        Else
            'List
            nRows = var.Count
            ReDim rng(1 To nRows)
            For i = 1 To nRows
                rng(i) = var.Item(i)
            Next i
        End If
        JsonToRange = rng
        
    Else
        JsonToRange = var
    
    End If
End Function

Private Function PythonCall(FuncName As String, Args As Collection) As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    Dim i, j As Integer
    Dim d As Dictionary
    Set d = New Dictionary
    d.Add "foo", FuncName
    
    If Args.Count > 0 Then
        Dim vArgs As Variant
        ReDim vArgs(1 To Args.Count)
        For i = 1 To Args.Count
            vArgs(i) = Args.Item(i)
        Next i
        d.Add "args", vArgs
    End If
    
    LetSet PythonCall, JsonToRange(QueryAPI("eval", d))
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Function

Private Function SanitizeName(s As String) As String
    SanitizeName = UCase(Left(s, 1)) & Mid(s, 2)
End Function

Private Function FunctionVBACode(PyFuncName As String, cArgs As Collection) As String()

    'Dim dTypes As Dictionary
    'Set dTypes = New Dictionary
    'dTypes.Add "int", "Integer"
    'dTypes.Add "double", "Double"
    'dTypes.Add "bool", "Bolean"
    'dTypes.Add "str", "String"
    
    Dim i As Integer
    Dim s, sVBAFuncName, res() As String
    ReDim res(1 To 5 + cArgs.Count)
    Dim bFirstArg As Boolean
    Dim dArg As Dictionary
    
    sVBAFuncName = SanitizeName(PyFuncName)
    
    'Function declaration
    s = "Function " & sVBAFuncName & "("
    bFirstArg = True
    For Each dArg In cArgs
        If Not bFirstArg Then
            s = s & ", "
        End If
        s = s & dArg.Item("name")
        bFirstArg = False
    Next dArg
    
    s = s & ") As Variant"
    res(1) = s
    
    'Collection Objection creation
    res(2) = "    Dim c As Collection : Set c = New Collection"
    
    i = 2
    For Each dArg In cArgs
        i = i + 1
        res(i) = "    c.Add " & dArg.Item("name")
    Next dArg
    
    'PythonCall function
    s = "    " & sVBAFuncName
    s = s & " = Application.Run(""modMain.PythonCall"", """ & PyFuncName & """, c)"
    res(i + 1) = s
    
    'Footer
    res(i + 2) = "End Function"
    res(i + 3) = ""
    
    FunctionVBACode = res
End Function


Private Sub GenerateUDFs()

    On Error Resume Next

    Dim dRes As Dictionary
    Set dRes = QueryAPI("signatures")

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(modUDF)
    If Err <> 0 Then
        Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
        VBComp.Name = modUDF
        Err.Clear
    End If
    Set CodeMod = VBComp.CodeModule
    
    Dim cArgs As Collection
    Dim dArg, dFoo As Dictionary
    Dim s, sArg, sCategory As String
    Dim key As Variant
    Dim ArgDesc() As String
    Dim i, j As Long
    Dim bFirstArg As Boolean
    
    With CodeMod

        .DeleteLines 1, .CountOfLines
        
        i = 0
        For Each key In dRes
            Set dFoo = dRes.Item(key)
            Set cArgs = dFoo.Item("args")
            
            For Each s In FunctionVBACode(CStr(key), cArgs)
                i = i + 1
                .InsertLines i, s
            Next s
            
            If dFoo.Item("category") Is Nothing Then
                sCategory = defaultCategory
            Else
                sCategory = dFoo.Item("category")
            End If
            
            Application.MacroOptions _
                Macro:=SanitizeName(CStr(key)), _
                Description:=dFoo.Item("doc"), _
                Category:=sCategory
            
            If cArgs.Count > 0 Then
                ReDim sArgDesc(1 To cArgs.Count)
                For j = 1 To cArgs.Count
                    sArgDesc(j) = cArgs.Item(j).Item("desc")
                Next j
                Application.MacroOptions ArgumentDescriptions:=sArgDesc
            End If

        Next key
    End With

End Sub

