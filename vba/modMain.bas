Attribute VB_Name = "modMain"
Sub Test()
    callPython "sub", 1, 2
End Sub

Function add(x, y) As String

    Dim vArgs
    vArgs = Array(x, y)
    
    Dim sUrl, sJson, sArgs As String
    sUrl = "http://127.0.0.1:5000/" + "query"
    
    sArgs = ConvertToJson(vArgs)
    sJson = ""
    sJson = sJson & """foo"": """ & "add" & ""","
    sJson = sJson & """args"": " & sArgs
    sJson = "{" & sJson & "}"
    
    Dim httpClient As New XMLHTTP60
    httpClient.Open "POST", sUrl, False
    httpClient.setRequestHeader "Content-type", "application/json"
    httpClient.send (sJson)
    add = httpClient.responseText
    
End Function


Function callPython(foo As String, x, y) As String

    Dim vArgs
    vArgs = Array(x, y)
    
    Dim sUrl, sJson, sArgs As String
    sUrl = "http://127.0.0.1:5000/" + "query"
    
    Set sArgs = ConvertToJson(vArgs)
    sJson = ""
    sJson = sJson & """foo"": """ & foo & ""","
    sJson = sJson & """args"": " & sArgs
    sJson = "{" & sJson & "}"
    
    Dim httpClient As New XMLHTTP60
    httpClient.Open "POST", sUrl, False
    httpClient.setRequestHeader "Content-type", "application/json"
    httpClient.send (sJson)
    callPython = httpClient.responseText
    
End Function


Function ExportedFunctions() As String
    Dim sUrl, sResponse As String
    sUrl = "http://127.0.0.1:5000/" + "list_udf"
    
    Dim httpClient As New XMLHTTP60
    httpClient.Open "GET", sUrl, False
    httpClient.send
    'Dim jsonRows As Collection
    Dim dUDF As Dictionary
    Set dUDF = JsonConverter.ParseJson(httpClient.responseText)
    
    Dim arr() As String
    ReDim arr(dUDF.Count - 1)
    For i = 0 To dUDF.Count - 1
        'arr(i) = dUDF.keys(i)
    Next i
    
    ExportedFunctions = "" 'arr
End Function
