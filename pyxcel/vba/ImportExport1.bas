Attribute VB_Name = "ImportExport1"
Option Explicit

Public Sub ImportThis()
    ImportModules ThisWorkbook
End Sub

Public Sub ExportThis()
    ExportModules ThisWorkbook
End Sub

Public Sub ExportModules(wkb As Workbook)
    Dim Path As String
    Path = wkb.Path
    
    ' Confirm if the folder exists. If not, creates it
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.FolderExists(Path) = False Then
        On Error Resume Next
        MkDir Path
        On Error GoTo 0
    End If
    
    Dim bExport As Boolean
    Dim sFileName As String
    Dim cmp As VBIDE.VBComponent
    
    For Each cmp In wkb.VBProject.VBComponents
        bExport = True
        ''' Concatenate the correct filename for export.
        Select Case cmp.Type
            Case vbext_ct_ClassModule
                sFileName = cmp.Name & ".cls"
            Case vbext_ct_MSForm
                sFileName = cmp.Name & ".frm"
            Case vbext_ct_StdModule
                sFileName = cmp.Name & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        If bExport Then
            ''' Export the component to a text file.
            cmp.Export Path & "\" & sFileName
        End If
    Next cmp

    MsgBox "Export is ready"
End Sub



Public Sub ImportModules(wkb As Workbook)
    Dim Path As String
    Path = wkb.Path

    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim sFileName, ext As String

        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(Path).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModules(wkb)
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(Path).Files
        sFileName = objFile.Name
        ext = objFSO.GetExtensionName(sFileName)
        If (ext = "cls") Or (ext = "frm") Or (ext = "bas") Then
            wkb.VBProject.VBComponents.Import objFile.Path
        End If
    Next objFile
    
    MsgBox "Import is ready"
End Sub


Function DeleteVBAModules(wkb As Workbook)
    Dim cmp As VBIDE.VBComponent
    Dim components As VBIDE.VBComponents
    Set components = wkb.VBProject.VBComponents
    For Each cmp In components
        If cmp.Type <> vbext_ct_Document Then
            components.Remove cmp
        End If
    Next cmp
End Function
