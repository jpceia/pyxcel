Attribute VB_Name = "Utils"
Option Explicit

Function ConcatRows( _
    range1 As Range, _
    range2 As Range) As Variant()
    
    Dim nCols, nRows, nRows1, nRows2 As Long
    nCols = range1.Columns.Count
    nRows1 = range1.Rows.Count
    nRows2 = range2.Rows.Count
    nRows = nRows1 + nRows2
    
    If nCols <> range2.Columns.Count Then
        Err.Raise 0, "", "Number of columns does not match"
    End If
    
    Dim i, j As Long
    Dim arr1, arr2, res As Variant
    arr1 = range1
    arr2 = range2
    ReDim res(1 To nRows, 1 To nCols)

    For i = 1 To nRows1
        For j = 1 To nCols
            res(i, j) = arr1(i, j)
        Next j
    Next i

    For i = 1 To nRows2
        For j = 1 To nCols
            res(i + nRows1, j) = arr2(i, j)
        Next j
    Next i
    
    ConcatRows = res
End Function

