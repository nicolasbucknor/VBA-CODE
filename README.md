# VBA-CODE

Option Explicit

Function LoadArray()

Dim wb As Workbook
Dim ws As Worksheet
Dim InputValue As String
Dim States() As String

ClearImmediateWindow

On Error GoTo ErrHandler

Set wb = ThisWorkbook
Set ws = wb.Worksheets("States")

InputValue = ws.Cells(1, 1).Value

States = Split(InputValue, ",")

Debug.Print UBound(States)

    Call removeduplicatevalues(States)
    
    Debug.Print UBound(States)

    Exit Function
    
ErrHandler:

    Debug.Print "Error nr: " & Err.Number & " Error Description: " & Err.Description
    
End Function


Function removeduplicatevalues(ByRef tmpvalues() As String)

On Error GoTo ErrHandler

Dim i As Long, n As Long

Dim temparr() As String: ReDim temparr(UBound(tmpvalues))
Dim d As New dictionary



    For i = 0 To UBound(tmpvalues)
        
        If Not d.exists(tmpvalues(i)) Then
        
            d.Add tmpvalues(i), tmpvalues(i)
            temparr(n) = tmpvalues(i)
            n = n + 1
            
        End If
        
    Next
    
    Debug.Print UBound(temparr)
    ReDim Preserve temparr(n - 1)
    tmpvalues = temparr
        
    Exit Function
    
ErrHandler:

    Debug.Print "Error nr: " & Err.Number & " Error Description: " & Err.Description
    
End Function

