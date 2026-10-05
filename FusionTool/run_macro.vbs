Option Explicit

Dim objFSO, objExcel, objWorkbook
Dim excelPath, macroName, namaHasil
Dim errorNumber, errorDescription

WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Echo "Mulai menjalankan skrip VBScript..."
WScript.Echo ""

If WScript.Arguments.Count < 3 Then
    Fail "validasi argumen", 0, "Diperlukan 3 argumen: path workbook, nama makro, dan nama hasil."
End If

macroName = WScript.Arguments.Item(1)
namaHasil = WScript.Arguments.Item(2)

On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    Fail "membuat FileSystemObject", errorNumber, errorDescription
End If

excelPath = objFSO.GetAbsolutePathName(WScript.Arguments.Item(0))
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    Fail "membaca path workbook", errorNumber, errorDescription
End If

If Not objFSO.FileExists(excelPath) Then
    On Error GoTo 0
    Fail "memeriksa workbook", 0, "File Excel tidak ditemukan: " & excelPath
End If
On Error GoTo 0

WScript.Echo "Nama File Excel: " & objFSO.GetFileName(excelPath)
WScript.Echo "Nama Makro: " & macroName
WScript.Echo ""

On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    Fail "membuka Microsoft Excel", errorNumber, errorDescription
End If

objExcel.Visible = False
If Err.Number = 0 Then objExcel.DisplayAlerts = False
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    Fail "mengatur opsi Microsoft Excel", errorNumber, errorDescription
End If

Set objWorkbook = objExcel.Workbooks.Open(excelPath, 0, True)
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    Fail "membuka workbook", errorNumber, errorDescription
End If
If objWorkbook Is Nothing Then
    On Error GoTo 0
    Fail "membuka workbook", 0, "Excel tidak mengembalikan workbook."
End If

Err.Clear
objExcel.Run macroName, namaHasil
If Err.Number <> 0 Then
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    Fail "menjalankan makro '" & macroName & "'", errorNumber, errorDescription
End If
On Error GoTo 0

Cleanup
WScript.Echo "Success: Proses konversi selesai."
WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Quit 0

Sub Fail(stage, number, description)
    WScript.Echo "Error saat " & stage & "."
    If number <> 0 Then
        WScript.Echo "Nomor error: " & number
    End If
    WScript.Echo "Detail: " & description
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    Cleanup
    WScript.Quit 1
End Sub

Sub Cleanup
    On Error Resume Next
    If Not objWorkbook Is Nothing Then objWorkbook.Close False
    If Not objExcel Is Nothing Then objExcel.Quit
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
End Sub