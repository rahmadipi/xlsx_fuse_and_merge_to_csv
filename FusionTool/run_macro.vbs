Set objFSO = CreateObject("Scripting.FileSystemObject")

WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Echo "Mulai menjalankan skrip VBScript..."
WScript.Echo ""

If WScript.Arguments.Count < 3 Then
    WScript.Echo "Error: Argumen yg dikirimkan kurang!"
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    WScript.Quit 1
End If

Dim excelPath, macroName, namaHasil
excelPath  = objFSO.GetAbsolutePathName(WScript.Arguments.Item(0))
macroName  = WScript.Arguments.Item(1)
namaHasil = WScript.Arguments.Item(2)

WScript.Echo "Nama File Excel: " & objFSO.GetFileName(excelPath)
WScript.Echo "Nama Makro: " & macroName
WScript.Echo ""

If Not objFSO.FileExists(excelPath) Then
    WScript.Echo "Error: File Excel tidak ditemukan di: "
    WScript.Echo "" & excelPath
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    WScript.Quit 1
End If

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(excelPath, 0, False)

If objWorkbook Is Nothing Then
    WScript.Echo "Error: Gagal membuka Workbook."
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    objExcel.Quit
    WScript.Quit 1
End If

On Error Resume Next
objExcel.Run macroName, namaHasil

If Err.Number <> 0 Then
    WScript.Echo "Error saat menjalankan Makro: " & Err.Description
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit 1
End If
On Error GoTo 0

objWorkbook.Close False
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "Success: Proses konversi selesai."
WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Quit 0