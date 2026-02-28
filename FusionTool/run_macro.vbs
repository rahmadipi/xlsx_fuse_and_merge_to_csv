On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim excelPath, macroName, outputPath
excelPath  = WScript.Arguments.Item(0)
macroName  = WScript.Arguments.Item(1)
namaHasil = WScript.Arguments.Item(2)

WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Echo "Mulai menjalankan skrip VBScript..."
WScript.Echo ""
WScript.Echo "Nama File Excel: " & objFSO.GetFileName(excelPath)
WScript.Echo "Nama Makro: " & macroName
WScript.Echo ""

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(excelPath, , True)

If objWorkbook Is Nothing Then
    WScript.Echo "Error: File tidak ditemukan atau tidak bisa dibuka."
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    WScript.Quit 1
End If

objExcel.Application.Run "'" & objWorkbook.Name & "'!" & macroName, namaHasil

objWorkbook.Close False
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "Success: Proses konversi selesai."
WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Quit 0