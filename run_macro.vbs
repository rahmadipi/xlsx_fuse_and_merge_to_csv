On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")

WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Echo "Mulai menjalankan skrip VBScript..."
WScript.Echo ""
WScript.Echo "Nama File Excel: " & objFSO.GetFileName(WScript.Arguments.Item(0))
WScript.Echo "Nama Makro: " & WScript.Arguments.Item(1)
WScript.Echo ""

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(WScript.Arguments.Item(0), , True)

If objWorkbook Is Nothing Then
    WScript.Echo "Error: File tidak ditemukan atau tidak bisa dibuka."
    WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
    WScript.Quit 1
End If

objExcel.Application.Run "'" & objWorkbook.Name & "'!" & WScript.Arguments.Item(1)

objWorkbook.Close False
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "Success: Proses konversi selesai."
WScript.Echo "+++++++++++++++++++++++++++++++++++++++++++++++++++"
WScript.Quit 0