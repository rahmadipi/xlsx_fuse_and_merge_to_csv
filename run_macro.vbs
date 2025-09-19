On Error Resume Next
WScript.Echo "Mulai menjalankan skrip VBScript..."
WScript.Echo "Argumen 1 (Nama File Excel): " & WScript.Arguments.Item(0)
WScript.Echo "Argumen 2 (Nama Makro): " & WScript.Arguments.Item(1)
WScript.Echo "----------------------------------------------------"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

Set objWorkbook = objExcel.Workbooks.Open(WScript.Arguments.Item(0))

If objWorkbook Is Nothing Then
    WScript.Echo "Error: File tidak ditemukan atau tidak dapat dibuka."
    WScript.Quit
End If

objExcel.Application.Run "'" & objWorkbook.Name & "'!" & WScript.Arguments.Item(1)

objWorkbook.Close
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Echo "Selesai menjalankan skrip VBScript."