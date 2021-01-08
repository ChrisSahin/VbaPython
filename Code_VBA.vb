Sub TestPythonFromVBA_Array()

Dim myArray
Dim returnValue
Dim myArrayStr As String
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1


myArray = Array(5, 2, 4, 6, -7)
myArrayStr = Join(myArray, ",")


returnValue = wsh.Run("C:\Users\CHRIS\Anaconda3\python.exe C:\Users\CHRIS\Documents\TestPythonVBA\SumValue.py " + myArrayStr, windowStyle, waitOnReturn)

Debug.Print (returnValue)
End Sub

Sub TestPythonFromVBA_Range()


Dim returnValue
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim FileAdress As String

FileAdress = Application.ActiveWorkbook.FullName
wsh.Run "C:\Users\CHRIS\Anaconda3\python.exe C:\Users\CHRIS\Documents\TestPythonVBA\SumValue_bis.py " + FileAdress, windowStyle, waitOnReturn

End Sub
