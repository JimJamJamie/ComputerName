'Header section
Option Explicit
On Error Resume Next

'Reference section
Dim objShell
Dim wmiQuery
Dim queryItem

'Worker section
Set objShell = CreateObject("WScript.Shell")
Set wmiQuery = GetObject("winmgmts:").ExecQuery _
	("select Name from Win32_ComputerSystem")

'Output section
For Each queryItem in wmiQuery
	WScript.Echo "This computer is called " & queryItem.Name
Next
