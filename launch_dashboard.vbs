' Placement Offers - silent launcher for the desktop shortcut.
' If the dashboard is already running, just open it. Otherwise start the
' server hidden (no console window) and then open the browser.
Option Explicit
Dim fso, sh, http, running, scriptDir, appPath, url
url = "http://127.0.0.1:5005/"
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appPath = scriptDir & "\app.py"

' --- Is the server already up? ---
running = False
On Error Resume Next
Set http = CreateObject("MSXML2.XMLHTTP")
http.open "GET", url, False
http.send
If Err.Number = 0 And http.status = 200 Then running = True
On Error GoTo 0

If running Then
    ' Already up: trigger a fresh rescan, then open the dashboard.
    On Error Resume Next
    Dim scanReq
    Set scanReq = CreateObject("MSXML2.XMLHTTP")
    scanReq.open "POST", "http://127.0.0.1:5005/api/scan?mode=new", False
    scanReq.send
    On Error GoTo 0
    sh.Run url, 1, False
Else
    ' Start hidden (window style 0). Try pythonw first, fall back to python.
    On Error Resume Next
    sh.Run "pythonw.exe """ & appPath & """ --no-open", 0, False
    If Err.Number <> 0 Then
        Err.Clear
        sh.Run "python.exe """ & appPath & """ --no-open", 0, False
    End If
    On Error GoTo 0
    ' Give Flask a moment to bind, then open the page.
    WScript.Sleep 3500
    sh.Run url, 1, False
End If
