Option Explicit

Main 

Function Update(host, api_key)
	Dim url
	url = "https://zonomi.com/app/dns/dyndns.jsp?host=<HOST>&api_key=<API_KEY>"

	Dim xhr
	set xhr = createobject("microsoft.xmlhttp")

	url = replace(url, "<HOST>", host)
	url = replace(url, "<API_KEY>", api_key)

	wscript.echo "Updating host " & host

	xhr.open "get",url,false
	xhr.setrequestheader "Pragma","no-cache"
	xhr.setrequestheader "Cache-control","no-cache"

	On Error Resume Next
	xhr.send
	if Err.Number<> 0 Then
		Update = false
		Exit Function
	End If
	On Error Goto 0

	wscript.echo "Got response: " & xhr.Status

	If (xhr.status = 200) then
		Update = true
	Else
		Update = false
	End If

End Function

Sub ShowUsage()
    wscript.echo ""
    wscript.echo "Usage: cscript zonomiUpdater.vbs -h<hostname> -k<api-key>"
    wscript.echo ""
    wscript.echo "Options:"
    wscript.echo "    -h<hostname>	The hostname to update"
    wscript.echo "    -k<api-key>	Your Zonomi API key"
    wscript.echo ""

End Sub

' Parse command line options
Function HandleCommandLineParameters(host, api_key)
    Dim objArgs, i, n

    n = 0
    Set objArgs = WScript.Arguments
    For i = 0 to objArgs.Count - 1
        Select case Mid(objArgs(i),2,1)
            case "h"
                n=n+1
                host=Mid(objArgs(i),3)
            case "k"
                n=n+1
                api_key=Mid(objArgs(i),3)
        End Select
    Next
    If(n < 2) Then
        ShowUsage()
        HandleCommandLineParameters = False
        Exit Function
    End If
    HandleCommandLineParameters = True
End Function

Sub Main
	Dim host, api_key

	wscript.echo "Zonomi.com DNS Updater"

	If Not HandleCommandLineParameters(host, api_key) Then
		wscript.quit
	End If

	If Update(host, api_key) Then
		wscript.echo "Successfully updated host " & host
	Else
		wscript.echo "Failed to update host " & host
	End If
End Sub
