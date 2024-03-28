Dim objShell
Dim godmode
Dim output
	godmode = False
	output = 0
Set objShell = CreateObject("WScript.Shell")
Dim net_sesh
	net_sesh = "net session"
	output = objShell.Run(net_sesh, 0, True) 
	If output = 0 Then
		godmode = True
	End If