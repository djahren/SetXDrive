Set args = Wscript.Arguments
Set objNetwork = WScript.CreateObject("WScript.Network") 

strPrefix = ""
strUserDrive = "X:"

If GetPath(strUserDrive) = "" Then
  strInputPath = args(0)
  'Wscript.Echo strInputPath
  strDriveLetter = Left(strInputPath, 2)
  'Wscript.Echo strDriveLetter
  If strDriveLetter <> "\\" Then
	strPrefix = GetPath(strDriveLetter)
	If strPrefix = "" Then
		'Local Drive
		Set objShell = CreateObject("WScript.Shell")
		strPrefix = "\\" & objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & "\"  & Left(strDriveLetter, 1) & "$" 
	End If
	
	'Append the rest of the location on
	strPath = Replace(strInputPath,strDriveLetter & "\",strPrefix & "\",1,1)
	
  Else
	strPath = args(0)
  End If
  WScript.Echo "Setting " + strUserDrive + " drive to: " + strPath
  objNetwork.MapNetworkDrive strUserDrive , strPath, True
Else
  'Set objNetwork = Nothing
	Wscript.Echo "Please right-click drive " + strUserDrive + " in My Computer and choose Disconnect before running this command again."
End If

Function GetPath(strDrive)
   Set objDrives = objNetwork.EnumNetworkDrives
   For i = 0 to objDrives.Count - 1 Step 2
      strNetDrive = objDrives.Item(i)
      strNetPath = objDrives.Item(i+1)
	  'Wscript.Echo strNetDrive
      If UCase(strNetDrive) = UCase(strDrive) Then
          GetPath = strNetPath
          Exit For
      End If
   Next
End Function

Set objNetwork = Nothing