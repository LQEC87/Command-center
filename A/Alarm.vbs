Dim m,a(4),x,y,z,qrfile,vwat,h,Uyck
msgbox"Please set an alarm."
For x = 0 to 4
  a(x) = 0
Next
msgbox "What time do you set the alarm?"

a(0) = Hour(Time)
a(1) = Minute(Time)

For i = 0 to 2 step 0
  y = msgbox("Now set hour:" & a(0) & vbCrLf & "May I do inclease? " & _
"If you want to do decrease, Please click the "  & Chr(34) & "No" & Chr(34)_
 & vbCrLf & "STOP -> [Cancel]",3,"Alarm")
  Select case y
    Case 6
      If a(0)=23 Then
        a(0) = 0
      Else
        a(0) = 1 + a(0)
      End If
    Case 7
      If a(0)=0 Then
        a(0) = 23
      Else
        a(0) = -1 + a(0)
      End If
    Case 2
      Exit For
  End Select
Next

If a(0)<=0 Then
  a(0) = 0
Else If a(0)>=24 Then
  a(0) = 0
End If

For j = 0 to 2 step 0
  z = msgbox("Now set minute:" & a(1) & vbCrLf & "May I inclease it? " & _
"If you want to decrease, Please click the "  & Chr(34) & "No" & Chr(34)_
 & vbCrLf & "STOP -> [Cancel]",3,"Alarm")
  Select case z
    Case 6
      If a(1)=59 Then
        a(1) = 0
      Else
        a(1) = 1 + a(1)
      End If
    Case 7
      If a(1)=0 Then
        a(1) = 59
      Else
        a(1) = -1 + a(1)
      End If
    Case 2
      Exit For
  End Select
Next

If a(1)<=0 Then
  a(1) = 0
Else If a(1)>=60 Then
  a(1) = 0
End If End If End If

Uyck = msgbox ("Your set time ... " & a(0) & ":" & a(1),1)

If Uyck<>1 Then
  msgbox "We stopped the alarm."
  WScript.Quit
End If

a(3) = Hour(Now) * 60 + Minute(Now)

a(4) = a(0) * 60 + a(1)
m = 0
m = a(4) - a(3)

If m<=0 Then
  msgbox "⚠NO⚠"
  WScript.Quit
End If

msgbox m & " minute."

h = 1000 * 60 * m

WScript.Sleep( h )

Dim fso,Shell
Set fso = CreateObject("Scripting.FileSystemObject")
qrfile = fso.getParentFolderName(WScript.ScriptFullName)
Set Shell = CreateObject("WScript.Shell")

Do
  Shell.Run "rundll32.exe url.dll,FileProtocolHandler " & qrfile & "\Timer.vbs"
  vwat = 1 + vwat
  WScript.Sleep(100)
Loop Until vwat=3

WScript.Sleep(1000)