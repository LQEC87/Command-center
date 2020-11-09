Option Explicit
Dim a,i,j,x,y,ext,T(1),Cnt,timer(1),fso,xyfile,Shell,sas,wsn,UName,objMain,objFolder,objFile,objOPTS,objInPut,FileName,ourfile,MusicFile,VideoFile,MFP,FEx,FEy,FEz,xc,yd,ze,FE,ExFi,CFE,INPIN(2)
Cnt = "Command Center [Built Version 1.0.0200]" & vbCrLf & "Copyright (c) 2020 Microsoft Corporation.  All rights reserved."
ext = 0
msgbox"Please tell your password."
For i = 1 to 5
  x = inputbox("put your pass" & vbCrLf & 6-i &" left.","please input")
  If x="Xword" Then
    Exit for
  End If
Next
If x="Xword" Then
  msgbox"Welcome to [command center]"
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set Shell = CreateObject("WScript.Shell")
  xyfile = fso.getParentFolderName(WScript.ScriptFullName)
  Do Until ext = 1
    y = inputbox("What do you want to do?" & vbCrLf & "You can use only { Lowercase }.","Command center")
    Select Case y
      case "exit"
        msgbox "please click {OK}",,"Command center"
        WScript.Quit
      case "memo"
        Shell.Run "rundll32.exe url.dll,FileProtocolHandler " & xyfile & "\A\gggggg.txt"
      case "time"
        msgbox "We are in" & vbCrLf & Now
      case "music"
        xc = false
        yd = false
        ze = false
        FE = false
        Set fso = CreateObject("Scripting.FileSystemObject")
        FEx = fso.FolderExists("c:\Users")
        If Not(FEx) Then
          FEy = fso.FolderExists("d:\Users")
          If Not(FEy) Then
            FEz = fso.FolderExists("e:\Users")
            If Not(FEz) Then
              ExFi = inputbox("Please input your main system DRIVE." & vbCrLf & "ex)  C Drive -> C")
              FE = true
            Else
              ze = true
            End If
          Else
            yd = true
          End If
        Else
        xc = true
        End If

        Set wsn = WScript.CreateObject("WScript.Network")
        UName = wsn.UserName
        ourfile = fso.getParentFolderName(WScript.ScriptFullName)
        MusicFile = "" & ourfile & "\A\System.MusicList.txt"
        Set objOPTS = fso.OpenTextFile(MusicFile,2,True)

        If xc Then
          CFE = "C"
          FileName = "C:\Users\" & UName & "\Music"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next 
        Else If yd Then
          CFE = "D"
          FileName = "D:\Users\" & UName & "\Music"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next
        Else If ze Then
          CFE = "E"
          FileName = "E:\Users\" & UName & "\Music"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next
        Else If FE Then
          CFE = "" & FE
          FileName = FE & ":\Users\" & UName & "\Music"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next
        End If End If
        End If End If

        objOPTS.Close

        Set objInPut = fso.OpenTextFile(MusicFile,1,False)
        INPIN(1) = ""
        Do While objInPut.AtEndOfStream<>True
          INPIN(0) = objInPut.ReadLine
          INPIN(1) = INPIN(1) & "" & INPIN(0) & vbCrLf & ""
        Loop
        objInPut.Close

        msgbox INPIN(1)
        INPIN(2) = ""
        Do While INPIN(2)=""
        INPIN(2) = inputbox( "Please Choose From The Files Below." & vbCrLf & INPIN(1) & vbCrLf & "If you want to exit, Please input {exit}or{Nothing}" )
        Loop
        MFP = "" & CFE & ":\Users\" & UName & "\Music\" & INPIN(2)

        FOR i = 1 to 1

        If INPIN(2)="exit" OR INPIN(2)="Nothing" Then
        EXIT FOR
        Else If fso.FolderExists(MFP) Then

        Set Shell = CreateObject("WScript.Shell")
        Shell.Run "rundll32.exe url.dll,FileProtocolHandler " & MFP

        Else
        Randomize
        WScript.Sleep( 1000*Int((3-1+1)*Rnd+1) )
        msgbox "An error has occurred."
        WScript.Quit
        End If End If

        NEXT
      case "movie"
        xc = false
        yd = false
        ze = false
        FE = false
        Set fso = CreateObject("Scripting.FileSystemObject")
        FEx = fso.FolderExists("c:\Users")
        If Not(FEx) Then
          FEy = fso.FolderExists("d:\Users")
          If Not(FEy) Then
            FEz = fso.FolderExists("e:\Users")
            If Not(FEz) Then
              ExFi = inputbox("Please input your main system DRIVE." & vbCrLf & "ex)  C Drive -> C")
              FE = true
            Else
              ze = true
            End If
          Else
            yd = true
          End If
        Else
        xc = true
        End If

        Set wsn = WScript.CreateObject("WScript.Network")
        UName = wsn.UserName
        ourfile = fso.getParentFolderName(WScript.ScriptFullName)
        VideoFile = "" & ourfile & "\A\System.VideoList.txt"
        Set objOPTS = fso.OpenTextFile(VideoFile,2,True)

        If xc Then
          CFE = "C"
          FileName = "C:\Users\" & UName & "\Videos"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next 
        Else If yd Then
          CFE = "D"
          FileName = "D:\Users\" & UName & "\Videos"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next
        Else If ze Then
          CFE = "E"
          FileName = "E:\Users\" & UName & "\Videos"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next
        Else If FE Then
          CFE = "" & FE
          FileName = FE & ":\Users\" & UName & "\Videos"
          Set objMain = fso.GetFolder(FileName)
          For Each objFile In objMain.Files
            objOPTS.WriteLine objFile.Name
          Next
          For Each objFolder In objMain.SubFolders
            objOPTS.WriteLine objFolder.Name
          Next
        End If End If
        End If End If

        objOPTS.Close

        Set objInPut = fso.OpenTextFile(VideoFile,1,False)
        INPIN(1) = ""
        Do While objInPut.AtEndOfStream<>True
          INPIN(0) = objInPut.ReadLine
          INPIN(1) = INPIN(1) & "" & INPIN(0) & vbCrLf & ""
        Loop
        objInPut.Close

        msgbox INPIN(1)
        INPIN(2) = ""
        Do While INPIN(2)=""
        INPIN(2) = inputbox( "Please Choose From The Files Below." & vbCrLf & INPIN(1) & vbCrLf & "If you want to exit, Please input {exit}or{Nothing}" )
        Loop
        MFP = "" & CFE & ":\Users\" & UName & "\Videos\" & INPIN(2)

        FOR i = 1 to 1

        If INPIN(2)="exit" OR INPIN(2)="Nothing" Then
        EXIT FOR
        Else If fso.FolderExists(MFP) Then

        Set Shell = CreateObject("WScript.Shell")
        Shell.Run "rundll32.exe url.dll,FileProtocolHandler " & MFP

        Else
        Randomize
        WScript.Sleep( 1000*Int((3-1+1)*Rnd+1) )
        msgbox "An error has occurred."
        WScript.Quit
        End If End If

        NEXT
      case "calculation"
        On Error Resume Next
        sas = true
        Do While sas
        For i = 0 to 2 step 0
        Err.Clear
        a = inputbox("You can use only [+]or[-]or[*]or[/] or number." & vbCrLf & _
        "If you want to stop this calculator, Please input {exit}","calculation")
        If a="" Then
          msgbox "Please input something."
        Else If a="exit" Then
          msgbox "Canceled."
          Exit Do
        Else
          Exit For
        End If End If
        Next
        msgbox "Answer: " & Eval(a)
        If Err.Number<>0 Then
          msgbox "Your formula cannot be calculated."
        Else
          Exit Do
        End If
        Loop
        On Error Goto 0
      case "help"
        msgbox"Command List",,"Command center"
        msgbox"alarm" & vbCrLf & _
              "calculation" & vbCrLf & _
              "command" & vbCrLf & _
              "exit" & vbCrLf & _
              "help" & vbCrLf & _
              "memo" & vbCrLf & _
              "movie" & vbCrLf & _
              "music" & vbCrLf & _
              "stopwatch" & vbCrLf & _
              "time" & vbCrLf & _
              "timer",,"Command center"
      case "stopwatch"
        msgbox"If you click {OK}, StopWatch starts.",,"Command_center Stopwatch"
        T(0) = ( Hour(Time) * 60 + Minute(Time) ) * 60 + Second(Time)
        msgbox"STOP!",,"Command center"
        T(1) = ( Hour(Time) * 60 + Minute(Time) ) * 60 + Second(Time)
        msgbox T(1) - T(0) & " second.",,"Command_center Stopwatch"
      case "timer"
        timer(0) = inputbox("What time do I wait?" & vbCrLf & "You can use only { Number }, and can set millisecond.","Timer")
        timer(1) = FormatNumber( timer(0) )
        msgbox"If you click {OK}, Timer starts."
        WScript.Sleep ( timer(1) )
        msgbox"It's Time!",48
      case "alarm"
        msgbox"#CAUTION!#" & vbCrLf & "When alarm is on, You can't check alarm on or off."
        Shell.Run "rundll32.exe url.dll,FileProtocolHandler " & xyfile & "\A\Alarm.vbs"
      case "command"
        msgbox Cnt
      case ""
        ext = 0
      case else
        msgbox Chr(34) & y & Chr(34) & " is not recognized as an internal or external command," & vbCrLf & _
               " operable program or batch file.",48,"Command center"
    End Select
  Loop
Else If Not(x="Xword") Then
  msgbox "can't login!",16
  WScript.Quit
End If End If