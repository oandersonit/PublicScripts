On Error Resume Next

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject ("Scripting.FileSystemObject")
strScriptPath = objFSO.GetParentFolderName(Wscript.ScriptFullName)
Set WshShell = CreateObject("Wscript.Shell")
Set WshEnv = WshShell.Environment("Process")
strComputer = WshEnv("COMPUTERNAME")
strDay = Day(Date)
If len(strDay) = 1 Then
   strDay = "0" & strDay
End If
strMonth = Month(Date)
If len(strMonth) = 1 Then
   strMonth = "0" & strMonth
End If
strYear = Year(Date)
strHour = Hour(Time)
If len(strHour) = 1 Then
   strHour = "0" & strHour
End If
strMinute = Minute(Time)
If len(strMinute) = 1 Then
   strMinute = "0" & strMinute
End If
strOUTFile = strComputer & "_" & strDay & "-" & strMonth & "-" & strYear & "--" & strHour & "-" & strMinute & "_CHK_Windows.HTML"

Set objOUTFile = objFSO.CreateTextFile (strOUTFile, ForWriting)

objOUTFile.WriteLine "<!DOCTYPE HTML PUBLIC " & """" & "-//W3C//DTD HTML 4.01 Transitional//EN" & """" & ">"
objOUTFile.WriteLine "<html>"
objOUTFile.WriteLine "<head>"
objOUTFile.WriteLine "  <title>CHECKLIST WINDOWS</title>"
objOUTFile.WriteLine "  <meta http-equiv=" & """" & "Content-Type" & """" & "content=" & """" & "text/html; charset=iso-8859-1" & """" & ">"
objOUTFile.WriteLine "  <STYLE>A {text-decoration: none;color: black; }"
objOUTFile.WriteLine "  </STYLE>"
objOUTFile.WriteLine "</head>"
objOUTFile.WriteLine ""
objOUTFile.WriteLine "<Body>"
objOUTFile.WriteLine "  <tr valign=" & """" & "top" & """" & ">"
objOUTFile.WriteLine "    <td> <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "    <td width=" & """" & "100%" & """" & " valign=" & """" & "top" & """" & " bgcolor=" & """" & "#000066" & """" & "> <div align=" & """" & "center" & """" & "><font color=" & """" & "#FFFFFF" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><strong>Validações Básicas do S.O.</strong></font></div></td>"
objOUTFile.WriteLine "  </tr>"
objOUTFile.WriteLine "  </Table>"

objOUTFile.WriteLine "    <tr>"
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">HOSTNAME</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><div align=" & """" & "center" & """" & ">" & strComputer & "</div></td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

Set objWMIService = GetObject("winmgmts://" & strComputer & "/root/cimv2")
Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS In colOSes
    strSO = objOS.Caption
    strSP = objOS.ServicePackMajorVersion
Next

objOUTFile.WriteLine "    <tr>"
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">VERSÃO DO SISTEMA OPERACIONAL</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><div align=" & """" & "center" & """" & ">" & strSO & " - Service Pack " & strSP & "</div></td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

Set objOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem", , 48)
For Each objI in objOS
    Dim sY, sM, sD, sH, sN, sS, dU, sT, iD, iH, iM, iSS
    sY = Left(objI.LastBootUpTime,4)
    sM = Mid(objI.LastBootUpTime, 5, 2)
    sD = Mid(objI.LastBootUpTime, 7, 2)
    sH = Mid(objI.LastBootUpTime, 9, 2)
    sN = Mid(objI.LastBootUpTime,11, 2)
    sS = Mid(objI.LastBootUpTime,13, 2)
    dU = DateSerial(sY, sM, sD) + TimeSerial(sH, sN, sS)
    iM = DateDiff("n", dU, Now)
    iD = iM \ 1440
    iM = iM - (iD * 1440)
    iH = iM \ 60
    iM = iM - (iH * 60)
Next
strUpTime = iD & " dias, " & iH & " horas e " & iM & " minutos"

objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">UPTIME</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><div align=" & """" & "center" & """" & ">" & strUpTime & "</div></td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">DATA/HORA DE EXECUÇÃO</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><div align=" & """" & "center" & """" & ">" & Now & "</div></td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">SERVIÇOS AUTOMÁTICOS</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Serviço</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Status</b></font></td>"
objOUTFile.WriteLine "            </tr>"
Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where StartMode='Auto'")
For Each objService In colListOfServices
    If objService.State = "Running" Then
       strColor = "#000000"
    Else
       strColor = "#FF0000"
    End if
    objOUTFile.WriteLine "            <tr>"
    objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & objService.Displayname & "</font></td>"
    objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & objService.State & "</font></td>"
    objOUTFile.WriteLine "            </tr>"
Next
objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "        </td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">DISCOS</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
objOUTFile.WriteLine "              <td width=" & """" & "40%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Disco</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Tamanho</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Utilizado</b></font></td>"
objOUTFile.WriteLine "            </tr>"
Set colItems = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk Where DriveType='3'")
For Each objItem in colItems
    strDiskDrive = objItem.Name
    strDrives = strDrives & strDiskDrive & ";"
    strDiskSize = Round(objItem.Size / 1024 / 1024 / 1024, 1)
    strDiskFree = Round(objItem.FreeSpace / 1024 / 1024 / 1024, 1)
    strDiskUsed = strDiskSize - strDiskFree
    strDiskUsedPCT = Round((strDiskUsed / strDiskSize)*100, 1)
    If strDiskUsedPCT < 90 Then
       strColor = "#000000"
    Else
       strColor = "#FF0000"
    End if
    objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
    objOUTFile.WriteLine "              <td width=" & """" & "40%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strDiskDrive & "</font></td>"
    objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strDiskSize & " GB</font></td>"
    objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strDiskUsed & " GB</font></td>"
    objOUTFile.WriteLine "            </tr>"
Next
strDrives = Left(strDrives, len(strDrives) - 1)
objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "        </td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

objOUTFile.WriteLine "  <tr valign=" & """" & "top" & """" & ">"
objOUTFile.WriteLine "    <td> <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "    <td width=" & """" & "100%" & """" & " valign=" & """" & "top" & """" & " bgcolor=" & """" & "#000066" & """" & "> <div align=" & """" & "center" & """" & "><font color=" & """" & "#FFFFFF" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><strong>Monitoração</strong></font></div></td>"
objOUTFile.WriteLine "  </tr>"
objOUTFile.WriteLine "  </Table>"
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">AGENTE OPEN VIEW</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "0" & """" & " CellSpacing=" & """" & "0" & """" & ">"
Set oExec = WshShell.Exec("cmd /c ovc -status")
strOut = oExec.StdOut.ReadAll
If instr(1, strOut, "OV Control") > 0 Then
   strOut = Split(strOut, Chr(10), -1)
   For Each Lin in strOut
       If instr(1, Lin, "Running") > 0 Then
          strColor = "#000000"
       Else
          strColor = "#FF0000"
       End if
       str2 = split(Lin, " ", -1)
       For Each str3 in str2
           str4 = str4 + 1
       Next
       strStatus = Split(Lin, " ")(str4 - 1)
       str5 = 0
       If str4 <> 0 Then
          Do Until str5 = str4 - 1
             strMSG = strMSG & Split(Lin, " ")(str5) & " "
             str5 = str5 + 1
          Loop
          objOUTFile.WriteLine "            <tr>"
          objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "Left" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strMSG & "</font></td>"
          objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "Left" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strStatus & "</font></td>"
          objOUTFile.WriteLine "            </tr>"
       End If
       strMSG = ""
       strStatus = ""
       str4 = 0
   Next
Else
   strColor = "#FF0000"
   objOUTFile.WriteLine "            <tr>"
   objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">Agente com problemas ou não instalado</font></td>"
   objOUTFile.WriteLine "            </tr>"
End If
objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

objOUTFile.WriteLine "  <tr valign=" & """" & "top" & """" & ">"
objOUTFile.WriteLine "    <td> <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "    <td width=" & """" & "100%" & """" & " valign=" & """" & "top" & """" & " bgcolor=" & """" & "#000066" & """" & "> <div align=" & """" & "center" & """" & "><font color=" & """" & "#FFFFFF" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><strong>Performance - Recursos Utilizados</strong></font></div></td>"
objOUTFile.WriteLine "  </tr>"
objOUTFile.WriteLine "  </Table>"
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
Set objInstance1 = objWMIService.Get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
N1 = objInstance1.PercentProcessorTime
D1 = objInstance1.TimeStamp_Sys100NS
WScript.Sleep(1000)
Set perf_instance2 = objWMIService.get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
N2 = perf_instance2.PercentProcessorTime
D2 = perf_instance2.TimeStamp_Sys100NS
PercentProcessorTime = Round((1 - ((N2 - N1)/(D2-D1)))*100, 1)
If Left(PercentProcessorTime, 1) = "-" Then
   PercentProcessorTime = 0
End If
If PercentProcessorTime < 90 Then
   strColor = "#000000"
Else
   strColor = "#FF0000"
End if
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">CPU</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "Center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & PercentProcessorTime & " %</font></td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine ""

Set objInstances = objWMIService.InstancesOf("Win32_PerfFormattedData_PerfOS_Memory",48)
For Each objInstance in objInstances
    strMemFree = objInstance.AvailableKBytes / 1024
    strNonPaged = Round(objInstance.PoolNonpagedBytes / 1024 / 1024, 1)
    strPaged = Round(objInstance.PoolPagedBytes / 1024 / 1024, 1)
Next
Set colComputer = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem") 
For Each objComputer in colComputer 
    strMemTotal = objComputer.TotalPhysicalMemory / 1024 / 1024
Next 
strMemUsed = Round(strMemTotal - strMemFree, 1)
strMemUsedPCT = Round((strMemUsed / strMemTotal)*100, 1)
If strMemUsedPCT < 90 Then
   strColor = "#000000"
Else
   strColor = "#FF0000"
End if
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">MEMÓRIA FÍSICA</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Tamanho Total</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Utilizada</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Pool Paged</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Pool Nonpaged</b></font></td>"
objOUTFile.WriteLine "            </tr>"
objOUTFile.WriteLine "            <tr>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & Round(strMemTotal) & " MB</font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & Round(strMemUsed, 1) & " MB</font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strPaged & " MB</font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strNonPaged & " MB</font></td>"
objOUTFile.WriteLine "            </tr>"
objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "        </td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"

Set colPageFiles = objWMIService.ExecQuery ("Select * from Win32_PageFileUsage")
For each objPageFile in colPageFiles
    strPGSize = strPGSize + objPageFile.AllocatedBaseSize
    strPGUsage = strPGUsage + objPageFile.CurrentUsage
Next
strPGUsagePCT = Round((strPGUsage / strPGSize)*100, 1)
If strPGUsagePCT < 90 Then
   strColor = "#000000"
Else
   strColor = "#FF0000"
End if
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">PAGE FILE</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
objOUTFile.WriteLine "              <td width=" & """" & "40%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Tamanho Total</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Utilizado</b></font></td>"
objOUTFile.WriteLine "            </tr>"
objOUTFile.WriteLine "            <tr>"
objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strPGSize & " MB</font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strPGUsage & " MB</font></td>"
objOUTFile.WriteLine "            </tr>"
objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "        </td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"

objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">I/O DOS DISCOS</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
objOUTFile.WriteLine "              <td width=" & """" & "40%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Disco</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "30%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>I/O</b></font></td>"
objOUTFile.WriteLine "            </tr>"
strDrives = Split(strDrives, ";", -1)
For Each strDrive in strDrives
    set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
    Set colItems = objRefresher.AddEnum (objWMIService, "Win32_PerfFormattedData_PerfDisk_PhysicalDisk").objectSet
    objRefresher.Refresh
    For i = 1 to 4
        For Each objItem in colItems
            If instr(1, objItem.Name, strDrive) > 0 Then
               If Not isnull(objItem.PercentDiskTime) Then
                  strTimeD = strTimeD + 1
                  strPCTTime = cint(strPCTTime) + cint(objItem.PercentDiskTime)
               End If
               If Not isnull(objItem.PercentIdleTime) Then
                  strIdleD = strIdleD + 1
                  strPCTIdle = cint(strPCTIdle) + cint(objItem.PercentIdleTime)
               End If
               Wscript.Sleep 2000
               objRefresher.Refresh
            End If
        Next
    Next
    strIO = Round((strPCTTime / strTimeD), 1)
    If strIO < 90 Then
       strColor = "#000000"
    Else
       strColor = "#FF0000"
    End if
    objOUTFile.WriteLine "            <tr>"
    objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strDrive & "</font></td>"
    objOUTFile.WriteLine "              <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=" & """" & strColor & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strIO & "</font></td>"
    objOUTFile.WriteLine "            </tr>"
Next
objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "        </td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"

Set strSortCPU = CreateObject("System.Collections.ArrayList") 
For Each objInstance1 in objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process Where Name <> 'Idle' And Name <> '_Total'")
    strProcName = objInstance1.Name
    N1 = objInstance1.PercentProcessorTime
    D1 = objInstance1.TimeStamp_Sys100NS
    WScript.Sleep(1000)
    For Each perf_instance2 in objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process where IDProcess = '" & objInstance1.IDProcess & "'")
        N2 = perf_instance2.PercentProcessorTime
        D2 = perf_instance2.TimeStamp_Sys100NS
        Exit For
    Next
    Nd = (N2 - N1)
    Dd = (D2 - D1)
    PercentProcessorTime = ((Nd / Dd)) * 100
    CPUUSage = Round(PercentProcessorTime ,0)
    If Len(CPUUSage) = 1 Then
       CPUUSage = "0" & CPUUSage
    End If
    If CPUUSage > 0 Then
       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where ProcessId ='" & objInstance1.IDProcess & "'")
       For Each objProcess In colProcesses
           Return = objProcess.GetOwner(strNameOfUser)
           If len(Round(objProcess.WorkingSetSize / 1024 / 1024)) = 1 Then
              strMem = "0" & Round(objProcess.WorkingSetSize / 1024 / 1024)
           Else
              strMem = Round(objProcess.WorkingSetSize / 1024 / 1024)
           End If
       Next
       strSortCPU.Add CPUUSage & "," & strProcName & "," & strNameOfUser & "," & strMem
    End If
    Nd = 0
    Dd = 0
    PercentProcessorTime = 0
    CPUUSage = 0
Next

objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">TOP 5 - CONSUMO DE RECURSOS</font></td>"
objOUTFile.WriteLine "        <td width=" & """" & "50%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & ">"
objOUTFile.WriteLine "          <table width=" & """" & "100%" & """" & " Border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "            <tr align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & ">"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Consumo CPU %</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Processo</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Owner</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Consumo de Memória (MB)</b></font></td>"
objOUTFile.WriteLine "            </tr>"
strSortCPU.Sort
strSortCPU.Reverse
strList = 0
For Each strProc in strSortCPU
    If strList <> 5 Then
       strProcs = strProcs & strProc & "|"
       strList = strList + 1
    Else
       Exit For
    End If
Next
strProcs = Split(Left(strProcs, len(strProcs) - 1), "|", - 1)
For Each strProc in strProcs
    strLinhas = Split(strProc, ",", -1)
    objOUTFile.WriteLine "            <tr>"
    For Each strLinha in strLinhas
        objOUTFile.WriteLine "              <td width=" & """" & "25%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strLinha & "</font></td>"
    Next
    objOUTFile.WriteLine "            </tr>"
Next

objOUTFile.WriteLine "          </table>"
objOUTFile.WriteLine "        </td>"
objOUTFile.WriteLine "      </tr>"
objOUTFile.WriteLine "    </table>"

objOUTFile.WriteLine "  <tr valign=" & """" & "top" & """" & ">"
objOUTFile.WriteLine "    <td> <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "    <td width=" & """" & "100%" & """" & " valign=" & """" & "top" & """" & " bgcolor=" & """" & "#000066" & """" & "> <div align=" & """" & "center" & """" & "><font color=" & """" & "#FFFFFF" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><strong>Event Viewer - System (Últimos 10 Eventos)</strong></font></div></td>"
objOUTFile.WriteLine "  </tr>"
objOUTFile.WriteLine "  </Table>"
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>EventID</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Date</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Source</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Type</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "60%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Message</b></font></td>"
objOUTFile.WriteLine "      </tr>"

strCount = 0
Set colLoggedEvents = objWMIService.ExecQuery ("Select * from Win32_NTLogEvent Where Logfile = 'System'")
For Each strEventSys in colLoggedEvents
    If strCount <> 11 Then
       If strEventSys.EventType = 1 Then
          strType = "Error"
       ElseIf strEventSys.EventType = 2 Then
          strType = "Warning"
       ElseIf strEventSys.EventType = 3 Then
          strType = "Information"
       End If
       strGenerated = Mid(strEventSys.TimeGenerated, 7, 2) & "/" & Mid(strEventSys.TimeGenerated, 5, 2) & "/" & Left(strEventSys.TimeGenerated, 4) & " " & Mid(strEventSys.TimeGenerated, 9, 2) & ":" & Mid(strEventSys.TimeGenerated, 11, 2) & ":" & Mid(strEventSys.TimeGenerated, 13, 2)
       If strType <> "Information" Then
          strCount = strCount + 1
          objOUTFile.WriteLine "      <tr>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strEventSys.EventCode & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strGenerated & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strEventSys.SourceName & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strType & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "60%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & Left(strEventSys.Message, 100) & "</font></td>"
          objOUTFile.WriteLine "      </tr>"
       End If
    Else
       Exit For
    End If
Next
objOUTFile.WriteLine "    </table>"

objOUTFile.WriteLine "  <tr valign=" & """" & "top" & """" & ">"
objOUTFile.WriteLine "    <td> <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "    <td width=" & """" & "100%" & """" & " valign=" & """" & "top" & """" & " bgcolor=" & """" & "#000066" & """" & "> <div align=" & """" & "center" & """" & "><font color=" & """" & "#FFFFFF" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><strong>Event Viewer - Application (Últimos 10 Eventos)</strong></font></div></td>"
objOUTFile.WriteLine "  </tr>"
objOUTFile.WriteLine "  </Table>"
objOUTFile.WriteLine "    <table width=" & """" & "100%" & """" & " border=" & """" & "1" & """" & " CellSpacing=" & """" & "0" & """" & ">"
objOUTFile.WriteLine "      <tr>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>EventID</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Date</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Source</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Type</b></font></td>"
objOUTFile.WriteLine "              <td width=" & """" & "60%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#000066" & """" & "><font color=" & """" & "#FFFFFF" & """" & " size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & "><b>Message</b></font></td>"
objOUTFile.WriteLine "      </tr>"

strCount = 0
Set colLoggedEvents = objWMIService.ExecQuery ("Select * from Win32_NTLogEvent Where Logfile = 'Application'")
For Each strEventApp in colLoggedEvents
    If strCount <> 11 Then
       If strEventApp.EventType = 1 Then
          strType = "Error"
       ElseIf strEventApp.EventType = 2 Then
          strType = "Warning"
       ElseIf strEventApp.EventType = 3 Then
          strType = "Information"
       End If
       strGenerated = Mid(strEventApp.TimeGenerated, 7, 2) & "/" & Mid(strEventApp.TimeGenerated, 5, 2) & "/" & Left(strEventApp.TimeGenerated, 4) & " " & Mid(strEventApp.TimeGenerated, 9, 2) & ":" & Mid(strEventApp.TimeGenerated, 11, 2) & ":" & Mid(strEventApp.TimeGenerated, 13, 2)
       If strType <> "Information" Then
          strCount = strCount + 1
          objOUTFile.WriteLine "      <tr>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strEventApp.EventCode & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strGenerated & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strEventApp.SourceName & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "10%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & strType & "</font></td>"
          objOUTFile.WriteLine "          <td width=" & """" & "60%" & """" & " align=" & """" & "center" & """" & " valign=" & """" & "middle" & """" & " bgcolor=" & """" & "#E2E2E2" & """" & "><font color=000000 size=" & """" & "2" & """" & " face=" & """" & "Arial, Helvetica, sans-serif" & """" & ">" & Left(strEventApp.Message, 100) & "</font></td>"
          objOUTFile.WriteLine "      </tr>"
       End If
    Else
       Exit For
    End If
Next
objOUTFile.WriteLine "    </table>"
objOUTFile.WriteLine "</body>"
objOUTFile.WriteLine "</html>"
ObjOUTFile.Close

Set objResFile = objFSO.CreateTextFile (strScriptPath & "\FTP", ForWriting)
objResFile.WriteLine "ogsuporte"
objResFile.WriteLine "Asdo" & "&" & "amudT05"
objResFile.WriteLine "cd Checklists"
objResFile.WriteLine "cd Windows"
objResFile.WriteLine "bi"
objResFile.WriteLine "put " & """" & strScriptPath & "\" & strOUTFile & """"
objResFile.WriteLine "bye"
objResFile.Close
WshShell.Run "cmd /c ftp -s:" & strScriptPath & "\" & "FTP" & " 200.185.21.10 >" & """" & strScriptPath & "\" & "FTP_LOG" & """", 0, True
objFSO.DeleteFile strScriptPath & "\" & "FTP", True
Set objFTPLog = objFSO.OpenTextFile (strScriptPath & "\" & "FTP_LOG", ForReading)
strRet = objFTPLog.ReadAll
objFTPLog.Close
objFSO.DeleteFile strScriptPath & "\" & "FTP_LOG", True
If instr(1, strRet, "Transfer complete.") > 0 Then
   MsgBox "Checklist gerado e copiado para o FTP 200.185.21.10 em 'Checklists\Windows'.",,"Checklist - Windows"
Else
   MsgBox "Checklist gerado. Não foi possível copiar para o FTP 200.185.21.10.",,"Checklist - Windows"
End If
