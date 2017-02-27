#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\OutlookReopen\icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include <AD.au3>
#include <Array.au3>


Local $oComError = ObjEvent("AutoIt.Error", "ComErrFunc")


Local $logFilePath = "C:\Temp\" & @ScriptName & " - " & @UserName & ".log"
If FileExists($logFilePath) Then FileDelete($logFilePath)

Local $aUsersToChange = ["s.v.antyuganova", _
						"v.v.artanova", _
						"u.b.afonina", _
						"o.n.baglay", _
						"e.e.bagrova", _
						"e.e.baranovskaya", _
						"j.a.biryukova", _
						"e.v.bochkareva", _
						"s.a.bysheva", _
						"n.a.veselov", _
						"s.a.volkova", _
						"l.p.volokitina", _
						"o.v.voronina", _
						"a.i.vybina", _
						"i.v.gavrikova", _
						"t.v.gorneeva", _
						"n.n.dorojkina", _
						"e.v.dybodelova", _
						"i.a.elagina", _
						"e.y.ermakov", _
						"l.p.ermakova", _
						"o.v.zhunina", _
						"n.g.zelenov", _
						"g.y.zolotova", _
						"e.g.ivancova", _
						"m.a.kazanceva", _
						"t.i.kapanadze", _
						"r.d.karakozova", _
						"i.g.karandashova", _
						"l.n.kataeva", _
						"y.v.kolosova", _
						"y.s.korotkova", _
						"n.v.kravchenko", _
						"a.v.kyklev", _
						"d.y.lebedinec", _
						"a.v.lipskiy", _
						"n.v.lobaskova", _
						"g.a.malanchak", _
						"n.v.maramzina", _
						"o.a.marinina", _
						"m.a.morozova", _
						"n.v.nesupravina", _
						"l.a.nikishina", _
						"e.p.osipov", _
						"i.u.pavlycheva", _
						"n.v.pankratova", _
						"k.u.polovinkina", _
						"e.a.popova", _
						"a.o.puzynya", _
						"o.n.pyrikova", _
						"n.a.razina", _
						"s.s.repina", _
						"s.l.reshteyn", _
						"a.f.ryazapova", _
						"n.v.smirnova", _
						"t.p.smirnova", _
						"n.g.sokolova", _
						"l.a.sorokina", _
						"a.o.spasskiy", _
						"e.a.tuzhilkina", _
						"d.v.tunin", _
						"e.v.firsova", _
						"m.n.hlytina", _
						"s.g.hodakova", _
						"t.a.chvanova", _
						"a.i.cheredilov", _
						"s.s.chernova", _
						"d.a.chizhikov", _
						"s.s.chukalina", _
						"o.v.shirshakova", _
						"y.v.shnayder", _
						"n.a.shokurova", _
						"i.a.shustova", _
						"i.v.schelokov", _
						"a.v.gorev", _
						"temp"]

_ArraySearch($aUsersToChange, @UserName)
If @error Then
	ToLog("user not in list to change")
	Exit
EndIf

ToLog("user in list to change")

Local $sMailDomain = "bzklinika.ru"

If Not _AD_Open() Then Exit
Local $sUserEmail = _AD_GetObjectAttribute(@UserName, "mail")
_AD_Close()

ToLog("$sUserEmail: " & $sUserEmail)

If Not StringInStr($sUserEmail, $sMailDomain) Then
	ToLog("user email in ad doesnt contain " & $sMailDomain)
	Exit
EndIf

Local $sExchangePath = "HKCU\Software\Microsoft\Exchange\Client\Options"
Local $sPickLogonProfile = "PickLogonProfile"

Local $sPLFValue = RegRead($sExchangePath, $sPickLogonProfile)
If Not @error Then
	ToLog("script has already completed")
	Exit
EndIf

ToLog("$sPLFValue: " & $sPLFValue)

Local $sOfficePath = "HKCU\Software\Microsoft\Office"
Local $sAutoDiscover = "\Outlook\AutoDiscover"
Local $sZCE = "ZeroConfigExchange"

Local $sSubKey = ""
Local $i = 1
Local $bOutlookKeyUpdated = False

While True
    $sSubKey = RegEnumKey($sOfficePath, $i)
    If @error Then ExitLoop

	$i += 1

	If Not StringInStr($sSubKey, ".0") Then ContinueLoop
	If $sSubKey < 14.0 Then ContinueLoop

	ToLog($sSubKey)

	Local $bWriteResult = RegWrite($sOfficePath & "\" & $sSubKey & $sAutoDiscover, $sZCE, "REG_DWORD", 1)
	If $bWriteResult Then $bOutlookKeyUpdated = True

	ToLog("result: " & $bWriteResult)
WEnd

If Not $bOutlookKeyUpdated Then ExitOnError()

ToLog("$bOutlookKeyUpdated: " & $bOutlookKeyUpdated)

RegWrite($sExchangePath, $sPickLogonProfile, "REG_SZ", "1")
If @error Then ExitOnError()

ToLog("$sPickLogonProfile updated")

Local $sOutlook = "outlook.exe"

Local $aProcessList = _ProcessListProperties($sOutlook)
;~ _ArrayDisplay($aProcessList)
If IsArray($aProcessList) And UBound($aProcessList, $UBOUND_COLUMNS) > 3 Then
	For $i = 1 To UBound($aProcessList, $UBOUND_ROWS) - 1
		If StringInStr($aProcessList[$i][3], @UserName) Then
			ProcessClose($aProcessList[$i][1])
		EndIf
	Next
EndIf

If ShellExecute($sOutlook) = -1 Then ExitOnError()
ToLog("outlook started")

Sleep(1000)
If WinExists("Microsoft Outlook", "Выполнить запуск в безопасном режиме?") Then
	ToLog("Error exist")
	Send("{ENTER}")
EndIf

Local $sConfigTitle = "Выбор конфигурации"
Local $sConfigText = "Имя конфигурации"

WinActivate($sConfigTitle, $sConfigText)
If Not WinWait($sConfigTitle, $sConfigText, 30) Then ExitOnError()

ToLog("Configuration select")

Local $sNewConfigName = "Exchange"

Local $sREComboBox20W1 = ControlGetText($sConfigTitle, $sConfigText, "REComboBox20W1")

If $sREComboBox20W1 <> $sNewConfigName Then
	ControlClick($sConfigTitle, $sConfigText, "[CLASS:Button; INSTANCE:1]")

	Local $sNewConfigTitle = "Новая конфигурация"
	Local $sNewConfigText = "Создание новой конфигурации"

	WinActivate($sNewConfigTitle, $sNewConfigText)
	If Not WinWait($sNewConfigTitle, $sNewConfigText, 30) Then ExitOnError()
	ToLog("new config")

	ControlSend($sNewConfigTitle, $sNewConfigText, "[CLASS:RichEdit20WPT; INSTANCE:1]", $sNewConfigName)
	ControlClick($sNewConfigTitle, $sNewConfigText, "[CLASS:Button; INSTANCE:1]")
EndIf

RegWrite($sExchangePath, $sPickLogonProfile, "REG_SZ", "0")

WinActivate($sConfigTitle, $sConfigText)
If Not WinWait($sConfigTitle, $sConfigText, 30) Then Exit

ControlClick($sConfigTitle, $sConfigText, "[CLASS:Button; INSTANCE:4]")
ControlClick($sConfigTitle, $sConfigText, "[CLASS:Button; INSTANCE:6]")
ControlClick($sConfigTitle, $sConfigText, "[CLASS:Button; INSTANCE:2]")


WinActivate("Входящие", "Строка состояния")
If Not WinWait("Входящие", "Строка состояния", 60) Then ExitOnError()

ToLog("Complete")

SendEmail()


Func ExitOnError()
	ToLog("!!! Exit on error")
	RegDelete($sExchangePath, $sPickLogonProfile)
	Exit
EndFunc


Func ComErrFunc()
    ConsoleWrite("We intercepted a COM Error !" & @CRLF & _
                "Number is: " & $oComError.number & @CRLF & _
                "WinDescription is: " & $oComError.windescription & @CRLF & _
                "Source is: " & $oComError.source & @CRLF & _
                "ScriptLine is: " & $oComError.scriptline)
Endfunc


;===============================================================================
; Function Name:    _ProcessListProperties()
; Description:   Get various properties of a process, or all processes
; Call With:       _ProcessListProperties( [$Process [, $sComputer]] )
; Parameter(s):  (optional) $Process - PID or name of a process, default is "" (all)
;          (optional) $sComputer - remote computer to get list from, default is local
; Requirement(s):   AutoIt v3.2.4.9+
; Return Value(s):  On Success - Returns a 2D array of processes, as in ProcessList()
;            with additional columns added:
;            [0][0] - Number of processes listed (can be 0 if no matches found)
;            [1][0] - 1st process name
;            [1][1] - 1st process PID
;            [1][2] - 1st process Parent PID
;            [1][3] - 1st process owner
;            [1][4] - 1st process priority (0 = low, 31 = high)
;            [1][5] - 1st process executable path
;            [1][6] - 1st process CPU usage
;            [1][7] - 1st process memory usage
;            [1][8] - 1st process creation date/time = "MM/DD/YYY hh:mm:ss" (hh = 00 to 23)
;            [1][9] - 1st process command line string
;            ...
;            [n][0] thru [n][9] - last process properties
; On Failure:      Returns array with [0][0] = 0 and sets @Error to non-zero (see code below)
; Author(s):        PsaltyDS at http://www.autoitscript.com/forum
; Date/Version:   12/01/2009  --  v2.0.4
; Notes:            If an integer PID or string process name is provided and no match is found,
;            then [0][0] = 0 and @error = 0 (not treated as an error, same as ProcessList)
;          This function requires admin permissions to the target computer.
;          All properties come from the Win32_Process class in WMI.
;            To get time-base properties (CPU and Memory usage), a 100ms SWbemRefresher is used.
;===============================================================================
Func _ProcessListProperties($Process = "", $sComputer = ".")
    Local $sUserName, $sMsg, $sUserDomain, $avProcs, $dtmDate
    Local $avProcs[1][2] = [[0, ""]], $n = 1

    ; Convert PID if passed as string
    If StringIsInt($Process) Then $Process = Int($Process)

    ; Connect to WMI and get process objects
    $oWMI = ObjGet("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy, (Debug)}!\\" & $sComputer & "\root\cimv2")
    If IsObj($oWMI) Then
        ; Get collection processes from Win32_Process
        If $Process == "" Then
            ; Get all
            $colProcs = $oWMI.ExecQuery("select * from win32_process")
        ElseIf IsInt($Process) Then
            ; Get by PID
            $colProcs = $oWMI.ExecQuery("select * from win32_process where ProcessId = " & $Process)
        Else
            ; Get by Name
            $colProcs = $oWMI.ExecQuery("select * from win32_process where Name = '" & $Process & "'")
        EndIf

        If IsObj($colProcs) Then
            ; Return for no matches
            If $colProcs.count = 0 Then Return $avProcs

            ; Size the array
            ReDim $avProcs[$colProcs.count + 1][10]
            $avProcs[0][0] = UBound($avProcs) - 1

            ; For each process...
            For $oProc In $colProcs
                ; [n][0] = Process name
                $avProcs[$n][0] = $oProc.name
                ; [n][1] = Process PID
                $avProcs[$n][1] = $oProc.ProcessId
                ; [n][2] = Parent PID
                $avProcs[$n][2] = $oProc.ParentProcessId
                ; [n][3] = Owner
                If $oProc.GetOwner($sUserName, $sUserDomain) = 0 Then $avProcs[$n][3] = $sUserDomain & "\" & $sUserName
                ; [n][4] = Priority
                $avProcs[$n][4] = $oProc.Priority
                ; [n][5] = Executable path
                $avProcs[$n][5] = $oProc.ExecutablePath
                ; [n][8] = Creation date/time
                $dtmDate = $oProc.CreationDate
                If $dtmDate <> "" Then
                    ; Back referencing RegExp pattern from weaponx
                    Local $sRegExpPatt = "\A(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(?:.*)"
                    $dtmDate = StringRegExpReplace($dtmDate, $sRegExpPatt, "$2/$3/$1 $4:$5:$6")
                EndIf
                $avProcs[$n][8] = $dtmDate
                ; [n][9] = Command line string
                $avProcs[$n][9] = $oProc.CommandLine

                ; increment index
                $n += 1
            Next
        Else
            SetError(2); Error getting process collection from WMI
        EndIf
        ; release the collection object
        $colProcs = 0

        ; Get collection of all processes from Win32_PerfFormattedData_PerfProc_Process
        ; Have to use an SWbemRefresher to pull the collection, or all Perf data will be zeros
        Local $oRefresher = ObjCreate("WbemScripting.SWbemRefresher")
        $colProcs = $oRefresher.AddEnum($oWMI, "Win32_PerfFormattedData_PerfProc_Process" ).objectSet
        $oRefresher.Refresh

        ; Time delay before calling refresher
        Local $iTime = TimerInit()
        Do
            Sleep(20)
        Until TimerDiff($iTime) >= 100
        $oRefresher.Refresh

        ; Get PerfProc data
        For $oProc In $colProcs
            ; Find it in the array
            For $n = 1 To $avProcs[0][0]
                If $avProcs[$n][1] = $oProc.IDProcess Then
                    ; [n][6] = CPU usage
                    $avProcs[$n][6] = $oProc.PercentProcessorTime
                    ; [n][7] = memory usage
                    $avProcs[$n][7] = $oProc.WorkingSet
                    ExitLoop
                EndIf
            Next
        Next
    Else
        SetError(1); Error connecting to WMI
    EndIf

    ; Return array
    Return $avProcs
EndFunc  ;==>_ProcessListProperties




Func SendEmail()
	Local $sMailServer = ""
	Local $sMailLogin = ""
	Local $sMailPassword = ""
	Local $sMailTo = ""

	Local $title = "Уведомление от " & @ScriptName
	Local $messageToSend = "Скрипт успешно отработал для пользователя " & @UserName & @CRLF & @CRLF & _
			"---------------------------------------" & @CRLF & _
			"Это автоматическое сообщение." & @CRLF & _
			"Пожалуйста, не отвечайте на него." & @CRLF & _
			"Имя системы: " & @ComputerName

	_INetSmtpMailCom($sMailServer, @ScriptName, $sMailLogin, $sMailTo, _
		$title, $messageToSend, "", "", "", $sMailLogin, $sMailPassword)
EndFunc   ;==>SendEmail


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", _
	$as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", _
	$s_Password = "", $IPPort = 25, $ssl = 0)

	Local $objEmail = ObjCreate("CDO.Message")
	Local $i_Error = 0
	Local $i_Error_desciption = ""

	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			If FileExists($S_Files2Attach[$x]) Then
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				$i_Error_desciption = $i_Error_desciption & @LF & 'File not found to attach: ' & $S_Files2Attach[$x]
				$as_Body &= $i_Error_desciption & @CRLF
			EndIf
		Next
	EndIf

	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
		$objEmail.TextBodyPart.Charset = "utf-8"
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf

	$objEmail.Configuration.Fields.Update
	$objEmail.Send

	If @error Then Return False
	Return True
EndFunc   ;==>_INetSmtpMailCom



Func ToLog($message)
	$message &= @CRLF
	ConsoleWrite($message)
	_FileWriteLog($logFilePath, $message)
EndFunc   ;==>ToLog