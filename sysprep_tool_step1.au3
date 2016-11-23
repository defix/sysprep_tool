#NoTrayIcon
#RequireAdmin
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GuiEdit.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <nd.Local.au3>
#include <UDF_AscII2Alpha_SnT_Exchange.au3>
#include <UDF_MSSQL.au3>

Global $department
Global $displayname
Global $username
Global $title
Global $mobile
Global $telephone
Global $sNicconfig
Global $sql_ip
Global $sql_user
Global $sql_pw
Global $sql_db
Global $cSQL
Global $idNext
Global $aNicconfig
Global $hostname
Global $oMyError

SJYHXX()
CLYHXX()
Exit

Func SJYHXX()
	#Region ### START Koda GUI section ### Form=
	$hForm1 = GUICreate("收集用户信息", 346, 499, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_TABSTOP))
	GUISetBkColor(0xFFFBF0)
	$idDepartment = GUICtrlCreateCombo("", 112, 120, 209, 25, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_SIMPLE))
	GUICtrlSetData(-1, $aCombo_ou_list)
	GUICtrlSetFont(-1, 11, 400, 0, "宋体")
	Global $idDisplayname = GUICtrlCreateEdit("", 112, 168, 209, 28, $ES_WANTRETURN)
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	Global $idUsername = GUICtrlCreateEdit("", 112, 216, 209, 28, $ES_WANTRETURN)
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	$idTitle = GUICtrlCreateEdit("", 112, 264, 209, 28, $ES_WANTRETURN)
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	$idHostname = GUICtrlCreateEdit("", 112, 312, 209, 28, $ES_WANTRETURN)
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	$idMobile = GUICtrlCreateEdit("", 112, 360, 209, 27, BitOR($ES_WANTRETURN, $ES_NUMBER))
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	$idTelephone = GUICtrlCreateEdit("", 161, 408, 161, 27, BitOR($ES_WANTRETURN, $ES_NUMBER))
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	$idNext = GUICtrlCreateButton("下一步", 248, 456, 75, 33)
	GUICtrlSetFont(-1, 12, 400, 0, "宋体")
	GUICtrlCreatePic("D:\Users\yang\Documents\GitHub\sysprep_tool\logo.bmp", 24, 8, 76, 100, BitOR($GUI_SS_DEFAULT_PIC, $WS_TABSTOP))
	GUICtrlCreateLabel("部　　门：", 24, 120, 84, 24, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("姓　　名：", 24, 168, 84, 24, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("账　　号：", 24, 216, 84, 24)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("职　　务：", 24, 264, 84, 24, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("主　　机：", 24, 312, 84, 24, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("手　　机：", 24, 360, 84, 24, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("固定电话：", 24, 408, 84, 24, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreateLabel("0951-", 112, 411, 49, 23, $WS_TABSTOP)
	GUICtrlSetFont(-1, 12, 400, 0, "Consolas")
	GUICtrlCreateEdit("请如实填写如下信息，此信息将用作使用人员与计算机的绑定。", 112, 8, 209, 100, BitOR($ES_READONLY, $ES_WANTRETURN))
	GUICtrlSetFont(-1, 10, 400, 0, "宋体")
	GUICtrlSetColor(-1, 0xFF0000)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
	While 1
		Local $nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $idNext
				$department = GUICtrlRead($idDepartment)
				$displayname = GUICtrlRead($idDisplayname)
				$username = GUICtrlRead($idUsername)
				$title = GUICtrlRead($idTitle)
				$hostname = GUICtrlRead($idHostname)
				$mobile = GUICtrlRead($idMobile)
				$telephone = "0951-" & GUICtrlRead($idTelephone)
				If ($department = "" Or $displayname = "" Or $username = "" Or $title = "" Or $mobile = "" Or $telephone = "0951-") Then
					MsgBox(0, "请填写完整信息", "请填写完整信息，不要留空，谢谢！")
				Else
					GUIDelete($hForm1)
					ExitLoop
				EndIf
			Case $idDisplayname
		EndSwitch
	WEnd
EndFunc   ;==>SJYHXX

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg
	Local $hWndFrom, $iIDFrom, $iCode, $hWndEdit
	$hWndEdit = GUICtrlGetHandle($idDisplayname)
	$hWndFrom = $ilParam
	$iIDFrom = _WinAPI_LoWord($iwParam)
	$iCode = _WinAPI_HiWord($iwParam)
	Switch $hWndFrom
		Case $idDisplayname, $hWndEdit
			Switch $iCode
				Case $EN_CHANGE
					_GUICtrlEdit_SetText($idUsername, _StringToLetter(GUICtrlRead($idDisplayname), 1))
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND


Func CLYHXX()
	#Region ### START Koda GUI section ### Form=
	$hForm2 = GUICreate("处理用户信息", 346, 451, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_TABSTOP))
	GUISetBkColor(0xFFFBF0)
	$idNext = GUICtrlCreateButton("提交", 248, 408, 75, 33)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlCreatePic("D:\Users\yang\Documents\GitHub\sysprep_tool\logo.bmp", 24, 8, 76, 100, BitOR($GUI_SS_DEFAULT_PIC, $WS_TABSTOP))
	GUICtrlCreateEdit("请如实填写如下信息，此信息将用作使用人员与计算机的绑定。", 112, 8, 209, 100, BitOR($ES_READONLY, $ES_WANTRETURN))
	GUICtrlSetFont(-1, 10, 400, 0, "宋体")
	GUICtrlSetColor(-1, 0xFF0000)
	Global $idOutput = GUICtrlCreateEdit("", 24, 120, 297, 281)
	GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	_Get_NIC_config()
	_UpdateOutput()
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $idNext
				If 0 Or $username = "" Then
					$username = InputBox("用户名冲突", "用户名 " & $username & " 与现有用户冲突，请输入新的用户名。", $username)
					_UpdateOutput()
				Else
					GUICtrlSetState($idNext, $GUI_DISABLE)
					_Submit()
				EndIf
		EndSwitch
	WEnd

EndFunc   ;==>CLYHXX

Func _UpdateOutput()
	$department = StringStripCR($aCombo_ou_list_to_OU[_ArraySearch($aCombo_ou_list_to_OU, $department)][1])
	GUICtrlSetData($idOutput, _
			"部　　门：" & StringTrimRight(StringReplace($department, "OU=", ""), 15) & @CRLF & _
			"姓　　名：" & $displayname & @CRLF & _
			"账　　号：" & $username & @CRLF & _
			"主　　机：" & $hostname & @CRLF & _
			"职　　务：" & $title & @CRLF & _
			"手　　机：" & $mobile & @CRLF & _
			"固定电话：" & $telephone & @CRLF & _
			$sNicconfig)
EndFunc   ;==>_UpdateOutput

Func _Get_NIC_config()
	Local $iPID = Run(@ComSpec & " /c wmic NICCONFIG WHERE IPEnabled=true GET IPaddress,MACAddress,DefaultIPGateway,IPSubnet,DNSServerSearchOrder,GatewayCostMetric", "", @SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($iPID)
	$sNicconfig = StdoutRead($iPID)
	Global $aNicconfig[10][6] = [ _
			[StringInStr($sNicconfig, "DefaultIPGateway"), StringInStr($sNicconfig, "DNSServerSearchOrder"), StringInStr($sNicconfig, "GatewayCostMetric"), StringInStr($sNicconfig, "IPAddress"), StringInStr($sNicconfig, "IPSubnet"), StringInStr($sNicconfig, "MACAddress")], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null], _
			[Null, Null, Null, Null, Null, Null]]
	Local $pCRLF = StringInStr($sNicconfig, @CRLF)
	$sNicconfig = StringTrimLeft($sNicconfig, $pCRLF)
	Local $i = 0
	While $pCRLF
		$i += 1
		$pCRLF = StringInStr($sNicconfig, @CRLF)
		$aNicconfig[$i][0] = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringMid($sNicconfig, $aNicconfig[0][0], $aNicconfig[0][1] - $aNicconfig[0][0]), '"', ''), '{', ''), '}', ''), Chr(13), ''), Chr(10), ''), Chr(32), '')
		$aNicconfig[$i][1] = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringMid($sNicconfig, $aNicconfig[0][1], $aNicconfig[0][2] - $aNicconfig[0][1]), '"', ''), '{', ''), '}', ''), Chr(13), ''), Chr(10), ''), Chr(32), '')
		$aNicconfig[$i][2] = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringMid($sNicconfig, $aNicconfig[0][2], $aNicconfig[0][3] - $aNicconfig[0][2]), '"', ''), '{', ''), '}', ''), Chr(13), ''), Chr(10), ''), Chr(32), '')
		$aNicconfig[$i][3] = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringMid($sNicconfig, $aNicconfig[0][3], $aNicconfig[0][4] - $aNicconfig[0][3]), '"', ''), '{', ''), '}', ''), Chr(13), ''), Chr(10), ''), Chr(32), '')
		$aNicconfig[$i][4] = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringMid($sNicconfig, $aNicconfig[0][4], $aNicconfig[0][5] - $aNicconfig[0][4]), '"', ''), '{', ''), '}', ''), Chr(13), ''), Chr(10), ''), Chr(32), '')
		$aNicconfig[$i][5] = StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringReplace(StringMid($sNicconfig, $aNicconfig[0][5], $pCRLF - $aNicconfig[0][5]), '"', ''), '{', ''), '}', ''), Chr(13), ''), Chr(10), ''), Chr(32), '')
		$sNicconfig = StringTrimLeft($sNicconfig, $pCRLF)
	WEnd
	$sNicconfig = ""
	$i = 1
	While $i <= 9
		If $aNicconfig[$i][3] <> "" Then
			$sNicconfig &= "网卡" & $i & ":　IP:" & $aNicconfig[$i][3] & @CRLF & "　　　　Subnet:" & $aNicconfig[$i][4] & @CRLF & "　　　　Gateway:" & $aNicconfig[$i][0] & "　Mitic:" & $aNicconfig[$i][2] & @CRLF & "　　　　DNSs:" & $aNicconfig[$i][1] & @CRLF & "　　　　MAC:" & $aNicconfig[$i][5] & @CRLF
		EndIf
		$i += 1
	WEnd
EndFunc   ;==>_Get_NIC_config

Func _Submit()
	$cSQL = _MSSQL_Con($sql_ip, $sql_user, $sql_pw, $sql_db)
	$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
	Local $existUsername = _MSSQL_GetRecord($cSQL, "allinfo", "department", "WHERE username='" & $username & "'", "department")
	If $existUsername <> "" Then
		$username = InputBox("账号重复", "账号" & $username & "已经存在在系统中，请换一个用户名。", $username)
		_UpdateOutput()
		GUICtrlSetState($idNext, $GUI_ENABLE)
	Else
		Local $bSQL = _MSSQL_AddRecord($cSQL, "allinfo", "'" & _
				$department & "', '" & _
				$displayname & "', '" & _
				$username & "', '" & _
				$hostname & "', '" & _
				$title & "', '" & _
				$mobile & "', '" & _
				$telephone & "', '" & _
				$aNicconfig[1][3] & "', '" & $aNicconfig[1][4] & "', '" & $aNicconfig[1][0] & "', '" & $aNicconfig[1][2] & "', '" & $aNicconfig[1][1] & "', '" & $aNicconfig[1][5] & "', '" & _
				$aNicconfig[2][3] & "', '" & $aNicconfig[2][4] & "', '" & $aNicconfig[2][0] & "', '" & $aNicconfig[2][2] & "', '" & $aNicconfig[2][1] & "', '" & $aNicconfig[2][5] & "', '" & _
				$aNicconfig[3][3] & "', '" & $aNicconfig[3][4] & "', '" & $aNicconfig[3][0] & "', '" & $aNicconfig[3][2] & "', '" & $aNicconfig[3][1] & "', '" & $aNicconfig[3][5] & "', '" & _
				$aNicconfig[4][3] & "', '" & $aNicconfig[4][4] & "', '" & $aNicconfig[4][0] & "', '" & $aNicconfig[4][2] & "', '" & $aNicconfig[4][1] & "', '" & $aNicconfig[4][5] & "', '" & _
				$aNicconfig[5][3] & "', '" & $aNicconfig[5][4] & "', '" & $aNicconfig[5][0] & "', '" & $aNicconfig[5][2] & "', '" & $aNicconfig[5][1] & "', '" & $aNicconfig[5][5] & "', '" & _
				$aNicconfig[6][3] & "', '" & $aNicconfig[6][4] & "', '" & $aNicconfig[6][0] & "', '" & $aNicconfig[6][2] & "', '" & $aNicconfig[6][1] & "', '" & $aNicconfig[6][5] & "', '" & _
				$aNicconfig[7][3] & "', '" & $aNicconfig[7][4] & "', '" & $aNicconfig[7][0] & "', '" & $aNicconfig[7][2] & "', '" & $aNicconfig[7][1] & "', '" & $aNicconfig[7][5] & "', '" & _
				$aNicconfig[8][3] & "', '" & $aNicconfig[8][4] & "', '" & $aNicconfig[8][0] & "', '" & $aNicconfig[8][2] & "', '" & $aNicconfig[8][1] & "', '" & $aNicconfig[8][5] & "', '" & _
				$aNicconfig[9][3] & "', '" & $aNicconfig[9][4] & "', '" & $aNicconfig[9][0] & "', '" & $aNicconfig[9][2] & "', '" & $aNicconfig[9][1] & "', '" & $aNicconfig[9][5] & "', '" & _
				@YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "'")
		If $bSQL Then
			MsgBox(0, "成功", "信息已写入数据库，请勿重复提交")
			FileDelete("info.ini")
			IniWrite("info.ini", "UserInfo", "department", $department)
			IniWrite("info.ini", "UserInfo", "displayname", $displayname)
			IniWrite("info.ini", "UserInfo", "username", $username)
			IniWrite("info.ini", "UserInfo", "hostname", $hostname)
			IniWrite("info.ini", "UserInfo", "title", $title)
			IniWrite("info.ini", "UserInfo", "mobile", $mobile)
			IniWrite("info.ini", "UserInfo", "elephone", $telephone)
			IniWrite("info.ini", "IP", "IP1", $aNicconfig[1][3])
			IniWrite("info.ini", "IP", "MASK1", $aNicconfig[1][4])
			IniWrite("info.ini", "IP", "GateWay1", $aNicconfig[1][0])
			IniWrite("info.ini", "IP", "GWMitic1", $aNicconfig[1][2])
			IniWrite("info.ini", "IP", "DNS1", $aNicconfig[1][1])
			IniWrite("info.ini", "IP", "MAC1", $aNicconfig[1][5])
			IniWrite("info.ini", "IP", "IP2", $aNicconfig[2][3])
			IniWrite("info.ini", "IP", "MASK2", $aNicconfig[2][4])
			IniWrite("info.ini", "IP", "GateWay2", $aNicconfig[2][0])
			IniWrite("info.ini", "IP", "GWMitic2", $aNicconfig[2][2])
			IniWrite("info.ini", "IP", "DNS2", $aNicconfig[2][1])
			IniWrite("info.ini", "IP", "MAC2", $aNicconfig[2][5])
			IniWrite("info.ini", "IP", "IP3", $aNicconfig[3][3])
			IniWrite("info.ini", "IP", "MASK3", $aNicconfig[3][4])
			IniWrite("info.ini", "IP", "GateWay3", $aNicconfig[3][0])
			IniWrite("info.ini", "IP", "GWMitic3", $aNicconfig[3][2])
			IniWrite("info.ini", "IP", "DNS3", $aNicconfig[3][1])
			IniWrite("info.ini", "IP", "MAC3", $aNicconfig[3][5])
			IniWrite("info.ini", "IP", "IP4", $aNicconfig[4][3])
			IniWrite("info.ini", "IP", "MASK4", $aNicconfig[4][4])
			IniWrite("info.ini", "IP", "GateWay4", $aNicconfig[4][0])
			IniWrite("info.ini", "IP", "GWMitic4", $aNicconfig[4][2])
			IniWrite("info.ini", "IP", "DNS4", $aNicconfig[4][1])
			IniWrite("info.ini", "IP", "MAC4", $aNicconfig[4][5])
			IniWrite("info.ini", "IP", "IP5", $aNicconfig[5][3])
			IniWrite("info.ini", "IP", "MASK5", $aNicconfig[5][4])
			IniWrite("info.ini", "IP", "GateWay5", $aNicconfig[5][0])
			IniWrite("info.ini", "IP", "GWMitic5", $aNicconfig[5][2])
			IniWrite("info.ini", "IP", "DNS5", $aNicconfig[5][1])
			IniWrite("info.ini", "IP", "MAC5", $aNicconfig[5][5])
			IniWrite("info.ini", "IP", "IP6", $aNicconfig[6][3])
			IniWrite("info.ini", "IP", "MASK6", $aNicconfig[6][4])
			IniWrite("info.ini", "IP", "GateWay6", $aNicconfig[6][0])
			IniWrite("info.ini", "IP", "GWMitic6", $aNicconfig[6][2])
			IniWrite("info.ini", "IP", "DNS6", $aNicconfig[6][1])
			IniWrite("info.ini", "IP", "MAC6", $aNicconfig[6][5])
			IniWrite("info.ini", "IP", "IP7", $aNicconfig[7][3])
			IniWrite("info.ini", "IP", "MASK7", $aNicconfig[7][4])
			IniWrite("info.ini", "IP", "GateWay7", $aNicconfig[7][0])
			IniWrite("info.ini", "IP", "GWMitic7", $aNicconfig[7][2])
			IniWrite("info.ini", "IP", "DNS7", $aNicconfig[7][1])
			IniWrite("info.ini", "IP", "MAC7", $aNicconfig[7][5])
			IniWrite("info.ini", "IP", "IP8", $aNicconfig[8][3])
			IniWrite("info.ini", "IP", "MASK8", $aNicconfig[8][4])
			IniWrite("info.ini", "IP", "GateWay8", $aNicconfig[8][0])
			IniWrite("info.ini", "IP", "GWMitic8", $aNicconfig[8][2])
			IniWrite("info.ini", "IP", "DNS8", $aNicconfig[8][1])
			IniWrite("info.ini", "IP", "MAC8", $aNicconfig[8][5])
			IniWrite("info.ini", "IP", "IP9", $aNicconfig[9][3])
			IniWrite("info.ini", "IP", "MASK9", $aNicconfig[9][4])
			IniWrite("info.ini", "IP", "GateWay9", $aNicconfig[9][0])
			IniWrite("info.ini", "IP", "GWMitic9", $aNicconfig[9][2])
			IniWrite("info.ini", "IP", "DNS9", $aNicconfig[9][1])
			IniWrite("info.ini", "IP", "MAC9", $aNicconfig[9][5])
			Exit
		Else
			MsgBox(0, "添加记录失败", "添加记录失败，请重试！")
			_UpdateOutput()
			GUICtrlSetState($idNext, $GUI_ENABLE)
		EndIf

	EndIf
	_MSSQL_End($cSQL)
EndFunc   ;==>_Submit

Func MyErrFunc()
	Local $HexNumber
	$HexNumber = Hex($oMyError.number, 8)
	MsgBox(0, "COM Error Test", "We intercepted a COM Error !" & @CRLF & @CRLF & _
			"err.description is: " & @TAB & $oMyError.description & @CRLF & _
			"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
			"err.number is: " & @TAB & $HexNumber & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
			"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oMyError.helpcontext _
			)
	SetError(1) ; to check for after this function returns
EndFunc   ;==>MyErrFunc

