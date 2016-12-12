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
Global $aNicconfig = _NetworkAdapterInfo()
Global $sNicconfig
Global $sql_ip
Global $sql_user
Global $sql_pw
Global $sql_db
Global $cSQL
Global $idNext
Global $hostname
Global $domain
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
	$idDomain = GUICtrlCreateCombo("", 24, 456, 200, 25, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_SIMPLE))
	GUICtrlSetData(-1, "nd.local|yzd.zhng.org")
	GUICtrlSetFont(-1, 11, 400, 0, "宋体")
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
				$domain = GUICtrlRead($idDomain)
				If ($department = "" Or $displayname = "" Or $username = "" Or $title = "" Or $mobile = "" Or $telephone = "0951-" Or $domain = "") Then
					MsgBox(0, "请填写完整信息", "请填写完整信息，不要留空，谢谢！")
				Else
					GUIDelete($hForm1)
					ExitLoop
				EndIf
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

	_UpdateOutput()
	
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $idNext
				GUICtrlSetState($idNext, $GUI_DISABLE)
				_Submit()
		EndSwitch
	WEnd

EndFunc   ;==>CLYHXX

Func _UpdateOutput()
	$department = StringStripCR($aCombo_ou_list_to_OU[_ArraySearch($aCombo_ou_list_to_OU, $department)][1])
	$sNicconfig = ""
	For $i = 1 To $aNicconfig[0][0] Step 1
		$sNicconfig &= "网卡" & $aNicconfig[$i][0] & "：" & @CRLF & "　　Model：" & $aNicconfig[$i][1] & @CRLF & "　　IP：" & $aNicconfig[$i][2] & @CRLF & "　　Subnet：" & $aNicconfig[$i][3] & @CRLF & "　　Gateway：" & $aNicconfig[$i][4] & "　　Mitic：" & $aNicconfig[$i][5] & @CRLF & "　　DNSs：" & $aNicconfig[$i][6] & @CRLF & "　　MAC：" & $aNicconfig[$i][7] & @CRLF
	Next
	GUICtrlSetData($idOutput, _
			"　域：" & $domain & @CRLF & _
			"部门：" & StringTrimRight(StringReplace($department, "OU=", ""), 15) & @CRLF & _
			"姓名：" & $displayname & @CRLF & _
			"账号：" & $username & @CRLF & _
			"主机：" & $hostname & @CRLF & _
			"职务：" & $title & @CRLF & _
			"手机：" & $mobile & @CRLF & _
			"座机：" & $telephone & @CRLF & _
			$sNicconfig)
EndFunc   ;==>_UpdateOutput

Func _NetworkAdapterInfo()
	Local $colItems = ""
	Local $objWMIService
	Local $NIC_ID = 0
	Local $NIC_Model = ""
	Local $NIC_Gateway = ""
	Local $NIC_HostName = ""
	Local $NIC_IP = ""
	Local $NIC_DNS1 = ""
	Local $NIC_DNS2 = ""
	Local $NIC_Subnet = ""
	Local $NIC_MAC = ""
	Local $NIC_NetConnectionID = ""
	Local $NIC_Info[10][8] ;最高9块网卡.
	$NIC_Info[0][0] = 0
	$NIC_Info[0][1] = "Model"
	$NIC_Info[0][2] = "IP"
	$NIC_Info[0][3] = "Subnet"
	$NIC_Info[0][4] = "Gateway"
	$NIC_Info[0][5] = "GWMetric"
	$NIC_Info[0][6] = "DNS"
	$NIC_Info[0][7] = "MAC"
	$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled != 0", "WQL", 0x10 + 0x20)
	If IsObj($colItems) Then
		For $objItem In $colItems
			$NIC_Model = $objItem.Description
			$NIC_IP = $objItem.IPAddress
			$NIC_Subnet = $objItem.IPSubnet
			$NIC_Gateway = $objItem.DefaultIPGateway
			$NIC_GWMetric = $objItem.GatewayCostMetric
			$NIC_DNS = $objItem.DNSServerSearchOrder
			$NIC_MAC = $objItem.MACAddress
			If IsArray($NIC_IP) Then
				Local $sTemp1 = ""
				For $sTemp2 In $NIC_IP
					If Not StringInStr($sTemp2, ":") Then
						$sTemp1 &= $sTemp2 & ","
					EndIf
				Next
				$NIC_IP = StringTrimRight($sTemp1, 1)
			EndIf
			If IsArray($NIC_Subnet) Then
				Local $sTemp1 = ""
				For $sTemp2 In $NIC_Subnet
					If StringInStr($sTemp2, ".") Then
						$sTemp1 &= $sTemp2 & ","
					EndIf
				Next
				$NIC_Subnet = StringTrimRight($sTemp1, 1)
			EndIf
			If IsArray($NIC_Gateway) Then
				Local $sTemp1 = ""
				For $sTemp2 In $NIC_Gateway
					If Not StringInStr($sTemp2, ":") Then
						$sTemp1 &= $sTemp2 & ","
					EndIf
				Next
				$NIC_Gateway = StringTrimRight($sTemp1, 1)
			EndIf
			If IsArray($NIC_GWMetric) Then
				Local $sTemp1 = ""
				For $sTemp2 In $NIC_GWMetric
					$sTemp1 &= $sTemp2 & ","
				Next
				$NIC_GWMetric = StringTrimRight($sTemp1, 1)
			EndIf
			If IsArray($NIC_DNS) Then
				Local $sTemp1 = ""
				For $sTemp2 In $NIC_DNS
					If Not StringInStr($sTemp2, ":") Then
						$sTemp1 &= $sTemp2 & ","
					EndIf
				Next
				$NIC_DNS = StringTrimRight($sTemp1, 1)
			EndIf
			$NIC_ID += 1
			$NIC_Info[0][0] = $NIC_ID
			$NIC_Info[$NIC_ID][0] = $NIC_ID
			$NIC_Info[$NIC_ID][1] = $NIC_Model
			$NIC_Info[$NIC_ID][2] = $NIC_IP
			$NIC_Info[$NIC_ID][3] = $NIC_Subnet
			$NIC_Info[$NIC_ID][4] = $NIC_Gateway
			$NIC_Info[$NIC_ID][5] = $NIC_GWMetric
			$NIC_Info[$NIC_ID][6] = $NIC_DNS
			$NIC_Info[$NIC_ID][7] = $NIC_MAC
		Next
		Return $NIC_Info
	Else
		Return $NIC_Info
	EndIf
EndFunc   ;==>_NetworkAdapterInfo

Func _Submit()
	$cSQL = _MSSQL_Con($sql_ip, $sql_user, $sql_pw, $sql_db)
	$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
	Local $existUsername = _MSSQL_GetRecord($cSQL, "allinfo", "department", "WHERE username='" & $username & "' AND domain='" & $domain & "'", "department")
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
				$domain & "', '" & _
				$aNicconfig[1][1] & "', '" & $aNicconfig[1][2] & "', '" & $aNicconfig[1][3] & "', '" & $aNicconfig[1][4] & "', '" & $aNicconfig[1][5] & "', '" & $aNicconfig[1][6] & "', '" & $aNicconfig[1][7] & "', '" & _
				$aNicconfig[2][1] & "', '" & $aNicconfig[2][2] & "', '" & $aNicconfig[2][3] & "', '" & $aNicconfig[2][4] & "', '" & $aNicconfig[2][5] & "', '" & $aNicconfig[2][6] & "', '" & $aNicconfig[2][7] & "', '" & _
				$aNicconfig[3][1] & "', '" & $aNicconfig[3][2] & "', '" & $aNicconfig[3][3] & "', '" & $aNicconfig[3][4] & "', '" & $aNicconfig[3][5] & "', '" & $aNicconfig[3][6] & "', '" & $aNicconfig[3][7] & "', '" & _
				$aNicconfig[4][1] & "', '" & $aNicconfig[4][2] & "', '" & $aNicconfig[4][3] & "', '" & $aNicconfig[4][4] & "', '" & $aNicconfig[4][5] & "', '" & $aNicconfig[4][6] & "', '" & $aNicconfig[4][7] & "', '" & _
				$aNicconfig[5][1] & "', '" & $aNicconfig[5][2] & "', '" & $aNicconfig[5][3] & "', '" & $aNicconfig[5][4] & "', '" & $aNicconfig[5][5] & "', '" & $aNicconfig[5][6] & "', '" & $aNicconfig[5][7] & "', '" & _
				$aNicconfig[6][1] & "', '" & $aNicconfig[6][2] & "', '" & $aNicconfig[6][3] & "', '" & $aNicconfig[6][4] & "', '" & $aNicconfig[6][5] & "', '" & $aNicconfig[6][6] & "', '" & $aNicconfig[6][7] & "', '" & _
				$aNicconfig[7][1] & "', '" & $aNicconfig[7][2] & "', '" & $aNicconfig[7][3] & "', '" & $aNicconfig[7][4] & "', '" & $aNicconfig[7][5] & "', '" & $aNicconfig[7][6] & "', '" & $aNicconfig[7][7] & "', '" & _
				$aNicconfig[8][1] & "', '" & $aNicconfig[8][2] & "', '" & $aNicconfig[8][3] & "', '" & $aNicconfig[8][4] & "', '" & $aNicconfig[8][5] & "', '" & $aNicconfig[8][6] & "', '" & $aNicconfig[8][7] & "', '" & _
				$aNicconfig[9][1] & "', '" & $aNicconfig[9][2] & "', '" & $aNicconfig[9][3] & "', '" & $aNicconfig[9][4] & "', '" & $aNicconfig[9][5] & "', '" & $aNicconfig[9][6] & "', '" & $aNicconfig[9][7] & "', '" & _
				@YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "'")
		If $bSQL Then
			MsgBox(0, "成功", "信息已写入数据库，请勿重复提交")
			FileDelete("info.ini")
			IniWrite("info.ini", "UserInfo", "domain", $domain)
			IniWrite("info.ini", "UserInfo", "department", $department)
			IniWrite("info.ini", "UserInfo", "displayname", $displayname)
			IniWrite("info.ini", "UserInfo", "username", $username)
			IniWrite("info.ini", "UserInfo", "hostname", $hostname)
			IniWrite("info.ini", "UserInfo", "title", $title)
			IniWrite("info.ini", "UserInfo", "mobile", $mobile)
			IniWrite("info.ini", "UserInfo", "elephone", $telephone)
			For $i = 1 To $aNicconfig[0][0] Step 1
				IniWrite("info.ini", "IP", "Model" & $aNicconfig[$i][0], $aNicconfig[$i][1])
				IniWrite("info.ini", "IP", "IP" & $aNicconfig[$i][0], $aNicconfig[$i][2])
				IniWrite("info.ini", "IP", "MASK" & $aNicconfig[$i][0], $aNicconfig[$i][3])
				IniWrite("info.ini", "IP", "GateWay" & $aNicconfig[$i][0], $aNicconfig[$i][4])
				IniWrite("info.ini", "IP", "GWMitic" & $aNicconfig[$i][0], $aNicconfig[$i][5])
				IniWrite("info.ini", "IP", "DNS" & $aNicconfig[$i][0], $aNicconfig[$i][6])
				IniWrite("info.ini", "IP", "MAC" & $aNicconfig[$i][0], $aNicconfig[$i][7])
			Next
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


