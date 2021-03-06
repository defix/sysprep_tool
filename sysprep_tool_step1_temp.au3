#NoTrayIcon
#RequireAdmin
Global $aNicconfig = _NetworkAdapterInfo()
Global $sNicconfig

DirCreate("D:\Users Documents\Desktop\")
DirCreate("D:\Users Documents\Documents\")
DirCreate("D:\Users Documents\Music\")
DirCreate("D:\Users Documents\Favorites\")
DirCreate("D:\Users Documents\Pictures\")
DirCreate("D:\Users Documents\Videos\")
FileDelete("info.ini")

For $i = 0 To $aNicconfig[0][0] Step 1
	IniWrite("info.ini", "IP", "Model" & $aNicconfig[$i][0], $aNicconfig[$i][1])
	IniWrite("info.ini", "IP", "IP" & $aNicconfig[$i][0], $aNicconfig[$i][2])
	IniWrite("info.ini", "IP", "MASK" & $aNicconfig[$i][0], $aNicconfig[$i][3])
	IniWrite("info.ini", "IP", "GateWay" & $aNicconfig[$i][0], $aNicconfig[$i][4])
	IniWrite("info.ini", "IP", "GWMitic" & $aNicconfig[$i][0], $aNicconfig[$i][5])
	IniWrite("info.ini", "IP", "DNS" & $aNicconfig[$i][0], $aNicconfig[$i][6])
	IniWrite("info.ini", "IP", "MAC" & $aNicconfig[$i][0], $aNicconfig[$i][7])
Next

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
	Local $NIC_Info[10][8] ;���9������
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
