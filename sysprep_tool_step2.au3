#NoTrayIcon
#RequireAdmin
#include <WinAPI.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <Array.au3>
#include <AutoItConstants.au3>
;~ #include <Constants.au3>
#include <Permissions.au3>

ToolTip("正在处理部署后的必要流程，请等待", @DesktopWidth - 500, 100, "必要的文件处理等", 2)
$sHostname = IniRead("info.ini", "UserInfo", "Hostname", " <Error> 请核实！")
$sIP = ""
For $i = 1 To 9
	If IniRead("info.ini", "IP", "IP" & $i, "") <> "" Then
		$sIP &= "网卡" & $i & "：" & @CR & _
				"ＩＰ地址" & ": " & IniRead("info.ini", "IP", "IP" & $i, "") & @CR & _
				"子网掩码" & ": " & IniRead("info.ini", "IP", "MASK" & $i, "") & @CR & _
				"网　　关" & ": " & IniRead("info.ini", "IP", "GateWay" & $i, "") & @CR & _
				"网关计数" & ": " & IniRead("info.ini", "IP", "GWMitic" & $i, "") & @CR & _
				"　ＤＮＳ" & ": " & IniRead("info.ini", "IP", "DNS" & $i, "") & @CR & _
				"　ＭＡＣ" & ": " & IniRead("info.ini", "IP", "MAC" & $i, "") & @CR
	EndIf
Next
Global $aDisk = DriveGetDrive($DT_FIXED)
Global $iPID[$aDisk[0]]
$iPID[0] = 0
For $i = 1 To $aDisk[0]
	If $aDisk[$i] <> @HomeDrive And DriveGetFileSystem($aDisk[$i]) = $DT_NTFS Then
		ToolTip("！！！正在处理 " & StringUpper($aDisk[$i]) & "\ 文件权限。请不要重启！！！" & @CR & @CR & "请设置主机名为：" & $sHostname & @CR & @CR & "请设置IP:" & @CR & $sIP, @DesktopWidth - 500, 100, "必要的文件处理等", 2)
		DirRemove($aDisk[$i] & "\$Recycle.Bin\", $DIR_REMOVE)
		DirRemove($aDisk[$i] & "\RECYCLER\", $DIR_REMOVE)
		_GrantAllAccess($aDisk[$i])
		DirRemove($aDisk[$i] & "\$Recycle.Bin", $DIR_REMOVE)
		DirRemove($aDisk[$i] & "\RECYCLER", $DIR_REMOVE)
	EndIf
Next
ToolTip("！！！正在处理 “Chrome绿色版” 文件权限。请不要重启！！！" & @CR & @CR & "请设置主机名为：" & $sHostname & @CR & @CR & "请设置IP:" & @CR & $sIP, @DesktopWidth - 500, 100, "必要的文件处理等", 2)
_GrantAllAccess("C:\Program Files\Chrome")
_GrantAllAccess("C:\Program Files (x86)\Chrome")
ToolTip("请设置主机名为：" & $sHostname & @CR & @CR & "请设置IP:" & @CR & $sIP & @CR & "文件权限处理完毕，设置完主机名和IP地址后，请重启后再安装软硬件。", @DesktopWidth - 500, 100, "必要的文件处理等", 2)
While 1
	Sleep(5000)
WEnd
