#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$form1 = GUICreate("收集用户信息", 355, 406, 739, 336, BitOR($GUI_SS_DEFAULT_GUI,$WS_MAXIMIZEBOX,$WS_TABSTOP))
GUISetBkColor(0xFFFBF0)
$department_gui = GUICtrlCreateCombo("", 112, 120, 209, 25, BitOR($GUI_SS_DEFAULT_COMBO,$CBS_SIMPLE))
GUICtrlSetData(-1, "宁东管委会|　├办公室|　├党群工作部|　│　├机关党委|　│　├机关工会|　│　└团工会|　├监察室|　│　└妇女联合会|　├督查室|　├经济技术合作局|　├经济发展局|　├规划建设土地局|　│　└建筑工程质量监督站|　├社会事务局|　├环境保护局|　│　└环境监察支队|　├财政局|　├审计局|　├战略规划局|　├循环经济研究院|　├安全生产监督管理局|　├综合执法局|　├政务服务中心|　├信息中心|　├环境监测站|　│　├综合科|　│　└业务科|　├宁东开发投资有限公司|　│　├总经办|　│　├综合管理部|　│　├投资发展部|　│　├财物部|　│　├工程建设部|　│　├资产运营部|　│　├司机室|　│　├宁东天然气公司|　│　├物业管理办公室|　│　├项目公司筹备组|　│　└污水处理公司|　├宁东管委会司机室|　├能化公司|　├维护中心|　├宁东创投|　└担保公司")
GUICtrlSetFont(-1, 10, 400, 0, "宋体")
$displayname = GUICtrlCreateInput("", 112, 168, 209, 28)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$title = GUICtrlCreateInput("", 112, 216, 209, 28)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$mobile = GUICtrlCreateInput("", 112, 264, 209, 27, BitOR($GUI_SS_DEFAULT_INPUT,$ES_NUMBER))
GUICtrlSetFont(-1, 12, 400, 0, "Consolas")
$telephone = GUICtrlCreateInput("", 161, 312, 161, 27, BitOR($GUI_SS_DEFAULT_INPUT,$ES_NUMBER))
GUICtrlSetFont(-1, 12, 400, 0, "Consolas")
$next = GUICtrlCreateButton("下一步", 248, 360, 75, 33)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$logo = GUICtrlCreatePic("D:\Users\yang\Documents\GitHub\sysprep_tool\logo.bmp", 24, 8, 76, 100, BitOR($GUI_SS_DEFAULT_PIC,$WS_TABSTOP))
GUICtrlCreateLabel("部　　门：", 24, 120, 84, 24, $WS_TABSTOP)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlCreateLabel("姓　　名：", 24, 168, 84, 24, $WS_TABSTOP)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlCreateLabel("职　　务：", 24, 216, 84, 24, $WS_TABSTOP)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlCreateLabel("手　　机：", 24, 264, 84, 24, $WS_TABSTOP)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlCreateLabel("固定电话：", 24, 312, 84, 24, $WS_TABSTOP)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUICtrlCreateLabel("0951-", 112, 315, 49, 23, $WS_TABSTOP)
GUICtrlSetFont(-1, 12, 400, 0, "Consolas")
GUICtrlCreateEdit("", 112, 8, 209, 100, BitOR($ES_WANTRETURN,$WS_BORDER))
GUICtrlSetData(-1, StringFormat("请如实填写，\r\n所填写信息将用于使用人员和计算\r\n机的绑定。"))
GUICtrlSetFont(-1, 10, 400, 0, "宋体")
GUICtrlSetColor(-1, 0xFF0000)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch
WEnd
