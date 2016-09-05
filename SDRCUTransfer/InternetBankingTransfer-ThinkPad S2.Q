[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=6165ca1f-c587-477a-bc0f-619b453d0810
Description=烟台农商银行网银自动转账脚本
Enable=1
AutoRun=0
[Repeat]
Type=1
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
'==========农商银行网银行内半自动转账==========
'==========开始前需先登录并到转账界面==========

Function RandDelay(a, b)
	Randomize
	RandTime = Rnd * (b - a) + a
	Delay RandTime * 1000
End Function

Call RandDelay(1, 2)

'==收款人账号==
MoveTo 1400, 460
Call RandDelay(1, 2)
LeftClick 1
MoveTo 265, 356
Call RandDelay(1, 2)
LeftClick 1
MoveTo 840, 470
Call RandDelay(0.5, 1)
LeftClick 1

'==转账金额==
MoveTo 360, 745
Call RandDelay(0.5, 1)
LeftClick 1
Call RandDelay(0.5, 1)
KeyPress "1", 1
KeyPress "5", 1

'==付款用途==
MoveTo 453, 820
Call RandDelay(0.5, 1)
LeftClick 1
MoveTo 453, 925
Call RandDelay(0.5, 1)
LeftClick 1

'==下一步==
MoveTo 1725, 990
Call RandDelay(0.5, 1)
LeftClick 1

Call RandDelay(1, 2)

'==图片验证码==
MoveTo 372, 860
Call RandDelay(1, 2)
LeftClick 1
Beep
Rem WaitForInputPicCode
PicCodeKey = WaitKey()
If PicCodeKey <> 13 Then 
	Goto WaitForInputPicCode
End If

'==交易密码==
MoveTo 1215, 700
Call RandDelay(0.5, 1)
LeftClick 1
KeyPress "1", 1
KeyPress "2", 1
KeyPress "3", 1
KeyPress "1", 1
KeyPress "2", 1
KeyPress "3", 1

'==短信验证码==
MoveTo 1400, 810
Call RandDelay(0.5, 1)
LeftClick 1
MoveTo 1024, 640
Call RandDelay(1, 2)
LeftClick 1
MoveTo 1215, 810
LeftClick 1
Rem WaitForInputSMCode
SMCodeKey = WaitKey()
If SMCodeKey <> 13 Then 
	Goto WaitForInputSMCode
End If

'==确认==
MoveTo 1575, 990
Call RandDelay(0.5, 1)
LeftClick 1

'==继续转账==
MoveTo 1727, 721
Call RandDelay(1, 2)
LeftClick 1
