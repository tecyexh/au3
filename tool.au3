#AutoIt3Wrapper_UseX64 =n
#include <GuiToolbar.au3>
#include <WinAPI.au3>
#include <Memory.au3>
#include <GuiButton.au3>  
#include <Array.au3>

;_ToolMain()
ConsoleWrite(_IPGIsFileEncrypt(@DesktopDir&"\12.txt"))
;测试main
Func _ToolMain()
   ;MsgBox(0,"test",_IPGGetSDAgentStatus())
   ;_IPGInstAgent("C:\Documents and Settings\VM-XP\桌面\3.59.138.513.exe")
	$apid = _IPGGetAgentPid()
	For $i = 1 To $apid[0]
		ConsoleWrite($apid[$i]&@crlf)
	Next
EndFunc   ;==>_Main

; #FUNCTIONS# ===========================================================================================================
; _ReadFile                 ; -->_ReadFile($sFileName)	读取指定文档，返回文档内容
; _RegDM                    ; -->_RegDM($dll_path)	注册大漠插件，返回大漠插件对象
;_GetLocalMAC   			; -->_GetLocalMAC()	获取本机mac，返回mac地址，无网卡返回0
;_FindWindowEx  			; -->_FindWindowEx($hPWnd,$hCWnd,$sClassName,$sWindowName)
;								该函数获得一个窗口的句柄，该窗口的类名和窗口名与给定的字符串相匹配。
;                 				这个函数查找子窗口，从排在给定的子窗口后面的下一个子窗口开始
;_IsToolBarTextExist		; -->_IsToolBarTextExist($hWnd,$sText)	判断工具栏按钮文本是否存在
;_GetToolBarText				; -->_GetToolBarText($hWnd)	获取工具栏文本
;_IsTrayTextExist			; -->_IsTrayTextExist($sText)	判断托盘文本是否存在
;_FileListEx				; -->_FileListEx($sDir)	查找指定路径下所有文件，返回文件路径,"|"分割文件路径
;_FindFile					; -->_FindFile($sDir,$sFileName)	找文件,返回文件名字符串，"|"分割
;_WriteFile					; -->_WriteFile($sDir,$sCon)	内容写入文件
;_WriteFileLine					; -->_WriteFileLine($sDir,$sCon)	添加一行将内容写入文件尾部
;_FileReadLineToArray		; -->_FileReadLineToArray($sDir)		按每行读取文档，并将写入数组
;_ProcesCmdline				; -->_ProcesCmdline($sPid)		根据pid获取进程cmdline
; =======================================================================================================================

; #IPG_FUNCTIONS# =======================================================================================================
;_IPGLoginConsole				; -->_IPGLoginConsole($sDir,$sIp,$sAccount,$sPasswd)	登录控制台
;_IPGGetAgentVersion            ; -->_IPGGetAgentVersion()		获取客户端版本号
;_IPGGetSDAgentStatus			; -->_IPGGetSDAgentStatus()		获取加密客户端状态
;_IPGInstAgent					; -->_IPGInstAgent($sDir,$bReboot)		安装客户端,安装后重启机器
;_IPGGetAgentPid				; -->_IPGGetAgentPid() 			获取客户端pid
;_IPGIsFileEncrypt				; -->_IPGIsFileEncrypt($sFileName)		判断文件是否加密
; =======================================================================================================================

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:读取指定文档，返回文档内容
; ===============================================================================================================================
Func _ReadFile($sFileName)
   Local $file = FileOpen($sFileName, 0)
   ; 检查以只读打开的文件
   If $file = -1 Then Return 0
   If @error Then Return SetError(@error, @extended, 0)
   Local $sContent = FileRead($file)
   FileClose($file)
   Return $sContent
EndFunc   ;==>_ReadFile

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:注册大漠插件
; ===============================================================================================================================
Func _RegDM($dll_path)
   Local $obj = ObjCreate("dm.dmsoft")
   If Not IsObj($obj) Then
	  RunWait(@ComSpec & ' /c regsvr32 /s ' & FileGetShortName($dll_path), '', @SW_HIDE)
	  $obj = ObjCreate("dm.dmsoft")
   EndIf
   Return $obj
EndFunc   ;==>_RegDM

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:获取本机mac，返回mac地址，无网卡返回0
; ===============================================================================================================================
Func _GetLocalMAC()
    Global Const $wbemFlagReturnImmediately = 0x10
	Global Const $wbemFlagForwardOnly = 0x20
    Local $strComputer, $objWMIService, $colItems, $Output,$sRet
    $colItems = ""
    $strComputer = "localhost"
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled != 0", "WQL", _
            $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
    Local $check2 = '', $check1 = ''
    If IsObj($colItems) Then
        For $objItem In $colItems
            $Output = "MAC地址： " & $objItem.MACAddress & @CRLF
			$sRet = $sRet&$Output
            $Output = ""
            $check1 = $objItem.MACAddress
            If $check1 = $check2 Then
                If $check1 = '' Then Return 0
                ExitLoop
            EndIf
        Next
	 EndIf
	 Return $sRet
EndFunc   ;==>_GetLocalMAC

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:该函数获得一个窗口的句柄，该窗口的类名和窗口名与给定的字符串相匹配。
;                 这个函数查找子窗口，从排在给定的子窗口后面的下一个子窗口开始
; ===============================================================================================================================
Func _FindWindowEx($hPWnd,$hCWnd,$sClassName,$sWindowName)
   	Local $aResult = DllCall("user32.dll", "hwnd", "FindWindowExW","hwnd",$hPWnd,"hwnd",$hCWnd,"wstr", $sClassName, "wstr", $sWindowName)
	If @error Then
	   Return SetError(@error, @extended, 0)
	EndIf
	Return $aResult[0]
EndFunc   ;==>_FindWindowEx

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:判断工具栏按钮文本是否存在
; ===============================================================================================================================
Func _IsToolBarTextExist($hWnd,$sText)
   Local $iButtonCount = _GUICtrlToolbar_ButtonCount($hWnd) ;返回工具栏按钮数
   Local $bRet
   For $i = 0 To $iButtonCount-1
	  $iIndex = _GUICtrlToolbar_IndexToCommand($hWnd,$i)
	  $sButtonText = _GUICtrlToolbar_GetButtonText ($hWnd,$iIndex)
	  If StringInStr($sButtonText,$sText) <> 0 Then
		 $bRet = True
		 ExitLoop
	  ElseIf (($sText == "") And ($i == $iButtonCount-1) )Then
		 $bRet = False
	  EndIf
   Next
   Return $bRet
EndFunc   ;==>_IsToolBarTextExist

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:获取工具栏按钮文本
; ===============================================================================================================================
Func _GetToolBarText($hWnd)
   Local $iButtonCount = _GUICtrlToolbar_ButtonCount($hWnd) ;返回工具栏按钮数
   Local $bRet
   For $i = 0 To $iButtonCount-1
	  $iIndex = _GUICtrlToolbar_IndexToCommand($hWnd,$i)
	  $sButtonText = _GUICtrlToolbar_GetButtonText ($hWnd,$iIndex)
	  $bRet = $bRet&"|"&$sButtonText
   Next
   Return $bRet
EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......:判断托盘文本是否存在
; ===============================================================================================================================
Func _IsTrayTextExist($sText)
   ;获取普通托盘区窗口句柄
   Local $hNTrayWnd = ControlGetHandle("[CLASS:Shell_TrayWnd]","","[CLASS:ToolbarWindow32; INSTANCE:1]")
   ;获取溢出托盘区窗口句柄
   Local $hOTrayWnd = ControlGetHandle("[CLASS:NotifyIconOverflowWindow]","","[CLASS:ToolbarWindow32; INSTANCE:1]")
   Local $bRet
   $bRet = _IsToolBarTextExist($hNTrayWnd,$sText) OR _IsToolBarTextExist($hOTrayWnd,$sText)
   Return $bRet
EndFunc   ;==>_IsTrayTextExist

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 查找指定路径下所有文件，返回文件路径,"|"分割文件路径
; ===============================================================================================================================
Func _FileListEx($sDir)
   If StringInStr(FileGetAttrib($sDir),"D")=0 Then Return SetError(1,0,"")

   Local $oFSO = ObjCreate("Scripting.FileSystemObject")
   Local $objDir
   Local $aDir = StringSplit($sDir, "|", 2)
   Local $iCnt = 0
   Local $sFiles = ""
   Do
	  $objDir = $oFSO.GetFolder($aDir[$iCnt])
	  For $aItem In $objDir.SubFolders
		 ;扩展应用改下这句, 如指定文件夹 If StringInStr($aItem.Name, "XXX") Then
		 $sDir &= "|" & $aItem.Path
		 ;文件夹层数可以通过 StringReplace($aItem.Path, "\", "", 0, 1)的@extended值来判断
	  Next
	  ;如果仅找文件夹,不找文件,$sFiles的语句都不用,最后是 Return $sDir
	  For $aItem In $objDir.Files
		 ;扩展应用改下面这句
		 $sFiles &= $aItem.Path & "|"
		 ;例如要找文件名中包含"kb"(不分大小写),改为: if StringInStr($aItem.Name, "kb") Then $sFiles &= $aItem.Path & "|"
		 ;其他应用请参照上例修改: $aItem.XXX
		 ;Attributes        设置或返回文件或文件夹的属性
		 ;DateCreated   返回指定的文件或文件夹的创建日期和时间。只读
		 ;DateLastAccessed 返回指定的文件或文件夹的上次访问日期(和时间)。只读
		 ;DateLastModified 返回指定的文件或文件夹的上次修改日期和时间。只读
		 ;ShortName   返回按照早期8.3文件命名约定转换的短文件名
		 ;ShortPath   返回按照早期8.3命名约定转换的短路径名
		 ;Size    对于文件返回指定文件的字节数；对于文件夹，返回文件夹所有的文件夹和子文件夹的字节数
		 ;Type    返回文件或文件夹的类型信息
	  Next
	  $iCnt += 1
	  If UBound($aDir) <= $iCnt Then $aDir = StringSplit($sDir, "|", 2)
   Until UBound($aDir) <= $iCnt
   If $sFiles Then $sFiles = StringTrimRight($sFiles, 1);去掉最右边"|"
   Return $sFiles
EndFunc   ;==>_FileListEx

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 找文件,返回文件名字符串，"|"分割
; ===============================================================================================================================
Func _FindFile($sDir,$sFileName)
   Local $sFiles = _FileListEx($sDir)
   Local $aFiles = StringSplit($sFiles, "|")
   Local $sFileNametemp
   Local $iFileLoca
   Local $iFileLen
   Local $sFileRet = ""
   For $iLen = 1 To $aFiles[0] Step 1
	  ;获取文件名(路径)长度
	  $iFileLen = StringLen($aFiles[$iLen])
	  ;获取文件名初始位置
	  $iFileLoca = StringInStr($aFiles[$iLen],"\",0,-1)
	  ;获取文件名
	  $sFileNametemp = StringRight($aFiles[$iLen],$iFileLen-$iFileLoca)
	  ;判断文件名是否存在
	  If(StringInStr($sFileNametemp,$sFileName)) Then $sFileRet &= $aFiles[$iLen] & "|"
   Next
   If $sFileRet Then $sFileRet = StringTrimRight($sFileRet, 1)
   Return $sFileRet
EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 内容写入文件
; ===============================================================================================================================
Func _WriteFile($sDir,$sCon)
   Local $file = FileOpen($sDir, 1)

   ; 检查文件是否以写入模式打开
   If $file = -1 Then
	  MsgBox(0, "错误", "无法打开文件.")
	  Exit
   EndIf

   FileWrite($file, $sCon)
   FileClose($file)
EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 添加一行将内容写入文件尾部
; ===============================================================================================================================
Func _WriteFileLine($sDir,$sCon)
   Local $file = FileOpen($sDir, 1)

   ; 检查文件是否以写入模式打开
   If $file = -1 Then
	  MsgBox(0, "错误", "无法打开文件.")
	  Exit
   EndIf

   FileWriteLine($file, $sCon)
   FileClose($file)
EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 按每行读取文档，并将写入数组
; ===============================================================================================================================
Func _FileReadLineToArray($sDir)
   Local $sText,$aTextArray
   Local $file = FileOpen($sDir, 0)
   ; 检查以只读打开的文件
   If $file = -1 Then
	  MsgBox(0, "错误", "无法打开文件.")
	  Exit
   EndIf
   ; 读入文本行直到文件结束(EOF)
   While 1
	  Local $line = FileReadLine($file)
	  If @error = -1 Then ExitLoop
	  $sText &= $line&"|"
   Wend
   FileClose($file)
   $sText = StringLeft($sText,StringLen($sText)-1)
   $aTextArray = StringSplit($sText,"|")
   Return $aTextArray
EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 复制制定文件N遍
; ===============================================================================================================================
Func _CopyFile($sFileName, $sCopyPath, $iTimes = 1)
	If FileExists($sFileName) Then
		Local $sName,$sType,$atemp
		$sName = StringRight($sFileName, StringLen($sFileName)-StringInStr($sFileName, "\" , 0, -1))
		$atemp = StringSplit($sName , ".")
		If $atemp[0] == 2 Then
			$sName = $atemp[1]
			$sType = $atemp[2]
		EndIf
		DirCreate($sCopyPath)
		For $i = 1 To $iTimes
			FileCopy($sFileName, $sCopyPath&"\"&$sName&"_"&$i&"."&$sType, 0)
		Next
	Else
		Return 0
	EndIf	
EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 修改保存word文档
; ===============================================================================================================================
Func _ModifyWord($sPath)
	For $i = 1 To 100
		Sleep(10000)
		ShellExecute($sPath&"modify_"&$i&".doc")
		Sleep(2000)
		WinWaitActive("modify_"&$i&".doc [兼容模式] - Microsoft Word")
		Sleep(500)
		Send($i)
		Send("{Enter}")
		Sleep(500)
		Send("!{F4}")
		WinWaitActive("[CLASS:#32770]")
		ControlClick("[CLASS:#32770]","","[CLASS:Button; INSTANCE:1]")
	Next
EndFunc
	
; #FUNCTION# ====================================================================================================================
; Author ........: yexh
; Modified.......: 2016.06.21
; Function.......: 根据pid获取进程cmdline
; ===============================================================================================================================
Func _ProcesCmdline($sPid)
	Local $cmdpath
	$strComputer = "."
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("Select * FROM Win32_Process Where ProcessId = "&$sPid)
	
	For $objItem In $colItems
		$cmdpath = $cmdpath& $objItem.CommandLine &"|"
		;$exepath = $objItem.ExecutablePath
	Next
	$cmdpath = StringLeft($cmdpath, StringLen($cmdpath)-1)
	Return $cmdpath
EndFunc

; #IPG_FUNCTION# ================================================================================================================
; Author ........: yexh
; Modified.......: 2016.04.28
; Function.......: 登录控制台
; ===============================================================================================================================
Func _IPGLoginConsole($sDir,$sIp,$sAccount,$sPasswd)
   Run($sDir)
   WinWaitActive("登录")
   ;输入IP，账号，密码。登录
   ControlSetText("登录","","[CLASS:Edit;INSTANCE:1]",$sIp)
   ControlSetText("登录","","[CLASS:Edit;INSTANCE:2]",$sAccount)
   ControlSetText("登录","","[CLASS:Edit;INSTANCE:3]",$sPasswd)
   ControlClick ( "登录", "确定","[CLASS:Button; INSTANCE:1]")
   Sleep(2000)
   If WinExists("IP-guard V3 控制台","产品剩余") Then
	  ControlClick("IP-guard V3 控制台", "确定", "[CLASS:Button; INSTANCE:1]")
	  Sleep(2000)
   EndIf
   If WinExists("IP-guard V3 控制台","发现新的控制台") Then
	  ControlClick("IP-guard V3 控制台", "", "[CLASS:Button; INSTANCE:2]")
	  Sleep(2000)
   EndIf
   WinWaitActive("IP-guard V3 控制台")
   Sleep(10000)
EndFunc

; #IPG_FUNCTION# ================================================================================================================
; Author ........: yexh
; Modified.......:
; Function.......: 获取客户端版本
; ===============================================================================================================================
Func _IPGGetAgentVersion()
   Local $sVersion = FileGetVersion(@SystemDir&"\winoav3.dll")
   Return $sVersion
EndFunc

; #IPG_FUNCTION# ================================================================================================================
; Author ........: yexh
; Modified.......: 2016.04.28
; Function.......: 获取加密客户端状态
; ===============================================================================================================================
Func _IPGGetSDAgentStatus()
   Local $sRet,$hNTrayWnd,$hOTrayWnd,$sText
   ;获取普通托盘区窗口句柄
   $hNTrayWnd = ControlGetHandle("[CLASS:Shell_TrayWnd]","","[CLASS:ToolbarWindow32; INSTANCE:1]")
   ;获取溢出托盘区窗口句柄
   $hOTrayWnd = ControlGetHandle("[CLASS:NotifyIconOverflowWindow]","","[CLASS:ToolbarWindow32; INSTANCE:1]")
   $sText = _GetToolBarText($hNTrayWnd)&_GetToolBarText($hOTrayWnd)
   If StringInStr($sText,"加密系统正在运行中") Then
	  $sRet = "自动加解密"
   ElseIf	StringInStr($sText,"加密系统未启动，可双击图标登入加密系统") Then
	  $sRet = "注销加密系统"
   ElseIf	StringInStr($sText,"加密系统未启动") Then
	  $sRet = "禁用加密系统"
   ElseIf	StringInStr($sText,"加密系统正在运行中(只读模式)") Then
	  $sRet = "只读模式"
   ElseIf	StringInStr($sText,"加密系统处于离线状态，加密功能暂停") Then
	  $sRet = "离线"
   ElseIf	StringInStr($sText,"加密系统已进入备用模式") Then
	  $sRet = "备用模式"
   ElseIf	StringInStr($sText,"加密系统处于离线授权状态") Then
	  $sRet = "离线授权登入"
   ElseIf	StringInStr($sText,"加密系统处于离线状态，加密功能暂停，可双击图标登入离线授权") Then
	  $sRet = "离线授权登出"
   Else
	  $sRet = "加密客户端未启动"
   EndIf
   Return $sRet
EndFunc

; #IPG_FUNCTION# ================================================================================================================
; Author ........: yexh
; Modified.......: 2016.06.20
; Function.......: 安装客户端
; ===============================================================================================================================
Func _IPGInstAgent($sDir,$bReboot = True)
	Local $sAgentType,$bAgent,$hPWnd,$hCWnd,$hGWnd
	$bAgent = False
	$sAgentType = StringRight($sDir, StringLen($sDir)-StringInStr($sDir, ".", 0, -1))
	If StringInStr($sAgentType, "exe") Then
		Run($sDir)
		$bAgent = True
	ElseIf StringInStr($sAgentType, "e32") Then
		Run(@ComSpec & ' /c "' & $sDir & '" -gui', "", @SW_HIDE)
		$bAgent = True
	EndIf
	If $bAgent Then
		WinWaitActive("安装 - 客户端")
		While WinExists("安装 - 客户端")
			If ControlCommand("安装 - 客户端", "", "[CLASS:Button; INSTANCE:2]", "IsEnabled") Then 
				If (StringInStr(ControlGetText("安装 - 客户端", "", "[CLASS:Button; INSTANCE:2]"), "完成") And (Not $bReboot)) Then
					$hPWnd = _WinAPI_FindWindow("#32770" , "安装 - 客户端")
					$hCWnd = _FindWindowEx($hPWnd,0,"#32770","")
					For $i = 1 To 3
						$hCWnd = _WinAPI_GetWindow($hCWnd, 2)
					Next
					$hGWnd = _FindWindowEx($hCWnd,0,"Button","不，稍后重新启动计算机。")	
					_GUICtrlButton_Click($hGWnd)
				EndIf
				ControlClick("安装 - 客户端", "", "[CLASS:Button; INSTANCE:2]")
			EndIf
			Sleep(500)
		WEnd
	EndIf
	Return $bAgent
EndFunc

; #IPG_FUNCTION# ================================================================================================================
; Author ........: yexh
; Modified.......: 2016.06.20
; Function.......: 获取客户端进程pid
; ===============================================================================================================================
Func _IPGGetAgentPid()
	Local $aTmp,$aPid[1],$aAgentMod[4] = ["WINRDLV3.EXE", "OSgwAgent.exe", "ONacAgent.exe", "sdhelper2.exe"],$smod,$scmdline,$icount
	For $i = 0 To 3
		If StringInStr($aAgentMod[$i], "WINRDLV3.EXE", 0) Then
			Local $aTmp = ProcessList($aAgentMod[$i])
			For $j = 1 To $aTmp [0][0]
				$scmdline = _ProcesCmdline($aTmp[$j][1])
				$smod = StringMid($scmdline, StringInStr($scmdline," ", 0)+1, StringInStr($scmdline,",", 0)-StringInStr($scmdline," ", 0)-1)
				_ArrayAdd($aPid, $aTmp [$j][1]&","&$smod)
				$icount += 1
			Next
		Else
			If ProcessExists($aAgentMod[$i]) Then
				_ArrayAdd($aPid, ProcessExists($aAgentMod[$i])&","&$aAgentMod[$i])
				$icount += 1
			EndIf
		EndIf
	Next
	$aPid[0] = $icount
	Return $aPid
EndFunc

; #IPG_FUNCTION# ================================================================================================================
; Author ........: yexh
; Modified.......: 2016.10.17
; Function.......: 判断文件是否加密
; ===============================================================================================================================
Func _IPGIsFileEncrypt($sFileName)
	Local $sContent,$sTsdHead,$bRet
	$bRet = False
	$sContent = _ReadFile($sFileName)
	$sTsdHead = StringLeft($sContent, 16)
	If (StringCompare($sTsdHead, "%TSD-Header-###%") == 0) Then $bRet = True
	Return $bRet
EndFunc