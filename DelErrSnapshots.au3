#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
Opt("TrayIconDebug",1)
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>

Global $sVmPath,$sVmrun,$sResult
Local $sIP,$sUser,$sPasswd
$sVmPath = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\vmware.exe", "Path")
$sVmrun = "C:\QTPJM\vmrun.exe"
$sResult = "C:\QTPJM\snapshots.txt"  ;保存获取到的快照内容

;MsgBox(1,"1",$sVmrun)

#Region ### START GUI section ### Form=
GUICreate("DelSnapshots") ; will create a dialog box that when displayed is centered
$Com = GUICtrlCreateCombo("[standard] WinXP_liyj_005/WinXP_liyj_005.vmx", 10, 10,370) ; create first item
;增加item，此处即虚拟机的vmx位置
GUICtrlSetData(-1, "[standard] Win7 x64_QTP_001/Windows 7 x64.vmx|[HDD225_3] win8x64_qtp_002/win8x64_qtp_002.vmx|[standard] win8_x64_QTP_003-1/win8_x64_QTP_003-1.vmx|[HDD223_2] Win7x64_Qtp2/Win7x64_Qtp2.vmx|预留别选", "item1") 
$Btn1 = GUICtrlCreateButton("ShowSnapshots",15,40)
$Btn2 = GUICtrlCreateButton("DelSnapshots",300,40)
$Lb1 = GUICtrlCreateEdit("",10,75,380,300,$ES_READONLY + $WS_VSCROLL,$WS_EX_CLIENTEDGE)
GUISetState(@SW_SHOW) ; will display an empty dialog box
#EndRegion ### END GUI section ###

; Run the GUI until the dialog is closed
While 1
   $msg = GUIGetMsg()
   Select
	  Case $msg = $GUI_EVENT_CLOSE
		 ExitLoop
	  Case $msg = $Btn1
		 $vmroute = GUICtrlRead($Com)
		 ShowVmSnaps($vmroute)
	  Case $msg = $Btn2
		 DelVmSnaps($vmroute)
   EndSelect
WEnd

Func ShowVmSnaps($vm)
   If StringInStr($vm,"standard") Then
	  $sIP = "ws-shared -h 192.168.2.13"
	  $sUser = "administrator"
	  $sPasswd = "123"
   Else
	  $sIP = "vc -h 192.168.2.1"
	  $sUser = "development\moyz"
	  $sPasswd = "123456"
   EndIf
   RunWait(@ComSpec & " /c " & $sVmrun &" -T "& $sIP &" -u "& $sUser &" -p "& $sPasswd &' listSnapshots "'& $vm &'" > '& $sResult, "", @SW_HIDE)
   ;Sleep(3000)
   $f = FileOpen($sResult)
   $snap = FileRead($f)
   FileClose($f)
   GUICtrlSetData($Lb1,$snap)
EndFunc

Func DelVmSnaps($vm)
   If StringInStr($vm,"standard") Then
	  $sIP = "ws-shared -h 192.168.2.13"
	  $sUser = "administrator"
	  $sPasswd = "123"
   Else
	  $sIP = "vc -h 192.168.2.1"
	  $sUser = "development\moyz"
	  $sPasswd = "123456"
   EndIf
   $f=FileOpen($sResult)
   $ln=0
   While 1    ;获取行数，即快照个数
	  If @error=-1 Then ExitLoop
	  FileReadLine($f)
	  $ln=$ln+1
   WEnd
   FileClose($f)
   $f=FileOpen($sResult)
   While $ln <>1    ;从尾倒回来循环删除 "Err_" 开头的快照
	  $sn=FileReadLine($f,$ln)
	  
	  If StringInStr($sn,"Err_") Then
		 RunWait(@ComSpec & " /c " & $sVmrun &" -T "& $sIP &" -u "& $sUser &" -p "& $sPasswd &' deleteSnapshot "'& $vm &'" '&$sn, "", @SW_HIDE)
		 Sleep(2000)
	  EndIf
	  $ln=$ln-1
   WEnd
   FileClose($f)
   ShowVmSnaps($vm)
   MsgBox(0,"Done","OK")
	  
   
EndFunc


