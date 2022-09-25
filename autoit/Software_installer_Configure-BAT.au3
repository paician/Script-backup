#include <file.au3>
#include <array.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <WinAPIIcons.au3>

#pragma compile(Out, Edit_Bat_Tool.exe)
#pragma compile(Icon, 'ico\Edit_bat.ico'); 設定 icon 路徑
#pragma compile(x64, True)
#pragma compile(UPX, True)
#pragma compile(ProductName, Edit Bat Tool For )
#pragma compile(ProductVersion, 1.1)
#pragma compile(FileVersion, 1.1)
#pragma compile(FileDescription, Edit Bat Tool For  BAT快速修改新增程式))
#pragma compile(ExecLevel, highestavailable)
#pragma compile(LegalCopyright, c )
#pragma compile(LegalTrademarks, '"by ')
#pragma compile(CompanyName, '')

$user = @SystemDir & "\user32.dll"
$shell = @SystemDir & "\shell32.dll"
$icoD = -4
Global $sVersion = "V1.1"
Global $iLogOnFlag = 0
Global $sParameters = ""
Global $Msg
_CheckUser()
Func _CheckUser()

Local $sMsg = ""

Switch @error
Case 0
   ;簡單用當前使用者帳號來區別才可使用
	  Switch @UserName
		  Case "d.duke" To "d.duke"
			  $sMsg = "OK"
			  Sleep(500)

		  Case Else
			  $sMsg = MsgBox(262144,"提示","未經許可使用者無法使用此功能")
			  Exit
	  EndSwitch
  Case Else
	  $sMsg = "Something went horribly wrong."
EndSwitch
MsgBox($MB_SYSTEMMODAL, "提示", $sMsg)
Configure()
EndFunc
;$lnk_file = ""

Func Configure()
$Form1 = GUICreate("快捷修改bat工具" & $sVersion, 742, 490, 717, 158)
$Input1 = GUICtrlCreateInput("", 150, 40, 256, 21);檔案名稱輸入框
GUICtrlSetTip($Input1, "輸入要儲存的bat檔名稱（不含.bat）")
$btn1 = GUICtrlCreateButton ( "(Browse)", 128, 80, 20, 20,$BS_ICON);目標程式圖示
GUICtrlSetTip($btn1, "請選擇目標程式")
GUICtrlSetImage($btn1, $shell, $icoD, 0)
$Input2 = GUICtrlCreateInput("", 150, 80, 256, 21) ;目標程式框
GUICtrlSetTip($Input2, "選擇要安裝的程式")
$btnsaved = GUICtrlCreateButton ( "(Browse)", 152, 120, 20, 20,$BS_ICON);要放置BAT路徑資料夾圖示按鈕
GUICtrlSetTip($btnsaved, "請選擇編輯存放位置")
GUICtrlSetImage($btnsaved, $shell, $icoD, 0)

$OK = GUICtrlCreateButton("新BAT存檔", 432, 79, 75, 25, $WS_GROUP);新BAT存檔按鈕
GUICtrlSetTip($OK, "確認內容無誤後會提示是否要覆蓋，如全新檔案則提示存檔成功")
$hInput3 = GUICtrlCreateInput("", 153, 160, 250, 18);SHOW出詳細路徑輸入框
GUICtrlSetTip(-1, "這裡會秀出上面的欄位")
$btn2 = GUICtrlCreateButton("生成code", 437, 32, 56, 34);要選定檔案名稱、目標程式、詳細路徑才會show預設bat程式碼於下方編輯
GUICtrlSetTip($btn2, "按此按鈕會生成預設程式碼在下方文字框裡")
$btn_check = GUICtrlCreateButton("檢查存檔是否存在", 589, 17, 110, 26);純粹檢查是否與檔案名稱相衝突bat檔
GUICtrlSetTip($btn_check, "這裡只是檢查當前是否有同名檔案存在")
$hButton4 = GUICtrlCreateButton("選要編輯的bat檔", 218, 215, 110, 30);選要編輯的bat檔按鈕
GUICtrlSetTip(-1, "按我選擇你要編輯的bat檔")
$hButton5 = GUICtrlCreateButton("舊存檔", 418, 217, 70, 30);舊存檔按鈕
$checkbat = GUICtrlCreateButton("檢查BAT執行語法",528,380,130,26)
GUICtrlSetTip($checkbat, "檢查編輯區相對路徑是否有問題--請注意，這個適用於新編輯的，不適合用於舊存檔區")
$Label1 = GUICtrlCreateLabel("檔案名稱", 8, 40, 99, 27)
$Label2 = GUICtrlCreateLabel("目標程式", 8, 80, 65, 27)
$Label3 = GUICtrlCreateLabel("要放置的BAT路徑資料夾", 8, 120, 132, 19)
;_ClearAndSetInputs($Input2)
$hEdit = GUICtrlCreateEdit("", 183, 273, 333, 151);新舊BAT存檔共用顯示編輯框
GUICtrlSetTip($hEdit, "按生成code後這裡會顯示出來，你可以在這文字框中編輯任何程式行，然後再去按新bat存檔按鈕")

$hLabel5 = GUICtrlCreateLabel("詳細路徑：", 8, 158, 102, 21)
GUICtrlSetTip(-1, "請確認路徑是否正確，")

$hGroup = GUICtrlCreateGroup("編輯區", 181, 197, 335, 60)
$hGroup2 = GUICtrlCreateGroup("新BAT建立區", 4, 5, 723, 186)
$hLabel5 = GUICtrlCreateLabel("編輯框為新bat和編輯舊bat共用區" & @CRLF & _
							  "但要注意的是，新建立要選定框內的按鈕", 528, 275, 196, 106)
$hInput4 = GUICtrlCreateInput("", 183, 430, 333, 19)
$hLabel6 = GUICtrlCreateLabel("相對路徑輸出", 33, 430, 83, 19)
GUICtrlSetTip(-1, "這裡會隨著點此產生選定完自動帶出來")
$hButton7 = GUICtrlCreateButton("產生", 123, 430, 53, 19)
GUICtrlSetTip(-1, "點此產生")
$hInput5 = GUICtrlCreateInput("", 183, 460, 333, 19)
GUICtrlSetTip($hInput5, "這裡會隨著產生按鈕自動帶資料,但要自行更改相對路徑位置，目前沒有另外寫出另一判定方法")
$runtest = GUICtrlCreateButton("測試", 528, 460, 53, 19)
GUICtrlSetTip(-1, "點此測試路徑是否能直接開啟")

GUISetState(@SW_SHOW)
$title = "Choose .exe or .msi file"

Local $nMsg = 0


While 1
$nMsg = GUIGetMsg()
Switch $nMsg
   Case $GUI_EVENT_CLOSE
		 ExitLoop

   Case $OK;新bat存檔觸發
	   If GUICtrlRead($Input1) And $Input2 And $hEdit = "" Then
		 MsgBox(0, "", "請填入上方欄位區")
   Else

	  If FileExists(GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat") Then
		 $check_file = MsgBox (4, "確認視窗", "是否要覆蓋？")
		 If $check_file = 6 Then
		 ;MsgBox(4096,"Prompt",GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat")

		 $fh = FileOpen(GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat", 2 + 512)
		 FileWriteLine($fh,GUICtrlRead($hEdit))
		 FileClose($fh)
		 MsgBox(4096,"Prompt",GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat" & "已覆蓋")
	  ElseIf $check_file = 7 Then
		 MsgBox(4096,"Prompt","取消存檔")
	  EndIf
      Else;判斷未有同名bat情況下存檔
		 $fh2 = FileOpen(GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat", 2 + 512)
		 FileWriteLine($fh2,GUICtrlRead($hEdit))
		 FileClose($fh2)
		 MsgBox(4096,"確認存檔","已存檔成功" & GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat")
	  EndIf
   EndIf
   Case	$hButton5;存檔
	  MsgBox(64,"Prompt","舊BAT存檔：" & GUICtrlRead($hInput3))
	  $fh = FileOpen(GUICtrlRead($hInput3), 2 + 512);開啟BAT
	  FileWriteLine($fh,GUICtrlRead($hEdit));寫入行
	  Sleep(100)
	  FileClose($fh);關閉檔案
   Case $btn1;這裡是選擇圖示icon事件
		 $trag_file = FileOpenDialog($title, @ScriptDir & "\..\", "Applications/Batch (*.exe; *.msi; *.bat; *.rar; *.zip; *.7z)")
		 GUICtrlSetData ( $Input2, $trag_file)
		 If @error Then
			 MsgBox(0,"","No file was chosen")
		  EndIf
	Case $btnsaved;這裡是選擇圖示icon事件
	  $sour_file = FileSelectFolder("請選定要放bat的路徑", @ScriptDir & "\Software", 2)
		 GUICtrlSetData ( $hInput3, $sour_file & "\")

	  If @error Then
		  MsgBox(0,"","No file was chosen")
	   EndIf
	Case $hButton7;點此產生
		 GUICtrlSetData($hInput4,"")
		 GUICtrlSetData($hInput4,"""%~dp0",1)
		 Local $string = GUICtrlRead($Input2)
		 Local $findText = "D:\TEST\DES\";"\\server path\Tools\xxx\"
		 Local $replaceText = ""
		 Local $newString = StringReplace($string,WildCardFindText($findText,$string),$replaceText)


		 Local $Str = StringReplace($string,WildCardFindText($findText,$string),$replaceText)
		 Local $ret = StringSplit ( $Str, '\', 2)
		 Local $string2 = GUICtrlRead($hInput3)
		 Local $replaceText = ""

		 Local $findText2 = "D:\TEST\DES";"\\server path\Tools\xxx\"
		 Local $newString2 = StringReplace($string2,WildCardFindText2($findText2,$string2),$replaceText)

		 Local $Str2 = StringReplace($string2,WildCardFindText2($findText2,$string2),$replaceText)
		 Local $ret2 = StringSplit ( $Str2, '\', 2)
		 $Msg = ""
		 For $i = 0 to UBound ($ret2)-2
				  $Msg = $Msg & $ret2[$i]
		 ;MsgBox(0,"123",StringRegExpReplace($Msg, ".*(?i)[A-Z]", $msg ))
		 ;MsgBox(0,"str2",$Str2)
		 ;$path = StringRegExpReplace($Msg, ".*(?i)[A-Z][^0-9]", "..\\")
		 GUICtrlSetData($hInput4,StringRegExpReplace($Msg, ".*(?i)[A-Z].*", "..\\"),1)
		 GUICtrlSetData($hInput5,StringRegExpReplace($Msg, ".*(?i)[A-Z].*", "..\\"),1)
		 Next
		 GUICtrlSetData($hInput4,$newString & """",1)
		 GUICtrlSetData($hInput5,$newString & """",1)
   Case $btn2;生成編輯框裡的code
		 GUICtrlSetData($hEdit,"");
		 Local $string = GUICtrlRead($Input2)
		 Local $findText = "D:\TEST\DES";"\\server path\Tools\xxx\"
		 Local $replaceText = ""
		 Local $newString = StringReplace($string,WildCardFindText($findText,$string),$replaceText)


		 Local $Str = StringReplace($string,WildCardFindText($findText,$string),$replaceText)
		 Local $ret = StringSplit ( $Str, '\', 2)

		 Local $string2 = GUICtrlRead($hInput3)
		 Local $findText2 = "D:\TEST\DES";"\\server path\Tools\xxx\"
		 Local $newString2 = StringReplace($string2,WildCardFindText2($findText2,$string2),$replaceText)


		 Local $Str2 = StringReplace($string2,WildCardFindText2($findText2,$string2),$replaceText)
		 Local $ret2 = StringSplit ( $Str2, '\', 2)

		 $Msg = ""
		 GUICtrlSetData($hEdit,"@echo off" & @CRLF & "echo 切換cmd語系中.." & @CRLF & _
"chcp 437" & @CRLF & "echo Installing...please wait a program running" & @CRLF & _
"""" & "%~dp0",1)

		 For $i = 0 to UBound ($ret2)-2
				  $Msg = $Msg & $ret2[$i]
		 GUICtrlSetData($hEdit,StringRegExpReplace($Msg, ".*(?i)[A-Z].*", "..\\"),1)
		 Next
		 If GUICtrlRead($Input2) = "" Then
		   MsgBox(0, "", "請填入目標程式路徑")
   Else
		 GUICtrlSetData($hEdit,$newString & """" & @CRLF & "echo Install has been done" & @CRLF & _
"timeout /t 5" & @CRLF & "exit",1)
   EndIf
   Case $btn_check;檢查存檔是否存在
	  If GUICtrlRead($hInput3) = "" Then
		MsgBox(0, "", "未填入/選擇目標bat路徑")
	  ElseIf FileExists(GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat") Then
		 MsgBox(4096,"Check FileExists","已有與檔案名稱相同檔案")

	  Else
		 MsgBox(4096,"Check FileExists","沒有與檔案名稱相同檔案")
	  EndIf
   Case $hButton4;選要編輯的bat檔
		 $edit_file = FileOpenDialog($title, @ScriptDir & "\Software", "Applications/Batch (*.bat)")
		 GUICtrlSetData ( $hInput3, $edit_file)
		 GUICtrlSetData ( $Input2, "")
		 GUICtrlSetData ( $Input1, "")
		 GUICtrlSetData ( $hInput4, "")
		 GUICtrlSetData ( $hInput5, "")
		 GUICtrlSetData($hEdit,FileRead($edit_file))
		 If @error Then
			 MsgBox(0,"","No file was chosen")
		  EndIf
   Case $checkbat;檢查BAT執行語法
		 $sText = GUICtrlRead($hEdit)
		 $aText = StringSplit($sText, @CRLF, 1)
		 For $i =5 To $aText[0]
			 $replace = StringReplace($aText[$i],"%~dp0",GUICtrlRead($hInput3) & "\")

			 If FileExists(GUICtrlRead($hInput3) & "\" & GUICtrlRead($Input1) & ".bat") Then
				  MsgBox(4096, "檢查相對路徑", "bat檔存在，接著會測試是否能執行第五行code，路徑為:" & $replace)
				  Runwait($replace)
  			 Else
				MsgBox(4096,"",GUICtrlRead($newString) & "NONO")
			 EndIf
			 ExitLoop
		 Next
		 Case $runtest

			   MsgBox(64,"檢查路徑", """" & @ScriptDir & GUICtrlRead($hInput5))
			   ShellExecuteWait("""" & @ScriptDir & GUICtrlRead($hInput5))
EndSwitch
WEnd

EndFunc
;正則處理
Func WildCardFindText($findText, $string)
If StringInStr($findText,"*") > 0 Then
  $WildCardPosition = StringInStr($findText,"*")
  If $WildCardPosition = StringLen($findText) Then
   $findText = StringLeft($findText,(StringLen($findText)-1))
   $findText = StringRight($string,(StringLen($string)-StringInStr($string,$findText))+1)
  Else
   $findText = StringRight($findText,(StringLen($findText)-1))
   $findText = StringLeft($string,(StringLen($string) - (StringLen($string)-StringInStr($string,$findText))))
  EndIf
EndIf
Return $findText
EndFunc
Func WildCardFindText2($findText2, $string2)
If StringInStr($findText2,"*") > 0 Then
  $WildCardPosition = StringInStr($findText2,"*")
  If $WildCardPosition = StringLen($findText2) Then
   $findText = StringLeft($findText2,(StringLen($findText2)-1))
   $findText = StringRight($string2,(StringLen($string2)-StringInStr($string2,$findText2))+1)
  Else
   $findText = StringRight($findText2,(StringLen($findText2)-1))
   $findText = StringLeft($string2,(StringLen($string2) - (StringLen($string2)-StringInStr($string2,$findText2))))
  EndIf
EndIf
Return $findText2
EndFunc
