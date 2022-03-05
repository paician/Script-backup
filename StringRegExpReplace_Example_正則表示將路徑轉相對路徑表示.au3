#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
Dim $Msg
$hGUI = GUICreate("Window", 500, 400, -1, -1)
$hEdit = GUICtrlCreateEdit("Edit", 173, 126, 220, 144)
$hButton = GUICtrlCreateButton("Button", 175, 50, 93, 29)



GUISetState()

While 1
	$hMsg = GUIGetMsg()
	Switch $hMsg
		Case $GUI_EVENT_CLOSE
			Exit
	    Case $hButton

   $Str = "Command\AutoIT\ChooseFileFolder\test"
   $ret = StringSplit ( $Str, '\', 2)

   For $i = 0 to UBound ($ret)-2
           $Msg = $Msg & $ret[$i]
		   ;MsgBox(0,"123",StringRegExpReplace($ret[$i], ".*(?i)[A-Z]", $msg ))

   GUICtrlSetData($hEdit,StringRegExpReplace($Msg, ".*(?i)[A-Z][^0-9]", "..\\"),1)
   Next
	EndSwitch
WEnd
