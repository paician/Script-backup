#include <Date.au3>
#include <File.au3>
#include <String.au3>

$date = _Nowdate()
$Date = @YEAR & "." & @MON & "." & @MDAY
$chc =1

Global $current = envget("current")

Func ExitFunction()
  Exit
EndFunc

Switch $chc
case 1
    $hc = ipaddress

Local $sStringip1 = StringInstr (@IPADDRESS1, $hc , 0)
Local $sStringip2 = StringInstr (@IPADDRESS4, $hc , 0)
EndSwitch
If $sStringip1 = 0 Then

    $fho = FileOpen("Path" & $date & ".csv", $FO_APPEND)
    FileWriteLine($fho,@YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & "," & @IPAddress1 & "," & @IPAddress2 & "," & @IPAddress3 & "," & @IPAddress4 & "," & @UserName & "," & @ComputerName & "," & @CRLF)
    FileClose($fho)
    Runwait('"' & $current & @ScriptDir & '' & '"')
    ExitFunction()
EndIf

$fhtwo = FileOpen("Path" & $date & ".csv", $FO_APPEND)
FileWriteLine($fhtwo,@YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & "," & @IPAddress1 & "," & @IPAddress2 & "," & @IPAddress3 & "," & @IPAddress4 & "," & @UserName & "," & @ComputerName & "," & @CRLF)
FileClose($fhtwo)
Runwait('"' & $current & @ScriptDir & '' & '"')
