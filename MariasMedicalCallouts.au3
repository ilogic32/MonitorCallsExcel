#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiMenu.au3>
#include <Date.au3>
#include <Array.au3>

; Create application object
Local $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

; Open an existing workbook and return its object identifier.
$sWorkbook = FileOpenDialog("Open Glocal Mind - Interaction Report",@MyDocumentsDir, "Excel (*.xlsx)", 1)
;~ Local $sWorkbook = @ScriptDir & "\Glocal Mind - Interaction Report.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example 1", "Error opening '" & $sWorkbook & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
ConsoleWrite("Workbook '" & $sWorkbook & "' has been opened successfully." & @CRLF & @CRLF & "Creation Date: " & $oWorkbook.BuiltinDocumentProperties("Creation Date").Value)

_Excel_RangeWrite($oWorkbook, 1, _NowDate(), "G6")

; Find the next empty line for Survey Nr
Local $aInput        = _Excel_RangeRead($oWorkbook, 2, "B1:B100") ; Interaction
;~ _ArrayDisplay($aInput)
Local $index = 0
For $a in $aInput
	$index = $index + 1
	ConsoleWrite($a & @CRLF)
	If $a = "" Then Exitloop
Next

GUI($oWorkbook, $index)

MsgBox(0,"Document saved","Document saved " & _NowDate())

_Excel_Close($oExcel)

Exit

Func GUI($oWorkbook, $index)
        ; Create a GUI with various controls.
        Local $hGUI          = GUICreate("Interaction Report " & _NowDate(), 400, 100)
		Local $idInput       = GUICtrlCreateInput("index is " & $index, 10, 8, 185, 25)
		Local $idSend        = GUICtrlCreateButton("Send", 195, 8, 185, 25)
        Local $idAttempt     = GUICtrlCreateButton("Attempt", 10, 70, 85, 25)
		Local $idNotReached  = GUICtrlCreateButton("NotReached", 110, 70, 85, 25)
		Local $idSecretary   = GUICtrlCreateButton("Secretary", 210, 70, 85, 25)
		Local $idInteraction = GUICtrlCreateButton("Interaction", 310, 70, 85, 25)

		Local $sAttemps      = _Excel_RangeRead($oWorkbook, 1, "H6") ; Total Nr of Attempts
		Local $sNotReached   = _Excel_RangeRead($oWorkbook, 1, "J6") ; NotReached
		Local $sSecretary    = _Excel_RangeRead($oWorkbook, 1, "K6") ; Secretary
		Local $sInteraction  = _Excel_RangeRead($oWorkbook, 1, "L6") ; Interaction

		Local $lAttemps      = GUICtrlCreateLabel($sAttemps, 42, 40) ; Total Nr of Attempts label
 		Local $lNotReached   = GUICtrlCreateLabel($sNotReached, 42*3.5, 40) ; NotReached label
		Local $lSecretary    = GUICtrlCreateLabel($sSecretary, 42*6, 40) ; Secretary label
		Local $lInteraction  = GUICtrlCreateLabel($sInteraction, 42*8.2, 40) ; Interaction label


        ; Display the GUI.
        GUISetState(@SW_SHOW, $hGUI)

		If Not IsDeclared("iMsgBoxAnswer") Then Local $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(36,"Reset?","Do you want to reset the call values?")
		Select
			Case $iMsgBoxAnswer = 6 ;Yes
				$sAttemps = 0
				$sNotReached = 0
				$sSecretary = 0
				$sInteraction = 0
				GUICtrlSetData($lAttemps, $sAttemps )
				GUICtrlSetData($lNotReached, $sNotReached)
				GUICtrlSetData($lSecretary, $sSecretary )
				GUICtrlSetData($lInteraction, $sInteraction )
				_Excel_RangeWrite($oWorkbook, 1, $sAttemps, "H6")
				_Excel_RangeWrite($oWorkbook, 1, $sNotReached, "J6")
				_Excel_RangeWrite($oWorkbook, 1, $sSecretary, "K6")
				_Excel_RangeWrite($oWorkbook, 1, $sInteraction, "L6")

			Case $iMsgBoxAnswer = 7 ;No

		EndSelect

        ; Retrieve the handle of the active window.
        Local $hWnd = WinGetHandle($hGUI)

        ; Set the active window as being ontop using the handle returned by WinGetHandle.
        WinSetOnTop($hWnd, "", $WINDOWS_ONTOP)

        ; Loop until the user exits.
        While 1
                Switch GUIGetMsg()
						Case $GUI_EVENT_CLOSE
							ExitLoop
						Case $idAttempt
							$sAttemps = $sAttemps + 1
							GUICtrlSetData($lAttemps, $sAttemps )
							_Excel_RangeWrite($oWorkbook, 1, $sAttemps, "H6")
						Case $idNotReached
							$sNotReached = $sNotReached + 1
							GUICtrlSetData($lNotReached, $sNotReached)
							_Excel_RangeWrite($oWorkbook, 1, $sNotReached, "J6")
						Case $idSecretary
							$sSecretary = $sSecretary + 1
							GUICtrlSetData($lSecretary, $sSecretary )
							_Excel_RangeWrite($oWorkbook, 1, $sSecretary, "K6")
						Case $idInteraction
							$sInteraction = $sInteraction + 1
							GUICtrlSetData($lInteraction, $sInteraction )
							_Excel_RangeWrite($oWorkbook, 1, $sInteraction, "L6")
						case $idSend
							$sinput = GUICtrlRead($idInput)
;~ 							MsgBox(0,"","Send " & $sinput & " to B" & $index)
							_Excel_RangeWrite($oWorkbook, 2, $sinput, "B" & $index)
							_Excel_RangeWrite($oWorkbook, 2, _nowdate(), "A" & $index)
							_Excel_RangeWrite($oWorkbook, 2,"P2203400102", "C" & $index)
							_Excel_RangeWrite($oWorkbook, 2,"CANADA", "D" & $index)
							$index = $index + 1
							GUICtrlSetData($idInput, "" )

                EndSwitch
        WEnd

        ; Delete the previous GUI and all controls.
        GUIDelete($hGUI)
EndFunc   ;==>Example
