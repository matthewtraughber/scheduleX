;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; IMPORTS
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
#include <IE.au3>
#include <Excel.au3>
#include <Date.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; GLOBALS
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Global $iHours = 0, $iMins = 0, $iSecs = 0
Global $xlsDocSched
Global $xlsDocEE
Global $assignEE
Global $fromDate
Global $toDate
Global $verifyDate = False
Global $iResponse

; Binds ESC as universal stop/exit of program
HotKeySet("{Esc}", "EscScheduleX")

; Sets GUI to read button input
Opt("GUIOnEventMode", 1)

; Calls main function
_Main()

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Main function
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func _Main()
	$Logo_jpg = @TempDir & "\scheduleX_LOGO.jpg"
	FileInstall("img\scheduleX.jpg", $Logo_jpg)

	; Creates GUI, sets buttons and styles
	#Region ### START Koda GUI section ### Form=
		$Form1 = GUICreate("<schedule> X v2.0  � MET", 315, 160)
			GUISetFont(6, 400, 0, "Terminal")
			GUISetBkColor(0xFFFFFF)

		$Button1 = GUICtrlCreateButton("Import/Assign/Generate", 5, 125, 165, 30, $WS_GROUP)
			GUICtrlSetFont(-1, 10, 800, 0, "Consolas")
			GUICtrlSetOnEvent($Button1, "ImportStartEE")

		$Button2 = GUICtrlCreateButton("Import ONLY", 175, 125, 90, 30, $WS_GROUP)
			GUICtrlSetFont(-1, 10, 800, 0, "Consolas")
			GUICtrlSetOnEvent($Button2, "ImportStartNoEE")

		$Button3 = GUICtrlCreateButton("Exit", 270, 125, 40, 30, $WS_GROUP)
			GUICtrlSetFont(-1, 10, 800, 0, "Consolas")
			GUICtrlSetOnEvent($Button3, "ExitScheduleX")

		GUISetOnEvent( $GUI_EVENT_CLOSE, "ExitScheduleX")
		$Pic1 = GUICtrlCreatePic($Logo_jpg, 10, 0, 300, 123, BitOR($SS_NOTIFY,$WS_GROUP,$WS_CLIPSIBLINGS))

		; Displays GUI
		GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###
EndFunc

; Initiates wait for user input
$Var = False
Call ("Wait")

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Import schedules with employee assignments
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func ImportStartEE()
    $assignEE = True
    $Var = True

	$iResponse = MsgBox(4, "<schedule>X ", "scheduleX � 2011 Matthew Traughber" & " " & "[matthewtraughber@gmail.com]" & _
		@CRLF & @CRLF & "This work is licensed under the Creative Commons Attribution-NoDerivs 3.0 Unported License." & _
		@CRLF & "To view a copy of this license, visit http://creativecommons.org/licenses/by-nd/3.0/" & _
		@CRLF & @CRLF & "By using this software you agree to the terms above." & _
		@CRLF &	"IF YOU DO NOT AGREE, DO NOT COPY OR USE THE SOFTWARE." & _
		@CRLF & @CRLF & @CRLF & "                                                                    " & "AGREE?", 0)

	If $iResponse = 7 Then
		Call ("ExitScheduleX")
	Else
	EndIf

	Do
		$fromDate = InputBox("<schedule>X", "Generate schedules FROM what date?"  & @CRLF & @CRLF & _
			"Format:  MM/DD/YYYY" & @CRLF & _
			@CRLF & "You cannot generate schedules more than four weeks at a time.", "1/1/2011", "")

		If StringInStr($fromDate, "/", 2, -1, 10, 5) Then
			$verifyDate = True
		Else
			MsgBox(0, "ERROR!", "DATE FORMAT INCORRECT" & @CRLF & @CRLF & @CRLF & "Verify date input matches correct format", 0)
		EndIf
	Until $verifyDate = True

	$verifyDate = False

	Do
		$toDate = InputBox("<schedule>X", "Generate schedules TO what date?"  & @CRLF & @CRLF & _
			"Formatt:  MM/DD/YYYY" & @CRLF & _
			@CRLF & "You cannot generate schedules more than four weeks at a time.", "1/1/2011", "")

		If StringInStr($toDate, "/", 2, -1, 10, 5) Then
			$verifyDate = True
		Else
			MsgBox(0, "ERROR!", "DATE FORMAT INCORRECT" & @CRLF & @CRLF & @CRLF & "Verify date input matches correct format", 0)
		EndIf
	Until $verifyDate = True

EndFunc

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Import without employee assignments
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func ImportStartNoEE()
    $assignEE = False
	$Var = True

	$iResponse = MsgBox(4, "<schedule>X", "scheduleX � 2011 Matthew Traughber" & " " & "[matthewtraughber@gmail.com]" & _
		@CRLF & @CRLF & "This work is licensed under the Creative Commons Attribution-NoDerivs 3.0 Unported License." & _
		@CRLF & "To view a copy of this license, visit http://creativecommons.org/licenses/by-nd/3.0/" & _
		@CRLF & @CRLF & "By using this software you agree to the terms above." & _
		@CRLF &	"IF YOU DO NOT AGREE, DO NOT COPY OR USE THE SOFTWARE." & _
		@CRLF & @CRLF & @CRLF & "                                                                    " & "AGREE?", 0)

	If $iResponse = 7 Then
		Call ("ExitScheduleX")
	Else
	EndIf

EndFunc

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Initialize excel documents
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func ExcelInitialize()
	$Var = False

	$xlsDocSched = _ExcelBookOpen(@ScriptDir & "\schedules.xls", 0)
	$xlsDocEE = _ExcelBookOpen(@ScriptDir & "\employees.xls", 0)
		; Checks to ensure file exists
		If @error = 1 Then
			MsgBox(0, "ERROR!", "UNABLE TO CREATE EXCEL OBJECT",0)
			Call ("Wait")
		ElseIf @error = 2 Then
			MsgBox(0, "ERROR!", "FILE DOES NOT EXIST" & @CRLF & @CRLF & @CRLF & "Confirm schedules.xls is located in the same directory as scheduleX.exe", 0)
			Call ("Wait")
		EndIf

		; Proceeds to next stage
		Call ("AliasVerify")
EndFunc

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Verifies TLO alias is correct
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func AliasVerify()
	; Opens TLO schedules URL
	Global $ieSite = _IECreate("https://timeandlabor.paychex.com/secure/old/TLOHome/ScheduleTemplates.asp")

	; Strips whitespace
	Global $xlsTLOAlias = StringStripWS(_ExcelReadCell($xlsDocSched, 1, 1), 8)
	Global $xlsTLOAlias2 = StringStripWS(_ExcelReadCell($xlsDocEE, 1, 1), 8)

	; Scans site for TLO alias
	Global $ieTLOSiteTXT = _IEBodyReadText($ieSite)
		; Compares scraped alias to XLS alias
		If StringInStr($ieTLOSiteTXT, $xlsTLOAlias, 2) Then
			If $assignEE Then
				If StringInStr($ieTLOSiteTXT, $xlsTLOAlias2, 2) Then
					MsgBox(0, "<schedule>X", "TLO SITE ALIAS MATCHES EXCEL DOCUMENTS", 0)
					Call ("ScheduleInput")
				Else
					MsgBox(0, "ERROR!", "TLO SITE ALIAS DOES NOT MATCH EMPLOYEES EXCEL DOCUMENT" & @CRLF & @CRLF & @CRLF & "Verify TLO site matches " & _
					"the alias in employees.xls, and retry", 0)
					_IEQuit ($ieSite)
				EndIf
			Else
				MsgBox(0, "<schedule>X", "TLO SITE ALIAS MATCHES EXCEL DOCUMENT", 0)
				Call ("ScheduleInput")
			EndIf
		ElseIf StringInStr($ieTLOSiteTXT, "Your session has timed out", 2) Then
			MsgBox(0, "ERROR!", "YOU ARE NOT LOGGED INTO TLO" & @CRLF & @CRLF & @CRLF & "Please log into correct TLO alias and retry", 0)
			_IEQuit ($ieSite)
		Else
			MsgBox(0, "ERROR!", "TLO SITE ALIAS DOES NOT MATCH SCHEDULES EXCEL DOCUMENT" & @CRLF & @CRLF & @CRLF & "Verify TLO site matches " & _
			"the alias in schedules.xls, and retry", 0)
			_IEQuit ($ieSite)

		EndIf
EndFunc

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Inputs schedules
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func ScheduleInput()
	; Defines start row
	$xlsRow = 2
	$xlsSchedName = _ExcelReadCell($xlsDocSched, $xlsRow, 1)

	$xlsCol2 = 2

	; Starts timer for input
	$begin = TimerInit()

	; Loops for multiple schedules
	Do
		; Defines start row
		$dayCounter = 1
		$xlsCol = 2

		; Searches for pre-existing schedules
		If StringInStr($ieTLOSiteTXT, "Search for", 2) Then
			$ieSchedSearch = _IEFormGetObjByName ($ieSite, "Form4")
				$ieSchedNameSearch = _IEFormElementGetObjByName ($ieSchedSearch, "txtSearchFor")
				_IEFormElementSetValue ($ieSchedNameSearch, $xlsSchedName, 1)

			$ieUpdSchedSearch = _IEGetObjByName ($ieSite, "Submit5")
				_IEAction ($ieUpdSchedSearch, "click")
				_IELoadWait ($ieSite)


			$ieSchedProp = _IEFormGetObjByName ($ieSite, "Form3")
				$ieSchedName = _IEFormElementGetObjByName ($ieSchedProp, "txtTemplateName")

			$ieSchedNameValue = _IEFormElementGetValue ($ieSchedName)
			$schedNameSearchCompare = StringInStr(StringStripWS($xlsSchedName, 8), StringStripWS($ieSchedNameValue, 8), 2)

			While $schedNameSearchCompare = 0
				$ieSchedSearch = _IEFormGetObjByName ($ieSite, "Form4")
					$ieSchedNameSearch = _IEFormElementGetObjByName ($ieSchedSearch, "txtSearchFor")
					_IEFormElementSetValue ($ieSchedNameSearch, "", 1)

				$ieUpdSchedSearch = _IEGetObjByName ($ieSite, "Submit5")
					_IEAction ($ieUpdSchedSearch, "click")
					_IELoadWait ($ieSite)

				_IELinkClickByText ($ieSite, "Add")

				$schedNameSearchCompare = 1
			WEnd
		Else
			_IELinkClickByText ($ieSite, "Add")
		EndIf

		; Sets schedule properties
		$ieSchedProp = _IEFormGetObjByName ($ieSite, "Form3")
			$ieSchedName = _IEFormElementGetObjByName ($ieSchedProp, "txtTemplateName")
				_IEFormElementSetValue ($ieSchedName, $xlsSchedName, 1)

			$ieSchedType = _IEFormElementGetObjByName ($ieSchedProp, "selAccessType")
				_IEFormElementSetValue ($ieSchedType, "2")

			$ieSchedOwner = _IEFormElementGetObjByName ($ieSchedProp, "selUserID")
				_IEFormElementSetValue ($ieSchedOwner, "0")

		$ieUpdSchedProp = _IEGetObjByName ($ieSite, "Submit3")
			_IEAction ($ieUpdSchedProp, "click")
			_IELoadWait ($ieSite)


		; Loop to input days / times
		While $dayCounter <= 8

			$xlsDay = _ExcelReadCell($xlsDocSched, $xlsRow, $xlsCol)
			$dayArrayCellCheck = StringInStr($xlsDay, "/", 2, 1)

				; Check # of shifts / apply settings
				If $xlsDay = "" Then
					; (do nothing - not scheduled)
				ElseIf $dayArrayCellCheck > 0 AND $dayArrayCellCheck < 6  Then
					MsgBox(0, "ERROR!", "INVALID CELL FORMAT DETECTED" & @CRLF & @CRLF & @CRLF & _
					"Verify schedule format and retry:" & @CRLF &  @CRLF & "  -  " & $xlsSchedName, 0)
					_ExcelBookClose($xlsDocSched)
					Call ("Wait")
				Else
					$dayArray = StringSplit($xlsDay, '/', 1)
					$ieSchedShifts = _IEFormGetObjByName ($ieSite, "Form2")

					; Input start / end times
					$dayShift1CellCheck = StringInStr($dayArray[1], "-", 2, 1)
						If $dayShift1CellCheck < 3 Then
							MsgBox(0, "ERROR!", "INVALID CELL FORMAT DETECTED" & @CRLF & @CRLF & @CRLF & _
							"Verify schedule format and retry:" & @CRLF &  @CRLF & "  -  " & $xlsSchedName, 0)
							_ExcelBookClose($xlsDocSched)
							Call ("Wait")
						Else
							$dayShift1Array = StringSplit($dayArray[1], '-', 1)

							$ieSchedTime = _IEFormGetObjByName ($ieSite, "Form1")
							$dayStart = _IEFormElementGetObjByName ($ieSchedTime, "txtStartTime" & $dayCounter)
							_IEFormElementSetValue ($dayStart, $dayShift1Array[1])

							$dayEnd = _IEFormElementGetObjByName ($ieSchedTime, "txtEndTime" & $dayCounter)
							_IEFormElementSetValue ($dayEnd, $dayShift1Array[2])
						EndIf
				EndIf

			; Increments counters
			$dayCounter = $dayCounter + 1
			$xlsCol = $xlsCol + 1
		WEnd


		; Applies shift times to schedules
		$ieUpdSchedTime = _IEGetObjByName ($ieSite, "Submit1")
		_IEAction ($ieUpdSchedTime, "click")
		_IELoadWait ($ieSite)

		; Scans site for error code
		Global $ieTLOSiteTXTError = _IEBodyReadText($ieSite)

		If StringInStr($ieTLOSiteTXTError, "Error: -1", 2) Then
			MsgBox(0, "ERROR!", "INVALID CELL FORMAT DETECTED" & @CRLF & @CRLF & @CRLF & _
			"Verify schedule format and retry:" & @CRLF &  @CRLF & "  -  " & $xlsSchedName, 0)
			_ExcelBookClose($xlsDocSched)
			Call ("Wait")
		Else
		EndIf

		; Assigns employees to schedulees
		If $assignEE Then
			$xlsRow2 = 2

			$xlsEEName = _ExcelReadCell($xlsDocEE, $xlsRow2, $xlsCol2)

			_IELinkClickByText ($ieSite, "Assign Multiple Employees")

			$ieEE2 = _IEAttach ("Schedule Template Assignment Page")
			_IELoadWait ($ieEE2)

			$ieEEAssignForm = _IEFormGetObjByName ($ieEE2, "Form1")

			$selectEE = _IEFormElementGetObjByName ($ieEEAssignForm, "selEmployeeSelect1")

			Do
				$xlsEEName = _ExcelReadCell($xlsDocEE, $xlsRow2, $xlsCol2)
				_IEFormElementOptionSelect ($selectEE, $xlsEEName, 1, "byText")

				_IELoadWait ($ieEE2)

				$xlsRow2 = $xlsRow2 + 1

			Until $xlsEEName = ""

			$schedName = _ExcelReadCell($xlsDocEE, 1, $xlsCol2)
			$xlsCol2 = $xlsCol2 + 1

			$ieMoveEE = _IEGetObjByName ($ieEE2, "rightMove")
			_IEAction ($ieMoveEE, "click")
			_IELoadWait ($ieEE2)

			_IELinkClickByText ($ieSite, "Return to ' " & $schedName &  " ' Schedule Template")

			; Generates schedules
			_IELinkClickByText ($ieSite, "Generate schedules")

			$ieEE3 = _IEAttach ("Generate schedules")
			_IELoadWait ($ieEE3)

			$ieGenerateForm = _IEFormGetObjByName ($ieEE3, "mainForm")

			_IEFormElementCheckboxSelect ($ieGenerateForm, 0, "", 1, "byIndex")

			_IELoadWait ($ieEE3)
			$ieEE3 = _IEAttach ("Generate schedules")
			_IELoadWait ($ieEE3)

			$ieGenerateForm = _IEFormGetObjByName ($ieEE3, "mainForm")

			$fromDate2 = _IEFormElementGetObjByName ($ieGenerateForm, "_txtGenerateFrom")
			$toDate2 = _IEFormElementGetObjByName ($ieGenerateForm, "_txtGenerateTo")
			_IEFormElementSetValue ($fromDate2, $fromDate)
			_IEFormElementSetValue ($toDate2, $toDate)

			_IEFormElementCheckboxSelect ($ieGenerateForm, 1, "", 1, "byIndex")

			_IELoadWait ($ieEE3)
			$ieEE3 = _IEAttach ("Generate schedules")
			_IELoadWait ($ieEE3)

			$ieEE3 = _IEAttach ("Generate schedules")
			$oSubmit = _IEGetObjByName ($ieEE3, "_btnSubmit")
			_IEAction ($oSubmit, "click")
			_IELoadWait ($ieEE3)

			Else

		EndIf

		; Increments counter
		$xlsRow = $xlsRow + 1
		$xlsSchedName = _ExcelReadCell($xlsDocSched, $xlsRow, 1)

	Until $xlsSchedName = ""

	; Ends / exits import (stops timer and gets values)
	$dif = TimerDiff($begin)
	_TicksToTime($dif, $iHours, $iMins, $iSecs)
	_ExcelBookClose($xlsDocSched)

	; Resets search box and shows all schedules in site
	$ieSchedSearch = _IEFormGetObjByName ($ieSite, "Form4")
		$ieSchedNameSearch = _IEFormElementGetObjByName ($ieSchedSearch, "txtSearchFor")
		_IEFormElementSetValue ($ieSchedNameSearch, "", 1)

	$ieUpdSchedSearch = _IEGetObjByName ($ieSite, "Submit5")
		_IEAction ($ieUpdSchedSearch, "click")
		_IELoadWait ($ieSite)

	; Reports stats
	MsgBox(0, "<schedule>X", "TLO SCHEDULE UPLOAD / CONFIGURATION COMPLETE" & @CRLF & @CRLF & _
	"Import took (hr, min, sec):  " & StringFormat("%02d:%02d:%02d", $iHours, $iMins, $iSecs) & @CRLF & _
	"Schedules imported/updated:  " & $xlsRow - 2 & @CRLF & @CRLF & @CRLF & _
	"Remember to check for multiple shifts - " & @CRLF & "You will need to enter 2nd/3rd/4th... shifts MANUALLY", 0)
EndFunc

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Exits script
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func ExitScheduleX()
	_ExcelBookClose($xlsDocSched)
    _ExcelBookClose($xlsDocEE)
	Exit
EndFunc


; Exits script due to escape key
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func EscScheduleX()
	MsgBox(0, "<schedule>X", "ESC KEY DETECTED" & @CRLF & @CRLF & "Program has been terminated.", 0)
	_ExcelBookClose($xlsDocSched)
    _ExcelBookClose($xlsDocEE)
	Exit
EndFunc

;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
; Waits for user input
;::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Func Wait()
	$Var = False

	While 1
		Sleep(1000)
		If $Var Then
			ExcelInitialize()
		EndIf
	WEnd
EndFunc
