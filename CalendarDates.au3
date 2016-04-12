#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=CalendarDates.ico
#AutoIt3Wrapper_Outfile=CalendarDates32.exe
#AutoIt3Wrapper_Outfile_x64=CalendarDates64.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Inserts date numbers into a calendar table.
#AutoIt3Wrapper_Res_Description=Insert dates in calendar
#AutoIt3Wrapper_Res_Fileversion=1.1.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2016 TNG Consulting Inc. All rights reserved.
#AutoIt3Wrapper_Res_Language=4105
#AutoIt3Wrapper_Res_Field=ProductName|Calendar Dates
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; This file is part of Calendar Dates.
;
; Calendar Dates is free software: You can redistribute it and/or modify
; it under the terms of the GNU General Public License, version 3,
; as published by the Free Software Foundation.
;
; Calendar Dates is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
; GNU General Public License for more details.
;
; You should have received a copy of the GNU General Public License
; along with Calendar Dates.  If not, see <http://www.gnu.org/licenses/>.
;

; Version information for Catalogue.
;
; @package    CalendarDates
; @copyright  2016 TNG Consulting Inc. - www.tngconsulting.ca
; @author     Michael Milette
; @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3.
; @compiler   AutoIt Version: 3.3.14.2
; @purpose    Insert dates into a calendar table.
;

#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <GuiComboBoxEx.au3>
#include <WinAPI.au3>

Global $help, $copyright, $gitURL, $aStrings

; Language strings.
$aString = ObjCreate("Scripting.Dictionary")
$aString('apptitle') = "Calendar Dates"
$aString('appversion') = "1.1.0 - beta"
$aString('help') = $aString('apptitle') & " inserts date numbers into a 7 column table." & Chr(13) & Chr(13) & "INSTRUCTIONS" & Chr(13) & Chr(13) & "1. Open the document containing your calendar table." & Chr(13) & "2. Position your cursor in the table cell of the first day of the month." & Chr(13) & "3. Select the number of days for that month and the open document in which you want to insert the dates." & Chr(13) & "4. Press the 'Insert' button to add dates or 'Delete' to remove them." & Chr(13) & Chr(13) & "This tool supports tables in the following applications:" & Chr(13) & Chr(13) & "- Microsoft Word" & Chr(13) & "- Microsoft Word Online" & Chr(13) & "- Google Docs" & Chr(13) & "- Microsoft OneNote" & Chr(13) & "- Microsoft OneNote Online" & Chr(13) & "- OpenOffice Writer" & Chr(13) & "- LibreOffice Writer" & Chr(13) & "- WordPerfect" & Chr(13) & "- Lotus Word Pro"
$aString('copyright') = "Copyright © 2016 TNG Consulting Inc. All rights reserved." & Chr(13) & "Visit www.tngconsulting.ca" & Chr(13) & "Written by Michael Milette" & Chr(13) & Chr(13) & "This is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License, version 3, as published by the Free Software Foundation." & Chr(13) & Chr(13) & "This software is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License at http://www.gnu.org/licenses/ for more details."
$aString('git_url') = "https://github.com/michael-milette/calendar-dates"
$aString('file') = "&File"
$aString('refresh') = "&Refresh"
$aString('exit') = "E&xit"
$aString('help') = "&Help"
$aString('checkforupdates') = "&Check for updates"
$aString('about') = "&About"
$aString('daysinmonth') = "Days in month:"
$aString('selectandclick') = "Select a document and press 'Insert' or 'Delete':"
$aString('delete') = "&Delete"
$aString('insert') = "&Insert"
$aString('close') = "&Close"
$aString('firstselectdocument') = "You must first select a document. Press OK and try again or press F5 to refresh the list."
$aString('docnotavailable') = "The document you selected is no longer available. Please re-select your document or press F5 to refresh the list."
$aString('apps') = ".*- Microsoft Word.*|.*- Word.*|.*- OpenOffice Writer.*|.*- LibreOffice Writer.*|.*- Writer.*|.*- Google Docs.*|.*- OneNote.*|.*- Microsoft OneNote.*|.*- AbiWord.*|.*- Kingsoft Writer.*|WordPerfect .*|.*Word Pro -.*"
$aString('apps_wordperfect') = "WordPerfect "
$aString('confirmremove') = "Are you sure you want to remove the dates from the calendar?"
$aString('done') = "Done"

; Set application title.
$aString('apptitle') = $aString('apptitle') & " - v" & $aString('appversion')
If Not @Compiled Then
	; If running from within the IDE.
    $aString('apptitle') = $aString('apptitle') & " (dev)"
EndIf

Opt("GUIOnEventMode", 1)
MainGUI()

; ----- GUIs
Func MainGUI()
	Global $openDocuments, $cmbLastDay, $cmbCellTabs, $listGUI

	; Create dialogue box.
	$listGUI = GUICreate($aString('apptitle'), 400, 220, -1, -1, BitXOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX), $WS_EX_TOPMOST)

	; Add a menu to it.
    Local $idFilemenu = GUICtrlCreateMenu($aString('file'))
    Local $idExititem = GUICtrlCreateMenuItem($aString('refresh') & @TAB & "F5", $idFilemenu)
	GUICtrlSetOnEvent(-1, "mnuFileRefresh")
	HotKeySet("{F5}", "mnuFileRefresh")
    Local $idExititem = GUICtrlCreateMenuItem($aString('exit'), $idFilemenu)
	GUICtrlSetOnEvent(-1, "btnClose")
    GUICtrlSetState(-1, $GUI_DEFBUTTON)
    Local $idHelpmenu = GUICtrlCreateMenu($aString('help'))
    Local $idHelpitem = GUICtrlCreateMenuItem($aString('help') & @TAB & "F1", $idHelpmenu)
	GUICtrlSetOnEvent(-1, "mnuHelpHelp")
	HotKeySet("{F1}", "mnuHelpHelp")
    Local $idExititem = GUICtrlCreateMenuItem($aString('checkforupdates'), $idHelpmenu)
	GUICtrlSetOnEvent(-1, "mnuHelpUpdateCheck")
    GUICtrlCreateMenuItem("", $idHelpmenu, 2) ; create a separator line
    Local $idExititem = GUICtrlCreateMenuItem($aString('about'), $idHelpmenu)
	GUICtrlSetOnEvent(-1, "mnuHelpAbout")

	; Set ESC key and X button action.
	GUISetOnEvent($GUI_EVENT_CLOSE, "btnClose")

	; Create drop down combo box listing possible number of days in a month.
	$hLabel = GUICtrlCreateLabel("", 10, 15, 80, 20)
	GUICtrlSetData($hLabel, $aString('daysinmonth'))
	$cmbLastDay = GUICtrlCreateCombo("", 85, 12, 50, 20)
	GUICtrlSetData($cmbLastDay, "28|29|30|31")
	; Default is 31 because there are more months with this number of days than any other number in a year.
	_GUICtrlComboBox_SetCurSel($cmbLastDay, 3)
	GUICtrlSetState($cmbLastDay, $GUI_FOCUS)

	; Display selection list of Documents that are currently open.
	; User will select the one of these into which the date numbers will be inserted.
	$openDocuments = GUICtrlCreateListView($aString('selectandclick'), 10, 40, 380, 115)
	_GUICtrlListView_SetColumnWidth($openDocuments, 0, 376)
	mnuFileRefresh()

	; Set Delete button and action.
	; This deletes the current dates in the calendar cells.
	$BtnSelect = GUICtrlCreateButton($aString('delete'), 10, 165, 60, 25, $BS_DEFPUSHBUTTON)
	GUICtrlSetOnEvent(-1, "btnDelete")

	; Set Insert button and action. Make it the default action when Enter is pressed.
	$BtnSelect = GUICtrlCreateButton($aString('insert'), 80, 165, 60, 25, $BS_DEFPUSHBUTTON)
	GUICtrlSetOnEvent(-1, "btnInsert")

	; Set the close button and action.
	$BtnSelect = GUICtrlCreateButton($aString('close'), 330, 165, 60, 25)
	GUICtrlSetOnEvent(-1, "btnClose")

	; Make this application modeless and ontop.
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)

	; Display GUI Dialogue box.
	GUISetState(@SW_SHOW)

	; Loop to prevent unnecessart and excessive usage of CPU.
	While 1
		Sleep(10)
	WEnd
EndFunc   ;==>MainGUI

; ////////////////////////////////// Functions //////////////////////////////////

;
; @Function: btnInsert()
; @Purpose: Insert the date numbers into the table cells.
; @Parameters: None.
; @Return: Nothing.
;
Func btnInsert()
	; Release this window as topmost while we take some action.
	WinSetOnTop($listGUI, "", 0)

	; Get the currently selected document windows.
	$sItem = GUICtrlRead(GUICtrlRead($openDocuments))
	$sItem = StringTrimRight($sItem, 1) ; Will remove the pipe "|" from the end of the string.

	; Get the currently selected last day.
	$iLastDay = GUICtrlRead($cmbLastDay)

	; If the user did not select a document in the list, remind them.
	If $sItem = "" Then
		MsgBox($MB_OK + $MB_ICONERROR, $aString('apptitle'), $aString('firstselectdocument'))
		mnuFileRefresh()
	Else
		; Otherwise, activate the Window and start sending the numbers followed by a tab keypress.
		If WinActivate($sItem) = 0 Then
			; Let user know that the document is no longer available and then refresh the list.
			MsgBox($MB_OK + $MB_ICONERROR, $aString('apptitle'), $aString('docnotavailable'))
			mnuFileRefresh()
		Else
			; Insert the dates.
			For $i = 1 To $iLastDay
				If $i = 1 Then
					Send("{TAB}+{TAB}")
				ElseIf $i > 1 Then
					Send("{TAB}")
				EndIf
				Send($i)
			Next
			Sleep(2000)
			; We're done!
			MsgBox($MB_OK + $MB_ICONINFORMATION, $aString('apptitle'), $aString('done'), 10)
		EndIf
	EndIf

	; Activate this application's window and make it top most again.
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>btnInsert

;
; @Function: btnDelete()
; @Purpose: Delete the date numbers from each of the table cells.
; @Parameters: None.
; @Return: Nothing.
;
Func btnDelete()
	; Release this window as topmost while we take some action.
	WinSetOnTop($listGUI, "", 0)

	; Get the currently selected document windows.
	$sItem = GUICtrlRead(GUICtrlRead($openDocuments))
	$sItem = StringTrimRight($sItem, 1) ; Will remove the pipe "|" from the end of the string

	; Get the currently selected last day.
	$iLastDay = GUICtrlRead($cmbLastDay)

	; If the user did not select a document in the list, remind them.
	If $sItem = "" Then
		MsgBox($MB_OK + $MB_ICONERROR, $aString('apptitle'), $aString('firstselectdocument'))
		mnuFileRefresh()
	Else
		; Otherwise, activate the Window and start deleting the numbers in each cell followed by a tab keypress.
		If $IDYES = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_DEFBUTTON2, $aString('apptitle'), $aString('confirmremove')) Then
			If WinActivate($sItem) = 0 Then
				; Let user know that the document is no longer available and then refresh the list.
				MsgBox($MB_OK + $MB_ICONERROR, $aString('apptitle'),  $aString('docnotavailable'))
				mnuFileRefresh()
			Else
				; Delete the dates in the calendar.
				For $i = 1 To $iLastDay
					If $i = 1 Then
						Send("{TAB}+{TAB}")
					ElseIf $i > 1 Then
						Send("{TAB}")
						If StringInStr($sItem, $aString('apps_wordperfect')) > 0 Then
							Send("+{END}") ; Select text in cell - WordPerfect doesn't do that automatically.
						EndIf
					EndIf
					Send("{DELETE}")
				Next
				Send("+{TAB " & ($iLastDay - 1) & "}")
				Sleep(2000)
				; We're done!
				MsgBox($MB_OK + $MB_ICONINFORMATION, $aString('apptitle'), $aString('done'), 10)
			EndIf
		EndIf
	EndIf

	; Activate this application's window and make it top most again.
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>btnDelete

;
; @Function: btnClose()
; @Purpose: Release this window from being topmost and close it.
; @Parameters: None.
; @Return: Nothing.
;
Func btnClose()
	WinSetOnTop($listGUI, "", 0)
	GUIDelete()
	Exit
EndFunc   ;==>btnClose

;
; @Function: mnuFileRefresh()
; @Purpose: Load the list of documents into the picklist.
; @Parameters: None.
; @Return: Nothing.
;
Func mnuFileRefresh()
	_GUICtrlListView_DeleteAllItems($openDocuments)
	Local $aWinList = WinList("[REGEXPTITLE:(?i)(" & $aString('apps') & ")]")
	For $i = 1 To $aWinList[0][0]
		If $aWinList[$i][0] <> "" And BitAND(WinGetState($aWinList[$i][1]), 2) Then
			GUICtrlCreateListViewItem($aWinList[$i][0], $openDocuments)
		EndIf
	Next

EndFunc   ;==>mnuFuleRefresh

;
; @Function: mnuHelpHelp()
; @Purpose: Display help.
; @Parameters: None.
; @Return: Nothing.
;
Func mnuHelpHelp()
	WinSetOnTop($listGUI, "", 0)
	MsgBox($MB_OK + $MB_ICONINFORMATION, $aString('apptitle'), $aString('help'))
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>mnuHelpHelp

;
; @Function: mnuHelpUpdateCheck()
; @Purpose: Open GitHub page in default web browser.
; @Parameters: None.
; @Return: Nothing.
;
Func mnuHelpUpdateCheck()
    ShellExecute($gitURL)
EndFunc   ;==>mnuHelpUpdateCheck

;
; @Function: mnuHelpAbout()
; @Purpose: Display copyright notice.
; @Parameters: None.
; @Return: Nothing.
;
Func mnuHelpAbout()
	WinSetOnTop($listGUI, "", 0)
	MsgBox($MB_OK + $MB_ICONINFORMATION, $aString('apptitle'), $aString('copyright'))
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>mnuHelpAbout
