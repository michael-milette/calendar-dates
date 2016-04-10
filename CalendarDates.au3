#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=CalendarDates.ico
#AutoIt3Wrapper_Outfile=CalendarDates32.exe
#AutoIt3Wrapper_Outfile_x64=CalendarDates64.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Inserts date numbers into a calendar table.
#AutoIt3Wrapper_Res_Description=Insert dates in calendar
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2016 TNG Consulting Inc. All rights reserved.
#AutoIt3Wrapper_Res_Language=4105
#AutoIt3Wrapper_Res_Field=ProductName|Calendar Dates
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; This file is part of Calendar Dates.
;
; Calendar Dates is free software: You can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.
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
; @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later.
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

; Set the application's title bar.
Global $appTitle
$appTitle = "Calendar Dates"
If @Compiled Then
    $appTitle = $appTitle & " - v" & FileGetVersion ( @ScriptFullPath )
Else
	; If running from within the IDE.
    $appTitle = $appTitle & " - v" & '1.0.0 (dev)'
EndIf

Opt("GUIOnEventMode", 1)
MainGUI()

; ----- GUIs
Func MainGUI()
	Global $openDocuments, $cmbLastDay, $cmbCellTabs, $listGUI

	; Create dialogue box.
	$listGUI = GUICreate($appTitle, 400, 200, -1, -1, BitXOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX), $WS_EX_TOPMOST)

	; Set ESC key and X button action.
	GUISetOnEvent($GUI_EVENT_CLOSE, "btnClose")

	; Create drop down combo box listing possible number of days in a month.
	$hLabel = GUICtrlCreateLabel("", 10, 15, 80, 20)
	GUICtrlSetData($hLabel, "Days in month:")
	$cmbLastDay = GUICtrlCreateCombo("", 85, 12, 50, 20)
	GUICtrlSetData($cmbLastDay, "28|29|30|31")
	; Default is 31 because there are more months with this number of days than any other number in a year.
	_GUICtrlComboBox_SetCurSel($cmbLastDay, 3)
	GUICtrlSetState($cmbLastDay, $GUI_FOCUS)

	; Display selection list of Word documents that are currently open.
	; User will select the one to insert the numbers into.
	$openDocuments = GUICtrlCreateListView("Select a Word Document window and click 'Insert':", 10, 40, 380, 115)
	_GUICtrlListView_SetColumnWidth($openDocuments, 0, 376)
	Local $aWinList = WinList("[REGEXPTITLE:(?i)(.*- Microsoft Word.*|.*- Word.*|.*- OpenOffice Writer.*|.*- LibreOffice Writer.*|.*- Writer.*|.*- Google Docs.*|.*- OneNote.*|.*- AbiWord.*|.*- Kingsoft Writer.*|WordPerfect .*|.*Word Pro -.*)]")
	If $aWinList[0][0] = 0 Then
		; There were no open documents detected.
		MsgBox($MB_OK + $MB_ICONERROR, $appTitle, "First open the document containing your calendar and then launch this tool again. This tool supports tables in the following applications:" & Chr(13) & Chr(13) & "- Microsoft Word" & Chr(13) & "- Microsoft Word Online" & Chr(13) & "- OpenOffice Writer" & Chr(13) & "- LibreOffice Writer" & Chr(13) & "- Google Docs" & Chr(13) & "- WordPerfect" & Chr(13) & "- Lotus Word Pro" & Chr(13) & "- OneNote")
		btnClose()
	EndIf
	; Load the list of documents into the picklist.
	For $i = 1 To $aWinList[0][0]
		If $aWinList[$i][0] <> "" And BitAND(WinGetState($aWinList[$i][1]), 2) Then
			GUICtrlCreateListViewItem($aWinList[$i][0], $openDocuments)
		EndIf
	Next

	; Set Help button and action.
	$BtnAdd = GUICtrlCreateButton("Help", 10, 165, 60, 25)
	GUICtrlSetOnEvent(-1, "btnHelp")

	; Set Insert button and action. Make it the default action when Enter is pressed.
	$BtnSelect = GUICtrlCreateButton("Insert", 80, 165, 60, 25, $BS_DEFPUSHBUTTON)
	GUICtrlSetOnEvent(-1, "btnInsert")

	; Set Delete button and action.
	; This deletes the current dates in the calendar cells.
	$BtnSelect = GUICtrlCreateButton("Delete", 150, 165, 60, 25, $BS_DEFPUSHBUTTON)
	GUICtrlSetOnEvent(-1, "btnDelete")

	; Set the close button and action.
	$BtnSelect = GUICtrlCreateButton("Close", 330, 165, 60, 25)
	GUICtrlSetOnEvent(-1, "btnClose")

	; Make this application modeless and ontop.
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)

	GUISetState(@SW_SHOW)

	While 1
		Sleep(10)
	WEnd
EndFunc   ;==>MainGUI

; ////////////////////////////////// Functions //////////////////////////////////

Func btnInsert()
	; Release this window as topmost while we take some action.
	WinSetOnTop($listGUI, "", 0)

	; Get the currently selected document windows.
	$sItem = GUICtrlRead(GUICtrlRead($openDocuments))
	$sItem = StringTrimRight($sItem, 1) ; Will remove the pipe "|" from the end of the string

	; Get the currently selected last day.
	$iLastDay = GUICtrlRead($cmbLastDay)

	; If the user did not select a document in the list, remind them.
	If $sItem = "" Then
		MsgBox($MB_OK + $MB_ICONINFORMATION, $appTitle, "You must select a document. Click OK to try again.")
	Else
		; Otherwise, activate the Window and start sending the numbers followed by a tab keypress.
		WinActivate($sItem)
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
		MsgBox($MB_OK + $MB_ICONINFORMATION, $appTitle, "Done", 10)
	EndIf

	; Activate this application's window and make it top most again.
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>btnInsert

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
		MsgBox($MB_OK + $MB_ICONINFORMATION, $appTitle, "You must select a document. Click OK to try again.")
	Else
		; Otherwise, activate the Window and start deleting the numbers in each cell followed by a tab keypress.
		If $IDYES = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_DEFBUTTON2, $appTitle, "Are you sure you want to remove the dates from the calendar?") Then
			WinActivate($sItem)
			For $i = 1 To $iLastDay
				If $i = 1 Then
					Send("{TAB}+{TAB}")
				ElseIf $i > 1 Then
					Send("{TAB}")
				EndIf
				Send("{DELETE}")
			Next
			Send("+{TAB " & ($iLastDay - 1) & "}")
			Sleep(2000)
			; We're done!
			MsgBox($MB_OK + $MB_ICONINFORMATION, $appTitle, "Done", 10)
		EndIf
	EndIf

	; Activate this application's window and make it top most again.
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>btnDelete

Func btnHelp()
	WinSetOnTop($listGUI, "", 0)
	; Display help and copyright notice.
	MsgBox($MB_OK + $MB_ICONINFORMATION, $appTitle, "The purpose of this tool is to insert numbers from 1 to 28-31 in a Word table, pressing a tab key between each of them in order to populate a calendar." & Chr(13) & Chr(13) & "INSTRUCTIONS" & Chr(13) & Chr(13) & "1. Position your cursor in the starting table cell." & Chr(13) & "2. Select the number of days and the document." & Chr(13) & "3. Click the Insert button." & Chr(13) & Chr(13) & "Copyright © 2016 TNG Consulting Inc. All rights reserved." & Chr(13) & "Written by: Michael Milette" & Chr(13) & Chr(13) & "This is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version." & Chr(13) & Chr(13) & "This software is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License at http://www.gnu.org/licenses/ for more details.")
	WinActivate($listGUI)
	WinSetOnTop($listGUI, "", $WINDOWS_ONTOP)
EndFunc   ;==>btnHelp

Func btnClose()
	; Release this window from being topmost and close it.
	WinSetOnTop($listGUI, "", 0)
	Exit
EndFunc   ;==>btnClose
