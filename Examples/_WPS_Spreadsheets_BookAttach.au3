#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include "..\WPS Office Adapter.au3"

_Example()

Func _Example()
	Local $oWPS_Spreadsheet, $oWorkbook

	; Create a new WPS Spreadsheets Application.
	$oWPS_Spreadsheet = _WPS_Spreadsheets_Open(True, False, True, True, True)
	If @error <> 0 Then Exit MsgBox($MB_OK, "", "Error creating the WPS Spreadsheets application object." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	; Create a new Workbook.
	$oWorkbook = _Excel_BookNew($oWPS_Spreadsheet)
	If @error Then
		MsgBox($MB_OK, "", "Error creating new workbook." & @CRLF & "Error:" & @error & ", Extended:" & @extended)
		_Excel_Close($oWPS_Spreadsheet, False, True)
		Exit
	EndIf

	; Attach to the first Workbook where the Title matches
	$oWorkbook = _WPS_Spreadsheets_BookAttach($oWorkbook.Application.Caption, $WPS_SEARCH_TITLE)
	If @error Then Exit MsgBox($MB_OK, "", "Error attaching to Workbook." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	MsgBox($MB_OK, "", "Search by Title Successful, attached to Workbook. " & $oWorkbook.Application.Caption)

	; Close the Spreadsheet Application.
	_Excel_Close($oWPS_Spreadsheet, False, True)

EndFunc
