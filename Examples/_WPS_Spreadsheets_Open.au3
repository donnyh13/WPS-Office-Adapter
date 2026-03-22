#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include "..\WPS Office Adapter.au3"
_Example()

Func _Example()
	Local $oWPS_Spreadsheet

	; Create a new WPS Spreadsheets Application.
	$oWPS_Spreadsheet = _WPS_Spreadsheets_Open(True, False, True, True, True)
	If @error <> 0 Then Exit MsgBox($MB_OK, "", "Error creating the WPS Spreadsheets application object." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	; Create a new Workbook
	_Excel_BookNew($oWPS_Spreadsheet)
	If @error Then
		MsgBox($MB_OK, "", "Error creating new workbook." & @CRLF & "Error:" & @error & ", Extended:" & @extended)
		_Excel_Close($oWPS_Spreadsheet, False, True)
		Exit
	EndIf

	MsgBox($MB_OK, "", "I successfully created a new WPS Spreadsheets Application instance.")

	; Close the Spreadsheets Application.
	_Excel_Close($oWPS_Spreadsheet, False, True)

EndFunc
