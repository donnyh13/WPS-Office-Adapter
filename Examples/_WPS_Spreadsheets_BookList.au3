#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include "..\WPS Office Adapter.au3"

_Example()

Func _Example()
	Local $oWPS_Spreadsheet
	Local $avWorkbooks[0][3]

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

	; Create another new Workbook
	_Excel_BookNew($oWPS_Spreadsheet)
	If @error Then
		MsgBox($MB_OK, "", "Error creating new workbook." & @CRLF & "Error:" & @error & ", Extended:" & @extended)
		_Excel_Close($oWPS_Spreadsheet, False, True)
		Exit
	EndIf

	; Attach to the first Workbook where the Title matches
	$avWorkbooks = _WPS_Spreadsheets_BookList()
	If @error Then Exit MsgBox($MB_OK, "", "Error retrieving array of Workbooks." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	_ArrayDisplay($avWorkbooks)

	_Excel_Close($oWPS_Spreadsheet, False, True)

EndFunc
