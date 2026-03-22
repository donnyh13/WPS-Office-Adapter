#include <MsgBoxConstants.au3>
#include <Word.au3>
#include "..\WPS Office Adapter.au3"

Func _Example()
	Local $oWriter

	; Create a new WPS Writer object
	$oWriter = _WPS_Writer_Create(True, True)
	If @error <> 0 Then Exit MsgBox($MB_OK, "", "Error creating a new WPS Writer application object." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	MsgBox($MB_OK, "", "I successfully created a new Instance of WPS Writer instance.")

	; Close the WPS Writer Application.
	_Word_Quit($oWriter)

EndFunc
