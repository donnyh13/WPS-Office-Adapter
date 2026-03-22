#include "..\PowerPoint.au3"
#include "..\WPS Office Adapter.au3"
#include <MsgBoxConstants.au3>

_Example()

Func _Example()
	Local $oWPS_Pres

	; Create a new WPS Presentation Application.
	$oWPS_Pres = _WPS_Presentation_Open(True, True)
	If @error Then Exit MsgBox($MB_OK, "", "Failed to create the WPS Presentation application object." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	MsgBox($MB_OK, "", "Successfully Opened a new WPS Office Presentation Object. ")

	; Close the Presentation Application.
	_PPT_Close($oWPS_Pres)

EndFunc
