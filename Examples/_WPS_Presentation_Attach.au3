#include "..\PowerPoint.au3"
#include "..\WPS Office Adapter.au3"
#include <MsgBoxConstants.au3>

_Example()

Func _Example()
	Local $oWPS_Pres, $oPresentation

	; Create a new WPS Presentation Object.
	$oWPS_Pres = _WPS_Presentation_Open(True, True)
	If @error Then Exit MsgBox($MB_OK, "", "Failed to create the WPS Presentation application object." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	; Create an empty Presentation
	$oPresentation = _PPT_PresentationNew($oWPS_Pres)
	If @error Then
		MsgBox($MB_OK, "", "Error Creating New presentation" & @CRLF & "Error:" & @error & ", Extended:" & @extended)
		_PPT_Close($oWPS_Pres)
		Exit
	EndIf

	; Attach to the new, Empty Presentation.
	$oPresentation = _WPS_Presentation_Attach($oPresentation.Windows(1).Caption(), $WPS_SEARCH_TITLE)
	If @error Then Exit MsgBox($MB_OK, "", "Error attaching to Presentation." & @CRLF & "Error:" & @error & ", Extended:" & @extended)

	MsgBox($MB_OK, "", "Search by Title Successfully attached to the following presentation. " & $oPresentation.Windows(1).Caption())

	; Close the Presentation.
	_PPT_Close($oWPS_Pres)

EndFunc
