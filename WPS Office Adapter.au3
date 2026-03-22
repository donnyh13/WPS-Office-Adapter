#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#Tidy_Parameters=/sf /reel /tcl=1

#include-once

#include <AutoItConstants.au3>

; #INDEX# =======================================================================================================================
; Title .........: WPS Office Adapter
; AutoIt Version : v3.3.16.1
; UDF Version    : 0.0.1-alpha
; Description ...: Provides functions for creating or connecting to a WPS application instance.
; Author(s) .....: donnyh13
; Dll ...........:
; Remarks .......: Once you retrieve the necessary Object, you can use the functions from water's UDFs for the corresponding Object type.
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _WPS_ComError_UserFunction
; _WPS_Presentation_Attach
; _WPS_Presentation_Open
; _WPS_Spreadsheets_BookAttach
; _WPS_Spreadsheets_BookList
; _WPS_Spreadsheets_Open
; _WPS_Writer_Create
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __WPS_InternalComErrorHandler
; __WPS_Obj_Get
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
Global Enum $WPS_SEARCH_FILE_NAME, _ ; Search by the File Name.
		$WPS_SEARCH_FILE_PATH, _     ; Search by the File Save Path.
		$WPS_SEARCH_TITLE            ; Search by the Document Title.
; ===============================================================================================================================

Global Enum $__WPS_TYPE_PRESENTATION, $__WPS_TYPE_DOCUMENT, $__WPS_TYPE_SPREADSHEETS
Global Enum $__WPS_RETURN_SUCCESS, $__WPS_RETURN_INPUT_ERROR, $__WPS_RETURN_RETRIEVAL_ERROR, $__WPS_RETURN_PROCESSING_ERROR

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_ComError_UserFunction
; Description ...: Set a UserFunction to receive the Fired COM Error Error outside of the UDF.
; Syntax ........: _WPS_ComError_UserFunction([$vUserFunction = Default[, $vParam1 = Null[, $vParam2 = Null[, $vParam3 = Null[, $vParam4 = Null[, $vParam5 = Null]]]]]])
; Parameters ....: $vUserFunction       - [optional] a Function or Keyword. Default value is Default. Accepts a Function, or the Keyword Default and Null. If set to a User function, the function may have up to 5 required parameters.
;                  $vParam1             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam2             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam3             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam4             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
;                  $vParam5             - [optional] a variant value. Default is Null. Any optional parameter to be called with the user function.
; Return values .: Success: 1 or UserFunction.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $vUserFunction Not a Function, or Default keyword, or Null Keyword.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Successfully set the UserFunction.
;                  @Error 0 @Extended 0 Return 2 = Successfully cleared the set UserFunction.
;                  @Error 0 @Extended 0 Return Function = Returning the set UserFunction.
; Author ........: mLipok
; Modified ......: donnyh13 - Added a clear UserFunction without error option. Also added parameters option.
; Remarks .......: The first parameter passed to the User function will always be the COM Error object. See below.
;                  Every COM Error will be passed to that function. The user can then read the following properties. (As Found in the COM Reference section in AutoIt Help File.) Using the first parameter in the UserFunction.
;                  For Example MyFunc($oMyError)
;                    $oMyError.number The Windows HRESULT value from a COM call
;                    $oMyError.windescription The FormatWinError() text derived from .number
;                    $oMyError.source Name of the Object generating the error (contents from ExcepInfo.source)
;                    $oMyError.description Source Object's description of the error (contents from ExcepInfo.description)
;                    $oMyError.helpfile Source Object's help file for the error (contents from ExcepInfo.helpfile)
;                    $oMyError.helpcontext Source Object's help file context id number (contents from ExcepInfo.helpcontext)
;                    $oMyError.lastdllerror The number returned from GetLastError()
;                    $oMyError.scriptline The script line on which the error was generated
;                    NOTE: Not all properties will necessarily contain data, some will be blank.
;                  If MsgBox or ConsoleWrite functions are passed to this function, the error details will be displayed using that function automatically.
;                  If called with Default keyword, the current UserFunction, if set, will be returned.
;                  If called with Null keyword, the currently set UserFunction is cleared and only the internal ComErrorHandler will be called for COM Errors.
;                  The stored UserFunction (besides MsgBox and ConsoleWrite) will be called as follows: UserFunc($oComError,$vParam1,$vParam2,$vParam3,$vParam4,$vParam5)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_ComError_UserFunction($vUserFunction = Default, $vParam1 = Null, $vParam2 = Null, $vParam3 = Null, $vParam4 = Null, $vParam5 = Null)
	#forceref $vParam1, $vParam2, $vParam3, $vParam4, $vParam5
	; If user does not set a function, UDF must use internal function to avoid AutoItError.
	Local Static $vUserFunction_Static = Default
	Local $avUserFuncWParams[@NumParams]

	If $vUserFunction = Default Then
		; just return stored static User Function variable
		Return $vUserFunction_Static
	ElseIf IsFunc($vUserFunction) Then
		; If User called Parameters, then add to array.
		If @NumParams > 1 Then
			$avUserFuncWParams[0] = $vUserFunction
			For $i = 1 To @NumParams - 1
				$avUserFuncWParams[$i] = Eval("vParam" & $i)
				; set static variable
			Next
			$vUserFunction_Static = $avUserFuncWParams
		Else
			$vUserFunction_Static = $vUserFunction
		EndIf
		Return SetError($__WPS_RETURN_SUCCESS, 0, 1)
	ElseIf $vUserFunction = Null Then
		; Clear User Function.
		$vUserFunction_Static = Default
		Return SetError($__WPS_RETURN_SUCCESS, 0, 2)
	Else
		; return error as an incorrect parameter was passed to this function
		Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
	EndIf
EndFunc   ;==>_WPS_ComError_UserFunction

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_Presentation_Attach
; Description ...:  Attach to a WPS Presentation matching a search string.
; Syntax ........: _WPS_Presentation_Attach($sString[, $iMode = $WPS_SEARCH_FILE_PATH[, $bPartialMatch = False]])
; Parameters ....: $sString             - a string value. The String to search for.
;                  $iMode               - [optional] an integer value. Default is $WPS_SEARCH_FILE_PATH. The Search mode. See Constants $WPS_SEARCH_*.
;                  $bPartialMatch       - [optional] a boolean value. Default is False. If True, and if Search mode is set to Title, a partial Title match will be accepted as a match.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sString not a String.
;                  @Error 1 @Extended 2 Return 0 = $iMode not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $bPartialMatch not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $iMode not one of the predefined Constants. See Constants $WPS_SEARCH_*.
;                  --Retrieval Errors--
;                  @Error 2 @Extended ? Return 0 = Failed to retrieve Presentation Object.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to identify requested Presentation.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success. Returning requested Presentation Object matched by the File Name.
;                  @Error 0 @Extended 2 Return Object = Success. Returning requested Presentation Object matched by the File Path.
;                  @Error 0 @Extended 3 Return Object = Success. Returning requested Presentation Object matched by a partial Title search.
;                  @Error 0 @Extended 4 Return Object = Success. Returning requested Presentation Object matched by the Title.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Replaces ......: _PPT_PresentationAttach
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_Presentation_Attach($sString, $iMode = $WPS_SEARCH_FILE_PATH, $bPartialMatch = False)
	Local $oPresentation
	Local $iCount = 1
	Local $sCLSID_Presentation = "{91493444-5A91-11CF-8700-00AA0060263B}" ; WPS Office Presentation CLSID

	If Not IsString($sString) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
	If Not IsInt($iMode) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)
	If Not IsBool($bPartialMatch) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 3, 0)

	While True
		$oPresentation = ObjGet("", $sCLSID_Presentation, $iCount)
		If @error Then Return SetError($__WPS_RETURN_RETRIEVAL_ERROR, @error, 0)
		$iCount += 1
		If Not StringInStr(ObjName($oPresentation, $OBJ_FILE), "\Kingsoft\") Then ContinueLoop
		Switch $iMode

			Case $WPS_SEARCH_FILE_NAME
				If $oPresentation.Name = $sString Then Return SetError($__WPS_RETURN_SUCCESS, 1, $oPresentation)

			Case $WPS_SEARCH_FILE_PATH
				If $oPresentation.FullName = $sString Then Return SetError($__WPS_RETURN_SUCCESS, 2, $oPresentation)

			Case $WPS_SEARCH_TITLE
				If $bPartialMatch Then
					If StringInStr($oPresentation.Windows(1).Caption, $sString) > 0 Then Return SetError($__WPS_RETURN_SUCCESS, 3, $oPresentation)
				Else
					If $oPresentation.Windows(1).Caption = $sString Then Return SetError($__WPS_RETURN_SUCCESS, 4, $oPresentation)
				EndIf

			Case Else
				Return SetError($__WPS_RETURN_INPUT_ERROR, 4, 0)
		EndSwitch
	WEnd
	Return SetError($__WPS_RETURN_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>_WPS_Presentation_Attach

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_Presentation_Open
; Description ...: Connect to an existing Presentation Application, or open a new one.
; Syntax ........: _WPS_Presentation_Open([$bVisible = True[, $bForceNew = False]])
; Parameters ....: $bVisible            - [optional] a boolean value. Default is True. If True, the Application is visible.
;                  $bForceNew           - [optional] a boolean value. Default is False. If True, a new Application instance is created.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bForceNew not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended ? Return 0 = Failed to create WPS Presentation Object. @Extended set to ObjCreate error value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened a WPS Presentation instance. A Instance already existed, and was connected to.
;                  @Error 0 @Extended 1 Return Object = Success. Successfully opened a WPS Presentation instance. A Instance did not already exist, or $bForceNew was set to True, a new instance was created.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If a instance is created, and the script/Function finishes which created the Object, without creating or opening a Presentation for that instance, the instance will be closed, this is beyond my control.
; Related .......:
; Replaces ......: _PPT_Open
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_Presentation_Open($bVisible = True, $bForceNew = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __WPS_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oKWPP
	Local $bApplOpened = False

	If Not IsBool($bVisible) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
	If Not IsBool($bForceNew) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)

	$oKWPP = __WPS_Obj_Get($__WPS_TYPE_PRESENTATION)
	If @error Or Not IsObj($oKWPP) Then
		$oKWPP = ObjCreate("KWPP.Application")
		If @error Or Not IsObj($oKWPP) Then Return SetError($__WPS_RETURN_PROCESSING_ERROR, @error, 0)
		$bApplOpened = True
	EndIf

	$oKWPP.Visible = $bVisible

	Return SetError($__WPS_RETURN_SUCCESS, $bApplOpened, $oKWPP)
EndFunc   ;==>_WPS_Presentation_Open

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_Spreadsheets_BookAttach
; Description ...: Attach to a workbook matching a search string.
; Syntax ........: _WPS_Spreadsheets_BookAttach($sString[, $iMode = $WPS_SEARCH_FILE_PATH[, $oInstance = Default]])
; Parameters ....: $sString             - a string value. The String to search for.
;                  $iMode               - [optional] an integer value. Default is $WPS_SEARCH_FILE_PATH. The Search mode. See Constants $WPS_SEARCH_*.
;                  $oInstance           - [optional] an object. Default is Default. The WPS Instance to search. Default keyword = all instances.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $sString not a String.
;                  @Error 1 @Extended 2 Return 0 = $iMode not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMode not one of the predefined Constants. See Constants $WPS_SEARCH_*.
;                  --Retrieval Errors--
;                  @Error 2 @Extended ? Return 0 = Failed to retrieve Workbook Object. @Extended set to ObjGet Error value.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to find requested Workbook.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success. Successfully match workbook by File Name.
;                  @Error 0 @Extended 2 Return Object = Success. Successfully match workbook by File Path.
;                  @Error 0 @Extended 3 Return Object = Success. Successfully match workbook by Title.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Replaces ......: _Excel_BookAttach
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_Spreadsheets_BookAttach($sString, $iMode = $WPS_SEARCH_FILE_PATH, $oInstance = Default)
	Local $oError = ObjEvent("AutoIt.Error", __WPS_InternalComErrorHandler)
	#forceref $oError

	Local $oWorkbook
	Local $iCount = 1
	Local $sCLSID_Workbook = "{00020819-0000-0000-C000-000000000046}" ; WPS Office Workbook CLSID/ MS Office Workbook CLSID

	If Not IsString($sString) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
	If Not IsInt($iMode) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)

	While True
		$oWorkbook = ObjGet("", $sCLSID_Workbook, $iCount)
		If @error Then Return SetError($__WPS_RETURN_RETRIEVAL_ERROR, @error, 0)
		$iCount += 1
		If Not StringInStr(ObjName($oWorkbook, $OBJ_FILE), "\Kingsoft\") Then ContinueLoop
		If $oInstance <> Default And $oInstance <> $oWorkbook.Parent Then ContinueLoop
		Switch $iMode

			Case $WPS_SEARCH_FILE_NAME
				If $oWorkbook.Name = $sString Then Return SetError($__WPS_RETURN_SUCCESS, 1, $oWorkbook)

			Case $WPS_SEARCH_FILE_PATH
				If $oWorkbook.FullName = $sString Then Return SetError($__WPS_RETURN_SUCCESS, 2, $oWorkbook)

			Case $WPS_SEARCH_TITLE
				If $oWorkbook.Application.Caption = $sString Then Return SetError($__WPS_RETURN_SUCCESS, 3, $oWorkbook)

			Case Else
				Return SetError($__WPS_RETURN_INPUT_ERROR, 3, 0)
		EndSwitch
	WEnd

	Return SetError($__WPS_RETURN_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>_WPS_Spreadsheets_BookAttach

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_Spreaksheets_BookList
; Description ...: Retrieve an array of Workbooks for all Spreadsheet instances or for a specific one.
; Syntax ........: _WPS_Spreaksheets_BookList([$oInstance = Default])
; Parameters ....: $oInstance           - [optional] an object. Default is Default. The Spreadsheet Object to retrieve a list of Workbooks for. Default = List for all Spreadsheet instances.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oInstance not a WPS Object.
;                  @Error 1 @Extended 2 Return 0 = $oInstance not set to Default.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returning a three column 2D array of Workbooks. @Extended is set to the number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array columns contain the following information:
;                  Column 0 - Object of the workbook
;                  Column 1 - Name of the workbook/file
;                  Column 2 - Complete path to the workbook/file
; Related .......:
; Replaces ......: _Excel_BookList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_Spreadsheets_BookList($oInstance = Default)
	Local $oError = ObjEvent("AutoIt.Error", __WPS_InternalComErrorHandler)
	#forceref $oError

	Local $aBooks[0][3]
	Local $iIndex = 0, $iTemp, $iCount = 1
	Local $oWorkbook
	Local $sCLSID_Workbook = "{00020819-0000-0000-C000-000000000046}" ; WPS Office Workbook CLSID

	If IsObj($oInstance) Then
		If Not StringInStr(ObjName($oWorkbook, $OBJ_FILE), "\Kingsoft\") Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
		$iTemp = $oInstance.Workbooks.Count
		ReDim $aBooks[$iTemp][3]
		For $iIndex = 0 To $iTemp - 1
			$aBooks[$iIndex][0] = $oInstance.Workbooks($iIndex + 1)
			$aBooks[$iIndex][1] = $oInstance.Workbooks($iIndex + 1).Name
			$aBooks[$iIndex][2] = $oInstance.Workbooks($iIndex + 1).Path
		Next
	Else
		If $oInstance <> Default Then Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)

		While True
			$oWorkbook = ObjGet("", $sCLSID_Workbook, $iCount)
			If @error Then ExitLoop
			$iCount += 1
			If Not StringInStr(ObjName($oWorkbook, $OBJ_FILE), "\Kingsoft\") Then ContinueLoop
			ReDim $aBooks[$iIndex + 1][3]
			$aBooks[$iIndex][0] = $oWorkbook
			$aBooks[$iIndex][1] = $oWorkbook.Name
			$aBooks[$iIndex][2] = $oWorkbook.Path
			$iIndex += 1
		WEnd
	EndIf
	Return SetError($__WPS_RETURN_SUCCESS, UBound($aBooks), $aBooks)
EndFunc   ;==>_WPS_Spreadsheets_BookList

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_Spreadsheets_Open
; Description ...: Connect to an existing Spreadsheets Application, or open a new one.
; Syntax ........: _WPS_Spreadsheets_Open([$bVisible = True[, $bDisplayAlerts = False[, $bScreenUpdating = True[, $bInteractive = True[, $bForceNew = False]]]]])
; Parameters ....: $bVisible            - [optional] a boolean value. Default is True. If True, the Application is visible.
;                  $bDisplayAlerts      - [optional] a boolean value. Default is False. If True, prompts and error prompts are shown.
;                  $bScreenUpdating     - [optional] a boolean value. Default is True. If True, Screen updating is enabled. If False, Screen updating is disabled, which may speed up the script.
;                  $bInteractive        - [optional] a boolean value. Default is True. If True, user input by the mouse and keyboard is accepted.
;                  $bForceNew           - [optional] a boolean value. Default is False. If True, a new Application instance is created.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bDisplayAlerts not a Boolean.
;                  @Error 1 @Extended 3 Return 0 = $bScreenUpdating not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $bInteractive not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bForceNew not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended ? Return 0 = Failed to create WPS Writer Object. @Extended set to ObjCreate error value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened a WPS Writer instance. A Instance already existed, and was connected to.
;                  @Error 0 @Extended 1 Return Object = Success. Successfully opened a WPS Writer instance. A Instance did not already exist, or $bForceNew was set to True, a new instance was created.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If a instance is created, and the script/Function finishes which created the Object, without creating or opening a Spreadsheet for that instance, the instance will be closed, this is beyond my control.
;                  When closing the object created using this function, you will need to call ForceClose with True, when using _Excel_Close, or use water's internal function (__Excel_CloseOnQuit) as in the Excel UDF.
; Related .......:
; Replaces ......: _Excel_Open
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_Spreadsheets_Open($bVisible = True, $bDisplayAlerts = False, $bScreenUpdating = True, $bInteractive = True, $bForceNew = False)
	Local $oError = ObjEvent("AutoIt.Error", __WPS_InternalComErrorHandler)
	#forceref $oError

	Local $oKET
	Local $bApplOpened = False

	If Not IsBool($bVisible) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
	If Not IsBool($bDisplayAlerts) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)
	If Not IsBool($bScreenUpdating) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 3, 0)
	If Not IsBool($bInteractive) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 4, 0)
	If Not IsBool($bForceNew) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 5, 0)

	If Not $bForceNew Then $oKET = __WPS_Obj_Get($__WPS_TYPE_SPREADSHEETS)
	If $bForceNew Or @error Then
		$oKET = ObjCreate("KET.Application")
		If @error Or Not IsObj($oKET) Then Return SetError($__WPS_RETURN_PROCESSING_ERROR, @error, 0)
		$bApplOpened = True
	EndIf

	$oKET.Visible = $bVisible
	$oKET.DisplayAlerts = $bDisplayAlerts
	$oKET.ScreenUpdating = $bScreenUpdating
	$oKET.Interactive = $bInteractive

	Return SetError($__WPS_RETURN_SUCCESS, $bApplOpened, $oKET)
EndFunc   ;==>_WPS_Spreadsheets_Open

; #FUNCTION# ====================================================================================================================
; Name ..........: _WPS_Writer_Create
; Description ...: Connect to an existing Writer Application, or open a new one.
; Syntax ........: _WPS_Writer_Create([$bVisible = True[, $bForceNew = False]])
; Parameters ....: $bVisible            - [optional] a boolean value. Default is True. If True, the Application is visible.
;                  $bForceNew           - [optional] a boolean value. Default is False. If True, a new Application instance is created.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $bVisible not a Boolean.
;                  @Error 1 @Extended 2 Return 0 = $bForceNew not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended ? Return 0 = Failed to create WPS Writer Object. @Extended set to ObjCreate error value.
;                  --Success--
;                  @Error 0 @Extended 0 Return Object = Success. Successfully opened a WPS Writer instance. A Instance already existed, and was connected to.
;                  @Error 0 @Extended 1 Return Object = Success. Successfully opened a WPS Writer instance. A Instance did not already exist, or $bForceNew was set to True, a new instance was created.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Replaces ......: _Word_Create
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _WPS_Writer_Create($bVisible = True, $bForceNew = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __WPS_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler
_excel_close
	Local $oKWPS
	Local $bApplOpened = False

	If Not IsBool($bVisible) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)
	If Not IsBool($bForceNew) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)

	If Not $bForceNew Then $oKWPS = __WPS_Obj_Get($__WPS_TYPE_DOCUMENT)
	If $bForceNew Or @error Then
		$oKWPS = ObjCreate("KWPS.Application")
		If @error Or Not IsObj($oKWPS) Then Return SetError($__WPS_RETURN_PROCESSING_ERROR, @error, 0)
		$bApplOpened = True
	EndIf

	$oKWPS.Visible = $bVisible
	Return SetError($__WPS_RETURN_SUCCESS, $bApplOpened, $oKWPS)
EndFunc   ;==>_WPS_Writer_Create

; ============================ INTERNAL FUNCTIONS ==============================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __WPS_InternalComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __WPS_InternalComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added parameters option. Also added MsgBox & ConsoleWrite options.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __WPS_InternalComErrorHandler(ByRef $oComError)
	; If not defined ComError_UserFunction then this function does nothing, in which case you can only check @error / @extended after suspect functions.
	Local $avUserFunction = _WPS_ComError_UserFunction(Default), $avUserParams[2] = ["CallArgArray", $oComError]
	Local $vUserFunction

	If IsArray($avUserFunction) Then
		$vUserFunction = $avUserFunction[0]
		ReDim $avUserParams[UBound($avUserFunction) + 1]
		For $i = 1 To UBound($avUserFunction) - 1
			$avUserParams[$i + 1] = $avUserFunction[$i]
		Next
	Else
		$vUserFunction = $avUserFunction
	EndIf
	If IsFunc($vUserFunction) Then
		Switch $vUserFunction
			Case ConsoleWrite
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)
			Case MsgBox
				MsgBox(0, "COM Error", "Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline)
			Case Else
				Call($vUserFunction, $avUserParams)
		EndSwitch
	EndIf
EndFunc   ;==>__WPS_InternalComErrorHandler

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __WPS_Obj_Get
; Description ...: Retrieve a WPS Application Object.
; Syntax ........: __WPS_Obj_Get($iType)
; Parameters ....: $iType               - an integer value. The WPS Application type. See constants $__WPS_TYPE_*.
; Return values .: Success: Object
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $iType not an Integer.
;                  @Error 1 @Extended 2 Return 0 = $iType doesn't match predefined constants. See constants $__WPS_TYPE_*.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve requested Object.
;                  --Success--
;                  @Error 0 @Extended 1 Return Object = Success. returning requested Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __WPS_Obj_Get($iType)
	Local $iCount = 1, $iError
	Local $oAppl
	Local $sType

	If Not IsInt($iType) Then Return SetError($__WPS_RETURN_INPUT_ERROR, 1, 0)

	Switch $iType

		Case $__WPS_TYPE_DOCUMENT
			$sType = "Word.Application"

		Case $__WPS_TYPE_SPREADSHEETS
			$sType = "Excel.Application"

		Case $__WPS_TYPE_PRESENTATION
			$sType = "PowerPoint.Application"

		Case Else
			Return SetError($__WPS_RETURN_INPUT_ERROR, 2, 0)

	EndSwitch

	Do
		$oAppl = ObjGet("", $sType, $iCount)
		$iError = @error
		If IsObj($oAppl) And StringInStr(ObjName($oAppl, $OBJ_FILE), "\Kingsoft\") Then SetError($__WPS_RETURN_SUCCESS, 1, $oAppl)
		$iCount += 1
	Until $iError

	Return SetError($__WPS_RETURN_PROCESSING_ERROR, 1, 0)
EndFunc   ;==>__WPS_Obj_Get
