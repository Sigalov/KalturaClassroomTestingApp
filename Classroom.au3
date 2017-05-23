#include <GuiComboBox.au3>
#include <Array.au3>
#include <File.au3>
#include <Date.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=c:\work\kalturaclassroomtestingapp\form\classroom.kxf
$Form1_1 = GUICreate("Kaltura Classroom QA Automation Testing Tool", 358, 784, -1, -1)
$menuFile = GUICtrlCreateMenu("&File")
$menuImport = GUICtrlCreateMenuItem("Import Scenario", $menuFile)
$menuExport = GUICtrlCreateMenuItem("Export Scenario", $menuFile)
$menuExit = GUICtrlCreateMenuItem("Exit", $menuFile)
GUISetIcon(".\Form\app.ico", -1)
GUISetBkColor(0xFFFFFF)
$RunApp = GUICtrlCreateButton("Start Kaltura Classroom", 12, 97, 162, 40)
$ExitButton = GUICtrlCreateButton("Exit", 16, 723, 74, 25)
GUICtrlSetFont(-1, 9, 800, 0, "MS Sans Serif")
$CloseApp = GUICtrlCreateButton("Close Kaltura Classroom", 189, 97, 160, 40)
$Pic1 = GUICtrlCreatePic(".\Form\logo.jpg", 9, 1, 333, 88)
$Group2 = GUICtrlCreateGroup("Generic Flow", 16, 200, 305, 473)
$cmbAction1 = GUICtrlCreateCombo("", 32, 280, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction2 = GUICtrlCreateCombo("", 32, 312, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay1 = GUICtrlCreateInput("", 184, 280, 121, 24)
$txtActionDelay2 = GUICtrlCreateInput("", 184, 312, 121, 24)
$cmbAction3 = GUICtrlCreateCombo("", 32, 344, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay3 = GUICtrlCreateInput("", 184, 344, 121, 24)
$cmbAction4 = GUICtrlCreateCombo("", 32, 376, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay4 = GUICtrlCreateInput("", 184, 376, 121, 24)
$cmbAction5 = GUICtrlCreateCombo("", 32, 408, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay5 = GUICtrlCreateInput("", 184, 408, 121, 24)
$cmbAction6 = GUICtrlCreateCombo("", 32, 440, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay6 = GUICtrlCreateInput("", 184, 440, 121, 24)
$cmbAction7 = GUICtrlCreateCombo("", 32, 472, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction8 = GUICtrlCreateCombo("", 32, 504, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction9 = GUICtrlCreateCombo("", 32, 536, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction10 = GUICtrlCreateCombo("", 32, 568, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction11 = GUICtrlCreateCombo("", 32, 600, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay7 = GUICtrlCreateInput("", 184, 472, 121, 24)
$txtActionDelay8 = GUICtrlCreateInput("", 184, 504, 121, 24)
$txtActionDelay9 = GUICtrlCreateInput("", 184, 536, 121, 24)
$txtActionDelay10 = GUICtrlCreateInput("", 184, 568, 121, 24)
$txtActionDelay11 = GUICtrlCreateInput("", 184, 600, 121, 24)
$chkDeleteAfterRec = GUICtrlCreateCheckbox("Delete Recordings after save", 33, 639, 238, 21)
$Label1 = GUICtrlCreateLabel("Action:", 32, 248, 44, 20)
$Label2 = GUICtrlCreateLabel("Delay:", 184, 248, 43, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$chkStartGeneric = GUICtrlCreateCheckbox("START/STOP Generic Flow", 16, 149, 201, 16)
$btnTest = GUICtrlCreateButton(".", 296, 723, 27, 25)
$chkStartAfterCrash = GUICtrlCreateCheckbox("Resume After Crash", 16, 176, 185, 17)
$txtFrameTitle = GUICtrlCreateInput("", 112, 688, 121, 24)
$Label3 = GUICtrlCreateLabel("Frame Title:", 16, 688, 87, 20)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$btnUpdateFrameTitle = GUICtrlCreateButton("Update", 240, 688, 83, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $iPID = ""
Global $iMax=11
Global $arrActions[$iMax][2]
Global $mainWindowName = "Classroom Capture - "
Global $KalturaKlassroomProcessName = "KalturaClassroom.exe"

;Create actions
GUICtrlSetData($cmbAction1, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction2, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction3, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction4, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction5, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction6, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction7, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction8, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction9, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction10, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")
GUICtrlSetData($cmbAction11, "Select Action|Click Record|Click Pause|Click Stop|Click Abort - X|Click Save|Enter Title|Enter Username|Enter Description|Enter Tags|Click Cancel|Click Yes|Click No", "Select Action")

; Set Frame Title name
GUICtrlSetData($txtFrameTitle, $mainWindowName)


Opt("SendKeyDelay", 1)
While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		Exit
	  Case $RunApp
		StartKalturaClassroom(True)
	  Case $CloseApp
		CloseKalturaClassroom()
	  Case $chkStartGeneric
		 FlowRecordLoop()
	  Case $btnTest
;~ 		testme()
	  Case $ExitButton
		Exit
	  Case $btnUpdateFrameTitle
		 UpdateFrameTitle()
	  Case $menuExit
		Exit
	  Case $menuExport
		 ExportScenario()
	  Case $menuImport
		 ImportScenario()
   EndSwitch
WEnd

Func StartKalturaClassroom($isPopMsg)
   DeletePersistencyFile();Work around - remove it after fix
   $iPID = Run('C:\Program Files\Kaltura\Classroom\CaptureApp\KalturaClassroom.exe', 'C:\Program Files\Kaltura\Classroom\CaptureApp\')
   ProcessWait($KalturaKlassroomProcessName)
   ;Focus on windows - click on the top
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 242, 39)
   WriteToLog("Classroom session was started")
EndFunc

Func RunAppIfNotRunning()
   If NOT processexists($KalturaKlassroomProcessName) Then
	  WriteToLog("CRASHED: Kaltura Klassroom Application not running")
	  WriteToLog("Going to start Kaltura Klassroom Application")
	  StartKalturaClassroom(False)
	  Sleep(8000)
	  Return True
   Else
	  Return False
   Endif
Endfunc

Func FlowRecordLoop()
   ; Auto-Save current flow
   If _IsChecked($chkStartGeneric) Then
	  $NOW = _NowCalc()
	  $NOW = StringReplace($NOW, " ", "-")
	  $NOW = StringReplace($NOW, "/", "-")
	  $NOW = StringReplace($NOW, ":", "-")
	  $fileName = "./" & $NOW & "_currentScen.csv"
	  AutoSave($fileName)
   EndIf
   While 1
	  If _IsChecked($chkStartGeneric) Then
		 StartGenericFlow()
	  Else
		 ExitLoop
	  EndIf
   WEnd
EndFunc

Func StartGenericFlow()
   CreateGenericActionsArray()
   WriteToLog("Going to Start the flow")
   For $i = 0 to UBound($arrActions) - 1
	  If $arrActions[$i][0] <> "Select Action" Then
		 If _IsChecked($chkStartGeneric) Then
			$comboNameStr = "$cmbAction" & $i + 1
			Eval(SetObjectBkColorToBlue(Execute($comboNameStr)))
			ClickGeneric($arrActions[$i][0])
		 EndIf
	  Endif
	  If $arrActions[$i][1] <> "" Then
		 If _IsChecked($chkStartGeneric) Then
			$txtBoxNameStr = "$txtActionDelay" & $i + 1
			Eval(SetObjectBkColorToBlue(Execute($txtBoxNameStr)))
			WriteToLog("Delay " & $arrActions[$i][1])
			Sleep($arrActions[$i][1] * 1000 + 1000)
		 EndIf
	  Endif
	  ; Check if App crashed
	  If _IsChecked($chkStartAfterCrash) Then
		 If RunAppIfNotRunning() = True Then
			Eval(SetObjectBkColorToDefault(Execute($comboNameStr)))
			Eval(SetObjectBkColorToDefault(Execute($txtBoxNameStr)))
			Return
		 EndIf
	  EndIf
	  Eval(SetObjectBkColorToDefault(Execute($comboNameStr)))
	  Eval(SetObjectBkColorToDefault(Execute($txtBoxNameStr)))
   Next
   WriteToLog("Flow Cycle ENDED")
EndFunc

Func CreateGenericActionsArray()
   Global $arrActions[$iMax][2] = [ [GUICtrlRead($cmbAction1), GUICtrlRead($txtActionDelay1)], _
									[GUICtrlRead($cmbAction2), GUICtrlRead($txtActionDelay2)], _
									[GUICtrlRead($cmbAction3), GUICtrlRead($txtActionDelay3)], _
									[GUICtrlRead($cmbAction4), GUICtrlRead($txtActionDelay4)], _
									[GUICtrlRead($cmbAction5), GUICtrlRead($txtActionDelay5)], _
									[GUICtrlRead($cmbAction6), GUICtrlRead($txtActionDelay6)], _
									[GUICtrlRead($cmbAction7), GUICtrlRead($txtActionDelay7)], _
									[GUICtrlRead($cmbAction8), GUICtrlRead($txtActionDelay8)], _
									[GUICtrlRead($cmbAction9), GUICtrlRead($txtActionDelay9)], _
									[GUICtrlRead($cmbAction10), GUICtrlRead($txtActionDelay10)], _
									[GUICtrlRead($cmbAction11), GUICtrlRead($txtActionDelay11)] _
								  ]
EndFunc

Func ClickGeneric($ClickType)
   Switch $ClickType
	  Case "Click Record"
		 ClickRecord()
	  Case "Click Pause"
		 ClickPause()
	  Case "Click Stop"
		 ClickStop()
	  Case "Click Abort - X"
		 ClickX()
	  Case "Click Save"
		 ClickCloseHintHoveringOverSaveButton()
		 ClickSave()
	  Case "Enter Title"
		 ClickOnTitleField()
		 Sleep(1000)
		 SetFileName("Title")
	  Case "Enter Username"
		 ClickOnUsername()
		 Sleep(1000)
		 SetFileName("Username")
	  Case "Enter Description"
		 ClickOnDescription()
		 Sleep(1000)
		 SetFileName("Description")
	  Case "Enter Tags"
 		 ClickOnTags()
		 Sleep(1000)
		 SetFileName("Tags")
	  Case "Click Cancel"
		 ClickCancel()
	  Case "Click Yes"
		 ClickYes()
	  Case "Click No"
		 ClickNo()
   EndSwitch
EndFunc

Func ClickRecord()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 543, 250)
   WriteToLog("Clicked on RECORD")
EndFunc

Func ClickStop()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 384, 259)
   WriteToLog("Clicked on STOP")
EndFunc

Func ClickOnTitleField()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 181, 225)
   WriteToLog("Clicked on the TITLE filed - focus on")
EndFunc

Func ClickOnUsername()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 181, 338)
   WriteToLog("Clicked on the USERNAME filed - focus on")
EndFunc

Func ClickOnDescription()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 181, 442)
   WriteToLog("Clicked on the DESCRIPTION filed - focus on")
EndFunc

Func ClickOnTags()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 181, 555)
   WriteToLog("Clicked on the TAGS filed - focus on")
EndFunc

Func ClickSave()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 934, 137)
   WriteToLog("Clicked on SAVE")
EndFunc

Func ClickPause()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 543, 250)
   WriteToLog("Clicked on PAUSE")
EndFunc

Func ClickX()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 704, 247)
   WriteToLog("Clicked on X - Abort")
EndFunc

Func ClickCancel()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 832, 137)
   WriteToLog("Clicked on Cancel")
EndFunc

Func ClickYes()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 483, 317)
   WriteToLog("Clicked on Yes")
EndFunc

Func ClickNo()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 586, 317)
   WriteToLog("Clicked on No")
EndFunc

Func ClickCloseHintHoveringOverSaveButton()
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 1005, 125)
   Sleep(1500)
   WriteToLog("Clicked 'X' - Closed hint hovering over save button (click if exist and if not)")
EndFunc

Func SetFileName($prefix)
   $NOW = _NowCalc()
   $logname = $prefix & "-" & $NOW
   Send($logname)
   WriteToLog("Text entered: " & String($logname))
EndFunc

Func CloseKalturaClassroom()
	ProcessClose($iPID)
	WriteToLog("Classroom session was closed")
EndFunc

Func MsgBoxCustom($title, $msg)
   ; Display a message box with a nested variable in its text.
   MsgBox($MB_OK, $title, $msg)
EndFunc

Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc

Func WriteToLog($text)
   Local $hFile = FileOpen(@ScriptDir & "\KClassroomQaLog.log", 1)
   _FileWriteLog($hFile, $text) ; Write to the logfile passing the filehandle returned by FileOpen.
   FileClose($hFile) ; Close the filehandle to release the file.
EndFunc

; Temp workaround 21/2/2017 18:47
Func DeletePersistencyFile()
   Local $sFilePath = "C:\Program Files\Kaltura\Classroom\Settings\appPersistency.json"
   DeleteFile($sFilePath)
EndFunc

Func DeleteRecording()
	If _IsChecked($chkDeleteAfterRec) Then
		Local $sFilePath = "C:\Program Files\Kaltura\Classroom\Recordings\"
		DeleteFile($sFilePath)
	EndIf
EndFunc

Func DeleteFile($sFilePath)
   ; Delete the temporary file.
   Local $iDelete = FileDelete($sFilePath)

   ; Display a message of whether the file was deleted.
   If $iDelete Then
	 WriteToLog("File was successfuly deleted: " & $sFilePath)
  Else
	 WriteToLog("An error occurred whilst deleting the file: " & $sFilePath)
  EndIf
EndFunc

Func SetObjectBkColorToBlue($objectId)
   GUICtrlSetBkColor($objectId, 0xA6CAF0)
   GUICtrlSetState($objectId, $GUI_HIDE)
   GUICtrlSetState($objectId, $GUI_SHOW)
EndFunc

Func SetObjectBkColorToDefault($objectId)
   GUICtrlSetBkColor($objectId, 0xFFFFFF)
   GUICtrlSetState($objectId, $GUI_HIDE)
   GUICtrlSetState($objectId, $GUI_SHOW)
EndFunc

Func UpdateFrameTitle()
   Global $mainWindowName = GUICtrlRead($txtFrameTitle)
EndFunc

Func ImportScenario()
   ; Choose scenario to load
   $filePath = LoadFilePicker()
   Local $hFile = FileOpen($filePath)
   If $hFile = -1 Then
	  MsgBox(0, "", "Unable to open file")
	  Exit
   EndIf

   Local $sCSV = FileRead($hFile)
   If @error Then
	  MsgBox(0, "", "Unable to read file")
	  FileClose($hFile)
	  Exit
   EndIf
   FileClose($hFile)
   Local $aCSV = _CSVSplit($sCSV) ; Create the main array

   ; Update the GUI with the steps from loaded scenario
   For $i = 0 to UBound($aCSV) - 1
	  Eval(_GUICtrlComboBox_SelectString(Execute("$cmbAction" & $i + 1), $aCSV[$i][0]))
	  Eval(GUICtrlSetData(Execute("$txtActionDelay" & $i + 1), $aCSV[$i][1]))
   Next
EndFunc

Func ExportScenario()
   ; Create an array from the form
   CreateGenericActionsArray()
   ; Create temp array
   $exportArr = $arrActions
   ; Select file to save
   $filePath = SaveAsPicker()
   ; Make sure file exists/create if not. Overwrite if needed
   Local $hFile = FileOpen($filePath, 2)
   If $hFile = -1 Then
	  MsgBox(0, "", "Unable to open file")
	  Exit
   EndIf

   Local $sCSV = FileRead($hFile)
   If @error Then
	  MsgBox(0, "", "Unable to read file")
	  FileClose($hFile)
	  Exit
   EndIf
   FileClose($hFile)

   ; Write to file
   FileWrite($filePath,  _ArrayToCSV($exportArr, Default, @CRLF, False))

   ; Display the saved file.
   MsgBox($MB_SYSTEMMODAL, "", "You exported current scenario to file:" & @CRLF & $filePath)
EndFunc

Func AutoSave($filePath)
   ; Create an array from the form
   CreateGenericActionsArray()
   ; Create temp array
   $exportArr = $arrActions
   ; Make sure file exists/create if not. Overwrite if needed
   Local $hFile = FileOpen($filePath, 2)
   If $hFile = -1 Then
	  MsgBox(0, "", "Unable to open file")
	  Exit
   EndIf

   Local $sCSV = FileRead($hFile)
   If @error Then
	  MsgBox(0, "", "Unable to read file")
	  FileClose($hFile)
	  Exit
   EndIf
   FileClose($hFile)

   ; Write to file
   FileWrite($filePath,  _ArrayToCSV($exportArr, Default, @CRLF, False))
   WriteToLog("You exported current scenario to file:" & $filePath)
EndFunc

Func SaveAsPicker()
   ; Create a constant variable in Local scope of the message to display in FileSaveDialog.
   Local Const $sMessage = "Choose a filename."

   ; Display a save dialog to select a file.
   Local $sFileSaveDialog = FileSaveDialog($sMessage, "::{450D8FBA-AD25-11D0-98A8-0800361B1103}", "Scenario (*.csv)", 18)
   If @error Then
	  ; Display the error message.
	  MsgBox($MB_SYSTEMMODAL, "", "No file was saved.")
   Else
	  ; Retrieve the filename from the filepath e.g. Example.csv.
	  Local $sFileName = StringTrimLeft($sFileSaveDialog, StringInStr($sFileSaveDialog, "\", $STR_NOCASESENSE, -1))

	  ; Check if the extension .csv is appended to the end of the filename.
	  Local $iExtension = StringInStr($sFileName, ".", $STR_NOCASESENSE)

	  ; If a period (dot) is found then check whether or not the extension is equal to .csv
	  If $iExtension Then
		 ; If the extension isn't equal to .csv then append to the end of the filepath.
		 If Not (StringTrimLeft($sFileName, $iExtension - 1) = ".csv") Then $sFileSaveDialog &= ".csv"
	  Else
		 ; If no period (dot) was found then append to the end of the file.
		 $sFileSaveDialog &= ".csv"
	  EndIf
	  Return $sFileSaveDialog
   EndIf
EndFunc

Func LoadFilePicker()
   ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
   Local Const $sMessage = "Hold down Ctrl or Shift to choose multiple files."

   ; Display an open dialog to select a list of file(s).
   Local $sFileOpenDialog = FileOpenDialog($sMessage, @WindowsDir & "\", "Scenario (*.csv)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
   If @error Then
      ; Display the error message.
      MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")

      ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
      FileChangeDir(@ScriptDir)
   Else
      ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
      FileChangeDir(@ScriptDir)

      ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
      $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

	  Return $sFileOpenDialog
   EndIf
EndFunc


; #INDEX# =======================================================================================================================
; Title .........: CSVSplit
; AutoIt Version : 3.3.8.1
; Language ......: English
; Description ...: CSV related functions
; Notes .........: CSV format does not have a general standard format, however these functions allow some flexibility.
;                  The default behaviour of the functions applies to the most common formats used in practice.
; Author(s) .....: czardas
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;_ArrayToCSV
;_ArrayToSubItemCSV
;_CSVSplit
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
; __GetSubstitute
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _ArrayToCSV
; Description ...: Converts a two dimensional array to CSV format
; Syntax.........: _ArrayToCSV ( $aArray [, $sDelim [, $sNewLine [, $bFinalBreak ]]] )
; Parameters ....: $aArray      - The array to convert
;                  $sDelim      - Optional - Delimiter set to comma by default (see comments)
;                  $sNewLine    - Optional - New Line set to @LF by default (see comments)
;                  $bFinalBreak - Set to true in accordance with common practice => CSV Line termination
; Return values .: Success  - Returns a string in CSV format
;                  Failure  - Sets @error to:
;                 |@error = 1 - First parameter is not a valid array
;                 |@error = 2 - Second parameter is not a valid string
;                 |@error = 3 - Third parameter is not a valid string
;                 |@error = 4 - 2nd and 3rd parameters must be different characters
; Author ........: czardas
; Comments ......; One dimensional arrays are returned as multiline text (without delimiters)
;                ; Some users may need to set the second parameter to semicolon to return the prefered CSV format
;                ; To convert to TSV use @TAB for the second parameter
;                ; Some users may wish to set the third parameter to @CRLF
; ===============================================================================================================================

Func _ArrayToCSV($aArray, $sDelim = Default, $sNewLine = Default, $bFinalBreak = True)
    If Not IsArray($aArray) Or Ubound($aArray, 0) > 2 Or Ubound($aArray) = 0 Then Return SetError(1, 0 ,"")
    If $sDelim = Default Then $sDelim = ","
    If $sDelim = "" Then Return SetError(2, 0 ,"")
    If $sNewLine = Default Then $sNewLine = @LF
    If $sNewLine = "" Then Return SetError(3, 0 ,"")
    If $sDelim = $sNewLine Then Return SetError(4, 0, "")

    Local $iRows = UBound($aArray), $sString = ""
    If Ubound($aArray, 0) = 2 Then ; Check if the array has two dimensions
        Local $iCols = UBound($aArray, 2)
        For $i = 0 To $iRows -1
            For $j = 0 To $iCols -1
                If StringRegExp($aArray[$i][$j], '["\r\n' & $sDelim & ']') Then
                    $aArray[$i][$j] = '"' & StringReplace($aArray[$i][$j], '"', '""') & '"'
                EndIf
                $sString &= $aArray[$i][$j] & $sDelim
            Next
            $sString = StringTrimRight($sString, StringLen($sDelim)) & $sNewLine
        Next
    Else ; The delimiter is not needed
        For $i = 0 To $iRows -1
            If StringRegExp($aArray[$i], '["\r\n' & $sDelim & ']') Then
                $aArray[$i] = '"' & StringReplace($aArray[$i], '"', '""') & '"'
            EndIf
            $sString &= $aArray[$i] & $sNewLine
        Next
    EndIf
    If Not $bFinalBreak Then $sString = StringTrimRight($sString, StringLen($sNewLine)) ; Delete any newline characters added to the end of the string
    Return $sString
EndFunc ;==> _ArrayToCSV

; #FUNCTION# ====================================================================================================================
; Name...........: _ArrayToSubItemCSV
; Description ...: Converts an array to multiple CSV formated strings based on the content of the selected column
; Syntax.........: _ArrayToSubItemCSV($aCSV, $iCol [, $sDelim [, $bHeaders [, $iSortCol [, $bAlphaSort ]]]])
; Parameters ....: $aCSV       - The array to parse
;                  $iCol       - Array column used to search for unique content
;                  $sDelim     - Optional - Delimiter set to comma by default
;                  $bHeaders   - Include csv column headers - Default = False
;                  $iSortCol   - The column to sort on for each new CSV (sorts ascending) - Default = False
;                  $bAlphaSort - If set to true, sorting will be faster but numbers won't always appear in order of magnitude.
; Return values .: Success   - Returns a two dimensional array - col 0 = subitem name, col 1 = CSV data
;                  Failure   - Returns an empty string and sets @error to:
;                 |@error = 1 - First parameter is not a 2D array
;                 |@error = 2 - Nothing to parse
;                 |@error = 3 - Invalid second parameter Column number
;                 |@error = 4 - Invalid third parameter - Delimiter is an empty string
;                 |@error = 5 - Invalid fourth parameter - Sort Column number is out of range
; Author ........: czardas
; Comments ......; @CRLF is used for line breaks in the returned array of CSV strings.
;                ; Data in the sorting column is automatically assumed to contain numeric values.
;                ; Setting $iSortCol equal to $iCol will return csv rows in their original ordered sequence.
; ===============================================================================================================================

Func _ArrayToSubItemCSV($aCSV, $iCol, $sDelim = Default, $bHeaders = Default, $iSortCol = Default, $bAlphaSort = Default)
    If Not IsArray($aCSV) Or UBound($aCSV, 0) <> 2 Then Return SetError(1, 0, "") ; Not a 2D array

    Local $iBound = UBound($aCSV), $iNumCols = UBound($aCSV, 2)
    If $iBound < 2 Then Return SetError(2, 0, "") ; Nothing to parse

    If IsInt($iCol) = 0 Or $iCol < 0 Or $iCol > $iNumCols -1 Then Return SetError(3, 0, "") ; $iCol is out of range

    If $sDelim = Default Then $sDelim = ","
    If $sDelim = "" Then Return SetError(4, 0, "") ; Delimiter can not be an empty string

    If $bHeaders = Default Then $bHeaders = False

    If $iSortCol = Default Or $iSortCol == False Then $iSortCol = -1
    If IsInt($iSortCol) = 0 Or $iSortCol < -1 Or $iSortCol > $iNumCols -1 Then Return SetError(5, 0, "") ; $iSortCol is out of range

    If $bAlphaSort = Default Then $bAlphaSort = False

    Local $iStart = 0

    If $bHeaders Then
        If $iBound = 2 Then Return SetError(2, 0, "") ; Nothing to parse
        $iStart = 1
    EndIf

    Local $sTestItem, $iNewCol = 0

    If $iSortCol <> -1 And ($bAlphaSort = False Or $iSortCol = $iCol) Then ; In this case we need an extra Column for sorting
        ReDim $aCSV [$iBound][$iNumCols +1]

        ; Populate column
        If $iSortCol = $iCol Then
            For $i = $iStart To $iBound -1
                $aCSV[$i][$iNumCols] = $i
            Next
        Else
            For $i = $iStart To $iBound -1
                $sTestItem = StringRegExpReplace($aCSV[$i][$iSortCol], "\A\h+", "") ; Remove leading horizontal WS
                If StringIsInt($sTestItem) Or StringIsFloat($sTestItem) Then
                    $aCSV[$i][$iNumCols] = Number($sTestItem)
                Else
                    $aCSV[$i][$iNumCols] = $aCSV[$i][$iSortCol]
                EndIf
            Next
        EndIf
        $iNewCol = 1
        $iSortCol = $iNumCols
    EndIf

    _ArraySort($aCSV, 0, $iStart, 0, $iCol) ; Sort on the selected column

    Local $aSubItemCSV[$iBound][2], $iItems = 0, $aTempCSV[1][$iNumCols + $iNewCol], $iTempIndex
    $sTestItem = Not $aCSV[$iBound -1][$iCol]

    For $i = $iBound -1 To $iStart Step -1
        If $sTestItem <> $aCSV[$i][$iCol] Then ; Start a new csv instance
            If $iItems > 0 Then ; Write to main array
                ReDim $aTempCSV[$iTempIndex][$iNumCols + $iNewCol]

                If $iSortCol <> -1 Then _ArraySort($aTempCSV, 0, $iStart, 0, $iSortCol)
                If $iNewCol Then ReDim $aTempCSV[$iTempIndex][$iNumCols]

                $aSubItemCSV[$iItems -1][0] = $sTestItem
                $aSubItemCSV[$iItems -1][1] = _ArrayToCSV($aTempCSV, $sDelim, @CRLF)
            EndIf

            ReDim $aTempCSV[$iBound][$iNumCols + $iNewCol] ; Create new csv template
            $iTempIndex = 0
            $sTestItem = $aCSV[$i][$iCol]

            If $bHeaders Then
                For $j = 0 To $iNumCols -1
                    $aTempCSV[0][$j] = $aCSV[0][$j]
                Next
                $iTempIndex = 1
            EndIf
            $iItems += 1
        EndIf

        For $j = 0 To $iNumCols + $iNewCol -1 ; Continue writing to csv
            $aTempCSV[$iTempIndex][$j] = $aCSV[$i][$j]
        Next
        $iTempIndex += 1
    Next

    ReDim $aTempCSV[$iTempIndex][$iNumCols + $iNewCol]

    If $iSortCol <> -1 Then _ArraySort($aTempCSV, 0, $iStart, 0, $iSortCol)
    If $iNewCol Then ReDim $aTempCSV[$iTempIndex][$iNumCols]

    $aSubItemCSV[$iItems -1][0] = $sTestItem
    $aSubItemCSV[$iItems -1][1] = _ArrayToCSV($aTempCSV, $sDelim, @CRLF)

    ReDim $aSubItemCSV[$iItems][2]
    Return $aSubItemCSV
EndFunc ;==> _ArrayToSubItemCSV

; #FUNCTION# ====================================================================================================================
; Name...........: _CSVSplit
; Description ...: Converts a string in CSV format to a two dimensional array (see comments)
; Syntax.........: CSVSplit ( $aArray [, $sDelim ] )
; Parameters ....: $aArray  - The array to convert
;                  $sDelim  - Optional - Delimiter set to comma by default (see 2nd comment)
; Return values .: Success  - Returns a two dimensional array or a one dimensional array (see 1st comment)
;                  Failure  - Sets @error to:
;                 |@error = 1 - First parameter is not a valid string
;                 |@error = 2 - Second parameter is not a valid string
;                 |@error = 3 - Could not find suitable delimiter replacements
; Author ........: czardas
; Comments ......; Returns a one dimensional array if the input string does not contain the delimiter string
;                ; Some CSV formats use semicolon as a delimiter instead of a comma
;                ; Set the second parameter to @TAB To convert to TSV
; ===============================================================================================================================

Func _CSVSplit($string, $sDelim = ",") ; Parses csv string input and returns a one or two dimensional array
    If Not IsString($string) Or $string = "" Then Return SetError(1, 0, 0) ; Invalid string
    If Not IsString($sDelim) Or $sDelim = "" Then Return SetError(2, 0, 0) ; Invalid string

    $string = StringRegExpReplace($string, "[\r\n]+\z", "") ; [Line Added] Remove training breaks
    Local $iOverride = 63743, $asDelim[3] ; $asDelim => replacements for comma, new line and double quote
    For $i = 0 To 2
        $asDelim[$i] = __GetSubstitute($string, $iOverride) ; Choose a suitable substitution character
        If @error Then Return SetError(3, 0, 0) ; String contains too many unsuitable characters
    Next
    $iOverride = 0

    Local $aArray = StringRegExp($string, '\A[^"]+|("+[^"]+)|"+\z', 3) ; Split string using double quotes delim - largest match
    $string = ""

    Local $iBound = UBound($aArray)
    For $i = 0 To $iBound -1
        $iOverride += StringInStr($aArray[$i], '"', 0, -1) ; Increment by the number of adjacent double quotes per element
        If Mod ($iOverride +2, 2) = 0 Then ; Acts as an on/off switch
            $aArray[$i] = StringReplace($aArray[$i], $sDelim, $asDelim[0]) ; Replace comma delimeters
            $aArray[$i] = StringRegExpReplace($aArray[$i], "(\r\n)|[\r\n]", $asDelim[1]) ; Replace new line delimeters
        EndIf
        $aArray[$i] = StringReplace($aArray[$i], '""', $asDelim[2]) ; Replace double quote pairs
        $aArray[$i] = StringReplace($aArray[$i], '"', '') ; Delete enclosing double quotes - not paired
        $aArray[$i] = StringReplace($aArray[$i], $asDelim[2], '"') ; Reintroduce double quote pairs as single characters
        $string &= $aArray[$i] ; Rebuild the string, which includes two different delimiters
    Next
    $iOverride = 0

    $aArray = StringSplit($string, $asDelim[1], 2) ; Split to get rows
    $iBound = UBound($aArray)
    Local $aCSV[$iBound][2], $aTemp
    For $i = 0 To $iBound -1
        $aTemp = StringSplit($aArray[$i], $asDelim[0]) ; Split to get row items
        If Not @error Then
            If $aTemp[0] > $iOverride Then
                $iOverride = $aTemp[0]
                ReDim $aCSV[$iBound][$iOverride] ; Add columns to accomodate more items
            EndIf
        EndIf
        For $j = 1 To $aTemp[0]
            If StringLen($aTemp[$j]) Then
                If Not StringRegExp($aTemp[$j], '[^"]') Then ; Field only contains double quotes
                    $aTemp[$j] = StringTrimLeft($aTemp[$j], 1) ; Delete enclosing double quote single char
                EndIf
                $aCSV[$i][$j -1] = $aTemp[$j] ; Populate each row
            EndIf
        Next
    Next

    If $iOverride > 1 Then
        Return $aCSV ; Multiple Columns
    Else
        For $i = 0 To $iBound -1
            If StringLen($aArray[$i]) And (Not StringRegExp($aArray[$i], '[^"]')) Then ; Only contains double quotes
                $aArray[$i] = StringTrimLeft($aArray[$i], 1) ; Delete enclosing double quote single char
            EndIf
        Next
        Return $aArray ; Single column
    EndIf

EndFunc ;==> _CSVSplit

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __GetSubstitute
; Description ...: Searches for a character to be used for substitution, ie one not contained within the input string
; Syntax.........: __GetSubstitute($string, ByRef $iCountdown)
; Parameters ....: $string   - The string of characters to avoid
;                  $iCountdown - The first code point to begin checking
; Return values .: Success   - Returns a suitable substitution character not found within the first parameter
;                  Failure   - Sets @error to 1 => No substitution character available
; Author ........: czardas
; Comments ......; This function is connected to the function _CSVSplit and was not intended for general use
;                  $iCountdown is returned ByRef to avoid selecting the same character on subsequent calls to this function
;                  Initially $iCountown should be passed with a value = 63743
; ===============================================================================================================================

Func __GetSubstitute($string, ByRef $iCountdown)
    If $iCountdown < 57344 Then Return SetError(1, 0, "") ; Out of options
    Local $sTestChar
    For $i = $iCountdown To 57344 Step -1
        $sTestChar = ChrW($i)
        $iCountdown -= 1
        If Not StringInStr($string, $sTestChar) Then
            Return $sTestChar
        EndIf
    Next
    Return SetError(1, 0, "") ; Out of options
EndFunc ;==> __GetSubstitute

Func testme()
   ; Choose scenario to load
   $filePath = LoadFilePicker()
   Local $hFile = FileOpen($filePath)
   If $hFile = -1 Then
	  MsgBox(0, "", "Unable to open file")
	  Exit
   EndIf

   Local $sCSV = FileRead($hFile)
   If @error Then
	  MsgBox(0, "", "Unable to read file")
	  FileClose($hFile)
	  Exit
   EndIf
   FileClose($hFile)
   Local $aCSV = _CSVSplit($sCSV) ; Create the main array
   _ArrayDisplay($aCSV)

   For $i = 0 to UBound($aCSV) - 1
	  Eval(_GUICtrlComboBox_SelectString(Execute("$cmbAction" & $i + 1), $aCSV[$i][0]))
	  Eval(GUICtrlSetData(Execute("$txtActionDelay" & $i + 1), $aCSV[$i][1]))
   Next
EndFunc