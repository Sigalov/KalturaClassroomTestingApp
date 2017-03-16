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
#Region ### START Koda GUI section ### Form=C:\work\KalturaClassroomTestingApp\Form\Classroom.kxf
$Form1_1 = GUICreate("Kaltura Classroom QA Automation Testing Tool", 355, 748, -1, -1)
$menuFile = GUICtrlCreateMenu("File")
$menuImport = GUICtrlCreateMenuItem("Import Scenario", $menuFile)
$menuExport = GUICtrlCreateMenuItem("Export Scenario", $menuFile)
$menuExit = GUICtrlCreateMenuItem("Exit", $menuFile)
GUISetIcon(".\Form\app.ico", -1)
GUISetBkColor(0xFFFFFF)
$RunApp = GUICtrlCreateButton("Start Kaltura Classroom", 12, 97, 162, 40)
$ExitButton = GUICtrlCreateButton("Exit", 16, 691, 74, 25)
GUICtrlSetFont(-1, 9, 800, 0, "MS Sans Serif")
$CloseApp = GUICtrlCreateButton("Close Kaltura Classroom", 189, 97, 160, 40)
$Pic1 = GUICtrlCreatePic(".\Form\logo.jpg", 9, 1, 333, 88)
$Group2 = GUICtrlCreateGroup("Generic Flow", 16, 200, 321, 473)
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
$btnTest = GUICtrlCreateButton(".", 312, 691, 27, 25)
$chkStartAfterCrash = GUICtrlCreateCheckbox("Resume After Crash", 16, 176, 185, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

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

Global $iPID = ""
Global $iMax=11
Global $arrActions[$iMax][2]
Global $mainWindowName = "Classroom Capture - "
Global $KalturaKlassroomProcessName = "KalturaClassroom.exe"

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
;~ 		RunAppIfNotRunning()
;~ 		testme()
	  Case $ExitButton
		Exit
	 Case $menuExit
		Exit
   EndSwitch
WEnd

Func testme()
   $aArray_2D = _ParseCSV("./scenario1.csv",",","message if error happens",true)
   ConsoleWrite($aArray_2D[0][0]& @CRLF)
   ConsoleWrite($aArray_2D[0][1]& @CRLF)
;~    ConsoleWrite($aArray_2D[1][2]& @CRLF)
;~    ConsoleWrite($aArray_2D[1][1]& @CRLF)
;~    ConsoleWrite($aArray_2D[2][2]& @CRLF)
;~    ConsoleWrite($aArray_2D[2][1]& @CRLF)
   ConsoleWrite(UBound($aArray_2D))
   For $i = 1 To UBound($aArray_2D) - 1
	  ConsoleWrite($aArray_2D[$i][1]& @CRLF)
	  ConsoleWrite($aArray_2D[$i][2]& @CRLF)
   Next


   _FileWriteFromArray("./scenario2.csv", $aArray_2D, 1)
;~    SetObjectBkColorToDefault($cmbAction1)
;~    GUICtrlSetBkColor($cmbAction1, 0xFFFFFF)
;~    GUICtrlSetState($cmbAction1, $GUI_HIDE)
;~    GUICtrlSetState($cmbAction1, $GUI_SHOW)

;~    GUICtrlSetColor($cmbAction1, 0xA6CAF0)
;~    GUICtrlSetBkColor(-1, 0xA6CAF0)
;~    If NOT ProcessExists("KalturaClassroom.exe") Then
;~ 	  MsgBoxCustom("Wait...", "Kaltura Klassroom Application is NOT running ")
;~    Else
;~ 	  MsgBoxCustom("Wait...", "Kaltura Klassroom Application is running ")
;~    Endif
EndFunc

Func StartKalturaClassroom($isPopMsg)
   DeletePersistencyFile();Work around - remove it after fix
   $iPID = Run('C:\Program Files\Kaltura\Classroom\CaptureApp\KalturaClassroom.exe', 'C:\Program Files\Kaltura\Classroom\CaptureApp\')
   ProcessWait($KalturaKlassroomProcessName)
   Sleep(10000)
   ;Close the developers tools
   ;Focus on windows - click on the top
   ControlClick($mainWindowName, "Chrome Legacy Window", "", "left", 1, 242, 39)
   WriteToLog("Classroom session was started")
   If $isPopMsg Then
	  MsgBoxCustom("Wait...", "Please focus on the window after launching the flow")
   EndIf
EndFunc

Func RunAppIfNotRunning()
   If NOT processexists($KalturaKlassroomProcessName) Then
	  WriteToLog("CRASHED: Kaltura Klassroom Application not running")
	  WriteToLog("Going to start Kaltura Klassroom Application")
	  StartKalturaClassroom(False)
   Endif
Endfunc

Func FlowRecordLoop()
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
		 $comboNameStr = "$cmbAction" & $i + 1
		 Eval(SetObjectBkColorToBlue(Execute($comboNameStr)))
		 ClickGeneric($arrActions[$i][0])
	  Endif
	  If $arrActions[$i][1] <> "" Then
		 $txtBoxNameStr = "$txtActionDelay" & $i + 1
		 Eval(SetObjectBkColorToBlue(Execute($txtBoxNameStr)))
		 Sleep($arrActions[$i][1] * 1000 + 1000)
	  Endif
	  ; Check if App crashed
	  If _IsChecked($chkStartAfterCrash) Then
		 RunAppIfNotRunning()
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
   Local $hFile = FileOpen(@ScriptDir & "\KKlassroomQaLog.log", 1)
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

Func _ParseCSV($f,$Dchar,$error,$skip)

  Local $array[$iMax][$iMax]
  Local $line = ""

  $i = 0
  $file = FileOpen($f,0)
  If $file = -1 Then
    MsgBox(0, "Error", $error)
    Return False
   EndIf

  ;skip 1st line
  If $skip Then $line = FileReadLine($file)

  While 1
       $i = $i + 1
       Local $line = FileReadLine($file)
       If @error = -1 Then ExitLoop
       $row_array = StringSplit($line,$Dchar)
        If $i == 1 Then $row_size = UBound($row_array)
        If $row_size <> UBound($row_array) Then  MsgBox(0, "Error", "Row: " & $i & " has different size ")
        $row_size = UBound($row_array)
        $array = _arrayAdd_2d($array,$i,$row_array,$row_size)

   WEnd
  FileClose($file)
  $array[0][0] = $i-1 ;stores number of lines
   $array[0][1] = $row_size -1  ; stores number of data in a row (data corresponding to index 0 is the number of data in a row actually that's way the -1)
   Return $array

EndFunc
Func _arrayAdd_2d($array,$inwhich,$row_array,$row_size)
    For $i=1 To $row_size -1 Step 1
        $array[$inwhich][$i] = $row_array[$i]
  Next
  Return $array
EndFunc
