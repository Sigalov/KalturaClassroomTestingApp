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
$Form1_1 = GUICreate("Kaltura Classroom QA Automation Testing Tool", 360, 706, -1, -1)
GUISetBkColor(0xFFFFFF)
$RunApp = GUICtrlCreateButton("Start Kaltura Classroom", 12, 105, 162, 40)
$ExitButton = GUICtrlCreateButton("Exit", 16, 667, 74, 25)
GUICtrlSetFont(-1, 9, 800, 0, "MS Sans Serif")
$CloseApp = GUICtrlCreateButton("Close Kaltura Classroom", 181, 105, 160, 40)
$Pic1 = GUICtrlCreatePic("C:\work\KalturaClassroomTestingApp\Form\logo.jpg", 9, 9, 333, 88)
$Group2 = GUICtrlCreateGroup("Generic Flow", 16, 184, 321, 473)
$cmbAction1 = GUICtrlCreateCombo("", 32, 264, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction2 = GUICtrlCreateCombo("", 32, 296, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay1 = GUICtrlCreateInput("", 184, 264, 121, 24)
$txtActionDelay2 = GUICtrlCreateInput("", 184, 296, 121, 24)
$cmbAction3 = GUICtrlCreateCombo("", 32, 328, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay3 = GUICtrlCreateInput("", 184, 328, 121, 24)
$cmbAction4 = GUICtrlCreateCombo("", 32, 360, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay4 = GUICtrlCreateInput("", 184, 360, 121, 24)
$cmbAction5 = GUICtrlCreateCombo("", 32, 392, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay5 = GUICtrlCreateInput("", 184, 392, 121, 24)
$cmbAction6 = GUICtrlCreateCombo("", 32, 424, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay6 = GUICtrlCreateInput("", 184, 424, 121, 24)
$cmbAction7 = GUICtrlCreateCombo("", 32, 456, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction8 = GUICtrlCreateCombo("", 32, 488, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction9 = GUICtrlCreateCombo("", 32, 520, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction10 = GUICtrlCreateCombo("", 32, 552, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$cmbAction11 = GUICtrlCreateCombo("", 32, 584, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$txtActionDelay7 = GUICtrlCreateInput("", 184, 456, 121, 24)
$txtActionDelay8 = GUICtrlCreateInput("", 184, 488, 121, 24)
$txtActionDelay9 = GUICtrlCreateInput("", 184, 520, 121, 24)
$txtActionDelay10 = GUICtrlCreateInput("", 184, 552, 121, 24)
$txtActionDelay11 = GUICtrlCreateInput("", 184, 584, 121, 24)
$chkDeleteAfterRec = GUICtrlCreateCheckbox("Delete Recordings after save", 33, 623, 238, 21)
$Label1 = GUICtrlCreateLabel("Action:", 32, 232, 44, 20)
$Label2 = GUICtrlCreateLabel("Delay:", 184, 232, 43, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$chkStartGeneric = GUICtrlCreateCheckbox("START/STOP Generic Flow", 16, 157, 201, 16)
$btnTest = GUICtrlCreateButton(".", 312, 667, 27, 25)
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
Opt("SendKeyDelay", 1)
While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		Exit
	  Case $RunApp
		StartKalturaClassroom()
	  Case $CloseApp
		CloseKalturaClassroom()
	  Case $chkStartGeneric
		FlowRecordLoop()
	 Case $btnTest
		CreateGenericActionsArray()
	  Case $ExitButton
		Exit
   EndSwitch
WEnd

Func StartKalturaClassroom()
   DeletePersistencyFile();Work around - remove it after fix
   $iPID = Run('C:\Program Files\Kaltura\Classroom\CaptureApp\KalturaClassroom.exe', 'C:\Program Files\Kaltura\Classroom\CaptureApp\')
   Sleep(5000)
   ;Close the developers tools
   ;Focus on windows - click on the top
   MsgBoxCustom("Wait...", "Please focus on the window after launching the flow")
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 242, 39)
   WriteToLog("Classroom session was started")
EndFunc

Func FlowRecordLoop()
   While 1
	  If _IsChecked($chkStartGeneric) Then
		 CreateGenericActionsArray()
	  Else
		 ExitLoop
	  EndIf
   WEnd
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
   For $i = 0 to UBound($arrActions) - 1
	  If $arrActions[$i][0] <> "Select Action" Then
		 ClickGeneric($arrActions[$i][0])
	  Endif
	  If $arrActions[$i][1] <> "" Then
		 Sleep($arrActions[$i][1])
	  Endif
   Next
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
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 543, 250)
   WriteToLog("Clicked on RECORD")
EndFunc

Func ClickStop()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 384, 259)
   WriteToLog("Clicked on STOP")
EndFunc

Func ClickOnTitleField()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 181, 225)
   WriteToLog("Clicked on the TITLE filed - focus on")
EndFunc

Func ClickOnUsername()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 181, 338)
   WriteToLog("Clicked on the USERNAME filed - focus on")
EndFunc

Func ClickOnDescription()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 181, 442)
   WriteToLog("Clicked on the DESCRIPTION filed - focus on")
EndFunc

Func ClickOnTags()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 181, 555)
   WriteToLog("Clicked on the TAGS filed - focus on")
EndFunc

Func ClickSave()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 934, 137)
   WriteToLog("Clicked on SAVE")
EndFunc

Func ClickPause()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 543, 250)
   WriteToLog("Clicked on PAUSE")
EndFunc

Func ClickX()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 704, 247)
   WriteToLog("Clicked on X - Abort")
EndFunc

Func ClickCancel()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 832, 137)
   WriteToLog("Clicked on Cancel")
EndFunc

Func ClickYes()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 483, 317)
   WriteToLog("Clicked on Yes")
EndFunc

Func ClickNo()
   ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 586, 317)
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
   Local $sFilePath = "C:\Program Files\Kaltura\Classroom\Settings\persistency.json"
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