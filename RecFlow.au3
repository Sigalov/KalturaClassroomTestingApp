#include <File.au3>
#include <Date.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=.\form\recflowform.kxf
$Form1_1 = GUICreate("Kaltura Classroom QA Automation Testing Tool", 288, 430, -1, -1)
GUISetBkColor(0xFFFFFF)
$RunApp = GUICtrlCreateButton("Start Kaltura Classroom", 10, 85, 131, 33)
$ExitButton = GUICtrlCreateButton("Exit", 15, 386, 60, 20)
GUICtrlSetFont(-1, 7, 800, 0, "MS Sans Serif")
$CloseApp = GUICtrlCreateButton("Close Kaltura Classroom", 147, 85, 130, 33)
$Pic1 = GUICtrlCreatePic(".\Form\logo.jpg", 7, 7, 271, 72)
$chkStartFlow1 = GUICtrlCreateCheckbox("START", 13, 137, 53, 13)
$Group1 = GUICtrlCreateGroup("Record Flow:", 13, 156, 257, 188)
$Label2 = GUICtrlCreateLabel("Click Record", 26, 176, 65, 17)
$Label3 = GUICtrlCreateLabel("Record Sleep", 26, 195, 69, 17)
$Label4 = GUICtrlCreateLabel("Click Stop", 26, 215, 52, 17)
$txtFlowRecInterval = GUICtrlCreateInput("", 124, 189, 98, 21)
$Label5 = GUICtrlCreateLabel("Wait for title Sleep", 26, 234, 90, 17)
$txtTitleSleep = GUICtrlCreateInput("", 124, 228, 98, 21)
$Label6 = GUICtrlCreateLabel("Click On Title Field", 26, 254, 92, 17)
$Label7 = GUICtrlCreateLabel("Click Save", 26, 273, 55, 17)
$Label8 = GUICtrlCreateLabel("Sleep after Save", 26, 293, 83, 17)
$txtSleepAfterSave = GUICtrlCreateInput("", 124, 286, 98, 21)
$chkDeleteAfterRec = GUICtrlCreateCheckbox("Delete Recordings after save", 26, 312, 193, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnTest = GUICtrlCreateLabel("", 212, 372, 52, 36)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


Global $iPID = ""

While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
   Case $GUI_EVENT_CLOSE
		Exit
	  Case $RunApp
		StartKalturaClassroom()
	  Case $CloseApp
		CloseKalturaClassroom()
	  Case $chkStartFlow1
		FlowRecordLoop(GUICtrlRead($txtFlowRecInterval))
	   Case $btnTest
		DeleteRecording()
		WriteToLog("TEST")
	  Case $ExitButton
		Exit
	EndSwitch
WEnd

Func StartKalturaClassroom()
	DeletePersistencyFile();Work around - remove it after fix
	$iPID = Run('C:\Program Files\Kaltura\Classroom\CaptureApp\KalturaClassroom.exe', 'C:\Program Files\Kaltura\Classroom\CaptureApp\')
	Sleep(5000)
	;Close the developers tools

	WriteToLog("Classroom session was started")
EndFunc

Func FlowRecordLoop($recordTime)
   While 1
	  If _IsChecked($chkStartFlow1) Then
		 FlowRecord($recordTime)
	  Else
		 ExitLoop
	  EndIf
	  ;Sleep(GUICtrlRead($txtSleepAfterSave))
   WEnd
EndFunc

Func FlowRecord($recordTime)
   ClickRecord()
   Sleep($recordTime + 1000)
   ClickStop()
   DeletePersistencyFile();Work around - remove it after fix
   Sleep(GUICtrlRead($txtTitleSleep))
   ClickOnTitleField()
   SetFileName()
   ClickSave()
   Sleep(GUICtrlRead($txtSleepAfterSave))
   DeleteRecording()
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

Func ClickSave()
	ControlClick("Classroom Capture - IL-IGORS-RND", "Chrome Legacy Window", "", "left", 1, 934, 137)
	WriteToLog("Clicked on SAVE")
EndFunc

Func SetFileName()
   $NOW = _NowCalc()
   $logname = $NOW
   Send($NOW)
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