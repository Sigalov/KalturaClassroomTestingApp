#include <File.au3>
#include <Date.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=C:\Program Files (x86)\AutoIt3\koda_1.7.3.0\Forms\Classroom.kxf
$Form1_1 = GUICreate("Kaltura Classroom QA Automation Testing Tool", 414, 578, -1, -1)
GUISetBkColor(0xFFFFFF)
$RunApp = GUICtrlCreateButton("Start Kaltura Classroom", 42, 104, 161, 41)
$ExitButton = GUICtrlCreateButton("Exit", 8, 544, 75, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$CloseApp = GUICtrlCreateButton("Close Kaltura Classroom", 210, 104, 161, 41)
$Pic1 = GUICtrlCreatePic("C:\Program Files (x86)\AutoIt3\koda_1.7.3.0\Forms\logo.jpg", 8, 8, 393, 89)
$chkStartFlow1 = GUICtrlCreateCheckbox("START", 16, 168, 65, 17)
$Group1 = GUICtrlCreateGroup("Record Flow:", 16, 192, 297, 281)
$Label2 = GUICtrlCreateLabel("Click Record", 32, 216, 81, 20)
$Label3 = GUICtrlCreateLabel("Record Sleep", 32, 240, 88, 20)
$Label4 = GUICtrlCreateLabel("Click Stop", 32, 264, 64, 20)
$txtFlowRecInterval = GUICtrlCreateInput("", 152, 232, 121, 24)
$Label5 = GUICtrlCreateLabel("Wait for title Sleep", 32, 288, 111, 20)
$txtTitleSleep = GUICtrlCreateInput("", 152, 280, 121, 24)
$Label6 = GUICtrlCreateLabel("Click On Title Field", 32, 312, 115, 20)
$Label7 = GUICtrlCreateLabel("Click Save", 32, 336, 68, 20)
$Label8 = GUICtrlCreateLabel("Sleep after Save", 32, 360, 104, 20)
$txtSleepAfterSave = GUICtrlCreateInput("", 152, 352, 121, 24)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnTest = GUICtrlCreateLabel("", 344, 536, 60, 36)
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
		 WriteToLog("TEST")
	  Case $ExitButton
		 Exit
	EndSwitch
WEnd

Func StartKalturaClassroom()
   $iPID = Run('C:\Program Files\Kaltura\Classroom\CaptureApp\KalturaClassroom.exe', 'C:\Program Files\Kaltura\Classroom\CaptureApp\')
   Sleep(7000)
   WriteToLog("Classroom session was started")
EndFunc

Func FlowRecordLoop($recordTime)
   While 1
	  If _IsChecked($chkStartFlow1) Then
		 FlowRecord($recordTime)
	  Else
		 ExitLoop
	  EndIf
	  Sleep(GUICtrlRead($txtSleepAfterSave))
   WEnd
EndFunc

Func FlowRecord($recordTime)
   ClickRecord()
   Sleep($recordTime + 1000)
   ClickStop()
   Sleep(GUICtrlRead($txtTitleSleep))
   ClickOnTitleField()
   SetFileName()
   ClickSave()
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