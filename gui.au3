#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>

Global $Files[0]
Global $File_Attr[0][3]

Global $pathToDB = @ScriptDir & "\" &"Database4.accdb"
Global $pathToFiles = @ScriptDir & "\" &"files"
Global $pathToDownloads = @UserProfileDir & "\Downloads"

Global $Table_Name = "DB1"

Global $Attr_Name[3] = ["", "", ""]

Global $PathsForDownload[0]
Global $ArrayForDownload[0]
Global $sFileToBeDownloaded[0]
Global $result[0]
Global $rating[0]
_DBUpdate()
_CalculateRatings()
$GroupHeight = 1
$GroupWidth = 1
$GroupsFromTop = 160
$FilesGroupFromLeft = 23
$RatingsGroupFromLeft = 123

$ListWidth = 95
$ListHeight = 195
$ListsFromTop = 180
$FileListFromLeft = 27
$RatingsListFromLeft = 125

$Form_Main = GUICreate("GUI managing Database", 250, 380)
$Group_Attributes = GUICtrlCreateGroup("Attributes", 20, 20, 200, 130)
$Checkbox_1 = GUICtrlCreateCheckbox($Attr_Name[0], 40, 50)
$Checkbox_2 = GUICtrlCreateCheckbox($Attr_Name[1], 40, 80)
$Checkbox_3 = GUICtrlCreateCheckbox($Attr_Name[2], 40, 110)
$idAddFile = GUICtrlCreateButton("Add", 140, 70, 60, 20)
$idDownloadFile = GUICtrlCreateButton("Download", 140, 110, 60, 20)
$Group_Files = GUICtrlCreateGroup("Files", $FilesGroupFromLeft, $GroupsFromTop, $GroupWidth ,$GroupHeight )
$Group_Ratings = GUICtrlCreateGroup("Ratings", $RatingsGroupFromLeft, $GroupsFromTop, $GroupWidth, $GroupHeight)
Global $List = GUICtrlCreateList("", $FileListFromLeft, $ListsFromTop, $ListWidth, $ListHeight)
Global $ListRatings = GUICtrlCreateList("", $RatingsListFromLeft, $ListsFromTop, $ListWidth, $ListHeight)

GUISetState(@SW_SHOW)

Global $sItems

; GUI loop
While 1
    $msg = GUIGetMsg()
    Switch $msg
        Case $Checkbox_1, $Checkbox_2, $Checkbox_3
            GUICtrlSetData($List, "")
			GUICtrlSetData($ListRatings, "")
            Access($Checkbox_1)
            Access($Checkbox_2)
            Access($Checkbox_3)
        Case $GUI_EVENT_CLOSE ; Close GUI
            ExitLoop
        Case $idAddFile
			$check1 = CheckboxCheck($Checkbox_1)
			$check2 = CheckboxCheck($Checkbox_2)
			$check3 = CheckboxCheck($Checkbox_3)
			; this check prevents uploading files if all 3 checkboxes are not checked
			If $check1 = false and $check2 = false and $check3 = false Then
			ContinueLoop
			Else 
				$sFileToBeCopied = FileOpenDialog("Select Files", @ScriptDir, "Text Files(*.txt)", 5)
				FileCopy($sFileToBeCopied, $pathToFiles)
				If @error Then ContinueLoop	
				$AdoCon = _DBConnect()
				$AdoRs = _DBCreateObject()
				$AdoRs.Open("SELECT * FROM " & $Table_Name, $AdoCon) 
				$aFiles = StringSplit($sFileToBeCopied, "|")
				Switch $aFiles[0]
					Case 1
						$AdoRs.AddNew
						$AdoRs.Fields("Feld1").value = StringTrimLeft($aFiles[1], StringInStr($aFiles[1], "\", 0, -1))
						If BitAnd(GUICtrlRead($Checkbox_1),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld2").value = GUICtrlRead($Checkbox_1, 1)
						If BitAnd(GUICtrlRead($Checkbox_2),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld3").value = GUICtrlRead($Checkbox_2, 1)
						If BitAnd(GUICtrlRead($Checkbox_3),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld4").value = GUICtrlRead($Checkbox_3, 1)
						$AdoRs.Fields("FilePath").value = $pathToFiles&"\"&StringTrimLeft($aFiles[1], StringInStr($aFiles[1], "\", 0, -1))
						$AdoRs.Update
					Case 2 To $aFiles[0]
						For $i = 2 To $aFiles[0]
							$AdoRs.AddNew
							$AdoRs.Fields("Feld1").value = $aFiles[$i]
							If BitAnd(GUICtrlRead($Checkbox_1),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld2").value = GUICtrlRead($Checkbox_1, 1)
							If BitAnd(GUICtrlRead($Checkbox_2),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld3").value = GUICtrlRead($Checkbox_2, 1)
							If BitAnd(GUICtrlRead($Checkbox_3),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Feld4").value = GUICtrlRead($Checkbox_3, 1)
							$AdoRs.Fields("FilePath").value = $pathToFiles&"\"&StringTrimLeft($aFiles[$i], StringInStr($aFiles[$i], "\", 0, -1))
							$AdoRs.Update
						Next
				EndSwitch
				$AdoRs.close
				$AdoCon.Close
				_DBUpdate()
				GUICtrlSetData($List, "")
				GUICtrlSetData($ListRatings, "")
				Access($Checkbox_1)
				Access($Checkbox_2)
				Access($Checkbox_3)
			Endif
		Case $idDownloadFile
			; getting selected items
			$aTextItems = _GUICtrlListBox_GetSelItemsText($List)
			
			; getting paths for downloading
			$sFileToBeDownloaded = getFilesForDownload($aTextItems)
		
			; copying files from files folder to downloads folder
			For $j = 1 to Ubound($sFileToBeDownloaded)-1
				FileCopy($sFileToBeDownloaded[$j], $pathToDownloads)
				If @error Then showMessage(@error)
			Next
			; clearing values for additional downloads
			$sFileToBeDownloaded = ""
	EndSwitch
WEnd

###########################################################################
;																		  ;
;'''''''''''''''''''''''''''''' FUNCTIONS '''''''''''''''''''''''''''''''';
;																		  ;
###########################################################################

Func Access($Checkbox)
    If GUICtrlRead($Checkbox) = $GUI_CHECKED Then
        Local $Chkbox_label = GUICtrlRead($Checkbox, 1)
        For $i = 0 To UBound($Files) - 1 Step 1
            If $File_Attr[$i][0] = $Chkbox_label Or $File_Attr[$i][1] = $Chkbox_label Or $File_Attr[$i][2] = $Chkbox_label Then
            _GUICtrlListBox_AddString($List, $Files[$i])
			_GUICtrlListBox_AddString($ListRatings, $rating[$i+1])
            Endif
        Next
    Endif
EndFunc
 
Func CheckboxCheck($Checkbox)
	Local $boolean = false
	If GUICtrlRead($Checkbox) = $GUI_CHECKED Then
	$boolean = true
	Endif
	Return $boolean
Endfunc 

Func _DBUpdate()
    $AdoCon = _DBConnect()
	$AdoRs = _DBCreateObject()
    $AdoRs.Open("SELECT COUNT(*) FROM " & $Table_Name, $AdoCon)
    $dimension = $AdoRs.Fields(0).Value
	$AdoRs2 = _DBCreateObject()
	$AdoRs2.Open("SELECT * FROM " & $Table_Name & " ORDER BY ID DESC", $AdoCon)
	;$max = $AdoRs2.Fields(0).Value
    ReDim $Files[$dimension]
    ReDim $File_Attr[$dimension][3]
    For $i = 0 To $dimension-1 Step 1
        $AdoRs = _DBCreateObject()
        $AdoRs.Open("SELECT * FROM " & $Table_Name & " WHERE ID = " & ($i + 1), $AdoCon)
		;While $AdoRs.Fields(0).Value <> null		
        $Files[$i] = $AdoRs.Fields(1).Value 
		if @error Then 
		showMessage(@error)
		else
        $File_Attr[$i][0] = $AdoRs.Fields(2).Value
        $File_Attr[$i][1] = $AdoRs.Fields(3).Value
        $File_Attr[$i][2] = $AdoRs.Fields(4).Value
		endif
    Next
    $AdoRs.Close
    $AdoCon.Close

    Local $a = 0
    For $i = 0 To UBound($Files) - 1 Step 1
        For $j = 0 To 2 Step 1
            If $a < 3 And Not $File_Attr[$i][$j] = "" Then
                For $k = $a To 2 Step 1
                    If $Attr_Name[$k] = $File_Attr[$i][$j] Then
                        ContinueLoop 2
                    EndIf
                Next
                $Attr_Name[$a] = $File_Attr[$i][$j]
                $a = $a + 1
            EndIf
        Next
    Next
EndFunc

Func getFilesForDownload($ArrayForDownload)
	$AdoCon = _DBConnect()
	$AdoRs = _DBCreateObject()
	$AdoRs.Open("SELECT COUNT(*) FROM " & $Table_Name, $AdoCon)
	
	; redimensioning $pathsForDownload variable so that it can store all paths 
	Redim $PathsForDownload[Ubound($ArrayForDownload)]
	
	; getting all paths in this loop using file names
	For $i = 1 To UBound($ArrayForDownload)-1
		$AdoRs = _DBCreateObject()
		$AdoRs.Open("SELECT * FROM " & $Table_Name & " WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)		
		$PathsForDownload[$i] = $AdoRs.Fields(5).Value
		
		; updating NumberOfDownloads value
		if $AdoRs.Fields(6).Value = null Then
		showMessage("in IF")
		$AdoRs3 = _DBCreateObject()
		$AdoRs3.Open("UPDATE " & $Table_Name & " SET NumberOfDownloads=1 WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)
		else 
		$AdoRs4 = _DBCreateObject()
		$AdoRs4.Open("SELECT * FROM " & $Table_Name & " WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)
		$DLNumber = $AdoRs.Fields(6).Value
		$DLNumber +=1
		showMessage($DLNumber)
		$AdoRs5 = _DBCreateObject()
		$AdoRs5.Open("UPDATE " & $Table_Name & " SET NumberOfDownloads="&$DLNumber&" WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)
		endif
	Next
	$AdoRs.Close
    $AdoCon.Close
	Return $PathsForDownload
EndFunc

Func _CalculateRatings()
	$AdoCon = _DBConnect()
	$AdoRs = _DBCreateObject()
    $AdoRs.Open("SELECT COUNT(*) FROM " & $Table_Name, $AdoCon)
    $dimension2 = $AdoRs.Fields(0).Value
	Global $DLTotal = 0
	Redim $result[$dimension2+1]
	Redim $rating[$dimension2+1]
	For $i = 1 To $dimension2
	$AdoRs3 = _DBCreateObject()
	$AdoRs3.Open("SELECT * FROM " & $Table_Name & " WHERE ID="&$i&"", $AdoCon)
	$DLCurrentRow = $AdoRs3.Fields(6).Value
	$DLTotal += $DLCurrentRow
	Next
	For $i = 1 To $dimension2
		$AdoRs3 = _DBCreateObject()
		$AdoRs3.Open("SELECT * FROM " & $Table_Name & " WHERE ID="&$i&"", $AdoCon)
		$DLCurrentRow = $AdoRs3.Fields(6).Value
		$result[$i] = $DLCurrentRow*100/$DLTotal
		If $result[$i]>16 Then
		$rating[$i] = 5
		ElseIf $result[$i]>10 Then 
		$rating[$i] = 4
		ElseIf $result[$i]>7 Then
		$rating[$i] = 3
		ElseIf $result[$i]>2 Then
		$rating[$i] = 2
		Else
		$rating[$i] = 1
		Endif
	Next
EndFunc

Func showMessage($message)
	MsgBox(0, "Autoit", $message)
EndFunc

Func _DBConnect()
	$AdoCon = ObjCreate("ADODB.Connection")
    $AdoCon.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & $pathToDB)
	return $AdoCon
EndFunc

Func _DBCreateObject()
	$AdoRs = ObjCreate("ADODB.Recordset")
	$AdoRs.CursorType = 1
	$AdoRs.LockType = 3
	return $AdoRs
EndFunc