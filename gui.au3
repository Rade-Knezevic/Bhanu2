#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>

Global $Files[0]
Global $SortedFiles[0]

Global $result[0]
Global $SortedRatings[0]

Global $File_Attr[0][3]
Global $SortedArrays[0][4]

Global $pathToDB = @ScriptDir & "\" &"Database1.accdb"
Global $pathToFiles = @ScriptDir & "\" &"files"
Global $pathToDownloads = @UserProfileDir & "\Downloads"

Global $Table_Name = "DB1"

Global $Attr_Name[3] = ["", "", ""]

Global $PathsForDownload[0]
Global $NamesForDownload[0]
Global $ArrayForDownload[0]
Global $sFileToBeDownloaded[0]


_DBUpdate()
$result = getDownloads()
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
Global $Checkbox_1 = GUICtrlCreateCheckbox("Controller", 40, 50)
Global $Checkbox_2 = GUICtrlCreateCheckbox("Testing", 40, 80)
Global $Checkbox_3 = GUICtrlCreateCheckbox("Simulation", 40, 110)
$idAddFile = GUICtrlCreateButton("Add", 140, 70, 60, 20)
$idDownloadFile = GUICtrlCreateButton("Download", 140, 110, 60, 20)
$Group_Files = GUICtrlCreateGroup("Files", $FilesGroupFromLeft, $GroupsFromTop, $GroupWidth ,$GroupHeight)
$Group_Ratings = GUICtrlCreateGroup("Ratings", $RatingsGroupFromLeft, $GroupsFromTop, $GroupWidth, $GroupHeight)
Global $List = GUICtrlCreateList("", $FileListFromLeft, $ListsFromTop, $ListWidth, $ListHeight,$LBS_EXTENDEDSEL)
Global $ListRatings = GUICtrlCreateList("", $RatingsListFromLeft, $ListsFromTop, $ListWidth, $ListHeight,BitOR($LBS_EXTENDEDSEL, $LBS_DISABLENOSCROLL))

GUISetState(@SW_SHOW)

Global $sItems

; GUI loop
While 1
    $msg = GUIGetMsg()
    Switch $msg
        Case $Checkbox_1, $Checkbox_2, $Checkbox_3
			GUICtrlSetData($List, "")		
			GUICtrlSetData($ListRatings, "")
			AccessAll($Checkbox_1,$Checkbox_2,$Checkbox_3)
            ;Access($Checkbox_1)
            ;Access($Checkbox_2)
            ;Access($Checkbox_3)
        Case $GUI_EVENT_CLOSE ; Close GUI
            ExitLoop
        Case $idAddFile
			$check1 = CheckboxCheck($Checkbox_1)
			$check2 = CheckboxCheck($Checkbox_2)
			$check3 = CheckboxCheck($Checkbox_3)
			; this check prevents upload ing files if all 3 checkboxes are not checked
			If $check1 = 0 and $check2 = 0 and $check3 = 0 Then
			showMessage("At least 1 checkbox must be checked in order to download")
			ContinueLoop
			Else	 	
				$sFileToBeCopied = FileOpenDialog("Select Files", @ScriptDir, "Text Files(*.txt)", 5)
				$IsNameExisting = checkIfFileExists($sFileToBeCopied)
				if $IsNameExisting = true Then
				showMessage("Name already exists! Rename file and upload!")
				ContinueLoop
				Else 
					FileCopy($sFileToBeCopied, $pathToFiles)
					ReDim $result[Ubound($Files)+1]
					ReDim $Files[Ubound($Files)+1]	
					If @error Then 
					ContinueLoop
					Else											
						$AdoCon = _DBConnect()
						$AdoRs = _DBCreateObject()
						$AdoRs.Open("SELECT * FROM " & $Table_Name, $AdoCon) 
						$aFiles = StringSplit($sFileToBeCopied, "|")
						Switch $aFiles[0]
							Case 1
								$AdoRs.AddNew
								$AdoRs.Fields("Feld1").value = StringTrimLeft($aFiles[1], StringInStr($aFiles[1], "\", 0, -1))
								If BitAnd(GUICtrlRead($Checkbox_1),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Controller").value = "Controller"
								If BitAnd(GUICtrlRead($Checkbox_2),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Testing").value = "Testing"
								If BitAnd(GUICtrlRead($Checkbox_3),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Simulation").value = "Simulation"
								$AdoRs.Fields("FilePath").value = $pathToFiles&"\"&StringTrimLeft($aFiles[1], StringInStr($aFiles[1], "\", 0, -1))
								$AdoRs.Update
							Case 2 To $aFiles[0]
								For $i = 2 To $aFiles[0]
									$AdoRs.AddNew
									$AdoRs.Fields("Feld1").value = $aFiles[$i]
									If BitAnd(GUICtrlRead($Checkbox_1),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Controller").value = "Controller"
									If BitAnd(GUICtrlRead($Checkbox_2),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Testing").value = "Testing"
									If BitAnd(GUICtrlRead($Checkbox_3),$GUI_CHECKED) = $GUI_CHECKED Then $AdoRs.Fields("Simulation").value = "Simulation"
									$AdoRs.Fields("FilePath").value = $pathToFiles&"\"&StringTrimLeft($aFiles[$i], StringInStr($aFiles[$i], "\", 0, -1))
									$AdoRs.Update
								Next
						EndSwitch
						$AdoRs.close
						$AdoCon.Close
						_DBUpdate()
						GUICtrlSetData($List, "")
						GUICtrlSetData($ListRatings, "")
						AccessAll($Checkbox_1,$Checkbox_2,$Checkbox_3)
						$sFileToBeCopied = ""
					Endif	
				Endif
			Endif
		Case $idDownloadFile
		; getting selected items
		$aTextItems = _GUICtrlListBox_GetSelItemsText($List)
		Local $suggestedName = $aTextItems[1]
			$sDownloadDialog = FileSaveDialog("Download file(s)", $pathToDownloads, "Text files (*.txt)", 16, $suggestedName)
			If @error Then
				; Display the error message.
				MsgBox($MB_SYSTEMMODAL, "", "No file was saved.")
			Else		
				; getting paths for downloading
				$sFileToBeDownloaded = getFilesForDownload($aTextItems)
			
				; copying files from files folder to downloads folder
				For $j = 1 to Ubound($sFileToBeDownloaded)-1
					FileCopy($sFileToBeDownloaded[$j], $sDownloadDialog)
					If @error Then showMessage(@error)
				Next
				; clearing value for additional downloads
				$sFileToBeDownloaded = ""
			Endif
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
		
		$SortedFiles = _EmptyString($SortedFiles)
		$SortedRatings = _EmptyString($SortedRatings)

		; make this array same length as the unsorted one
		ReDim $SortedFiles[UBound($Files)]
		ReDim $SortedRatings[UBound($Files)]
		ReDim $SortedArrays[UBound($Files)][4]
		
		$result = _ArrayToNumber($result)
		; sorting ratings using function from highest to lowest
		$SortedFiles = _SortMySecondArray($result,$Files)
		$SortedRatings = _SortMyArray($result)
		
		For $i = 0 To Ubound($Files)-1
		$SortedArrays[$i][0] = $Files[$i]
		$SortedArrays[$i][1] = $result[$i]
		$SortedArrays[$i][2] = $SortedRatings[$i]
		$SortedArrays[$i][3] = $SortedFiles[$i]
		Next

		For $i = 0 To UBound($Files)-1 Step 1
            If $File_Attr[$i][0] = $Chkbox_label Or $File_Attr[$i][1] = $Chkbox_label Or $File_Attr[$i][2] = $Chkbox_label Then
			_GUICtrlListBox_AddString($List, $SortedFiles[$i])
			_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
            Endif
        Next
    Endif
EndFunc

Func AccessAll($Checkbox1,$Checkbox2,$Checkbox3)

	$check1 = CheckboxCheck($Checkbox_1)
	$check2 = CheckboxCheck($Checkbox_2)
	$check3 = CheckboxCheck($Checkbox_3)
	if $check1 or $check2 or $check3 Then
		$SortedFiles = _EmptyString($SortedFiles)
		$SortedRatings = _EmptyString($SortedRatings)

		; make this array same length as the unsorted one
		ReDim $SortedFiles[UBound($Files)]
		ReDim $SortedRatings[UBound($Files)]
		ReDim $SortedArrays[UBound($Files)][4]
			
		$result = _ArrayToNumber($result)
		; sorting ratings using function from highest to lowest
		$SortedFiles = _SortMySecondArray($result,$Files)
		$SortedRatings = _SortMyArray($result)
			
		For $i = 0 To Ubound($Files)-1
		$SortedArrays[$i][0] = $Files[$i]
		$SortedArrays[$i][1] = $result[$i]
		$SortedArrays[$i][2] = $SortedRatings[$i]
		$SortedArrays[$i][3] = $SortedFiles[$i]
		Next
			
		$accessSwitch = _AccessSwitch($check1,$check2,$check3)

		For $i = 0 To UBound($Files)-1 Step 1
			Switch $accessSwitch
				Case 1
				If $File_Attr[$i][0] = "Controller" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 2
				If $File_Attr[$i][1] = "Testing" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 3
				If $File_Attr[$i][2] = "Simulation" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 4
				If $File_Attr[$i][0] = "Controller" or $File_Attr[$i][1] = "Testing" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 5
				If $File_Attr[$i][0] = "Controller" or $File_Attr[$i][2] = "Simulation" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 6
				If $File_Attr[$i][1] = "Testing" or $File_Attr[$i][2] = "Simulation" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 7
				If $File_Attr[$i][0] = "Controller" or $File_Attr[$i][1] = "Testing" or $File_Attr[$i][2] = "Simulation" Then
				_GUICtrlListBox_AddString($List, $SortedFiles[$i])
				_GUICtrlListBox_AddString($ListRatings, $SortedRatings[$i])
				Endif
				Case 0
				GUICtrlSetData($List, "")		
				GUICtrlSetData($ListRatings, "")
			EndSwitch	
		Next
	Endif	
EndFunc

Func _AccessSwitch($check1,$check2,$check3)
	Global $accessSwitch
	If $check1 = 1 AND $check2 = 0 AND $check3 = 0 Then
		$accessSwitch = 1
	ElseIf $check1 = 0 AND $check2 = 1 AND $check3 = 0 Then
		$accessSwitch = 2
	ElseIf $check1 = 0 AND $check2 = 0 AND $check3 = 1 Then
		$accessSwitch = 3
	ElseIf $check1 = 1 AND $check2 = 1 AND $check3 = 0 Then
		$accessSwitch = 4
	ElseIf $check1 = 0 AND $check2 = 1 AND $check3 = 1 Then
		$accessSwitch = 5	
	ElseIf $check1 = 1 AND $check2 = 0 AND $check3 = 1 Then
		$accessSwitch = 6	
	ElseIf $check1 = 1 AND $check2 = 1 AND $check3 = 1 Then
		$accessSwitch = 7
	Else
		$accessSwitch = 0
	Endif	
	Return $accessSwitch	
EndFunc	
	
Func CheckboxCheck($Checkbox)
	Local $boolean = 0
	If GUICtrlRead($Checkbox) = $GUI_CHECKED Then
	$boolean = 1
	Endif
	Return $boolean
Endfunc 

Func _DBUpdate()
    $AdoCon = _DBConnect()
	$AdoRs = _DBCreateObject()
    $AdoRs.Open("SELECT COUNT(*) FROM " & $Table_Name, $AdoCon)
    $dimension = $AdoRs.Fields(0).Value	
	$AdoRs.Close
    ReDim $Files[$dimension+1]
    ReDim $File_Attr[$dimension+1][3]
    For $i = 0 To $dimension-1 Step 1
        $AdoRs2 = _DBCreateObject()
        $AdoRs2.Open("SELECT * FROM " & $Table_Name & " WHERE ID = "&($i+1), $AdoCon)
		$Files[$i] = $AdoRs2.Fields(1).Value 
		if @error Then 
		showMessage(@error)
		else
        $File_Attr[$i][0] = $AdoRs2.Fields(2).Value
        $File_Attr[$i][1] = $AdoRs2.Fields(3).Value
        $File_Attr[$i][2] = $AdoRs2.Fields(4).Value
		endif
		$AdoRs2.Close
    Next
    $AdoCon.Close

    Local $a = 0
    For $i = 0 To UBound($Files)-1 Step 1
        For $j = 0 To 2 Step 1
            If $a < 3 And Not $File_Attr[$i][$j] = "" Then
                For $k = $a To 2 Step 1
                    If $Attr_Name[$k] = $File_Attr[$i][$j] Then
                        ContinueLoop 2
                    EndIf
                Next
                $Attr_Name[$a] = $File_Attr[$i][$j]
                $a += 1
            EndIf
        Next
    Next
EndFunc

Func getFilesForDownload($ArrayForDownload)
	$AdoCon = _DBConnect()
	; redimensioning $pathsForDownload variable so that it can store all paths 
	Redim $PathsForDownload[Ubound($ArrayForDownload)]
	Redim $NamesForDownload[Ubound($ArrayForDownload)]
	; getting all paths in this loop using file names
	For $i = 1 To UBound($ArrayForDownload)-1
		$AdoRs = _DBCreateObject()
		$AdoRs.Open("SELECT * FROM " & $Table_Name & " WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)		

		$NamesForDownload[$i] = $AdoRs.Fields(1).Value
		$PathsForDownload[$i] = @ScriptDir &"\"&"files"&"\"&$NamesForDownload[$i]

		; updating NumberOfDownloads value 
		$AdoRs2 = _DBCreateObject()
		$AdoRs2.Open("SELECT * FROM " & $Table_Name & " WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)
		$DLNumber = $AdoRs2.Fields(6).Value
		$DLNumber +=1
		$AdoRs3 = _DBCreateObject()
		$AdoRs3.Open("UPDATE " & $Table_Name & " SET NumberOfDownloads="&$DLNumber&" WHERE Feld1 = '"&$ArrayForDownload[$i]&"'", $AdoCon)
	Next
    $AdoCon.Close
	Return $PathsForDownload
EndFunc

Func getDownloads()
	$AdoCon = _DBConnect()
	$AdoRs = _DBCreateObject()
    $AdoRs.Open("SELECT COUNT(*) FROM " & $Table_Name, $AdoCon)
    $dimension2 = $AdoRs.Fields(0).Value
	Redim $result[$dimension2+1]
	For $i = 0 To $dimension2-1
		$AdoRs3 = _DBCreateObject()
		$AdoRs3.Open("SELECT * FROM " & $Table_Name & " WHERE ID="&($i+1), $AdoCon)
		$result[$i] =  $AdoRs3.Fields(6).Value
	Next
	$AdoRs.Close
    $AdoCon.Close
	Return $result
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

Func checkIfFileExists($FileToBeAdded)
	$AdoCon = _DBConnect()
	$check = false 
	For $k = 1 To Ubound($Files)-1
		$AdoRs = _DBCreateObject()
		$AdoRs.Open("SELECT * FROM " & $Table_Name & " WHERE ID = "& ($k), $AdoCon)	
		$currentName = $AdoRs.Fields(1).Value
		$FileTrimmed = StringTrimLeft($FileToBeAdded, StringInStr($FileToBeAdded, "\", 0, -1))
		if $FileTrimmed = $currentName Then
		$check = true
		endif		
	Next
	$AdoRs.Close
    $AdoCon.Close
	Return $check
EndFunc

Func _SortMyArray($array)
	Local $temp
	For $i = 0 to UBound($array-1)
		For $k = 0 to UBound($array-1)
		If $k+1 < UBound($array-1) Then
			If $array[$k]<$array[$k+1] Then
				$temp = $array[$k]
				$array[$k] = $array[$k+1]
				$array[$k+1] = $temp  	
			Endif
		Endif	
		Next
	Next	
	Return $array
EndFunc

Func _SortMySecondArray($array, $array2)
	Local $temp
	Local $temp2
	For $i = 0 to UBound($array-1)
		For $k = 0 to UBound($array-1)
		If $k+1 < UBound($array-1) Then
			If $array[$k]<$array[$k+1] Then
				$temp = $array[$k]
				$array[$k] = $array[$k+1]
				$array[$k+1] = $temp  	
				$temp2 = $array2[$k]
				$array2[$k] = $array2[$k+1]
				$array2[$k+1] = $temp2  
			Endif
		Endif	
		Next
	Next	
	Return $array2
EndFunc

Func _ArrayToNumber($result)
	For $i = 0 To UBound($result)-1 Step 1
		$result[$i] = Number($result[$i])
	Next
	return $result
EndFunc

Func _EmptyString($result)
	For $i = 0 To UBound($result)-1 Step 1
		$result[$i] = ""
	Next
	return $result
EndFunc