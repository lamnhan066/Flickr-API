;--------------------------------------------------------------------------------------
; Author:  Ltnhan.st.94
; Website: ltnhanst94.name.vn
; Email:   ltnhan.st.94@gmail.com
;--------------------------------------------------------------------------------------
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Description=Flickr Images Downloader
#AutoIt3Wrapper_Res_LegalCopyright=2019 by Ltnhan.st.94
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/sf /sv /rm /pe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include-once

#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <WinAPIEx.au3>
#include <WinAPI.au3>
#include <InetConstants.au3>
#include "Forms\FormMain.isf"
#include "Forms\Infor.isf"
#include "Forms\FormOption.isf"
#include "Flickr_API.au3"

Global $ApiKey = ""
Global $Secret = ""

;--------------------------------------------------------------------------------------
_Flickr_Setup($ApiKey, $Secret)
;--------------------------------------------------------------------------------------

Global $SizeName = StringSplit("Original|X-Large 6K|X-Large 5K|X-Large 4K|X-Large 3K|Large 2048|Large 1600|Large 1024|Medium 800|Medium 640|Medium 500|Small 400|Small 320|Small 240|Thumbnail|Square 150|Square 75", "|", 2)
Global $SizeUrl  = StringSplit("o|6k|5k|4k|3k|k|h|l|c|z|m|n|w|s|t|q|sq", "|", 2)
;--------------------------------------------------------------------------------------
Global $IsCheckedNameStt    = _IniReadState("Setting", "NameStt", "True")
Global $IsCheckedNameId     = _IniReadState("Setting", "NameId", "False")
Global $IsCheckedNameSize   = _IniReadState("Setting", "NameSize", "True")
Global $IsCheckedAutoSize   = _IniReadState("Setting", "AutoSize", "True")
Global $IsCheckedVideo      = _IniReadState("Setting", "Video", "True")
Global $IsCheckedFromPage1  = _IniReadState("Setting", "FromPage1", "False")
Global $IsCheckedOpenFolder = _IniReadState("Setting", "OpenFolder", "True")
;~ Global $IsCheckedAutoUpdate = _IniReadState("Setting", "AutoUpdate", "False")
;--------------------------------------------------------------------------------------
Global $IsDebug = False
;--------------------------------------------------------------------------------------
For $s In $SizeName
	GUICtrlSetData($Cb_Size, $s, "Original")
Next
For $i = 1 To 100
	GUICtrlSetData($Cb_Thread, $i, 8)
Next
;--------------------------------------------------------------------------------------
_ArrayReverse($SizeName)
_ArrayReverse($SizeUrl)
;--------------------------------------------------------------------------------------
_SetStt("Đã sẵn sàng để sử dụng!")
GUISetState(@SW_SHOW)
;--------------------------------------------------------------------------------------
If Not FileExists(@ScriptDir & "\Configure.ini") Then
	Local $GuiInfor = _Infor()
	GUISetState(@SW_SHOW, $GuiInfor)
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete($GuiInfor)
				ExitLoop
			Case $Lb_EmailUrl
				; skip
		EndSwitch
		Sleep(10)
	WEnd
	FileWrite(@ScriptDir & "\Configure.ini", "")
EndIf
;--------------------------------------------------------------------------------------
While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Bt_Question
			Local $GuiInfor = _Infor()
			GUISetState(@SW_SHOW, $GuiInfor)
			While 1
				Switch GUIGetMsg()
					Case $GUI_EVENT_CLOSE
						GUIDelete($GuiInfor)
						ExitLoop
					Case $Lb_EmailUrl
						; skip
				EndSwitch
				Sleep(10)
			WEnd
		Case $Bt_Path
			GUICtrlSetData($Ip_Path, FileSelectFolder("Chọn đường dẫn lưu", @ScriptDir))
		Case $Download
			_SetStt("Đang kiểm tra..")
			Local $sID = GUICtrlRead($IpUrl)
			If $sID = "" and _IsChecked($Cb_Private) Then
				_Flickr_CheckToken()
				$sID = _Flickr_GetOAuthUserID()
			EndIf
			If StringLeft($sID, 4) == "http" Then $sID = _Flickr_GetIdFromUrl($sID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sID = ' & $sID & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			If $sID <> "" Then
				_FlickrDownload($sID, GUICtrlRead($Ip_Path), GUICtrlRead($Cb_Thread))
			Else
				_SetStt("UserID hoặc Url lỗi! Vui lòng kiểm tra lại!")
			EndIf
		Case $Bt_ExportTxt
			_SetStt("Đang kiểm tra Url")
			Local $sID = GUICtrlRead($IpUrl)
			If $sID = "" And _IsChecked($Cb_Private) Then
				_Flickr_CheckToken()
				$sID = _Flickr_GetOAuthUserID()
			EndIf
			If StringLeft($sID, 4) == "http" Then $sID = _Flickr_GetIdFromUrl($sID)
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sID = ' & $sID & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			If $sID <> "" Then
				_FlickrExport($sID, GUICtrlRead($Ip_Path))
			Else
				_SetStt("UserID hoặc Url lỗi! Vui lòng kiểm tra lại!")
			EndIf
		Case $Bt_Option
			Local $hFormOption = _FormOption()
			_SetOptionState()
			GUISetState(@SW_SHOW, $hFormOption)
			While 1
				Switch GUIGetMsg()
					Case $GUI_EVENT_CLOSE
						_SaveInfo()
						GUIDelete($hFormOption)
						ExitLoop
					Case $Cb_NameStt
						Local $hExample = GUICtrlRead($Lb_Example)
						If _IsChecked($Cb_NameStt) Then
							If Not StringInStr($hExample, "88") Then GUICtrlSetData($Lb_Example, "88. " & $hExample)
						Else
							If StringInStr($hExample, "88") Then GUICtrlSetData($Lb_Example, StringReplace($hExample, "88. ", ""))
						EndIf
					Case $Cb_NameId
						If _IsChecked($Cb_NameId) Then
							GUICtrlSetData($Lb_Example, StringReplace(GUICtrlRead($Lb_Example), "Tên", "ID"))
						Else
							GUICtrlSetData($Lb_Example, StringReplace(GUICtrlRead($Lb_Example), "ID", "Tên"))
						EndIf
					Case $Cb_NameSize
						Local $hExample = StringReplace(GUICtrlRead($Lb_Example), ".jpg", "")
						If _IsChecked($Cb_NameSize) Then
							If Not StringInStr($hExample, "Kích thước") Then GUICtrlSetData($Lb_Example, $hExample & " (Kích thước).jpg")
						Else
							If StringInStr($hExample, "Kích thước") Then GUICtrlSetData($Lb_Example, StringReplace($hExample, " (Kích thước)", ".jpg"))
						EndIf
				EndSwitch
				Sleep(10)
			WEnd
	EndSwitch
	Sleep(10)
WEnd

Func _FlickrDownload($ID, $SavePath, $MultiNumber)
	_SetStt("Đang kiểm tra ID..")
	GUICtrlSetData($Download, "Dừng")
	;--------------------------------------------------------------------------------------
	Local $Text = _HttpRequest(2, _UrlReplace($ID, "1"))
	;--------------------------------------------------------------------------------------
	If StringInStr($Text, '"stat":"fail"') Then
		_SetStt("Lỗi! Thông tin: " & StringRegExp($Text, '"message":"(.*?)"', 1)[0] & "!")
		Return
	EndIf
	;--------------------------------------------------------------------------------------
	Local $TotalImg = StringRegExp($Text, 'total="(.*?)"', 1)
	If @error Then
		Local $TotalImg = StringRegExp($Text, '"total":"(.*?)"', 1)
		If @error Then
			Local $TotalImg = StringRegExp($Text, '"total":(.*?),', 1)
			If @error Then
				_SetStt("Lỗi! Không xác định!")
				Local $ErrorInfo = StringRegExp($Text, '{"(.*?)\[{', 1)
				If not @error Then MsgBox(64, "Debug!", $ErrorInfo[0])
				Return
			EndIf
		EndIf
	EndIf
	$TotalImg = Number($TotalImg[0])
	If $TotalImg == 0 Then
		_SetStt("Không có ảnh hoặc không có quyền xem!")
		Return
	EndIf
	Local $TotalPages = Number(StringRegExp($Text, '"pages":(.*?),', 1)[0])
	;--------------------------------------------------------------------------------------
	If Not $IsCheckedAutoSize Then
		If Not StringInStr($Text, "url_") Then
			_SetStt("Không có ảnh với kích thước đã chọn!")
			Return
		EndIf
	EndIf
	;--------------------------------------------------------------------------------------
	_SetStt("Lấy danh sách url ảnh..")
	;--------------------------------------------------------------------------------------
	DirCreate($SavePath)
	;--------------------------------------------------------------------------------------
	If $IsDebug Then
		DirCreate($SavePath & "\FlickrImagesDownloaderLog")
	EndIf
	;--------------------------------------------------------------------------------------
	Local $Timer = TimerInit()
	Local $PageCount = $TotalPages, $TotalImgCount = 0, $NonImgCount = 0
	Local $Text, $i ; ,$AllPhotoID[0]
	
	If $IsCheckedFromPage1 Then $PageCount = 1	
	While 1
		Local $AllLink[0], $AllTitle[0], $AllInfo[0], $AllVideoPrivateLink[0]
		_SetStt("Lấy danh sách url ảnh (Trang " & $PageCount & "/" & $TotalPages & ")..")
		$Text = _HTMLDecode(_HttpRequest(2, _UrlReplace($ID, $PageCount)))
		$i = _ArraySearch($SizeName, GUICtrlRead($Cb_Size))
		_SetStt("Đang mã hóa lại thông tin..")
		$AllInfo = StringRegExp(StringReplace($Text, "\", ""), '{"id"(.*?)}', 3)
		If @error or UBound($AllInfo) == 0 Then 
			If $IsCheckedFromPage1 Then 
				If $PageCount = $TotalPages Then ExitLoop
				$PageCount += 1
			Else
				If $PageCount = 1 Then ExitLoop
				$PageCount -= 1
			EndIf
			ContinueLoop
		EndIf
		;--------------------------------------------------------------------------------------
		If not $IsCheckedFromPage1 Then _ArrayReverse($AllInfo)
		;--------------------------------------------------------------------------------------
		If $IsDebug Then FileWrite($SavePath & "\FlickrImagesDownloaderLog\page" & $PageCount & ".txt", "") ;;;;;;;;;;;;;;;;;;;;;;;;;;;;
		;--------------------------------------------------------------------------------------
		_SetStt("Xử lý lại dữ liệu..")
		Local $PreTitle, $PreLink, $PreID, $PreCount = 1
		For $i = 0 To UBound($AllInfo) - 1
			If $IsDebug Then FileWrite($SavePath & "\FlickrImagesDownloaderLog\page" & $PageCount & ".txt", "Đã xử lý dữ liệu: " & $i & " ảnh  => ") ;;;;;;;;;;;;;;;;;;;;;;
			GUICtrlSetData($ProgressBar, Round(($i + 1) / UBound($AllInfo) * 100))
			$PreID = StringRegExp($AllInfo[$i], ':"(.*?)"', 1)[0]
			If $IsCheckedNameId Then
				$PreTitle = $PreID
			Else
				$PreTitle = StringRegExp($AllInfo[$i], '"title":"(.*?)"', 1)[0]
				$PreTitle = StringReplace($PreTitle, "/", "")
				$PreTitle = StringReplace($PreTitle, ":", "")
				$PreTitle = StringReplace($PreTitle, "*", "")
				$PreTitle = StringReplace($PreTitle, "\", "")
				$PreTitle = StringReplace($PreTitle, "?", "")
				$PreTitle = StringReplace($PreTitle, "<", "")
				$PreTitle = StringReplace($PreTitle, ">", "")
				$PreTitle = StringReplace($PreTitle, "|", "")
			EndIf
			For $j = UBound($SizeName) - 1 To 0 Step -1
				$PreLink = StringRegExp($AllInfo[$i], '"url_' & $SizeUrl[$j] & '":"(.*?)"', 1)
				If Not @error and $PreLink[0] <> "" Then
					If $IsCheckedVideo And StringInStr($AllInfo[$i], '"media":"video"') Then
						Local $hData, $PreID = StringRegExp($AllInfo[$i], ':"(.*?)"', 1)[0]
						If _IsChecked($Cb_Private) Then
							$hData = _HttpRequest(2, _Flickr_GetSizes($PreID))
						Else
							$hData = _HttpRequest(2, 'https://api.flickr.com/services/rest/?method=flickr.photos.getSizes&api_key=' & $ApiKey & '&photo_id=' & $PreID & '&format=json&nojsoncallback=1')
						EndIf
						$hData = StringRegExp(StringReplace($hData, "\", ""), '"source":"(.*?)"', 3)
						If @error Then
							If $IsCheckedNameSize Then $PreTitle = $PreTitle & " (" & $SizeName[$j] & ")"
							_ArrayAdd($AllLink, $PreLink[0])
						Else
							If _IsChecked($Cb_Private) Then
								If $IsCheckedNameSize Then $PreTitle = $PreTitle & " (" & $SizeName[$j] & ")"
								_ArrayAdd($AllLink, $PreLink[0])
								_ArrayAdd($AllVideoPrivateLink, $hData[UBound($hData) - 1])
							Else
								If $IsCheckedNameSize Then $PreTitle = $PreTitle & " (Video)" & ".mp4"
								_ArrayAdd($AllLink, $hData[UBound($hData) - 1])
							EndIf
						EndIf
					Else
						If $IsCheckedNameSize Then $PreTitle = $PreTitle & " (" & $SizeName[$j] & ")"
						_ArrayAdd($AllLink, $PreLink[0])
					EndIf
					If $IsCheckedNameStt Then $PreTitle = $TotalImgCount + $PreCount & ". " & $PreTitle
					If not StringInStr($PreTitle, ".mp4") Then $PreTitle &= StringRight($PreLink[0], 4)
					_ArrayAdd($AllTitle, $PreTitle)
					$PreCount += 1
					ExitLoop
				EndIf
			Next
			If $IsDebug Then FileWrite($SavePath & "\FlickrImagesDownloaderLog\page" & $PageCount & ".txt", $PreTitle & @CRLF) ;;;;;;;;;;;;;;;;;;;;;;
		Next
		;--------------------------------------------------------------------------------------
		If UBound($AllLink) < $MultiNumber Then $MultiNumber = UBound($AllLink) ; Tranh loi do qua nhieu thread trong khi qua it hinh anh
		;--------------------------------------------------------------------------------------
		If UBound($AllVideoPrivateLink) > 0 Then
			Local $File1 = FileOpen($SavePath & "\FlickrUrlExportVideoPrivate.txt", 1)
			For $i = 0 To UBound($AllVideoPrivateLink) - 1
				_SetStt("Đang ghi (" & ($i + 1) & "/" & UBound($AllVideoPrivateLink) & " video riêng tư)..")
				FileWrite($File1, $AllVideoPrivateLink[$i] & @CRLF)
			Next
			FileClose($File1)
		EndIf
		;--------------------------------------------------------------------------------------
		If $IsDebug Then FileWrite($SavePath & "\FlickrImagesDownloaderLog\page" & $PageCount & ".txt", "Tổng link: " & UBound($AllLink) & @CRLF)
		If $IsDebug Then FileWrite($SavePath & "\FlickrImagesDownloaderLog\page" & $PageCount & ".txt", "Tổng tiêu đề: " & UBound($AllTitle) & @CRLF)
		;--------------------------------------------------------------------------------------
		_ArrayReverse($AllLink)
		_ArrayReverse($AllTitle)
		;--------------------------------------------------------------------------------------
		Local $MultiHandle[$MultiNumber]
		Local $MultiTitle[$MultiNumber]
		Local $EndArray = 0
		;--------------------------------------------------------------------------------------
		For $i = 0 To $MultiNumber - 1
			While 1
				Local $FileLink = _ArrayPop($AllLink)
				Local $FileTitle = _ArrayPop($AllTitle)
				
				$TotalImgCount += 1
			
				If not FileExists($SavePath & "\" & $FileTitle) Then ExitLoop ; Không tải những tệp đã tồn tại
				
				If UBound($AllLink) = 0 Then ExitLoop 2
			WEnd
			;--------------------------------------------------------------------------------------
			$MultiHandle[$i] = InetGet($FileLink, $SavePath & "\" & $FileTitle, 1, 1)
			$MultiTitle[$i] = $FileTitle
			;--------------------------------------------------------------------------------------
			ConsoleWrite("Downloading " & $TotalImgCount & "/" & $TotalImg & ": " & $FileTitle & ".jpg" & @CRLF)
			GUICtrlSetData($ProgressBar, 0)
			_SetStt("Tải " & $TotalImgCount & "/" & $TotalImg & "(" & Round($TotalImgCount / $TotalImg * 100, 2) & "%): Còn 0d 0h 0m 0s..")
		Next
		;--------------------------------------------------------------------------------------
		While 1
			If UBound($AllLink) = 0 Then ExitLoop
			;--------------------------------------------------------------------------------------
			For $i = 0 To $MultiNumber - 1
				If IsArray($AllLink) And UBound($AllLink) > 0 And InetGetInfo($MultiHandle[$i], $INET_DOWNLOADCOMPLETE) Then
					While 1
						If $IsDebug Then  FileWrite($SavePath & "\FlickrImagesDownloaderLog\page" & $PageCount & ".txt", "Đã tải: " & $MultiTitle[$i] & " => " & InetGetInfo($MultiHandle[$i], $INET_DOWNLOADERROR) & @CRLF)
						Local $FileLink = _ArrayPop($AllLink)
						Local $FileTitle = _ArrayPop($AllTitle)
						$TotalImgCount += 1
						;--------------------------------------------------------------------------------------
						If not FileExists($SavePath & "\" & $FileTitle) Then ExitLoop ; Không tải những tệp đã tồn tại
						;--------------------------------------------------------------------------------------
						If UBound($AllLink) = 0 Then ExitLoop 2
					WEnd
					;--------------------------------------------------------------------------------------
					InetClose($MultiHandle[$i])
					$MultiHandle[$i] = InetGet($FileLink, $SavePath & "\" & $FileTitle, 1, 1)
					$MultiTitle[$i] = $FileTitle
					;--------------------------------------------------------------------------------------
					ConsoleWrite("Downloading " & $TotalImgCount & "/" & $TotalImg & ": " & $FileTitle & ".jpg" & @CRLF)
					GUICtrlSetData($ProgressBar, $TotalImgCount / $TotalImg * 100)
					Local $TimeRemaining = TimerDiff($Timer) / $TotalImgCount * ($TotalImg - $TotalImgCount) / 1000
					Local $Day = Int($TimeRemaining / 60 / 60 / 24)
					Local $Hour = Int($TimeRemaining / 60 / 60 - $Day * 24)
					Local $Min = Int($TimeRemaining / 60 - $Hour * 60 - $Day * 24 * 60)
					Local $Sec = Int($TimeRemaining - $Min * 60 - $Hour * 60 * 60 - $Day * 24 * 60 * 60)
					_SetStt("Tải " & $TotalImgCount & "/" & $TotalImg & "(" & Round($TotalImgCount / $TotalImg * 100, 2) & "%): Còn " & $Day & "d " & $Hour & "h " & $Min & "m " & $Sec & "s..")
				EndIf
			Next
			;--------------------------------------------------------------------------------------
			Switch GUIGetMsg()
				Case $Download
					_SetStt("Hoàn thành tệp đang tải trước khi dừng..")
					Local $FileCompleteCount
					While 1
						$FileCompleteCount = 0
						For $i = 0 To $MultiNumber - 1
							If InetGetInfo($MultiHandle[$i], $INET_DOWNLOADCOMPLETE) Then $FileCompleteCount += 1
						Next
						If $FileCompleteCount = $MultiNumber Then 
							_SetStt("Đã dừng!")					
							GUICtrlSetData($Download, "Bắt đầu")
							Return
						EndIf
						Sleep(100)
					WEnd
				Case $GUI_EVENT_CLOSE
					_SetStt("Hoàn thành tệp đang tải trước khi thoát..")
					Local $FileCompleteCount
					While 1
						$FileCompleteCount = 0
						For $i = 0 To $MultiNumber - 1
							If not IsHWnd($MultiHandle[$i]) or InetGetInfo($MultiHandle[$i], $INET_DOWNLOADCOMPLETE) Then $FileCompleteCount += 1
						Next
						If $FileCompleteCount = $MultiNumber Then Exit
						Sleep(100)
					WEnd
			EndSwitch
			Sleep(1)
		WEnd
		If ($IsCheckedFromPage1 and $PageCount == $TotalPages) or (not $IsCheckedFromPage1 and $PageCount == 1) Then
			_SetStt("Đang kết thúc tiến trình..")
			Local $FileCompleteCount
			While 1
				$FileCompleteCount = 0
				For $i = 0 To $MultiNumber - 1
					If not IsHWnd($MultiHandle[$i]) or InetGetInfo($MultiHandle[$i], $INET_DOWNLOADCOMPLETE) Then $FileCompleteCount += 1
				Next
				If $FileCompleteCount = $MultiNumber Then ExitLoop
				Sleep(100)
			WEnd
			ExitLoop
		EndIf
		If $IsCheckedFromPage1 Then
			$PageCount +=1
		Else
			$PageCount -= 1
		EndIf
	WEnd
	;--------------------------------------------------------------------------------------
	If $TotalImgCount > 0 Then 
		If $IsCheckedOpenFolder Then
			ShellExecute($SavePath)
			If FileExists($SavePath & "\FlickrUrlExportVideoPrivate.txt") Then ShellExecute($SavePath & "\FlickrUrlExportVideoPrivate.txt")
		EndIf		
		_SetStt(StringFormat("Hoàn thành (%s/%s ảnh)!", $TotalImgCount, $TotalImg))
		GUICtrlSetData($Download, "Tải")
	EndIf
EndFunc   ;==>_FlickrDownload

Func _FlickrExport($ID, $SavePath)
	_SetStt("Đang kiểm tra ID..")
	Local $Text = _HttpRequest(2, _UrlReplace($ID, "1"))
	;--------------------------------------------------------------------------------------
	If StringInStr($Text, '"stat":"fail"') Then
		_SetStt("Lỗi! Thông tin: " & StringRegExp($Text, '"message":"(.*?)"', 1)[0] & "!")
		Return
	EndIf
	;--------------------------------------------------------------------------------------
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $Text = ' & $Text & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	Local $TotalImg = StringRegExp($Text, 'total="(.*?)"', 1)
	If @error Then
		Local $TotalImg = StringRegExp($Text, '"total":"(.*?)"', 1)
		If @error Then
			Local $TotalImg = StringRegExp($Text, '"total":(.*?),', 1)
			If @error Then
				_SetStt("Lỗi! Không xác định!")
				Local $ErrorInfo = StringRegExp($Text, '{"(.*?)\[{', 1)
				If not @error Then MsgBox(64, "Debug!", $ErrorInfo[0])
				Return
			EndIf
		EndIf
	EndIf
	$TotalImg = Number($TotalImg[0])
	If $TotalImg == 0 Then
		_SetStt("Không có ảnh hoặc không có quyền xem!")
		Return
	EndIf
	Local $TotalPages = Number(StringRegExp($Text, '"pages":(.*?),', 1)[0])
	;--------------------------------------------------------------------------------------
	If Not $IsCheckedAutoSize Then
		If Not StringInStr($Text, "url_") Then
			_SetStt("Không có ảnh với kích thước đã chọn!")
			Return
		EndIf
	EndIf
	;--------------------------------------------------------------------------------------
	_SetStt("Lấy danh sách url ảnh..")
	;--------------------------------------------------------------------------------------
	DirCreate($SavePath)
	;--------------------------------------------------------------------------------------
	Local $Timer = TimerInit()
	Local $PageCount = 1, $TotalImgCount = 0, $NonImgCount = 0
	While 1
		Local $Text, $i, $AllLink[0], $AllTitle[0], $AllInfo[0], $AllVideoPrivateLink[0] ;,$AllPhotoID[0]
		_SetStt("Lấy danh sách url ảnh (Trang " & $PageCount & "/" & $TotalPages & ")..")
		$Text = _HttpRequest(2, _UrlReplace($ID, $PageCount))
		$i = _ArraySearch($SizeName, GUICtrlRead($Cb_Size))
		_SetStt("Đang mã hóa lại thông tin..")
		$AllInfo = StringRegExp(StringReplace(_HTMLDecode($Text), "\", ""), '{"id"(.*?)}', 3)
		;--------------------------------------------------------------------------------------
		_SetStt("Xử lý lại dữ liệu..")
		Local $PreTitle, $PreLink, $PreID, $PreCount = 1
		For $i = 0 To UBound($AllInfo) - 1
			GUICtrlSetData($ProgressBar, Round(($i + 1) / UBound($AllInfo) * 100))
			;--------------------------------------------------------------------------------------
			$TotalImgCount += 1
			_SetStt("Đang xử lý " & ($i + 1) & "/" & UBound($AllInfo))
			;--------------------------------------------------------------------------------------
			For $j = UBound($SizeName) - 1 To 0 Step -1
				$PreLink = StringRegExp($AllInfo[$i], '"url_' & $SizeUrl[$j] & '":"(.*?)"', 1)
				If Not @error Then
					If $IsCheckedVideo And StringInStr($AllInfo[$i], '"media":"video"') Then
						Local $hData, $PreID = StringRegExp($AllInfo[$i], ':"(.*?)"', 1)[0]
						If _IsChecked($Cb_Private) Then
							ConsoleWrite("PreID = " & $PreID)
							$hData = _HttpRequest(2, _Flickr_GetSizes($PreID))
						Else
							$hData = _HttpRequest(2, 'https://api.flickr.com/services/rest/?method=flickr.photos.getSizes&api_key=' & $ApiKey & '&photo_id=' & $PreID & '&format=json&nojsoncallback=1')
						EndIf
						$hData = StringRegExp(StringReplace($hData, "\", ""), '"source":"(.*?)"', 3)
						If @error Then
							_ArrayAdd($AllLink, $PreLink[0])
						Else
							If _IsChecked($Cb_Private) Then
								_ArrayAdd($AllLink, $PreLink[0])
								;------------------------------------------------------------------- Video của Flickr riêng tư
								_ArrayAdd($AllVideoPrivateLink, $hData[UBound($hData) - 1])
							Else
								_HttpRequest(0, $hData[UBound($hData) - 1])
								_ArrayAdd($AllLink, _GetLocationRedirect())
							EndIf
						EndIf
					Else
						_ArrayAdd($AllLink, $PreLink[0])
					EndIf
					_ArrayAdd($AllTitle, $PreTitle)
					$PreCount += 1
					ExitLoop
				EndIf
			Next
			If GUIGetMsg() = $GUI_EVENT_CLOSE Then Exit
		Next
		;--------------------------------------------------------------------------------------
		_SetStt("Đang ghi..")
		Local $File = FileOpen($SavePath & "\FlickrUrlExport.txt", 1)
		For $i = 0 To UBound($AllLink) - 1
			_SetStt("Đang ghi " & ($i + 1) & "/" & UBound($AllLink) & " ảnh và video..")
			FileWrite($File, $AllLink[$i] & @CRLF)
			GUICtrlSetData($ProgressBar, ($i + 1) / UBound($AllLink) * 100)
		Next
		FileClose($File)
		If UBound($AllVideoPrivateLink) > 0 Then
			Local $File1 = FileOpen($SavePath & "\FlickrUrlExportVideoPrivate.txt", 1)
			For $i = 0 To UBound($AllVideoPrivateLink) - 1
				_SetStt("Đang ghi " & ($i + 1) & "/" & UBound($AllVideoPrivateLink) & " video riêng tư..")
				FileWrite($File1, $AllVideoPrivateLink[$i] & @CRLF)
			Next
			FileClose($File1)
		EndIf
		;--------------------------------------------------------------------------------------
		If GUIGetMsg() = $GUI_EVENT_CLOSE Then Exit
		;--------------------------------------------------------------------------------------
		If $PageCount >= Number($TotalPages) Then ExitLoop
		$PageCount += 1
	WEnd
	;--------------------------------------------------------------------------------------
	If $TotalImgCount > 0 Then _SetStt(StringFormat("Hoàn thành (%s/%s ảnh)!", $TotalImgCount, $TotalImg))
	;--------------------------------------------------------------------------------------
	If $IsCheckedOpenFolder Then
		ShellExecute($SavePath)
		If FileExists($SavePath & "\FlickrUrlExportVideoPrivate.txt") Then ShellExecute($SavePath & "\FlickrUrlExportVideoPrivate.txt")
	EndIf
EndFunc   ;==>_FlickrExport

Func _UrlReplace($ID, $Page = 1)
	Local $i = _ArraySearch($SizeName, GUICtrlRead($Cb_Size))
	If _IsGroup() Then
		If _IsChecked($Cb_Private) Then Return _Flickr_GetGroupPhoto($ID, $Page, _GetSizeParam(), True)
		Return _Flickr_GetGroupPhoto($ID, $Page, _GetSizeParam())
	ElseIf _IsPhotoSet() Then
		Local $iPhotoSetID = StringSplit(GUICtrlRead($IpUrl), "/")
		$iPhotoSetID = $iPhotoSetID[UBound($iPhotoSetID) - 1]
		If _IsChecked($Cb_Private) Then Return _Flickr_GetPhotosetsPhoto($ID, $Page, _GetSizeParam(), $iPhotoSetID, True)
		Return _Flickr_GetPhotosetsPhoto($ID, $Page, _GetSizeParam(), $iPhotoSetID)
	Else
		If _IsChecked($Cb_Private) Then Return _Flickr_GetPeoplePhoto($ID, $Page, _GetSizeParam(), True)
		Return _Flickr_GetPeoplePhoto($ID, $Page, _GetSizeParam())
	EndIf
EndFunc   ;==>_UrlReplace

Func _GetSizeParam()
	Local $iMaxSize = _ArraySearch($SizeName, GUICtrlRead($Cb_Size))
	;--------------------------------------------------------------------------------------
	If Not $IsCheckedAutoSize Then Return $SizeUrl[$iMaxSize]
	;--------------------------------------------------------------------------------------
	Local $SizeParam = ""
	For $i = 0 To $iMaxSize
		$SizeParam &= ",url_" & $SizeUrl[$i]
	Next
	;--------------------------------------------------------------------------------------
	Return StringTrimLeft($SizeParam, 5)
EndFunc   ;==>_GetSizeParam

Func _IsGroup()
	Return StringInStr(StringLower(GUICtrlRead($IpUrl)), "/groups/")
EndFunc   ;==>_IsGroup

Func _IsPhotoSet()
	Return StringInStr(StringLower(GUICtrlRead($IpUrl)), "/albums/")
EndFunc

Func _SetStt($String)
	GUICtrlSetData($Lb_Status, $String)
EndFunc   ;==>_SetStt

Func _IsChecked($hHandle)
	Return BitAND(GUICtrlRead($hHandle), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

Func _IsInternetConnected()
	Local $aReturn = DllCall('connect.dll', 'long', 'IsInternetConnected')
	If @error Then
		Return SetError(1, 0, False)
	EndIf
	Return $aReturn[0] = 0
EndFunc   ;==>_IsInternetConnected

Func _SaveInfo()
	$IsCheckedNameStt    = _IsChecked($Cb_NameStt)
	$IsCheckedNameId     = _IsChecked($Cb_NameId)
	$IsCheckedNameSize   = _IsChecked($Cb_NameSize)
	$IsCheckedAutoSize   = _IsChecked($Cb_AutoSize)
	$IsCheckedVideo      = _IsChecked($Cb_Video)
	$IsCheckedFromPage1  = _IsChecked($Cb_FromPage1)
	$IsCheckedOpenFolder = _IsChecked($Cb_OpenFolder)
;~ 	$IsCheckedAutoUpdate = _IsChecked($Cb_AutoUpdate)
	;--------------------------------------------------------------------------------------
	_IniWriteState("Setting", "NameStt",    $IsCheckedNameStt)
	_IniWriteState("Setting", "NameId",     $IsCheckedNameId)
	_IniWriteState("Setting", "NameSize",   $IsCheckedNameSize)
	_IniWriteState("Setting", "AutoSize",   $IsCheckedAutoSize)
	_IniWriteState("Setting", "Video",      $IsCheckedVideo)
	_IniWriteState("Setting", "FromPage1",  $IsCheckedFromPage1)
	_IniWriteState("Setting", "OpenFolder", $IsCheckedOpenFolder)
;~ 	_IniWriteState("Setting", "AutoUpdate", $IsCheckedAutoUpdate)
EndFunc   ;==>_SaveInfo

Func _SetOptionState()
	If $IsCheckedNameStt    Then GUICtrlSetState($Cb_NameStt, $GUI_CHECKED)
	If $IsCheckedNameId     Then GUICtrlSetState($Cb_NameId, $GUI_CHECKED)
	If $IsCheckedNameSize   Then GUICtrlSetState($Cb_NameSize, $GUI_CHECKED)
	If $IsCheckedAutoSize   Then GUICtrlSetState($Cb_AutoSize, $GUI_CHECKED)
	If $IsCheckedVideo      Then GUICtrlSetState($Cb_Video, $GUI_CHECKED)
	If $IsCheckedFromPage1  Then GUICtrlSetState($Cb_FromPage1, $GUI_CHECKED)
	If $IsCheckedOpenFolder Then GUICtrlSetState($Cb_OpenFolder, $GUI_CHECKED)
;~ 	If $IsCheckedAutoUpdate Then GUICtrlSetState($Cb_AutoUpdate, $GUI_CHECKED)
	;--------------------------------------------------------------------------------------
	Local $hExample = GUICtrlRead($Lb_Example)
	If $IsCheckedNameStt Then
		If Not StringInStr($hExample, "88") Then GUICtrlSetData($Lb_Example, "88. " & $hExample)
	Else
		If StringInStr($hExample, "88") Then GUICtrlSetData($Lb_Example, StringReplace($hExample, "88. ", ""))
	EndIf
	If $IsCheckedNameId Then
		GUICtrlSetData($Lb_Example, StringReplace(GUICtrlRead($Lb_Example), "Tên", "ID"))
	Else
		GUICtrlSetData($Lb_Example, StringReplace(GUICtrlRead($Lb_Example), "ID", "Tên"))
	EndIf
	Local $hExample = StringReplace(GUICtrlRead($Lb_Example), ".jpg", "")
	If $IsCheckedNameSize Then
		If Not StringInStr($hExample, "Kích thước") Then GUICtrlSetData($Lb_Example, $hExample & " (Kích thước).jpg")
	Else
		If StringInStr($hExample, "Kích thước") Then GUICtrlSetData($Lb_Example, StringReplace($hExample, " (Kích thước)", ".jpg"))
	EndIf
EndFunc   ;==>_SetOptionState

Func _IniReadState($Section, $Key, $Default)
	Return IniRead(@ScriptDir & "\Configure.ini", $Section, $Key, $Default) = "True"
EndFunc   ;==>_IniReadState

Func _IniWriteState($Section, $Key, $Default)
	Return IniWrite(@ScriptDir & "\Configure.ini", $Section, $Key, $Default)
EndFunc   ;==>_IniWriteState
