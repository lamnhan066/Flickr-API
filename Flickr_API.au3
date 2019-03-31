; Flickr API write by Ltnhan.st.94
; Email: ltnhan.st.94@gmail.com
#include-once
#include <_HttpRequest.au3> ; UDF _HttpRequest của tác giả Huân Hoàng
#include <Date.au3>
#include <Array.au3>
#include <Crypt.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#Include <GuiButton.au3>
#include <EditConstants.au3>

Global $__Flickr_ApiKey             = ""
Global $__Flickr_Secret             = ""
Global $__Flickr_IsSaveToken        = True
Global $__Flickr_PathTokenSaved     = @ScriptDir & "\FlickrOAutInfor.ini"

Global $__Flickr_Oauth_Fullname     = ""
Global $__Flickr_Oauth_Token        = ""
Global $__Flickr_Oauth_Token_Secret = ""
Global $__Flickr_Oauth_User_Nsid    = ""
Global $__Flickr_Oauth_Username     = ""

Global $__Flickr_IsTokenChecked     = False

Func _Flickr_SetUp($ApiKey, $Secret = "", $OAuthToken = "", $OAuthSecret = "", $IsSaveToken = True, $PathTokenSaved = @ScriptDir & "\FlickrOAutInfor.ini")
	If $ApiKey         = Default Then $ApiKey      = ""
	If $Secret         = Default Then $Secret      = ""
	If $OAuthToken     = Default Then $OAuthToken  = ""
	If $OAuthSecret    = Default Then $OAuthSecret = ""
	If $IsSaveToken    = Default Then $IsSaveToken = True
	If $PathTokenSaved = Default Then $OAuthSecret = @ScriptDir & "\FlickrOAutInfor.ini"

	If $ApiKey      <> "" Then $__Flickr_ApiKey             = $ApiKey
	If $Secret      <> "" Then $__Flickr_Secret   			= $Secret
	If $OAuthToken  <> "" Then $__Flickr_Oauth_Token 		= $OAuthToken
	If $OAuthSecret <> "" Then $__Flickr_Oauth_Token_Secret = $OAuthSecret

	$__Flickr_IsSaveToken    = $IsSaveToken
	$__Flickr_PathTokenSaved = $PathTokenSaved

	If $__Flickr_Oauth_Token = "" Then
		$__Flickr_Oauth_Fullname     = BinaryToString(IniRead($__Flickr_PathTokenSaved, "Info", "1", ""), 4) ; FullName
		$__Flickr_Oauth_Token        = BinaryToString(IniRead($__Flickr_PathTokenSaved, "Info", "2", ""), 4) ; Token
		$__Flickr_Oauth_Token_Secret = BinaryToString(IniRead($__Flickr_PathTokenSaved, "Info", "3", ""), 4) ; TokenSecret
		$__Flickr_Oauth_User_Nsid    = BinaryToString(IniRead($__Flickr_PathTokenSaved, "Info", "4", ""), 4) ; UserID
		$__Flickr_Oauth_Username     = BinaryToString(IniRead($__Flickr_PathTokenSaved, "Info", "5", ""), 4) ; User Name
	EndIf
EndFunc

Func _Flickr_CheckToken()
	If $__Flickr_ApiKey = "" or  $__Flickr_Secret = "" Then
		ConsoleWrite("!  Please use _Flickr_SetUp to setup Flickr_API UDF first!")
		Return
	EndIf
	If not $__Flickr_IsTokenChecked Then
		If $__Flickr_Oauth_Token <> "" and not _Flickr_OauthCheckToken() Then
			$__Flickr_Oauth_Fullname     = ""
			$__Flickr_Oauth_Token        = ""
			$__Flickr_Oauth_Token_Secret = ""
			$__Flickr_Oauth_User_Nsid    = ""
			$__Flickr_Oauth_Username     = ""
			FileDelete($__Flickr_PathTokenSaved)
		EndIf
		If $__Flickr_Oauth_Token  = "" Then _Flickr_GetAccessToken()
		If $__Flickr_Oauth_Token <> "" Then $__Flickr_IsTokenChecked = True
	EndIf
EndFunc

Func _Flickr_OauthCheckToken()
	Local $Rs = _HttpRequest(2, _Flickr_ApiGetUrl("flickr.auth.oauth.checkToken", "", False, True, True))
	Return StringInStr($Rs, '"stat":"ok"') <> 0
EndFunc

Func _Flickr_GetOAuthToken()
	Return $__Flickr_Oauth_Token
EndFunc

Func _Flickr_GetOAuthSecret()
	Return $__Flickr_Oauth_Token_Secret
EndFunc

Func _Flickr_GetOAuthUserName()
	Return $__Flickr_Oauth_Username
EndFunc

Func _Flickr_GetOAuthUserID()
	Return $__Flickr_Oauth_User_Nsid
EndFunc

Func _Flickr_GetOAuthFullName()
	Return $__Flickr_Oauth_Fullname
EndFunc

Func _Flickr_GetAccessToken($perm = "read")
	If $perm <> "read" and $perm <> "write" and $perm <> "delete" Then $perm = "read"
	; ------------------------------------------------------------------------------------------------------
	Local $Part1 = "GET"
	Local $Part2 = 'https://www.flickr.com/services/oauth/request_token'
	Local $Part3 = "oauth_callback=oob" _
				 & "&oauth_consumer_key=" & $__Flickr_ApiKey  _
				 & "&oauth_nonce=" & String(Random(11111111, 99999999, 1)) _
				 & "&oauth_signature_method=HMAC-SHA1&oauth_timestamp=" &  _TimeGetStamp() _
				 & "&oauth_version=1.0"
	Local $baseString = $Part1 & "&" & _URIEncode($Part2) & "&" & _URIEncode($Part3)
	Local $signature  = _URIEncode(base64(hmac($__Flickr_Secret & "&", $baseString, "sha1")))
	Local $RqData     = _HttpRequest(2, $Part2 & "?" & $Part3 & '&oauth_signature=' & $signature)
		  $RqData     = StringReplace($RqData, "oauth_callback_confirmed=true&oauth_token=", "")
		  $RqData     = StringReplace($RqData, "oauth_token_secret=", "")
	Local $Rs		  = StringSplit($RqData, "&", 2)
	; ------------------------------------------------------------------------------------------------------
	ShellExecute("https://www.flickr.com/services/oauth/authorize?oauth_token=" & $Rs[0] & "&perms=" & $perm)
	; ------------------------------------------------------------------------------------------------------
	Local $InputBox = _InputAuthorizeCode()
	; ------------------------------------------------------------------------------------------------------
	Local $Part1 = "GET"
	Local $Part2 = 'https://www.flickr.com/services/oauth/access_token'
	Local $Part3 = "oauth_callback=oob" _
				 & "&oauth_consumer_key=" & $__Flickr_ApiKey  _
				 & "&oauth_nonce=" & String(Random(11111111, 99999999, 1)) _
				 & "&oauth_signature_method=HMAC-SHA1&oauth_timestamp=" & _TimeGetStamp() _
				 & "&oauth_token=" & $Rs[0] _
				 & "&oauth_verifier=" & $InputBox _
				 & "&oauth_version=1.0"
	Local $baseString = $Part1 & "&" & _URIEncode($Part2) & "&" & _URIEncode($Part3)
	Local $signature  = _URIEncode(base64(hmac($__Flickr_Secret & "&" & $Rs[1], $baseString, "sha1")))
	Local $RqData     = _HttpRequest(2,$Part2 & "?" & $Part3 & '&oauth_signature=' & $signature)
		  $RqData 	  = StringReplace($RqData, "fullname=", "")
		  $RqData 	  = StringReplace($RqData, "oauth_token=", "")
		  $RqData 	  = StringReplace($RqData, "oauth_token_secret=", "")
		  $RqData	  = StringReplace($RqData, "user_nsid=", "")
		  $RqData  	  = StringReplace($RqData, "username=", "")
	; ------------------------------------------------------------------------------------------------------
	Local $Rs = StringSplit($RqData, "&", 2)
	For $i = 0 to UBound($Rs) - 1
		$Rs[$i] = _URIDecode($Rs[$i])
	Next
	; ------------------------------------------------------------------------------------------------------
	$__Flickr_Oauth_Fullname = _HTMLDecode($Rs[0])
	$__Flickr_Oauth_Token = $Rs[1]
	$__Flickr_Oauth_Token_Secret = $Rs[2]
	$__Flickr_Oauth_User_Nsid = $Rs[3]
	$__Flickr_Oauth_Username = _HTMLDecode($Rs[4])
	; ------------------------------------------------------------------------------------------------------
	If $__Flickr_IsSaveToken Then
		IniWrite($__Flickr_PathTokenSaved, "Info", "1", StringToBinary($__Flickr_Oauth_Fullname, 4))
		IniWrite($__Flickr_PathTokenSaved, "Info", "2", StringToBinary($__Flickr_Oauth_Token, 4))
		IniWrite($__Flickr_PathTokenSaved, "Info", "3", StringToBinary($__Flickr_Oauth_Token_Secret, 4))
		IniWrite($__Flickr_PathTokenSaved, "Info", "4", StringToBinary($__Flickr_Oauth_User_Nsid, 4))
		IniWrite($__Flickr_PathTokenSaved, "Info", "5", StringToBinary($__Flickr_Oauth_Username, 4))
	EndIf
	; ------------------------------------------------------------------------------------------------------
	Return $Rs
EndFunc

Func _Flickr_GetIdFromUrl($Url)
	If StringInStr($Url, "/groups/") Then
		Local $RqUrl = _Flickr_ApiGetUrl("flickr.urls.lookupGroup", "url=" & $Url)
	Else
		Local $RqUrl = _Flickr_ApiGetUrl("flickr.urls.lookupUser", "url=" & $Url)
	EndIf
	Local $RqData = _HttpRequest(2, $RqUrl)
	Local $Rs = StringRegExp($RqData, '"id":"(.*?)"', 1)
	If @error Then Return ""
	Return $Rs[0]
EndFunc   ;==>_GetIdFromUrl

Func _Flickr_GetPeoplePhoto($UserID, $Page, $Size, $IsOAutho = False)
	Local $ParamArray[4] = ["per_page=500", "user_id=" & $UserID, "page=" & $Page, "extras=media,url_" & $Size]
	Return _Flickr_ApiGetUrl("flickr.people.getPhotos", $ParamArray, $IsOAutho)
EndFunc

Func _Flickr_GetGroupPhoto($UserID, $Page, $Size, $IsOAutho = False)
	Local $ParamArray[4] = ["per_page=500", "group_id=" & $UserID, "page=" & $Page, "extras=media,url_" & $Size]
	Return _Flickr_ApiGetUrl("flickr.groups.pools.getPhotos", $ParamArray, $IsOAutho)
EndFunc

Func _Flickr_GetPhotosetsPhoto($UserID, $Page, $Size, $PhotosetID, $IsOAutho = False)
	Local $ParamArray[5] = ["per_page=500", "group_id=" & $UserID, "page=" & $Page, "extras=media,url_" & $Size, "photoset_id=" & $PhotosetID]
	Return _Flickr_ApiGetUrl("flickr.photosets.getPhotos", $ParamArray, $IsOAutho)
EndFunc

Func _Flickr_GetSizes($PhotoID, $IsOAutho = False)
	Local $ParamArray[1] = ["photo_id=" & $PhotoID]
	Return _Flickr_ApiGetUrl("flickr.photos.getSizes", $ParamArray, $IsOAutho)
EndFunc

Func _Flickr_ApiGetUrl($Method, $ParamArray, $OAuth = False, $IsSign = False, $IsSignWithOAuthSecret = False)
	Local $FlickrMethod = "GET"
	Local $FlickrUrl = "https://api.flickr.com/services/rest"
	; ------------------------------------------------------------------------------------------------------
	If $OAuth  = Default Then $OAuth = False
	If $IsSign = Default Then $IsSign = False
	If $IsSignWithOAuthSecret = Default Then $IsSignWithOAuthSecret = False
	; ------------------------------------------------------------------------------------------------------
	Local $Params = "method=" & $Method _
					&"&nojsoncallback=1" _
					&"&format=json" _
					&"&api_key=" & $__Flickr_ApiKey
	; ------------------------------------------------------------------------------------------------------
	If $ParamArray <> "" or IsArray($ParamArray) Then
		If not IsArray($ParamArray) Then $ParamArray = StringSplit($ParamArray, "&", 2)
		For $iParam in $ParamArray
			$Params &= "&" & $iParam
		Next
	EndIf
	; ------------------------------------------------------------------------------------------------------
	If $OAuth Then
		_Flickr_CheckToken()
		$IsSign = True
		$IsSignWithOAuthSecret = True
	EndIf
	; ------------------------------------------------------------------------------------------------------
	If $IsSign Then
		$Params = StringReplace($Params, "&api_key=", "&oauth_consumer_key=")
		$Params &= "&oauth_version=1.0" _
				   &"&oauth_signature_method=HMAC-SHA1" _
				   &"&oauth_nonce=" & String(Random(11111111, 99999999, 1)) _
				   &"&oauth_timestamp=" & _TimeGetStamp()
		If $IsSignWithOAuthSecret Then
			$Params &= "&oauth_token=" & $__Flickr_Oauth_Token
		EndIf
		$Params = _JsonSort($Params)
		Local $baseString = $FlickrMethod & "&" & _URIEncode($FlickrUrl) & "&" & _URIEncode($Params)
		If $IsSignWithOAuthSecret Then
			Local $signature = _URIEncode(base64(hmac($__Flickr_Secret & "&" & $__Flickr_Oauth_Token_Secret, $baseString, "sha1")))
		Else
			Local $signature = _URIEncode(base64(hmac($__Flickr_Secret & "&", $baseString, "sha1")))
		EndIf
		$Params &= "&oauth_signature=" & $signature
	EndIf
	; ------------------------------------------------------------------------------------------------------
	Return $FlickrUrl & "?" & $Params
EndFunc

Func _JsonSort($JsonRaw)
	Local $Split = StringSplit($JsonRaw, "&", 2)
	_ArraySort($Split)
	Local $Result= ""
	For $Json in $Split
		$Result &= $Json & "&"
	Next
	Return StringTrimRight($Result, 1)
EndFunc

Func _TimeGetStamp()
    Local $av_Time
    $av_Time = DllCall('CrtDll.dll', 'long:cdecl', 'time', 'ptr', 0)
    If @error Then
        SetError(99)
        Return False
    EndIf
    Return $av_Time[0]
EndFunc

Func sha1($message)
	Return _Crypt_HashData($message, $CALG_SHA1)
EndFunc

Func md5($message)
	Return _Crypt_HashData($message, $CALG_MD5)
EndFunc

Func hmac($key, $message, $hash="md5")
	Local $blocksize = 64
	Local $a_opad[$blocksize], $a_ipad[$blocksize]
	Local Const $oconst = 0x5C, $iconst = 0x36
	Local $opad = Binary(''), $ipad = Binary('')
	$key = Binary($key)
	If BinaryLen($key) > $blocksize Then
		If $hash = "md5" Then $key = md5($key)
		If $hash = "sha1" Then $key = sha1($key)
	EndIf
	For $i = 1 To BinaryLen($key)
		 $a_ipad[$i-1] = Number(BinaryMid($key, $i, 1))
		 $a_opad[$i-1] = Number(BinaryMid($key, $i, 1))
	Next
	For $i = 0 To $blocksize - 1
		 $a_opad[$i] = BitXOR($a_opad[$i], $oconst)
		 $a_ipad[$i] = BitXOR($a_ipad[$i], $iconst)
	Next
	For $i = 0 To $blocksize - 1
		 $ipad &= Binary('0x' & Hex($a_ipad[$i],2))
		 $opad &= Binary('0x' & Hex($a_opad[$i],2))
	Next
	If $hash = "md5" Then return md5($opad & md5($ipad & Binary($message)))
	If $hash = "sha1" Then return sha1($opad & sha1($ipad & Binary($message)))
EndFunc


;==============================================================================================================================
; Function:         base64($vCode [, $bEncode = True [, $bUrl = False]])
;
; Description:      Decode or Encode $vData using Microsoft.XMLDOM to Base64Binary or Base64Url.
;                   IMPORTANT! Encoded base64url is without @LF after 72 lines. Some websites may require this.
;
; Parameter(s):     $vData      - string or integer | Data to encode or decode.
;                   $bEncode    - boolean           | True - encode, False - decode.
;                   $bUrl       - boolean           | True - output is will decoded or encoded using base64url shema.
;
; Return Value(s):  On Success - Returns output data
;                   On Failure - Returns 1 - Failed to create object.
;
; Author (s):       (Ghads on Wordpress.com), Ascer
;===============================================================================================================================
Func base64($vCode, $bEncode = True, $bUrl = False)
    Local $oDM = ObjCreate("Microsoft.XMLDOM")
    If Not IsObj($oDM) Then Return SetError(1, 0, 1)
    Local $oEL = $oDM.createElement("Tmp")
    $oEL.DataType = "bin.base64"
    If $bEncode then
        $oEL.NodeTypedValue = Binary($vCode)
        If Not $bUrl Then Return $oEL.Text
        Return StringReplace(StringReplace(StringReplace($oEL.Text, "+", "-"),"/", "_"), @LF, "")
    Else
        If $bUrl Then $vCode = StringReplace(StringReplace($vCode, "-", "+"), "_", "/")
        $oEL.Text = $vCode
        Return $oEL.NodeTypedValue
    EndIf
EndFunc ;==>base64

Func _InputAuthorizeCode()
	Local $InputAuthorizeCode = GUICreate("Input Authorize Code",212,76,-1,-1,-1,BitOr($WS_EX_TOPMOST,$WS_EX_TOOLWINDOW))
	Local $Ip_Code = GUICtrlCreateInput("",20,10,178,24,$ES_CENTER,$WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1,11,400,0,"Arial")
	GUICtrlSendMsg(-1, $EM_SETCUEBANNER, True, "xxx-xxx-xxx")
	Local $Bt_Ok = GUICtrlCreateButton("OK",60,40,100,30,-1,-1)
	GUICtrlSetFont(-1,11,400,0,"Arial")
	GUISetState()
	While 1
		If GUIGetMsg() = $Bt_Ok Then
			Local $Rs = GUICtrlRead($Ip_Code)
			GUIDelete($InputAuthorizeCode)
			Return $Rs
		EndIf
		Sleep(10)
	Wend
EndFunc