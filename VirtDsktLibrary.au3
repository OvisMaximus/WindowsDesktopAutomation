#include <Constants.au3>
Opt("MustDeclareVars", True)

; Windows 11 currently registers as Windows 10 in the registry. This script is compatible with both OS. This could break for Windows 11 in the future if Microsoft updates the registry for Windows 11 to reflect it's OS
If @OSVersion <> "WIN_10" And  @OSVersion <> "WIN_11" Then Exit MsgBox($MB_SYSTEMMODAL, "", "This script only runs on Win 10 and Win 11 but your system announces itself as " & @OSVersion)

Local $OSBuild = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild")
Local Enum $windows10 = 21999, $windows11 = 22000
Local $IID_IVirtualDesktop, $IID_IVirtualDesktopManagerInternal, $tagIVirtualDesktopManagerInternal

If $OSBuild >= $windows11 Then
    $IID_IVirtualDesktop = "{536D3495-B208-4CC9-AE26-DE8111275BF8}"
    $IID_IVirtualDesktopManagerInternal = "{B2F925B9-5A0F-4D2E-9F4D-2B1507593C10}"

    If $OSBuild >= 22489 Then
        $tagIVirtualDesktopManagerInternal = _
        "GetCount hresult(ptr;int*);" & _
        "MoveViewToDesktop hresult(ptr;ptr);" & _
        "CanViewMoveDesktops hresult(ptr;bool*);" & _
        "GetCurrentDesktop hresult(ptr;ptr*);" & _
        "GetAllCurrentDesktops hresult();" & _
        "GetDesktops hresult(ptr;ptr*);" & _
        "GetAdjacentDesktop hresult(ptr;int;ptr*);" & _
        "SwitchDesktop hresult(ptr;ptr);" & _
        "CreateDesktopW hresult(ptr;int*);" & _
        "MoveDesktop hresult(ptr;int*;int);" & _
        "RemoveDesktop hresult(ptr;ptr);" & _
        "FindDesktop hresult(struct*;ptr*);"
    Else
        $tagIVirtualDesktopManagerInternal = _
        "GetCount hresult(ptr;int*);" & _
        "MoveViewToDesktop hresult(ptr;ptr);" & _
        "CanViewMoveDesktops hresult(ptr;bool*);" & _
        "GetCurrentDesktop hresult(ptr;ptr*);" & _
        "GetDesktops hresult(ptr;ptr*);" & _
        "GetAdjacentDesktop hresult(ptr;int;ptr*);" & _
        "SwitchDesktop hresult(ptr;ptr);" & _
        "CreateDesktopW hresult(ptr;int*);" & _
        "MoveDesktop hresult(ptr;int*;int);" & _
        "RemoveDesktop hresult(ptr;ptr);" & _
        "FindDesktop hresult(struct*;ptr*);"
    EndIf


ElseIf $OSBuild < $windows11 Then
    $IID_IVirtualDesktop = "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}"
    $IID_IVirtualDesktopManagerInternal = "{F31574D6-B682-4CDC-BD56-1827860ABEC6}"

    $tagIVirtualDesktopManagerInternal = _
    "GetCount hresult(int*);" & _
    "MoveViewToDesktop hresult(ptr;ptr);" & _
    "CanViewMoveDesktops hresult(ptr;bool*);" & _
    "GetCurrentDesktop hresult(ptr*);" & _
    "GetDesktops hresult(ptr*);" & _
    "GetAdjacentDesktop hresult(ptr;int;ptr*);" & _
    "SwitchDesktop hresult(ptr);" & _
    "CreateDesktopW hresult(int*);" & _
    "RemoveDesktop hresult(ptr;ptr);" & _
    "FindDesktop hresult(struct*;ptr*);"
EndIf

; Instanciation objects
Local $CLSID_ImmersiveShell = "{c2f03a33-21f5-47fa-b4bb-156362a2f239}"
Local $IID_IUnknown = "{00000000-0000-0000-c000-000000000046}"
Local $IID_IServiceProvider = "{6D5140C1-7436-11CE-8034-00AA006009FA}"
Local $tIID_IServiceProvider = __uuidof($IID_IServiceProvider)
Local $tagIServiceProvider  = _
    "QueryService hresult(struct*;struct*;ptr*);"

; VirtualDesktopManagerInternal object
Const Enum $eLeftDirection = 3, $eRightDirection
Local $CLSID_VirtualDesktopManagerInternal = "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}"
Local $tCLSID_VirtualDesktopManagerInternal = __uuidof($CLSID_VirtualDesktopManagerInternal)
Local $tIID_IVirtualDesktopManagerInternal = __uuidof($IID_IVirtualDesktopManagerInternal)

; ApplicationViewCollection object
Local $CLSID_IApplicationViewCollection = "{1841C6D7-4F9D-42C0-AF41-8747538F10E5}"
Local $tCLSID_IApplicationViewCollection = __uuidof($CLSID_IApplicationViewCollection)
Local $IID_IApplicationViewCollection = "{1841C6D7-4F9D-42C0-AF41-8747538F10E5}"
Local $tIID_IApplicationViewCollection = __uuidof($IID_IApplicationViewCollection)
Local $tagIApplicationViewCollection = _
    "GetViews hresult(struct*);" & _
    "GetViewsByZOrder hresult(struct*);" & _
    "GetViewsByAppUserModelId hresult(wstr;struct*);" & _
    "GetViewForHwnd hresult(hwnd;ptr*);" & _
    "GetViewForApplication hresult(ptr;ptr*);" & _
    "GetViewForAppUserModelId hresult(wstr;int*);" & _
    "GetViewInFocus hresult(ptr*);"

; ApplicationView object
Local $IID_IApplicationView = "{372E1D3B-38D3-42E4-A15B-8AB2B178F513}"
Local $tagIApplicationView = _
    "GetIids hresult(ulong*;ptr*);" & _
    "GetRuntimeClassName hresult(str*);" & _
    "GetTrustLevel hresult(int*);" & _
    "SetFocus hresult();" & _
    "SwitchTo hresult();" & _
    "TryInvokeBack hresult(ptr);" & _
    "GetThumbnailWindow hresult(hwnd*);" & _
    "GetMonitor hresult(ptr*);" & _
    "GetVisibility hresult(int*);" & _
    "SetCloak hresult(int;int);" & _
    "GetPosition hresult(clsid;ptr*);" & _
    "SetPosition hresult(ptr);" & _
    "InsertAfterWindow hresult(hwnd);" & _
    "GetExtendedFramePosition hresult(struct*);" & _
    "GetAppUserModelId hresult(wstr*);" & _
    "SetAppUserModelId hresult(wstr);" & _
    "IsEqualByAppUserModelId hresult(wstr;int*);" & _
    "GetViewState hresult(uint*);" & _
    "SetViewState hresult(uint);" & _
    "GetNeediness hresult(int*);"

; VirtualDesktopPinnedApps object
Local $CLSID_VirtualDesktopPinnedApps = "{b5a399e7-1c87-46b8-88e9-fc5747b171bd}"
Local $tCLSID_VirtualDesktopPinnedApps = __uuidof($CLSID_VirtualDesktopPinnedApps)
Local $IID_IVirtualDesktopPinnedApps = "{4ce81583-1e4c-4632-a621-07a53543148f}"
Local $tIID_IVirtualDesktopPinnedApps = __uuidof($IID_IVirtualDesktopPinnedApps)
Local $tagIVirtualDesktopPinnedApps = _
    "IsAppIdPinned hresult(wstr;bool*);" & _
    "PinAppID hresult(wstr);" & _
    "UnpinAppID hresult(wstr);" & _
    "IsViewPinned hresult(ptr;bool*);" & _
    "PinView hresult(ptr);" & _
    "UnpinView hresult(ptr);"

; Miscellaneous objects
Local $IID_IObjectArray = "{92ca9dcd-5622-4bba-a805-5e9f541bd8c9}"
Local $tagIObjectArray = _
    "GetCount hresult(int*);" & _
    "GetAt hresult(int;ptr;ptr*);"
Local $tIID_IVirtualDesktop = __uuidof($IID_IVirtualDesktop) ; Windows OS Specific Variable
Local $tagIVirtualDesktop = _
    "IsViewVisible hresult(ptr;bool*);" & _
    "GetId hresult(clsid*);"

; ActiveDesktop object
Global Const $AD_GETWP_IMAGE = 1
Global Const $AD_APPLY_ALL = 7
Global $CLSID_ActiveDesktop = "{75048700-ef1f-11d0-9888-006097deacf9}"
Global $IID_IActiveDesktop = "{f490eb00-1240-11d1-9888-006097deacf9}"

; implemented only first 3 methods
Global $tagIActiveDesktop = _
  "ApplyChanges hresult(dword);" & _
  "GetWallpaper hresult(struct*;uint;dword);" & _
  "SetWallpaper hresult(wstr;dword)"

; create object
Local $oActiveDesktop = ObjCreateInterface($CLSID_ActiveDesktop, $IID_IActiveDesktop, $tagIActiveDesktop)
If Not IsObj($oActiveDesktop) Then Exit MsgBox( 0, "Script Error", "Object of Active Desktop control is invalid.", 90)

; objects creation
Local $pService
Local $oImmersiveShell = ObjCreateInterface($CLSID_ImmersiveShell, $IID_IUnknown, "")
If @error Then Exit MsgBox("No Immersive shell available. Error: " & @error & ", extended: " & @extended & @CRLF)
$oImmersiveShell.QueryInterface($tIID_IServiceProvider, $pService)
If @error Then Exit MsgBox("No Query Interface available. Error: " & @error & ", extended: " & @extended & @CRLF)
Local $oService = ObjCreateInterface($pService, $IID_IServiceProvider, $tagIServiceProvider)
If @error Then Exit MsgBox("No Service available. Error: " & @error & ", extended: " & @extended & @CRLF)

Local $pApplicationViewCollection, $pVirtualDesktopManagerInternal, $pVirtualDesktopPinnedApps
$oService.QueryService($tCLSID_IApplicationViewCollection, $tIID_IApplicationViewCollection, $pApplicationViewCollection)
If @error Then Exit MsgBox("No Access to Application View Collection. Error: " & @error & ", extended: " & @extended & @CRLF)
Local $oApplicationViewCollection = ObjCreateInterface($pApplicationViewCollection, $IID_IApplicationViewCollection, $tagIApplicationViewCollection)
If @error Then Exit MsgBox("Cast of Application View Collection to Object failed. Error: " & @error & ", extended: " & @extended & @CRLF)
$oService.QueryService($tCLSID_VirtualDesktopManagerInternal, $tIID_IVirtualDesktopManagerInternal, $pVirtualDesktopManagerInternal)
If @error Then Exit MsgBox("No Access to Virtual Desktop Manager Internal. Error: " & @error & ", extended: " & @extended & @CRLF)
Local $oVirtualDesktopManagerInternal = ObjCreateInterface($pVirtualDesktopManagerInternal, $IID_IVirtualDesktopManagerInternal, $tagIVirtualDesktopManagerInternal)
If @error Then Exit MsgBox("Cast of Virtual Desktop Manager Internal to Object failed. Error: " & @error & ", extended: " & @extended & @CRLF)



; gives the number of virtual desktops
Func getNumDesktops() 
	Local $iCount, $iHresult
	$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.GetCount(0, $iCount) : $oVirtualDesktopManagerInternal.GetCount($iCount)
	If $iHresult <> 0 Then 
			SetError($iHresult)
			Return(-1)
		Else 
			Return($iCount)
	EndIf
EndFunc

; creates a new virtual desktop and returns its handle
Func createDesktop()
	Local $pNew
	($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.CreateDesktopW(0, $pNew) : $oVirtualDesktopManagerInternal.CreateDesktopW($pNew)
	If @error Then Return SetError(@error, @extended, -1)
	Return($pNew)
EndFunc 


; switch to a virtual desktop
Func switchDesktop(Const $pNew)
	($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.SwitchDesktop(0, $pNew) : $oVirtualDesktopManagerInternal.SwitchDesktop($pNew)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc

; enumerates all desktops
Func getDesktopArray(ByRef $oArray)
	Local $pArray
	($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.GetDesktops(0, $pArray) : $oVirtualDesktopManagerInternal.GetDesktops($pArray)
	If @error Then 
		ConsoleWriteError("GetDesktops failed. Error Code " & @error & ", Extended: " & @extended & @CRLF)
		Return SetError(@error, @extended, -1)
	EndIf
	$oArray = ObjCreateInterface($pArray, $IID_IObjectArray, $tagIObjectArray)
	If @error Then  
		ConsoleWriteError("Could not cast desktop list " & $pArray & " to array. Error Code " & @error & ", Extended: " & @extended & @CRLF)
		Return SetError(@error, @extended, -2)
	EndIf
	Return 0
EndFunc

;get the desktop handle at nth position
Func getDesktopAtPosition($n, ByRef $pDesktop)
	Local $oArray, $iCount
	getDesktopArray($oArray)
	If @error Then
		ConsoleWriteError("Can't access list of desktops. Error Code is " & @error & @CRLF)
		Return SetError(@error,  @extended, -1)
	EndIf
	$oArray.GetCount($iCount)
	If $iCount - 1 < $n Then
		ConsoleWriteError("Can't access desktop " & $n & " because only " & $iCount & " desktops exist" & @CRLF)
		Return SetError(-1, 0, -2)
	EndIf
	$oArray.GetAt($n, DllStructGetPtr($tIID_IVirtualDesktop), $pDesktop)
	If @error Then Return SetError(@error, @extended, -3)
	Return 0
EndFunc

; gives the current desktop id
Func getCurrentDesktopHandle(ByRef $pCurrent)
	$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.GetCurrentDesktop(0, $pCurrent) : $oVirtualDesktopManagerInternal.GetCurrentDesktop($pCurrent)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc

; returns the adjacent desktop id
Func getAdjacentDesktopHandle($pOriginDesktop, $direction, ByRef $pDesktop)
	$oVirtualDesktopManagerInternal.GetAdjacentDesktop($pOriginDesktop, $direction, $pDesktop)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc

; switches to a specific desktop
Func switchToDesktop($pDesktop)
	($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.SwitchDesktop(0, $pDesktop) : $oVirtualDesktopManagerInternal.SwitchDesktop($pDesktop)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc

Func GetViewHandleForWinHandle($hWin, ByRef $hView)
	$oApplicationViewCollection.GetViewForHwnd($hWin, $hView)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc
 
Func GetAppUserModelFromViewHandle($hView, ByRef $sApp)
	Local $oView
    $oView = ObjCreateInterface($hView, $IID_IApplicationView, $tagIApplicationView)
	$oView.GetAppUserModelId($sApp)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc

; move application to a specific desktop
Func MoveViewToDesktop($hView, $hDesktop)
	$oVirtualDesktopManagerInternal.MoveViewToDesktop($hView, $hDesktop)
	If @error Then Return SetError(@error, @extended, -1)
	Return 0
EndFunc

Func MoveWindowToDesktop($hWnd, $iDesktop) 
	Local $iHresult, $pView, $pDesktop
	$iHresult = GetViewHandleForWinHandle($hWnd, $pView)
	If @error Then Return SetError(@error, @extended, -1)
	$iHresult = getDesktopAtPosition($iDesktop, $pDesktop)
	If @error Then Return SetError(@error, @extended, -2)
	$iHresult = MoveViewToDesktop($pView, $pDesktop)
	If @error Then Return SetError(@error, @extended, -3)
	Return 0
EndFunc

Func MoveWindow($sQueryText, $sContentText, $iDesktop, $iXpos, $iYpos, $iWidth, $iHeight)
	Local $hWnd, $iHresult
	$hWnd = WinGetHandle($sQueryText, $sContentText)
	$iHresult = MoveWindowToDesktop($hWnd, $iDesktop)
	If @error Then Return SetError(@error, @extended, -1)
	$iHresult = WinMove($hWnd, "", $iXpos, $iYpos, $iWidth, $iHeight)
	If @error Then Return SetError(@error, @extended, -2)
	Return 0
EndFunc

;~ ; deletes an existing desktop
;~ $iHresult = $oVirtualDesktopManagerInternal.RemoveDesktop($pNew, $pDesktop)
;~ ConsoleWrite("Delete = " & $iHresult & @CRLF)

;~ $oService.QueryService($tCLSID_VirtualDesktopPinnedApps, $tIID_IVirtualDesktopPinnedApps, $pVirtualDesktopPinnedApps)
;~ ConsoleWrite("Virtual Desktop Pinned Apps = " & $pVirtualDesktopPinnedApps & @CRLF)
;~ Local $oVirtualDesktopPinnedApps = ObjCreateInterface($pVirtualDesktopPinnedApps, $IID_IVirtualDesktopPinnedApps, $tagIVirtualDesktopPinnedApps)
;~ ConsoleWrite("Virtual Desktop Pinned Apps = " & IsObj($oVirtualDesktopPinnedApps) & @CRLF)

;~ ; get current wallpaper
;~ Local $iHresult, $iLength = 100
;~ Local $tName = DllStructCreate("wchar string[" & $iLength & "]")
;~ $iHresult = $oActiveDesktop.GetWallpaper(DllStructGetPtr($tName), $iLength, $AD_GETWP_IMAGE)
;~ ConsoleWrite($tName.string & "/" & $iHresult & @CRLF)

;~ ; set new wallpaper
;~ $iHresult = $oActiveDesktop.SetWallpaper("C:\Windows\Web\Wallpaper\acer01.jpg", 0)  ; <<<<<<<<<<< set correct path to WP file
;~ ConsoleWrite("Set Wallpaper returned " & $iHresult & @CRLF)
;~ $iHresult = $oActiveDesktop.ApplyChanges($AD_APPLY_ALL)
;~ ConsoleWrite("Apply Changes returned " & $iHresult & @CRLF)

Func __uuidof($sGUID)
    Local $tGUID = DllStructCreate("ulong Data1;ushort Data2;ushort Data3;byte Data4[8]")
    DllCall("ole32.dll", "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $tGUID)
    If @error Then Return SetError(@error, @extended, 0)
    Return $tGUID
EndFunc   ;==>__uuidof



