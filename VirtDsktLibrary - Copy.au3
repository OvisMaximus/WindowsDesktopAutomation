#include <Constants.au3>

;https://github.com/MScholtes/VirtualDesktop/blob/master/VirtualDesktop11Insider.cs
;https://github.com/MScholtes/VirtualDesktop/blob/master/VirtualDesktop11.cs
;https://github.com/MScholtes/VirtualDesktop/blob/master/VirtualDesktop.cs

Opt("MustDeclareVars", True)

; Windows 11 currently registers as Windows 10 in the registry. This script is compatible with both OS. This could break for Windows 11 in the future if Microsoft updates the registry for Windows 11 to reflect it's OS
If @OSVersion <> "WIN_10" And  @OSVersion <> "WIN_11" Then Exit MsgBox($MB_SYSTEMMODAL, "", "This script only runs on Win 10 and Win 11" & @OSVersion)

Local $OSBuild = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild")
Local Enum $windows10 = 21999, $windows11 = 22000
;GetOS()

;OS Specific Variables
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



; objects creation
Local $pService
Local $oImmersiveShell = ObjCreateInterface($CLSID_ImmersiveShell, $IID_IUnknown, "")
ConsoleWrite("Immersive shell = " & IsObj($oImmersiveShell) & @CRLF)
$oImmersiveShell.QueryInterface($tIID_IServiceProvider, $pService)
ConsoleWrite("Service pointer = " & $pService & @CRLF)
Local $oService = ObjCreateInterface($pService, $IID_IServiceProvider, $tagIServiceProvider)
ConsoleWrite("Service = " & IsObj($oService) & @CRLF)

Local $pApplicationViewCollection, $pVirtualDesktopManagerInternal, $pVirtualDesktopPinnedApps
$oService.QueryService($tCLSID_IApplicationViewCollection, $tIID_IApplicationViewCollection, $pApplicationViewCollection)
ConsoleWrite("View collection pointer = " & $pApplicationViewCollection & @CRLF)
Local $oApplicationViewCollection = ObjCreateInterface($pApplicationViewCollection, $IID_IApplicationViewCollection, $tagIApplicationViewCollection)
ConsoleWrite("View collection = " & IsObj($oApplicationViewCollection) & @CRLF)

$oService.QueryService($tCLSID_VirtualDesktopManagerInternal, $tIID_IVirtualDesktopManagerInternal, $pVirtualDesktopManagerInternal)
ConsoleWrite("Virtual Desktop pointer = " & $pVirtualDesktopManagerInternal & @CRLF)
Local $oVirtualDesktopManagerInternal = ObjCreateInterface($pVirtualDesktopManagerInternal, $IID_IVirtualDesktopManagerInternal, $tagIVirtualDesktopManagerInternal)
ConsoleWrite("Virtual Desktop = " & IsObj($oVirtualDesktopManagerInternal) & @CRLF)

$oService.QueryService($tCLSID_VirtualDesktopPinnedApps, $tIID_IVirtualDesktopPinnedApps, $pVirtualDesktopPinnedApps)
ConsoleWrite("Virtual Desktop Pinned Apps = " & $pVirtualDesktopPinnedApps & @CRLF)
Local $oVirtualDesktopPinnedApps = ObjCreateInterface($pVirtualDesktopPinnedApps, $IID_IVirtualDesktopPinnedApps, $tagIVirtualDesktopPinnedApps)
ConsoleWrite("Virtual Desktop Pinned Apps = " & IsObj($oVirtualDesktopPinnedApps) & @CRLF)

Local $iCount, $pCurrent, $pLeft, $pNew, $iHresult, $hWnd, $pView, $pArray, $pDesktop, $oArray, $oView, $sView, $bValue

; gives the number of virtual desktops
$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.GetCount(0, $iCount) : $oVirtualDesktopManagerInternal.GetCount($iCount)
;$iCount = 1

ConsoleWrite("Number of Desktop = " & $iCount & "/" & $iHresult & @CRLF)
;If $iCount > 1 Then Exit MsgBox($MB_SYSTEMMODAL, "", "Please close all additional Virtual Desktops")

; creates a new virtual desktop
$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.CreateDesktopW(0, $pNew) : $oVirtualDesktopManagerInternal.CreateDesktopW($pNew)
ConsoleWrite("Create = " & $pNew & "/" & $iHresult & @CRLF)
$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.SwitchDesktop(0, $pNew) : $oVirtualDesktopManagerInternal.SwitchDesktop($pNew)
ConsoleWrite("Switch = " & $iHresult & @CRLF)
Run("Notepad.exe")
WinWait("[CLASS:Notepad]")

; enumerates all desktops
$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.GetDesktops(0, $pArray) : $oVirtualDesktopManagerInternal.GetDesktops($pArray)
$oArray = ObjCreateInterface($pArray, $IID_IObjectArray, $tagIObjectArray)
ConsoleWrite("Array = " & IsObj($oArray) & @CRLF)
$oArray.GetCount($iCount)
ConsoleWrite("Count = " & $iCount & @CRLF)
$oArray.GetAt(0, DllStructGetPtr($tIID_IVirtualDesktop), $pDesktop)
ConsoleWrite("Desktop 0 = " & $pDesktop & @CRLF)
$oArray.GetAt(1, DllStructGetPtr($tIID_IVirtualDesktop), $pCurrent)
ConsoleWrite("Desktop 1 = " & $pCurrent & @CRLF)

; gives the current desktop id
$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.GetCurrentDesktop(0, $pCurrent) : $oVirtualDesktopManagerInternal.GetCurrentDesktop($pCurrent)
ConsoleWrite("Current = " & $pCurrent & "/" & $iHresult & @CRLF)

; returns the adjacent desktop id
$iHresult = $oVirtualDesktopManagerInternal.GetAdjacentDesktop($pCurrent, $eLeftDirection, $pLeft)
ConsoleWrite("Get Left = " & $pLeft & "/" & $iHresult & @CRLF)

; switches to a specific desktop
MsgBox ($MB_SYSTEMMODAL,"","Now it will return to previous desktop")
Sleep(500) ; gives time for the msg box to close
$iHresult = ($OSBuild >= $windows11) ? $oVirtualDesktopManagerInternal.SwitchDesktop(0, $pLeft) : $oVirtualDesktopManagerInternal.SwitchDesktop($pLeft)
ConsoleWrite("Switch = " & $iHresult & @CRLF)

; get pointer to a view based on hwnd and create an application view
$hWnd = WinGetHandle("[CLASS:Notepad]")
$iHresult = $oApplicationViewCollection.GetViewForHwnd($hWnd, $pView)
ConsoleWrite("View from handle = " & $pView & "/" & $iHresult & @CRLF)
$oView = ObjCreateInterface($pView, $IID_IApplicationView, $tagIApplicationView)
ConsoleWrite("Application view = " & IsObj($oView) & @CRLF)
$iHresult = $oView.GetAppUserModelId($sView)
ConsoleWrite("Get App ID = " & $sView & "/" & $iHresult & @CRLF)

; verify if app is pinned with ptr and string
$iHresult = $oVirtualDesktopPinnedApps.IsViewPinned($pView, $bValue)
ConsoleWrite("Is View Pinned = " & $bValue & "/" & $iHresult & @CRLF)
$iHresult = $oVirtualDesktopPinnedApps.IsAppIdPinned($sView, $bValue)
ConsoleWrite("Is AppId Pinned = " & $bValue & "/" & $iHresult & @CRLF)

; move application to a specific desktop
$iHresult = $oVirtualDesktopManagerInternal.MoveViewToDesktop($pView, $pDesktop)
ConsoleWrite("Move = " & $iHresult & @CRLF)

Sleep(2000)

; deletes an existing desktop
$iHresult = $oVirtualDesktopManagerInternal.RemoveDesktop($pNew, $pDesktop)
ConsoleWrite("Delete = " & $iHresult & @CRLF)

Func __uuidof($sGUID)
    Local $tGUID = DllStructCreate("ulong Data1;ushort Data2;ushort Data3;byte Data4[8]")
    DllCall("ole32.dll", "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $tGUID)
    If @error Then Return SetError(@error, @extended, 0)
    Return $tGUID
EndFunc   ;==>__uuidof
