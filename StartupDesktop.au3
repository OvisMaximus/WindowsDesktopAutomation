#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         Laura

 Script Function:
	Clean up the desktop after boot by moving windows to the apropriate desktop, positioning and sizing windows.

#ce ----------------------------------------------------------------------------
; Script Start - Add your code below here

#include <MsgBoxConstants.au3>
#include "VirtDsktLibrary.au3"
#include "ToolsLibrary.au3"

Local $yOfs = 210
	
launchProgram("C:\Program Files\DesktopOK\DesktopOK_x64.exe /load C:\Users\lorenz\bin\DesktopOK\layout.dok")
	
deleteUserFile("Desktop\Checking your credentials.lnk")
deleteUserFile("Desktop\https   outlook.office365.com.lnk")
deleteUserFile("Desktop\Microsoft Teams.lnk")
deleteUserFile("Desktop\Adobe Acrobat.lnk")

MoveWindow("DPDHL Daily", "", 2, 0, 0, 1920, 1056)
MoveWindow("Posteingang",  "", 2, 3834, 0, 1549, 846) 
MoveWindow("Kalender -",  "", 2, 3834, 840 + $yOfs, 1549, 846 - $yOfs ) 
MoveWindow("[REGEXPTITLE: (.* Microsoft Teams)]",  "", 2, 5377, 0, 1535, 840) 
MoveWindow("Google Kalender ",  "", 2, 5370, 840, 1548, 848) 
MoveWindow("Signal", "", 0, 0, 0, 1920, 1056)
MoveWindow("Chrome Social", "", 0, 1913,0,1933,1060 -$yOfs)
MoveWindow("Telegram", "", 0, 0, 1055, 1920,1056)
MoveWindow("Phone Link", "", 0, 1913, 1055, 1933,1060 - 2 * $yOfs)

MoveWindow("Spotify Premium", "", 1, 692, -1436, 2550, 1385)
MoveWindow("Media", "", 1, 0, 0, 2200, 1600)

MoveWindow("OBS Studio", "", 2, 708, -1420, 2520, 1353)
