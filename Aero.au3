#include <GUIConstants.au3>

$Struct = DllStructCreate("int cxLeftWidth;int cxRightWidth;int cyTopHeight;int cyBottomHeight;")
$sStruct = DllStructCreate("dword;int;ptr;int")

Global $MyArea[4] = [50, 50, 50, 50]

$GUI = GUICreate("Windows Vista DWM", 243, 243)
$Apply = GUICtrlCreateButton("Apply", 80, 104, 83, 25, 0)
$ICE = GUICtrlCreateButton("DWM Check", 80, 134, 83, 25, 0)
GUISetState()

While 1
    $iMsg = GUIGetMsg()
    Switch $iMsg
        Case $GUI_EVENT_CLOSE
            Exit
        Case $Apply
            ;_Aero_ApplyGlass($GUI)
            ;_Aero_EnableBlurBehind($GUI)
            _Aero_EnableBlurBehind($GUI)
        Case $ICE
            If _Aero_ICE() Then
                MsgBox(0, "_Aero_ICE", "DWM is enabled!")
            Else
                MsgBox(0, "_Aero_ICE", "DWM is NOT enabled!")
            EndIf
    EndSwitch
WEnd

; #FUNCTION#;===============================================================================
;
; Name...........: _Aero_ApplyGlass
; Description ...: Applys glass effect to a window
; Syntax.........: _Aero_ApplyGlass($hWnd, [$bColor)
; Parameters ....: $hWnd - Window handle:
;                 $bColor  - Background color
; Return values .: Success - No return
;                 Failure - Returns 0
; Author ........: James Brooks
; Modified.......:
; Remarks .......: Thanks to weaponx!
; Related .......:
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================
Func _Aero_ApplyGlass($hWnd, $bColor = 0x000000)
    If @OSVersion <> "WIN_VISTA" Then
        MsgBox(16, "_Aero_ApplyGlass", "You are not running Vista!")
        Exit
    Else
        GUISetBkColor($bColor); Must be here!
        $Ret = DllCall("dwmapi.dll", "long", "DwmExtendFrameIntoClientArea", "hwnd", $hWnd, "long*", DllStructGetPtr($Struct))
        If @error Then
            Return 0
            SetError(1)
        Else
            Return $Ret
        EndIf
    EndIf
EndFunc ;==>_Aero_ApplyGlass

; #FUNCTION#;===============================================================================
;
; Name...........: _Aero_ApplyGlassArea
; Description ...: Applys glass effect to a window area
; Syntax.........: _Aero_ApplyGlassArea($hWnd, $Area, [$bColor)
; Parameters ....: $hWnd - Window handle:
;                  $Area - Array containing area points
;                 $bColor  - Background color
; Return values .: Success - No return
;                 Failure - Returns 0
; Author ........: James Brooks
; Modified.......:
; Remarks .......: Thanks to monoceres!
; Related .......:
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================
Func _Aero_ApplyGlassArea($hWnd, $Area, $bColor = 0x000000)
    If @OSVersion <> "WIN_VISTA" Then
        MsgBox(16, "_Aero_ApplyGlass", "You are not running Vista!")
        Exit
    Else
        If IsArray($Area) Then
            DllStructSetData($Struct, "cxLeftWidth", $Area[0])
            DllStructSetData($Struct, "cxRightWidth", $Area[1])
            DllStructSetData($Struct, "cyTopHeight", $Area[2])
            DllStructSetData($Struct, "cyBottomHeight", $Area[3])
            GUISetBkColor($bColor); Must be here!
            $Ret = DllCall("dwmapi.dll", "long*", "DwmExtendFrameIntoClientArea", "hwnd", $hWnd, "ptr", DllStructGetPtr($Struct))
            If @error Then
                Return 0
            Else
                Return $Ret
            EndIf
        Else
            MsgBox(16, "_Aero_ApplyGlassArea", "Area specified is not an array!")
        EndIf
    EndIf
EndFunc ;==>_Aero_ApplyGlassArea

; #FUNCTION#;===============================================================================
;
; Name...........: _Aero_EnableBlurBehind
; Description ...: Enables the blur effect on the provided window handle.
; Syntax.........: _Aero_EnableBlurBehind($hWnd)
; Parameters ....: $hWnd - Window handle:
; Return values .: Success - No return
;                 Failure - Returns 0
; Author ........: James Brooks
; Modified.......:
; Remarks .......: Thanks to komalo
; Related .......:
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================
Func _Aero_EnableBlurBehind($hWnd, $bColor = 0x000000)
    If @OSVersion <> "WIN_VISTA" Then
        MsgBox(16, "_Aero_ApplyGlass", "You are not running Vista!")
        Exit
    Else
        Const $DWM_BB_ENABLE = 0x00000001

        DllStructSetData($sStruct, 1, $DWM_BB_ENABLE)
        DllStructSetData($sStruct, 2, "1")
        DllStructSetData($sStruct, 4, "1")

        GUISetBkColor($bColor); Must be here!
        $Ret = DllCall("dwmapi.dll", "int", "DwmEnableBlurBehindWindow", "hwnd", $hWnd, "ptr", DllStructGetPtr($sStruct))
        If @error Then
            Return 0
        Else
            Return $Ret
        EndIf
    EndIf
EndFunc ;==>_Aero_EnableBlurBehind

; #FUNCTION#;===============================================================================
;
; Name...........: _Aero_ICE
; Description ...: Returns 1 if DWM is enabled or 0 if not
; Syntax.........: _Aero_ICE()
; Parameters ....:
; Return values .: Success - Returns 1
;                 Failure - Returns 0
; Author ........: James Brooks
; Modified.......:
; Remarks .......: Thanks to BrettF
; Related .......:
; Link ..........;
; Example .......; Yes
;
;;==========================================================================================
Func _Aero_ICE()
    $ICEStruct = DllStructCreate("int;")
    $Ret = DllCall("dwmapi.dll", "int", "DwmIsCompositionEnabled", "ptr", DllStructGetPtr($ICEStruct))
    If @error Then
        Return 0
    Else
        Return DllStructGetData($ICEStruct, 1)
    EndIf
EndFunc ;==>_Aero_ICE
