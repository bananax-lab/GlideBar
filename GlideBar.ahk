#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
; GlideBar v1.242
; F15 + mouse movement scroll utility
;
; v1.242 changes:
; - Parameter tuning only.
; - Excel horizontal: unchanged from v1.241.
; - Excel vertical: about 20x faster than v1.241.
; - Non-Excel apps such as JMP / Notepad / Chrome: unchanged.
; ============================================================


; ============================================================
; はじめに：パラメータ指定方法
; ============================================================
;
; 基本的に、ユーザーが触るのはこの設定ブロックだけ。
;
; ------------------------------------------------------------
; 共通設定
; ------------------------------------------------------------
;
; axisLockThreshold
;   横/縦の方向判定を始めるまでの移動量。
;   大きいほど誤爆しにくい。
;
; axisDominance
;   横/縦どちらが優勢かを判定する倍率。
;   大きいほど、明確に横/縦へ動かさないとロックされない。
;
; ------------------------------------------------------------
; Excel以外のアプリ設定
; ------------------------------------------------------------
;
; AddWheelApp(
;     "アプリ名.EXE",
;
;     縦スクロール感度,
;     縦最大ステップ,
;
;     横スクロール感度,
;     横最大ステップ
; )
;
; 感度：
;   小さいほど速い。
;   大きいほど遅い。
;   小数OK。例：1.5
;
; 最大ステップ：
;   1回のRaw Inputで送る最大Wheel数。
;   大きいほど速いが、スクロールの粒度は荒くなりやすい。
;
; v1.242の基本思想：
;   縦スクロールは横スクロールよりかなり速くする。
;   Excel以外はすべて WheelUp/Down / WheelLeft/Right で統一する。
;
; ------------------------------------------------------------
; Excel設定
; ------------------------------------------------------------
;
; AddExcelApp(
;     "EXCEL.EXE",
;
;     横：何pxで1ステップ,
;     横：1ステップで何列動かすか,
;     横：最大ステップ,
;
;     縦：何pxで1ステップ,
;     縦：1ステップで何行動かすか,
;     縦：最大ステップ
; )
;
; Excel横：
;   1px = 1列。v1.241から変更なし。
;
; Excel縦：
;   1px = 400行。v1.241比で約20倍速。
;
; ============================================================


; ============================================================
; 共通設定
; ============================================================

triggerKey := "F15"

; v1.21以降、軸ロックは問題なしとのことなので維持
axisLockThreshold := 8
axisDominance     := 1.25

; F15押下直後のRaw Inputノイズを無視
ignoreRawInputAfterStartMs := 100

; カーソルをその場に固定する
freezeCursorDuringDrag := true

; Excel誤作動の原因になりやすかったので標準OFF
parkCursorOnStart := false


; ============================================================
; アプリ別設定
; ============================================================

Profiles := Map()
DEFAULT_PROFILE := ""

; ------------------------------------------------------------
; Default / その他アプリ
; ------------------------------------------------------------
; v1.241から変更なし。
;
; 縦が速すぎる：
;   1.5 → 2 → 3
;
; 縦が荒い：
;   24 → 16 → 12
SetDefaultWheelApp(
    1.5,   ; 縦スクロール感度：小さいほど速い
    24,    ; 縦最大ステップ

    22,    ; 横スクロール感度：小さいほど速い
    4      ; 横最大ステップ
)

; ------------------------------------------------------------
; Excel
; ------------------------------------------------------------
; v1.242で縦だけさらに高速化。
;
; 横：
;   v1.241のまま。
;   1px = 1列。横は問題なしなので変更なし。
;
; 縦：
;   v1.241の 1px = 20行 から、
;   1px = 400行 に変更。
;   約20倍速。
;
; Excel縦が速すぎる：
;   縦 1,400,300 → 1,200,300
;   さらに控えめなら 1,100,300
;
; Excel縦をさらに速く：
;   縦 1,400,300 → 1,600,300
;
; JMP / メモ帳 / Chrome などExcel以外は変更なし。
AddExcelApp(
    "EXCEL.EXE",

    1,     ; 横：何pxで1ステップ
    1,     ; 横：1ステップで何列動かすか
    200,   ; 横：最大ステップ

    1,     ; 縦：何pxで1ステップ
    400,   ; 縦：1ステップで何行動かすか
    300    ; 縦：最大ステップ
)

; ------------------------------------------------------------
; JMP
; ------------------------------------------------------------
; v1.241から変更なし。
; Excel以外はすべて同じWheel方式に統一。
; カーソルキー方式は使わない。
;
; 横について：
;   JMP側が WheelRight 1回 = 3列 と解釈する場合、
;   AHK側だけでは完全な1列単位化は難しい。
;
; v1.242でも、
;   横最大ステップを 1 にして、
;   「3列 × 複数回送信」で6列/9列/12列飛ぶことを防ぐ。
AddWheelApp(
    "JMP.EXE",

    1.5,   ; 縦スクロール感度
    24,    ; 縦最大ステップ

    12,    ; 横スクロール感度
    1      ; 横最大ステップ
)

AddWheelApp(
    "JMPPRO.EXE",

    1.5,   ; 縦スクロール感度
    24,    ; 縦最大ステップ

    12,    ; 横スクロール感度
    1      ; 横最大ステップ
)

; ------------------------------------------------------------
; Chrome
; ------------------------------------------------------------
; v1.241から変更なし。
AddWheelApp(
    "CHROME.EXE",

    1.25,  ; 縦スクロール感度
    24,    ; 縦最大ステップ

    22,    ; 横スクロール感度
    4      ; 横最大ステップ
)

AddWheelApp(
    "MSEDGE.EXE",

    1.25,  ; 縦スクロール感度
    24,    ; 縦最大ステップ

    22,    ; 横スクロール感度
    4      ; 横最大ステップ
)

; ------------------------------------------------------------
; メモ帳
; ------------------------------------------------------------
; v1.241から変更なし。
AddWheelApp(
    "NOTEPAD.EXE",

    1.5,   ; 縦スクロール感度
    20,    ; 縦最大ステップ

    24,    ; 横スクロール感度
    3      ; 横最大ステップ
)

; ------------------------------------------------------------
; PDF系
; ------------------------------------------------------------
; v1.241から変更なし。
AddWheelApp(
    "ACROBAT.EXE",

    1.25,  ; 縦スクロール感度
    24,    ; 縦最大ステップ

    24,    ; 横スクロール感度
    3      ; 横最大ステップ
)

AddWheelApp(
    "ACRORD32.EXE",

    1.25,  ; 縦スクロール感度
    24,    ; 縦最大ステップ

    24,    ; 横スクロール感度
    3      ; 横最大ステップ
)


; ============================================================
; ここから下は基本的に触らない
; ============================================================

isDragging := false
axisLock := ""                ; "", "H", "V"

totalDx := 0
totalDy := 0

horizontalAccum := 0.0
verticalAccum   := 0.0

startMouseX := 0
startMouseY := 0

activeHwnd := 0
activeProcessName := ""
targetScrollHwnd := 0

currentProfile := DEFAULT_PROFILE

ignoreRawInputUntil := 0

WM_INPUT := 0x00FF
RID_INPUT := 0x10000003
RIM_TYPEMOUSE := 0

RAWINPUTHEADER_SIZE := (A_PtrSize = 8) ? 24 : 16


; ============================================================
; Startup
; ============================================================

SendMode "Event"
SetKeyDelay -1, -1
CoordMode "Mouse", "Screen"

RegisterRawMouse()
OnMessage(WM_INPUT, OnRawInput)
OnExit(OnScriptExit)


; ============================================================
; Hotkeys
; ============================================================

*F15::StartGlideBar()
*F15 Up::StopGlideBar()

; クリック無効化：GlideBar動作中のみ
#HotIf IsGlideBarActive()
*LButton::Return
*LButton Up::Return
*RButton::Return
*RButton Up::Return
*MButton::Return
*MButton Up::Return
*XButton1::Return
*XButton1 Up::Return
*XButton2::Return
*XButton2 Up::Return
#HotIf


; ============================================================
; Profile helper
; ============================================================

SetDefaultWheelApp(verticalUnit, verticalMax, horizontalUnit, horizontalMax) {
    global Profiles, DEFAULT_PROFILE

    DEFAULT_PROFILE := MakeWheelProfile(
        verticalUnit,
        verticalMax,
        horizontalUnit,
        horizontalMax
    )

    Profiles["DEFAULT"] := DEFAULT_PROFILE
}

AddWheelApp(processName, verticalUnit, verticalMax, horizontalUnit, horizontalMax) {
    global Profiles

    Profiles[StrUpper(processName)] := MakeWheelProfile(
        verticalUnit,
        verticalMax,
        horizontalUnit,
        horizontalMax
    )
}

AddExcelApp(
    processName,
    horizontalUnit,
    horizontalColumnsPerStep,
    horizontalMaxSteps,
    verticalUnit,
    verticalRowsPerStep,
    verticalMaxSteps
) {
    global Profiles

    Profiles[StrUpper(processName)] := Map(
        "mode", "EXCEL",

        "excelHorizontalUnit", horizontalUnit,
        "excelHorizontalColumnsPerStep", horizontalColumnsPerStep,
        "excelHorizontalMaxSteps", horizontalMaxSteps,

        "excelVerticalUnit", verticalUnit,
        "excelVerticalRowsPerStep", verticalRowsPerStep,
        "excelVerticalMaxSteps", verticalMaxSteps
    )
}

MakeWheelProfile(verticalUnit, verticalMax, horizontalUnit, horizontalMax) {
    return Map(
        "mode", "WHEEL",

        "verticalUnit", verticalUnit,
        "verticalMax", verticalMax,

        "horizontalUnit", horizontalUnit,
        "horizontalMax", horizontalMax
    )
}

LoadActiveProfile() {
    global activeHwnd, activeProcessName, currentProfile
    global Profiles, DEFAULT_PROFILE

    try {
        activeHwnd := WinGetID("A")
        activeProcessName := StrUpper(WinGetProcessName("ahk_id " activeHwnd))
    } catch {
        activeHwnd := 0
        activeProcessName := ""
    }

    if Profiles.Has(activeProcessName) {
        currentProfile := Profiles[activeProcessName]
    } else {
        currentProfile := DEFAULT_PROFILE
    }
}

GetProfileValue(key, fallback := "") {
    global currentProfile, DEFAULT_PROFILE

    if IsObject(currentProfile) && currentProfile.Has(key)
        return currentProfile[key]

    if IsObject(DEFAULT_PROFILE) && DEFAULT_PROFILE.Has(key)
        return DEFAULT_PROFILE[key]

    return fallback
}


; ============================================================
; Main control
; ============================================================

StartGlideBar(*) {
    global isDragging
    global axisLock, totalDx, totalDy, horizontalAccum, verticalAccum
    global startMouseX, startMouseY
    global targetScrollHwnd, activeHwnd
    global ignoreRawInputUntil
    global ignoreRawInputAfterStartMs
    global freezeCursorDuringDrag, parkCursorOnStart

    if isDragging
        return

    LoadActiveProfile()

    isDragging := true

    axisLock := ""
    totalDx := 0
    totalDy := 0
    horizontalAccum := 0.0
    verticalAccum := 0.0

    tmpWinHwnd := 0
    tmpCtrlHwnd := 0

    MouseGetPos &startMouseX, &startMouseY, &tmpWinHwnd, &tmpCtrlHwnd, 2

    targetScrollHwnd := tmpCtrlHwnd ? tmpCtrlHwnd : (tmpWinHwnd ? tmpWinHwnd : activeHwnd)

    ; F15押下直後のRaw Inputノイズを捨てる
    ignoreRawInputUntil := A_TickCount + ignoreRawInputAfterStartMs

    ; 標準OFF推奨。
    ; ExcelでONにすると、退避移動がRaw Inputに混ざることがある。
    if parkCursorOnStart {
        try MouseMove 0, A_ScreenHeight - 1, 0
        ignoreRawInputUntil := A_TickCount + ignoreRawInputAfterStartMs
    }

    if freezeCursorDuringDrag {
        try BlockInput "MouseMove"
    }

    SetTimer WatchdogGlideBar, 50
}

StopGlideBar(*) {
    global isDragging
    global axisLock, totalDx, totalDy, horizontalAccum, verticalAccum
    global startMouseX, startMouseY
    global ignoreRawInputUntil
    global freezeCursorDuringDrag, parkCursorOnStart

    if !isDragging
        return

    isDragging := false

    SetTimer WatchdogGlideBar, 0

    if freezeCursorDuringDrag {
        try BlockInput "MouseMoveOff"
    }

    if parkCursorOnStart {
        try MouseMove startMouseX, startMouseY, 0
    }

    axisLock := ""
    totalDx := 0
    totalDy := 0
    horizontalAccum := 0.0
    verticalAccum := 0.0
    ignoreRawInputUntil := 0
}

WatchdogGlideBar() {
    global isDragging, triggerKey

    ; F15 Upを取りこぼした場合の保険
    if isDragging && !GetKeyState(triggerKey, "P") {
        StopGlideBar()
    }
}

IsGlideBarActive(*) {
    global isDragging, triggerKey
    return isDragging && GetKeyState(triggerKey, "P")
}


; ============================================================
; Raw Input
; ============================================================

RegisterRawMouse() {
    ; RAWINPUTDEVICE
    ; USHORT usUsagePage
    ; USHORT usUsage
    ; DWORD  dwFlags
    ; HWND   hwndTarget

    ridSize := 8 + A_PtrSize
    rid := Buffer(ridSize, 0)

    RIDEV_INPUTSINK := 0x00000100

    NumPut("UShort", 1, rid, 0)                  ; Generic Desktop
    NumPut("UShort", 2, rid, 2)                  ; Mouse
    NumPut("UInt", RIDEV_INPUTSINK, rid, 4)
    NumPut("Ptr", A_ScriptHwnd, rid, 8)

    ok := DllCall(
        "RegisterRawInputDevices",
        "Ptr", rid.Ptr,
        "UInt", 1,
        "UInt", ridSize,
        "UInt"
    )

    if !ok {
        MsgBox "Raw Input の登録に失敗しました。GlideBarを終了します。", "GlideBar v1.242"
        ExitApp
    }
}

OnRawInput(wParam, lParam, msg, hwnd) {
    global isDragging
    global RID_INPUT, RIM_TYPEMOUSE, RAWINPUTHEADER_SIZE
    global ignoreRawInputUntil

    if !isDragging
        return

    if A_TickCount < ignoreRawInputUntil
        return

    Critical 5

    size := 0

    DllCall(
        "GetRawInputData",
        "Ptr", lParam,
        "UInt", RID_INPUT,
        "Ptr", 0,
        "UInt*", &size,
        "UInt", RAWINPUTHEADER_SIZE,
        "UInt"
    )

    if size <= 0
        return

    raw := Buffer(size, 0)

    result := DllCall(
        "GetRawInputData",
        "Ptr", lParam,
        "UInt", RID_INPUT,
        "Ptr", raw.Ptr,
        "UInt*", &size,
        "UInt", RAWINPUTHEADER_SIZE,
        "UInt"
    )

    if result = 0xFFFFFFFF
        return

    rawType := NumGet(raw, 0, "UInt")
    if rawType != RIM_TYPEMOUSE
        return

    ; RAWMOUSE begins after RAWINPUTHEADER.
    ; RAWMOUSE layout:
    ; USHORT usFlags        offset 0
    ; ULONG  ulButtons      offset 4
    ; ULONG  ulRawButtons   offset 8
    ; LONG   lLastX         offset 12
    ; LONG   lLastY         offset 16

    mouseOffset := RAWINPUTHEADER_SIZE

    dx := NumGet(raw, mouseOffset + 12, "Int")
    dy := NumGet(raw, mouseOffset + 16, "Int")

    if dx = 0 && dy = 0
        return

    ProcessMouseDelta(dx, dy)
}


; ============================================================
; Axis lock
; ============================================================

ProcessMouseDelta(dx, dy) {
    global axisLock
    global totalDx, totalDy

    if axisLock = "" {
        totalDx += dx
        totalDy += dy

        TryDecideAxisLock()

        ; まだ斜めっぽくて判定できない場合はスクロールしない
        if axisLock = ""
            return

        ; ロック確定時点までの移動量も、選ばれた軸のスクロールに使う
        if axisLock = "H" {
            HandleHorizontalDelta(totalDx)
        } else if axisLock = "V" {
            HandleVerticalDelta(totalDy)
        }

        totalDx := 0
        totalDy := 0
        return
    }

    if axisLock = "H" {
        HandleHorizontalDelta(dx)
    } else if axisLock = "V" {
        HandleVerticalDelta(dy)
    }
}

TryDecideAxisLock() {
    global axisLock
    global totalDx, totalDy
    global axisLockThreshold, axisDominance

    absDx := Abs(totalDx)
    absDy := Abs(totalDy)

    if absDx >= axisLockThreshold && absDx > absDy * axisDominance {
        axisLock := "H"
        return
    }

    if absDy >= axisLockThreshold && absDy > absDx * axisDominance {
        axisLock := "V"
        return
    }
}


; ============================================================
; Horizontal / Vertical handlers
; ============================================================

HandleHorizontalDelta(dx) {
    global horizontalAccum

    if dx = 0
        return

    mode := GetProfileValue("mode", "WHEEL")

    if mode = "EXCEL" {
        horizontalAccum += dx

        steps := TakeSteps(
            &horizontalAccum,
            GetProfileValue("excelHorizontalUnit", 1),
            GetProfileValue("excelHorizontalMaxSteps", 200)
        )

        if steps = 0
            return

        columns := steps * GetProfileValue("excelHorizontalColumnsPerStep", 1)

        if !ExcelScrollColumns(columns) {
            SendHorizontalWheel(steps)
        }

        return
    }

    horizontalAccum += dx

    steps := TakeSteps(
        &horizontalAccum,
        GetProfileValue("horizontalUnit", 22),
        GetProfileValue("horizontalMax", 4)
    )

    if steps = 0
        return

    SendHorizontalWheel(steps)
}

HandleVerticalDelta(dy) {
    global verticalAccum

    if dy = 0
        return

    mode := GetProfileValue("mode", "WHEEL")

    if mode = "EXCEL" {
        verticalAccum += dy

        steps := TakeSteps(
            &verticalAccum,
            GetProfileValue("excelVerticalUnit", 1),
            GetProfileValue("excelVerticalMaxSteps", 300)
        )

        if steps = 0
            return

        rows := steps * GetProfileValue("excelVerticalRowsPerStep", 400)

        if !ExcelScrollRows(rows) {
            SendVerticalWheel(steps)
        }

        return
    }

    verticalAccum += dy

    steps := TakeSteps(
        &verticalAccum,
        GetProfileValue("verticalUnit", 1.5),
        GetProfileValue("verticalMax", 24)
    )

    if steps = 0
        return

    SendVerticalWheel(steps)
}

TakeSteps(&accum, unit, maxSteps) {
    if unit <= 0
        return 0

    if Abs(accum) < unit
        return 0

    if accum > 0 {
        steps := Floor(accum / unit)
    } else {
        steps := Ceil(accum / unit)
    }

    if Abs(steps) > maxSteps {
        steps := steps > 0 ? maxSteps : -maxSteps
    }

    accum -= steps * unit
    return steps
}


; ============================================================
; Excel direct scroll
; ============================================================

ExcelScrollColumns(columns) {
    if columns = 0
        return true

    try {
        xl := ComObjActive("Excel.Application")
        win := xl.ActiveWindow

        currentColumn := win.ScrollColumn
        nextColumn := currentColumn + columns

        if nextColumn < 1
            nextColumn := 1

        win.ScrollColumn := nextColumn
        return true
    } catch {
        return false
    }
}

ExcelScrollRows(rows) {
    if rows = 0
        return true

    try {
        xl := ComObjActive("Excel.Application")
        win := xl.ActiveWindow

        currentRow := win.ScrollRow
        nextRow := currentRow + rows

        if nextRow < 1
            nextRow := 1

        win.ScrollRow := nextRow
        return true
    } catch {
        return false
    }
}


; ============================================================
; Wheel scroll
; ============================================================

SendHorizontalWheel(steps) {
    if steps = 0
        return

    key := steps > 0 ? "{WheelRight}" : "{WheelLeft}"
    count := Abs(steps)

    Loop count {
        SendEvent key
    }
}

SendVerticalWheel(steps) {
    if steps = 0
        return

    ; dy > 0 = mouse down = scroll down
    key := steps > 0 ? "{WheelDown}" : "{WheelUp}"
    count := Abs(steps)

    Loop count {
        SendEvent key
    }
}


; ============================================================
; Cleanup
; ============================================================

OnScriptExit(exitReason, exitCode) {
    ForceCleanup()
}

ForceCleanup() {
    global freezeCursorDuringDrag

    try SetTimer WatchdogGlideBar, 0

    if freezeCursorDuringDrag {
        try BlockInput "MouseMoveOff"
    }
}