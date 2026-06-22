#Requires AutoHotkey v2.0
#SingleInstance Force
#MaxThreadsPerHotkey 1
Persistent

; ============================================================
; GlideBar v1.344
; ------------------------------------------------------------
; Base:
;   v1.322 Excel behavior
;   v1.331 Non-Excel behavior
;
; v1.344 changes:
;   - Excel:
;       Parameter-only tuning from v1.343.
;       Keep minimum speed, maximum speed, and acceleration strength unchanged.
;       Widen only the smooth transition band to reduce stepped speed changes.
;       Excel horizontal and vertical remain isolated from Non-Excel tuning.
;
;   - Non-Excel:
;       Parameter-only tuning from v1.343.
;       Keep minimum and maximum vertical speed unchanged.
;       Start acceleration a little earlier near low speed.
;       Widen high-speed transition so JMP connects to top speed more mildly.
;       Split wheel messages into smaller chunks while keeping total per-tick maximum unchanged.
;
;   - Design:
;       Pointer acceleration behavior is split into two profiles:
;         1) Excel
;         2) Non-Excel apps
;       App-specific handling for Non-Excel remains parameter-only.
;
;   - Cursor jitter fix remains:
;       Raw Input + ClipCursor
;
;   - Freeze / "too many hotkeys" mitigation remains:
;       KeyWait-based F15 trigger.
;       A_HotkeyInterval / A_MaxHotkeysPerInterval.
; ============================================================

; AHK v2 hotkey storm protection
A_HotkeyInterval := 2000
A_MaxHotkeysPerInterval := 1000

CoordMode "Mouse", "Screen"
SendMode "Input"
SetMouseDelay -1

; ============================================================
; Startup
; ============================================================

global WM_INPUT := 0x00FF

OnMessage WM_INPUT, RawInputProc
RegisterRawMouse()
OnExit CleanupOnExit

; ============================================================
; User parameters
; ============================================================

global GB_VERSION := "v1.344"

; Trigger
global TriggerKey := "F15"

; Timer
global TickMs := 10

; Axis lock
global LockThreshold := 6
global AxisDominance := 1.25

; Deadzone
global DeadzoneX := 0.3
global DeadzoneY := 0.3

; ------------------------------------------------------------
; Acceleration profile split
; ------------------------------------------------------------
; v1.34+ separates pointer acceleration behavior into:
;   1) Excel
;   2) Non-Excel apps
;
; Reason:
;   Excel uses direct ScrollColumn / ScrollRow control.
;   Non-Excel uses wheel messages.
;   The same mouse delta can feel very different in these outputs.
; ------------------------------------------------------------

; ------------------------------------------------------------
; Excel acceleration profile
; ------------------------------------------------------------
; Reproduce v1.322 Excel behavior.
; Keep these independent from Non-Excel tuning.
; ------------------------------------------------------------

; Excel horizontal in v1.322 used:
;   AccelX := 1.0
;   AccelBase := 20.0
;   MaxAdjustedDeltaX := 40.0
global ExcelXAccel := 1.0
global ExcelXAccelBase := 20.0
global ExcelXMaxAdjustedDelta := 40.0

; Excel vertical dedicated acceleration from v1.322.
global ExcelYAccelStart := 2.0
global ExcelYFastAccel := 240.0
global ExcelYFastAccelBase := 20.0
global ExcelYMaxAdjustedDelta := 3600.0

global ExcelYTransitionRange := 240.0

; ------------------------------------------------------------
; Non-Excel acceleration profile
; ------------------------------------------------------------
; v1.344 keeps v1.331 Non-Excel output method, with parameter-only tuning.
; Non-Excel apps all use the same output method.
; App-specific handling is parameter-only.
; ------------------------------------------------------------

; Non-Excel horizontal in v1.331 used:
;   AccelX := 1.0
;   AccelBase := 20.0
;   NonExcelAccelBoostX := 18.0
;   NonExcelMaxAdjustedDeltaX := 100.0
global NonExcelXAccel := 1.0
global NonExcelXAccelBase := 20.0
global NonExcelAccelBoostX := 18.0
global NonExcelMaxAdjustedDeltaX := 100.0

; Non-Excel vertical base from v1.331, tuned in v1.341 / v1.342:
;   v1.331 / v1.34 NonExcelAccelBoostY := 13500.0
;   v1.341       NonExcelAccelBoostY := 20000.0
;   v1.342 / v1.343 NonExcelAccelBoostY := 30000.0
;   v1.344       NonExcelAccelBoostY := 50000.0
;
;   v1.331 / v1.34 NonExcelMaxAdjustedDeltaY := 50000.0
;   v1.341       NonExcelMaxAdjustedDeltaY := 500000.0
;   v1.342 / v1.343 NonExcelMaxAdjustedDeltaY := 1000000.0
;
;   v1.344 keeps the same maximum speed as v1.342 / v1.343,
;   but widens the high-speed transition and uses smaller wheel chunks:
;     - smoother connection near top speed
;     - less jumpy acceleration in JMP
;     - earlier low-speed acceleration start
global NonExcelYAccel := 1.0
global NonExcelAccelBoostY := 50000.0

; v1.33: 30000.0
; v1.331 / v1.34: 50000.0
; v1.341: 500000.0
; v1.342: 1000000.0
; With WheelYStepsPerPx := 0.60, this allows about 600000 wheel steps/tick.
global NonExcelMaxAdjustedDeltaY := 1000000.0

; Non-Excel vertical smooth transition.
; Larger = smoother / slower transition to high speed.
; Smaller = more aggressive.
; v1.344:
;   AccelStart 1.0 -> 0.5 for earlier low-speed acceleration start.
;   TransitionRange 420.0 -> 480.0 for smoother high-speed connection.
global NonExcelYAccelStart := 0.5
global NonExcelYTransitionRange := 480.0
global NonExcelYFastAccelBase := 20.0

; ============================================================
; Output parameters
; ============================================================

; ------------------------------------------------------------
; Excel direct scroll
; ------------------------------------------------------------
; Same as v1.321 / v1.322 / v1.33.
; ------------------------------------------------------------

global ExcelXStepsPerPx := 0.25
global ExcelYStepsPerPx := 6.0

global ExcelMaxStepsPerTickX := 60
global ExcelMaxStepsPerTickY := 15000

; ------------------------------------------------------------
; Non-Excel wheel scroll
; ------------------------------------------------------------
; All non-Excel apps use unified PostMessage wheel output.
;
; Horizontal:
;   Same basic speed as v1.33.
;
; Vertical:
;   Base speed remains 0.60, so very-low-speed feel is preserved.
;   v1.343 keeps the same high-speed maximum as v1.342.
;   Acceleration starts earlier, while top-speed transition is widened.
;
; Individual apps can be tuned only by these parameters.
; ------------------------------------------------------------

global DefaultWheelXStepsPerPx := 0.05
global DefaultWheelYStepsPerPx := 0.60

global ChromeWheelXStepsPerPx := 0.05
global ChromeWheelYStepsPerPx := 0.60

global JmpWheelXStepsPerPx := 0.05
global JmpWheelYStepsPerPx := 0.60

global NotepadWheelXStepsPerPx := 0.05
global NotepadWheelYStepsPerPx := 0.60

global WheelMaxStepsPerTickX := 25

; v1.33: 6000
; v1.331 / v1.34: 30000
; v1.341: 300000
; v1.342: 600000
global WheelMaxStepsPerTickY := 600000

; ------------------------------------------------------------
; Unified compressed PostMessage wheel output
; ------------------------------------------------------------
; All non-Excel apps use this method.
;
; WHEEL_DELTA is 120.
; Very large single wheel delta can overflow signed 16-bit high word,
; so split into compressed messages.
;
; 240 wheel steps/message = 28800 wheel delta.
;
; v1.33:
;   240 steps/message * 25 messages/tick = 6000 steps/tick
;
; v1.331:
;   240 steps/message * 125 messages/tick = 30000 steps/tick
;
; v1.341:
;   240 steps/message * 1250 messages/tick = 300000 steps/tick
;
; v1.342:
;   240 steps/message * 2500 messages/tick = 600000 steps/tick
;
; v1.343:
;   120 steps/message * 5000 messages/tick = 600000 steps/tick
;   Same maximum speed, smaller chunks for smoother JMP high-speed scrolling.
;
; v1.344:
;   80 steps/message * 7500 messages/tick = 600000 steps/tick
;   Same maximum speed, even smaller chunks for smoother JMP high-speed scrolling.
; ------------------------------------------------------------

global WheelPostMaxStepsPerMessage := 80
global WheelPostMaxMessagesPerTickX := 4
global WheelPostMaxMessagesPerTickY := 7500

; Safety
global MaxSessionMs := 30000

; ============================================================
; Runtime state
; ============================================================

global g_isDragging := false

global g_startX := 0
global g_startY := 0

global g_targetHwnd := 0
global g_targetTopHwnd := 0
global g_targetExe := ""

global g_axisLock := ""
global g_lockBufX := 0.0
global g_lockBufY := 0.0

global g_accumX := 0.0
global g_accumY := 0.0

; Raw Input pending deltas
global g_pendingDX := 0.0
global g_pendingDY := 0.0

global g_sessionStartTick := 0

global g_xl := ""
global g_excelWin := ""
global g_excelSheet := ""

; ============================================================
; Hotkeys
; ============================================================

; ------------------------------------------------------------
; KeyWait-based trigger.
; Avoids repeated StartGlideBar() by F15 auto-repeat.
; ------------------------------------------------------------

$*F15::
{
    if IsGlideBarActive() {
        return
    }

    StartGlideBar()

    ; Keep this hotkey thread alive until F15 is released.
    ; Repeated F15 down events are ignored by #MaxThreadsPerHotkey 1.
    KeyWait "F15"

    StopGlideBar()
}

; Emergency restore
^!F15::ForceStopGlideBar()

; Extra emergency restore not using F15
^!Esc::ForceStopGlideBar()

; Disable mouse buttons during GlideBar operation
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
; Main control
; ============================================================

StartGlideBar(*) {
    global g_isDragging
    global g_startX, g_startY
    global g_targetHwnd, g_targetTopHwnd, g_targetExe
    global g_axisLock, g_lockBufX, g_lockBufY
    global g_accumX, g_accumY
    global g_pendingDX, g_pendingDY
    global g_sessionStartTick
    global TickMs
    global g_xl, g_excelWin, g_excelSheet

    if g_isDragging {
        return
    }

    ; Capture both top-level window and control under cursor.
    ; Non-Excel output method is still unified:
    ;   compressed PostMessage wheel.
    MouseGetPos &x, &y, &topHwnd, &ctrlHwnd, 2

    g_startX := x
    g_startY := y
    g_targetTopHwnd := topHwnd

    if (ctrlHwnd) {
        g_targetHwnd := ctrlHwnd
    } else {
        g_targetHwnd := topHwnd
    }

    try {
        g_targetExe := WinGetProcessName("ahk_id " topHwnd)
    } catch {
        g_targetExe := ""
    }

    g_axisLock := ""
    g_lockBufX := 0.0
    g_lockBufY := 0.0
    g_accumX := 0.0
    g_accumY := 0.0
    g_pendingDX := 0.0
    g_pendingDY := 0.0

    g_xl := ""
    g_excelWin := ""
    g_excelSheet := ""

    if IsExcelTarget() {
        try {
            g_xl := ComObjActive("Excel.Application")
            g_excelWin := g_xl.ActiveWindow
            g_excelSheet := g_xl.ActiveSheet
        } catch {
            g_xl := ""
            g_excelWin := ""
            g_excelSheet := ""
        }
    }

    ; Cursor jitter fix:
    ; Keep cursor at the original point.
    ; Raw Input still receives physical mouse movement deltas.
    ClipCursorToPoint(g_startX, g_startY)

    g_sessionStartTick := A_TickCount
    g_isDragging := true

    SetTimer GlideBarTick, TickMs
}

StopGlideBar(*) {
    global g_isDragging
    global g_axisLock, g_lockBufX, g_lockBufY
    global g_accumX, g_accumY
    global g_pendingDX, g_pendingDY
    global g_xl, g_excelWin, g_excelSheet
    global g_startX, g_startY

    if !g_isDragging {
        return
    }

    SetTimer GlideBarTick, 0

    g_isDragging := false

    ReleaseCursorClip()

    ; Restore cursor to start point.
    try {
        DllCall("SetCursorPos", "Int", g_startX, "Int", g_startY)
    }

    g_axisLock := ""
    g_lockBufX := 0.0
    g_lockBufY := 0.0
    g_accumX := 0.0
    g_accumY := 0.0
    g_pendingDX := 0.0
    g_pendingDY := 0.0

    g_xl := ""
    g_excelWin := ""
    g_excelSheet := ""
}

ForceStopGlideBar(*) {
    global g_isDragging
    global g_axisLock, g_lockBufX, g_lockBufY
    global g_accumX, g_accumY
    global g_pendingDX, g_pendingDY
    global g_xl, g_excelWin, g_excelSheet

    try {
        SetTimer GlideBarTick, 0
    }

    g_isDragging := false

    ReleaseCursorClip()

    g_axisLock := ""
    g_lockBufX := 0.0
    g_lockBufY := 0.0
    g_accumX := 0.0
    g_accumY := 0.0
    g_pendingDX := 0.0
    g_pendingDY := 0.0

    g_xl := ""
    g_excelWin := ""
    g_excelSheet := ""
}

IsGlideBarActive(*) {
    global g_isDragging
    return g_isDragging
}

CleanupOnExit(ExitReason, ExitCode) {
    try {
        SetTimer GlideBarTick, 0
    }

    ReleaseCursorClip()
}

; ============================================================
; Raw Input
; ============================================================

RegisterRawMouse() {
    ; RAWINPUTDEVICE
    ; usUsagePage = 0x01 generic desktop controls
    ; usUsage     = 0x02 mouse
    ; dwFlags     = RIDEV_INPUTSINK
    ; hwndTarget  = A_ScriptHwnd

    RIDEV_INPUTSINK := 0x00000100

    ridSize := 8 + A_PtrSize
    rid := Buffer(ridSize, 0)

    NumPut("UShort", 0x01, rid, 0)
    NumPut("UShort", 0x02, rid, 2)
    NumPut("UInt", RIDEV_INPUTSINK, rid, 4)
    NumPut("Ptr", A_ScriptHwnd, rid, 8)

    ok := DllCall(
        "RegisterRawInputDevices",
        "Ptr", rid,
        "UInt", 1,
        "UInt", ridSize
    )

    if !ok {
        MsgBox "GlideBar: Raw Input registration failed."
    }
}

RawInputProc(wParam, lParam, msg, hwnd) {
    global g_isDragging
    global g_pendingDX, g_pendingDY

    if !g_isDragging {
        return
    }

    RID_INPUT := 0x10000003
    RIM_TYPEMOUSE := 0
    headerSize := 8 + A_PtrSize * 2

    size := 0

    DllCall(
        "GetRawInputData",
        "Ptr", lParam,
        "UInt", RID_INPUT,
        "Ptr", 0,
        "UInt*", &size,
        "UInt", headerSize
    )

    if (size <= 0) {
        return
    }

    raw := Buffer(size, 0)

    result := DllCall(
        "GetRawInputData",
        "Ptr", lParam,
        "UInt", RID_INPUT,
        "Ptr", raw,
        "UInt*", &size,
        "UInt", headerSize
    )

    if (result = -1) {
        return
    }

    type := NumGet(raw, 0, "UInt")

    if (type != RIM_TYPEMOUSE) {
        return
    }

    ; RAWMOUSE starts after RAWINPUTHEADER.
    ; RAWMOUSE layout:
    ;   usFlags:  offset + 0
    ;   buttons:  offset + 4
    ;   rawBtns:  offset + 8
    ;   lLastX:   offset + 12
    ;   lLastY:   offset + 16
    dx := NumGet(raw, headerSize + 12, "Int")
    dy := NumGet(raw, headerSize + 16, "Int")

    if (dx = 0 && dy = 0) {
        return
    }

    ; Safety clamp for raw accumulation.
    ; Prevent huge burst when the system was busy for a moment.
    g_pendingDX := ClampFloat(g_pendingDX + dx, -500.0, 500.0)
    g_pendingDY := ClampFloat(g_pendingDY + dy, -500.0, 500.0)
}

; ============================================================
; Cursor lock
; ============================================================

ClipCursorToPoint(x, y) {
    ; ClipCursor rectangle:
    ; left, top, right, bottom
    ;
    ; right/bottom are exclusive-ish, so x+1/y+1 gives a 1px box.

    rect := Buffer(16, 0)

    NumPut("Int", x, rect, 0)
    NumPut("Int", y, rect, 4)
    NumPut("Int", x + 1, rect, 8)
    NumPut("Int", y + 1, rect, 12)

    DllCall("ClipCursor", "Ptr", rect)
}

ReleaseCursorClip() {
    DllCall("ClipCursor", "Ptr", 0)
}

; ============================================================
; Tick loop
; ============================================================

GlideBarTick() {
    global g_isDragging
    global g_axisLock, g_lockBufX, g_lockBufY
    global g_sessionStartTick
    global MaxSessionMs
    global TriggerKey
    global LockThreshold, AxisDominance
    global g_pendingDX, g_pendingDY

    if !g_isDragging {
        return
    }

    ; Key-up取りこぼし対策
    if !GetKeyState(TriggerKey, "P") {
        StopGlideBar()
        return
    }

    ; 長時間押しっぱなし安全停止
    if (A_TickCount - g_sessionStartTick > MaxSessionMs) {
        StopGlideBar()
        return
    }

    dx := g_pendingDX
    dy := g_pendingDY

    g_pendingDX := 0.0
    g_pendingDY := 0.0

    if (dx = 0 && dy = 0) {
        return
    }

    ; Axis lock not decided yet
    if (g_axisLock = "") {
        g_lockBufX += dx
        g_lockBufY += dy

        absX := Abs(g_lockBufX)
        absY := Abs(g_lockBufY)

        if (absX >= LockThreshold && absX >= absY * AxisDominance) {
            g_axisLock := "H"

            delta := g_lockBufX
            g_lockBufX := 0.0
            g_lockBufY := 0.0

            ProcessHorizontal(delta)
            return
        }

        if (absY >= LockThreshold && absY >= absX * AxisDominance) {
            g_axisLock := "V"

            delta := g_lockBufY
            g_lockBufX := 0.0
            g_lockBufY := 0.0

            ProcessVertical(delta)
            return
        }

        return
    }

    ; Axis already locked
    if (g_axisLock = "H") {
        ProcessHorizontal(dx)
        return
    }

    if (g_axisLock = "V") {
        ProcessVertical(dy)
        return
    }
}

; ============================================================
; Axis processing
; ============================================================

ProcessHorizontal(rawDelta) {
    global DeadzoneX
    global ExcelXAccel, ExcelXAccelBase, ExcelXMaxAdjustedDelta
    global NonExcelXAccel, NonExcelXAccelBase, NonExcelAccelBoostX, NonExcelMaxAdjustedDeltaX
    global ExcelXStepsPerPx, ExcelMaxStepsPerTickX
    global WheelMaxStepsPerTickX
    global g_accumX

    if IsExcelTarget() {
        ; Excel horizontal:
        ; Reproduce v1.322 with an Excel-only acceleration profile.
        delta := PreprocessDelta(
            rawDelta,
            DeadzoneX,
            ExcelXAccel,
            ExcelXAccelBase,
            ExcelXMaxAdjustedDelta
        )

        if (delta = 0) {
            return
        }

        addSteps := delta * ExcelXStepsPerPx
        g_accumX += addSteps

        steps := TakeWholeSteps(g_accumX)
        g_accumX -= steps

        steps := ClampInt(steps, -ExcelMaxStepsPerTickX, ExcelMaxStepsPerTickX)

        if (steps != 0) {
            ExcelScrollHorizontal(steps)
        }

        return
    }

    ; Non-Excel horizontal:
    ; All non-Excel apps use unified PostMessage wheel output.
    delta := PreprocessDelta(
        rawDelta,
        DeadzoneX,
        NonExcelXAccel * NonExcelAccelBoostX,
        NonExcelXAccelBase,
        NonExcelMaxAdjustedDeltaX
    )

    if (delta = 0) {
        return
    }

    scale := GetWheelXStepsPerPx()
    addSteps := delta * scale
    g_accumX += addSteps

    steps := TakeWholeSteps(g_accumX)
    g_accumX -= steps

    steps := ClampInt(steps, -WheelMaxStepsPerTickX, WheelMaxStepsPerTickX)

    if (steps != 0) {
        SendHorizontalWheel(steps)
    }
}

ProcessVertical(rawDelta) {
    global DeadzoneY
    global NonExcelYAccel, NonExcelAccelBoostY, NonExcelMaxAdjustedDeltaY
    global NonExcelYAccelStart, NonExcelYTransitionRange, NonExcelYFastAccelBase
    global ExcelYStepsPerPx, ExcelMaxStepsPerTickY
    global ExcelYAccelStart, ExcelYFastAccel, ExcelYFastAccelBase, ExcelYMaxAdjustedDelta, ExcelYTransitionRange
    global WheelMaxStepsPerTickY
    global g_accumY

    if IsExcelTarget() {
        ; Excel vertical:
        ; Same as v1.321 / v1.322 / v1.33.
        delta := PreprocessExcelYDeltaSmooth(
            rawDelta,
            DeadzoneY,
            ExcelYFastAccel,
            ExcelYAccelStart,
            ExcelYFastAccelBase,
            ExcelYMaxAdjustedDelta,
            ExcelYTransitionRange
        )

        if (delta = 0) {
            return
        }

        addSteps := delta * ExcelYStepsPerPx
        g_accumY += addSteps

        steps := TakeWholeSteps(g_accumY)
        g_accumY -= steps

        steps := ClampInt(steps, -ExcelMaxStepsPerTickY, ExcelMaxStepsPerTickY)

        if (steps != 0) {
            ExcelScrollVertical(steps)
        }

        return
    }

    ; Non-Excel vertical:
    ; All non-Excel apps use unified PostMessage wheel output.
    ; Low-speed behavior starts rising earlier than v1.341.
    ; v1.343 keeps the high-speed maximum unchanged from v1.342.
    ; TransitionRange is widened further to reduce jumpy top-speed changes in JMP.
    delta := PreprocessNonExcelYDeltaSmooth(
        rawDelta,
        DeadzoneY,
        NonExcelYAccel * NonExcelAccelBoostY,
        NonExcelYAccelStart,
        NonExcelYFastAccelBase,
        NonExcelMaxAdjustedDeltaY,
        NonExcelYTransitionRange
    )

    if (delta = 0) {
        return
    }

    scale := GetWheelYStepsPerPx()
    addSteps := delta * scale
    g_accumY += addSteps

    steps := TakeWholeSteps(g_accumY)
    g_accumY -= steps

    steps := ClampInt(steps, -WheelMaxStepsPerTickY, WheelMaxStepsPerTickY)

    if (steps != 0) {
        SendVerticalWheel(steps)
    }
}

; ============================================================
; Preprocess layer
; ============================================================

PreprocessDelta(rawDelta, deadzone, accel, accelBase, maxAdjustedDelta) {
    ; 1. deadzone
    if (Abs(rawDelta) <= deadzone) {
        return 0.0
    }

    ; 2. acceleration
    adjusted := ApplyAcceleration(rawDelta, accel, accelBase)

    ; 3. clamp / safety limit
    adjusted := ClampFloat(adjusted, -maxAdjustedDelta, maxAdjustedDelta)

    return adjusted
}

PreprocessExcelYDeltaSmooth(rawDelta, deadzone, accel, accelStart, accelBase, maxAdjustedDelta, transitionRange) {
    ; Excel vertical dedicated path.
    ; Same as v1.321 / v1.322 / v1.33.

    if (Abs(rawDelta) <= deadzone) {
        return 0.0
    }

    adjusted := ApplySmoothAcceleration(
        rawDelta,
        accel,
        accelStart,
        accelBase,
        maxAdjustedDelta,
        transitionRange
    )

    adjusted := ClampFloat(adjusted, -maxAdjustedDelta, maxAdjustedDelta)

    return adjusted
}

PreprocessNonExcelYDeltaSmooth(rawDelta, deadzone, accel, accelStart, accelBase, maxAdjustedDelta, transitionRange) {
    ; Non-Excel vertical dedicated path.
    ;
    ; v1.331 concept:
    ;   - Keep low-speed movement close to v1.33.
    ;   - Keep smooth transition unchanged.
    ;   - Raise high-speed max only.
    ;   - Output method is unified for all non-Excel apps.

    if (Abs(rawDelta) <= deadzone) {
        return 0.0
    }

    adjusted := ApplySmoothAcceleration(
        rawDelta,
        accel,
        accelStart,
        accelBase,
        maxAdjustedDelta,
        transitionRange
    )

    adjusted := ClampFloat(adjusted, -maxAdjustedDelta, maxAdjustedDelta)

    return adjusted
}

ApplyAcceleration(delta, accel, accelBase) {
    if (delta = 0) {
        return 0.0
    }

    if (accel = 0) {
        return delta
    }

    absDelta := Abs(delta)

    adjustedAbs := absDelta + accel * absDelta * absDelta / accelBase

    if (delta > 0) {
        return adjustedAbs
    } else {
        return -adjustedAbs
    }
}

ApplySmoothAcceleration(delta, accel, accelStart, accelBase, maxAdjustedDelta, transitionRange) {
    if (delta = 0) {
        return 0.0
    }

    absDelta := Abs(delta)

    ; Very slow area remains linear.
    if (accel = 0 || absDelta <= accelStart) {
        adjustedAbs := absDelta
    } else {
        over := absDelta - accelStart

        ; Linear value keeps slow precision.
        linearAbs := absDelta

        ; Fast target acceleration, capped before blending.
        fastAbs := absDelta + accel * over * over / accelBase
        fastAbs := Min(fastAbs, maxAdjustedDelta)

        ; Smooth transition weight.
        ; Larger transitionRange means slower transition to fast speed.
        if (transitionRange <= 0) {
            weight := 1.0
        } else {
            t := ClampFloat(over / transitionRange, 0.0, 1.0)
            weight := SmoothStep01(t)
        }

        adjustedAbs := linearAbs + (fastAbs - linearAbs) * weight
    }

    adjustedAbs := Min(adjustedAbs, maxAdjustedDelta)

    if (delta > 0) {
        return adjustedAbs
    } else {
        return -adjustedAbs
    }
}

SmoothStep01(t) {
    ; SmoothStep:
    ;   0 -> 0
    ;   1 -> 1
    ;   middle transition is smooth
    ;
    ; formula:
    ;   t * t * (3 - 2 * t)

    t := ClampFloat(t, 0.0, 1.0)

    return t * t * (3 - 2 * t)
}

; ============================================================
; Excel output
; ============================================================

ExcelScrollHorizontal(steps) {
    global g_excelWin, g_excelSheet

    if (g_excelWin = "") {
        return
    }

    try {
        current := g_excelWin.ScrollColumn

        maxCol := 16384
        try {
            maxCol := g_excelSheet.Columns.Count
        }

        target := ClampInt(current + steps, 1, maxCol)

        if (target != current) {
            g_excelWin.ScrollColumn := target
        }
    } catch {
        ; Excel COM failed. Do nothing.
    }
}

ExcelScrollVertical(steps) {
    global g_excelWin, g_excelSheet

    if (g_excelWin = "") {
        return
    }

    try {
        current := g_excelWin.ScrollRow

        maxRow := 1048576
        try {
            maxRow := g_excelSheet.Rows.Count
        }

        target := ClampInt(current + steps, 1, maxRow)

        if (target != current) {
            g_excelWin.ScrollRow := target
        }
    } catch {
        ; Excel COM failed. Do nothing.
    }
}

; ============================================================
; Non-Excel output
; ============================================================

SendHorizontalWheel(steps) {
    ; Unified output method for all non-Excel apps.
    PostWheelMessage("H", steps)
}

SendVerticalWheel(steps) {
    ; Unified output method for all non-Excel apps.
    PostWheelMessage("V", steps)
}

PostWheelMessage(axis, steps) {
    global g_targetHwnd, g_targetTopHwnd
    global g_startX, g_startY
    global WheelPostMaxStepsPerMessage
    global WheelPostMaxMessagesPerTickX, WheelPostMaxMessagesPerTickY

    if (steps = 0) {
        return
    }

    if (g_targetHwnd = 0 && g_targetTopHwnd = 0) {
        return
    }

    WM_MOUSEWHEEL := 0x020A
    WM_MOUSEHWHEEL := 0x020E
    WHEEL_DELTA := 120

    isHorizontal := (axis = "H")
    msg := isHorizontal ? WM_MOUSEHWHEEL : WM_MOUSEWHEEL

    count := Abs(steps)

    if isHorizontal {
        maxMessages := WheelPostMaxMessagesPerTickX
    } else {
        maxMessages := WheelPostMaxMessagesPerTickY
    }

    maxTotal := WheelPostMaxStepsPerMessage * maxMessages
    count := Min(count, maxTotal)

    remaining := count
    sentMessages := 0

    while (remaining > 0 && sentMessages < maxMessages) {
        chunkSteps := Min(remaining, WheelPostMaxStepsPerMessage)

        ; Direction mapping:
        ;   Internal vertical steps > 0 means WheelDown.
        ;   WM_MOUSEWHEEL positive delta means WheelUp.
        ;
        ;   Internal horizontal steps > 0 means WheelRight.
        ;   WM_MOUSEHWHEEL positive delta generally means WheelRight.
        if isHorizontal {
            wheelDelta := (steps > 0) ? (chunkSteps * WHEEL_DELTA) : (-chunkSteps * WHEEL_DELTA)
        } else {
            wheelDelta := (steps > 0) ? (-chunkSteps * WHEEL_DELTA) : (chunkSteps * WHEEL_DELTA)
        }

        wParam := MakeWheelWParam(wheelDelta)
        lParam := MakeMouseLParam(g_startX, g_startY)

        target := g_targetHwnd
        ok := false

        if (target != 0) {
            ok := DllCall(
                "PostMessage",
                "Ptr", target,
                "UInt", msg,
                "Ptr", wParam,
                "Ptr", lParam,
                "Int"
            )
        }

        ; Same output method, fallback target only.
        ; This is not an app-specific scroll method.
        if (!ok && g_targetTopHwnd != 0 && g_targetTopHwnd != target) {
            DllCall(
                "PostMessage",
                "Ptr", g_targetTopHwnd,
                "UInt", msg,
                "Ptr", wParam,
                "Ptr", lParam,
                "Int"
            )
        }

        remaining -= chunkSteps
        sentMessages += 1
    }
}

MakeWheelWParam(wheelDelta) {
    ; wParam:
    ;   high word = signed wheel delta
    ;   low word  = key state, here 0
    ;
    ; Keep only lower 16 bits of signed delta for high word.
    return (wheelDelta & 0xFFFF) << 16
}

MakeMouseLParam(x, y) {
    ; WM_MOUSEWHEEL lParam uses screen coordinates.
    return (x & 0xFFFF) | ((y & 0xFFFF) << 16)
}

; ============================================================
; App profile
; ============================================================

IsExcelTarget() {
    global g_targetExe
    return (StrLower(g_targetExe) = "excel.exe")
}

IsJmpTarget() {
    global g_targetExe
    exe := StrLower(g_targetExe)

    ; JMPの実行ファイル名ゆれ対策
    return InStr(exe, "jmp")
}

IsChromeTarget() {
    global g_targetExe
    return (StrLower(g_targetExe) = "chrome.exe")
}

IsNotepadTarget() {
    global g_targetExe
    exe := StrLower(g_targetExe)
    return (exe = "notepad.exe" || exe = "notepad++.exe")
}

GetWheelXStepsPerPx() {
    global DefaultWheelXStepsPerPx
    global ChromeWheelXStepsPerPx
    global JmpWheelXStepsPerPx
    global NotepadWheelXStepsPerPx

    if IsJmpTarget() {
        return JmpWheelXStepsPerPx
    }

    if IsChromeTarget() {
        return ChromeWheelXStepsPerPx
    }

    if IsNotepadTarget() {
        return NotepadWheelXStepsPerPx
    }

    return DefaultWheelXStepsPerPx
}

GetWheelYStepsPerPx() {
    global DefaultWheelYStepsPerPx
    global ChromeWheelYStepsPerPx
    global JmpWheelYStepsPerPx
    global NotepadWheelYStepsPerPx

    if IsJmpTarget() {
        return JmpWheelYStepsPerPx
    }

    if IsChromeTarget() {
        return ChromeWheelYStepsPerPx
    }

    if IsNotepadTarget() {
        return NotepadWheelYStepsPerPx
    }

    return DefaultWheelYStepsPerPx
}

; ============================================================
; Utility
; ============================================================

TakeWholeSteps(value) {
    ; 小数累積から、0方向に整数分だけ取り出す。
    ; 例:
    ;   0.9  -> 0
    ;   1.2  -> 1
    ;  -0.9  -> 0
    ;  -1.2  -> -1

    if (value >= 0) {
        return Floor(value)
    } else {
        return Ceil(value)
    }
}

ClampInt(value, minValue, maxValue) {
    if (value < minValue) {
        return minValue
    }

    if (value > maxValue) {
        return maxValue
    }

    return value
}

ClampFloat(value, minValue, maxValue) {
    if (value < minValue) {
        return minValue
    }

    if (value > maxValue) {
        return maxValue
    }

    return value
}