;List of general improvements for Windows

    :*O:dat`t:: ;Current date expansion
        v := A_Now
        FormatTime, v, %v%, MM/dd/yyyy
        SendInput %v%
    return

    ^+v:: ; INCREMENTAL PASTE
        RegExMatch(clipboard,"(\d+)$",Number)
        NewNumber:=Number1+1
        If NewNumber < 10
            NewNumber:="0" NewNumber
        Clipboard:=RegExReplace(Clipboard,Number1 "$",NewNumber)
        send ^v
    Return

    ;Put in () Paranthese
    :?*:()::(){left}
    return

    ;Select all and calculate use TextMath shortcut
    ^NumpadEnter::
        Send, ^a
        Send, ^+=
    return

    ;CLOSE WITH CTRL SHIFT RBUTTON
    ^+RButton::
        Send {LButton}
        winclose, A
    return

    ;Window Drag
    Alt & LButton::
        CoordMode, Mouse  ; Switch to screen/absolute coordinates.
        MouseGetPos, EWD_MouseStartX, EWD_MouseStartY, EWD_MouseWin
        WinGetPos, EWD_OriginalPosX, EWD_OriginalPosY,,, ahk_id %EWD_MouseWin%
        WinGet, EWD_WinState, MinMax, ahk_id %EWD_MouseWin%
        if EWD_WinState = 0  ; Only if the window isn't maximized
            SetTimer, EWD_WatchMouse, 10 ; Track the mouse as the user drags it.
    return



; ## Application Specific ##

    ;Add easy keyboard navigation to File Explorer
    #IfWinActive ahk_exe Explorer.EXE
        up::SendInput, {Alt down}{Up}{Alt up} ;Up a folder
        left::SendInput, {Alt down}{Left}{Alt up} ; Back 
        right::SendInput, {Alt down}{Right}{Alt up} ; Forward
    #IfWinActive



; ## Functions ##

    EWD_WatchMouse: ; This is a function that Window Drag requires
        GetKeyState, EWD_LButtonState, LButton, P
        if EWD_LButtonState = U  ; Button has been released, so drag is complete.
        {
            SetTimer, EWD_WatchMouse, off
            return
        }
        GetKeyState, EWD_EscapeState, Escape, P
        if EWD_EscapeState = D  ; Escape has been pressed, so drag is cancelled.
        {
            SetTimer, EWD_WatchMouse, off
            WinMove, ahk_id %EWD_MouseWin%,, %EWD_OriginalPosX%, %EWD_OriginalPosY%
            return
        }
        ; Otherwise, reposition the window to match the change in mouse coordinates
        ; caused by the user having dragged the mouse:
        CoordMode, Mouse
        MouseGetPos, EWD_MouseX, EWD_MouseY
        WinGetPos, EWD_WinX, EWD_WinY,,, ahk_id %EWD_MouseWin%
        SetWinDelay, -1   ; Makes the below move faster/smoother.
        WinMove, ahk_id %EWD_MouseWin%,, EWD_WinX + EWD_MouseX - EWD_MouseStartX, EWD_WinY + EWD_MouseY - EWD_MouseStartY
        EWD_MouseStartX := EWD_MouseX  ; Update for the next timer-call to this subroutine.
        EWD_MouseStartY := EWD_MouseY
    return