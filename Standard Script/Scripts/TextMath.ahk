



; TextMath - evaluates math based on just highlighted text
^+=:: Send, % Eval(get_Highlight())
;-------------------------------------------------------------------------------
get_Highlight() { ; get highlighted text
;-------------------------------------------------------------------------------
    $_AutoTrim := A_AutoTrim            ; remember current setting
    AutoTrim, Off                       ; keep leading/trailing white spaces
    ClipStore := ClipboardAll           ; remember current content

    Clipboard := ""                     ; empty clipboard
    Send, ^c                            ; copy highlighted text
    ClipWait, 1                         ; wait for change
    Result := Clipboard                 ; store highlighted text

    Clipboard := ClipStore              ; restore previous content
    AutoTrim, % $_AutoTrim              ; restore previous setting
    VarSetCapacity(ClipStore, 0)        ; free memory

    Return, Result
}

;-------------------------------------------------------------------------------
Eval(Exp) { ; use Javascript/COM, by TLM and tidbit
;-------------------------------------------------------------------------------
    static Constants := "E|LN2|LN10|LOG2E|LOG10E|PI|SQRT1_2|SQRT2"
    static Functions := "abs|acos|asin|atan|atan2|ceil|cos|exp|floor"
                     .  "|log|max|min|pow|random|round|sin|sqrt|tan"

    Transform, Exp, Deref, %Exp% ; support variables

    Exp := Format("{:L}", Exp) ; everything lowercase
    Exp := RegExReplace(Exp, "i)(" Constants ")", "$U1") ; constants uppercase
    Exp := RegExReplace(Exp, "i)(" Constants "|" Functions ")", "Math.$1")

    Obj := ComObjCreate("HTMLfile")
    Obj.Write("<body><script>document.body.innerText=eval('" Exp "');</script>")
    Return, InStr(Result := Obj.body.innerText, "body") ? "ERROR" : Result
}

return


