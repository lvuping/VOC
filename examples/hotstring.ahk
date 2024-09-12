^g::
InputBox, userInput, Input Box, Please enter text:
    originalClipboard := ClipboardAll ; 현재 클립보드 내용을 저장

    if (userInput = "data")
    {
        Clipboard := "Dear [Name],`r`n`r`n" 
        . "The requested data modifications have been completed.`r`n"
        . "Please review the changes at your earliest convenience.`r`n"
        . "Should you find any issues or require further assistance, please do not hesitate to contact me via email.`r`n"
        . "Thank you for your attention to this matter.`r`n`r`n"
        . "Best regards,"
    }
    else if (userInput = "status")
    {
        Clipboard := "Dear Name,`r`n"
        . "The status has been updated as requested.`r`n"
        . "Please review the changes.`r`n"
        . "If there are any issues or further adjustments needed, please let me know via email."
    }
    else if (userInput = "anal")
    {
        Clipboard := "Hi [Recipient's Name],`r`n`r`n"
        . "I hope you're doing well.`r`n`r`n"
        . "I want to share the analysis we did on [specific topic or data]. I've attached the detailed report for you to check out.`r`n`r`n"
        . "Here's a quick summary:`r`n`r`n"
        . "[Key Point 1]`r`n"
        . "[Key Point 2]`r`n"
        . "[Key Point 3]`r`n"
        . "The report covers all our findings, including [any significant insights or conclusions]. I think you'll find this information useful for [specific purpose or project].`r`n`r`n"
        . "If you have any questions or need more details, just let me know.`r`n`r`n"
        . "Best regards,"
    }
    else if (userInput = "start")
    {
        Clipboard := "Hi ,`r`n`r`n"
        . "Thank you for reaching out to us with your requests.`r`n`r`n"
        . "We have received your VOC and will review it as soon as possible. You can expect a response from us shortly.`r`n`r`n"
        . "If you have any additional information or questions in the meantime, please feel free to let us know.`r`n`r`n"
        . "Thank you for your patience.`r`n`r`n"
        . "Best regards,"
    }
    else
    {
        MsgBox, Error: Unrecognized input.
        Clipboard := originalClipboard ; 클립보드 내용을 원래대로 복원
        return
    }

    ClipWait, 1 ; Clipboard 내용이 설정될 때까지 대기
    Send, ^v
    Sleep, 100 ; 붙여넣기가 완료될 때까지 잠시 대기
    Clipboard := originalClipboard ; 클립보드 내용을 원래대로 복원
return

; ::o1:: ①
; ::o2:: ②
; ::o3:: ③
; ::o4:: ④
; ::o5:: ⑤
; ::o6:: ⑥
; ::o7:: ⑦
; ::o8:: ⑧
; ::o9:: ⑨

