; This experimental file will attempt to download a lost file from the Google cache based on the URL in the
; clipboard.  It has been tested for about 5 minutes, so use at your own risk.
;
; To use:
;    Navigate to a page that has a link to a file you want
;    Click the link to see if it's been restored
;    If not, copy the URL from the address bar into the clipboard, and then execute this script
;    If all goes well, you'll end up with the cached version in your Temp folder (A_Temp) and open in Notepad


    ; Make sure we have a valid URL to get from the cache
    OriginalURL := Clipboard
    if (!Instr(OriginalURL, "http:")) {
        MsgBox The clipboard doesn't appear to contain a valid URL.  Copy the URL to the clipboard and try again
        ExitApp
    }

    CacheURL   := "http://webcache.googleusercontent.com/search?q=cache:" OriginalURL "&hl=en&prmd=imvns&strip=1"
    TargetFile := A_Temp "\" RegexReplace(OriginalURL, "^.*/", "")

    ; Use brute force to convert the HTML to text by reading it into an Internet Explorer control and letting
    ; it do the work.  Not very efficient, but it is easy.
    IEDoc := IEDoc(CacheURL, False)
    Text := ie.Document.body.innerText

    ; Get rid of the Google Cache header and write the text to a file
    Text := RegexReplace(Text, "`r", "")
    Text := RegexReplace(Text, "is).*(Text-only|Full) version", "")
    Text := RegexReplace(Text, "^[\r\n ]+", "")
    Text := RegexReplace(Text, "\n", "`r`n")
    IE.Quit()

    F := FileOpen(TargetFile, "w")
    F.Write(Text)
    F.Close()

    Run Notepad.exe "%TargetFile%"

    ExitApp

;=============================================================================================================
; Open a URL in Internet Explorer
;-------------------------------------------------------------------------------------------------------------
IEDoc(URL, Visible=False) {
    global IE
    ie := ComObjCreate("InternetExplorer.Application")

    ie.Visible := Visible
    ie.Navigate(URL)
    loop 100 {
        Sleep 100
        if ie.ReadyState == 4 {
            break
        }
    }

    doc := ie.Document
    loop 100 {
        Sleep 100
        if doc.readyState == "complete" {
            break
        }
    }

    return doc
}

