; =========================
; MASTER SCRIPT SETUP
; =========================

; ------------------------
; CONFIG SETUP
; ------------------------
global CONFIG := Map()
InitConfig()  ; initialize config values

InitConfig() {
    global CONFIG

    ; Base directory for config
    CONFIG["dir"]  := A_AppData "\DT_Scripts"
    CONFIG["file"] := CONFIG["dir"] "\config.ini"

    ; Ensure folder exists
    if !DirExist(CONFIG["dir"])
        DirCreate(CONFIG["dir"])

    ; If config file doesn't exist, prompt for first + last name
    if !FileExist(CONFIG["file"]) {

        firstBox := InputBox("Enter your first name:", "Setup")
        if (firstBox.Result != "OK")
            ExitApp()

        lastBox := InputBox("Enter your last name:", "Setup")
        if (lastBox.Result != "OK")
            ExitApp()

        IniWrite(firstBox.Value, CONFIG["file"], "User", "FirstName")
        IniWrite(lastBox.Value,  CONFIG["file"], "User", "LastName")
    }

    ; Load values into memory for fast access
    CONFIG["FirstName"] := IniRead(CONFIG["file"], "User", "FirstName", "Unknown")
    CONFIG["LastName"]  := IniRead(CONFIG["file"], "User", "LastName", "Unknown")

    ; Common network paths (centralized for future use)
    CONFIG["OpenIPSGate"]         := "O:\MAN_Engineering\Open IPS Gate\"
    CONFIG["OpenIPSGateDomestic"] := "O:\MAN_Engineering\Open_IPS_Gate_Domestic_Only\"
    CONFIG["ClosedIPSGate"]       := "O:\MAN_Engineering\Closed_IPS_Gate\"
	CONFIG["SharedTemplates"]     := "O:\MAN_Engineering\Shared Templates\"
    CONFIG["DataCheck"]           := CONFIG["SharedTemplates"] "IPS Data Check\"

}

; ------------------------
; UTILITY FUNCTIONS
; ------------------------

; Trim leading/trailing whitespace
Trim(str) {
    return RegExReplace(str, "^\s+|\s+$")
}

; Remove common titles from a line
StripTitle(line) {
    titles := ["Sales Associate", "Manager", "Director", "VP", "Vice President", "Executive", "Coordinator", "Assistant", "Engineer", "Specialist", "Senior Executive Sales Consultant", "Sales Consultant", "Mr."]
    for _, title in titles
        line := RegExReplace(line, "^\s*" . title . "\s+", "", "i")
    for _, title in titles
        line := RegExReplace(line, "\s+" . title . "\s*$", "", "i")
    line := RegExReplace(line, "[;,\s]+$", "")
    line := RegExReplace(line, "^\s+|\s+$")
    return line
}

; Helper to jump to cell in Excel template
FillCell(cellRef) {
    Send("{F5}")
    Sleep(50)
    Send(cellRef)
    Send("{Enter}")
}

; Centralized function to get first + last name
GetFullName() {
    global CONFIG
    return CONFIG["FirstName"] . " " . CONFIG["LastName"]
}

GetSelectedExplorerItems() {
    Items := []

    for Window in ComObject("Shell.Application").Windows {
        try {
            if WinActive("ahk_id " Window.HWND) {
                for Item in Window.Document.SelectedItems
                    Items.Push(Item.Path)
                break
            }
        }
    }

    return Items
}

ClearFolderContents(FolderPath) {

    ; Delete files
    Loop Files FolderPath "\*", "F"
        FileDelete(A_LoopFileFullPath)

    ; Delete folders recursively
    Loop Files FolderPath "\*", "D"
        DirDelete(A_LoopFileFullPath, true)
}

; ------------------------
; HOTKEYS
; ------------------------

; ------------------------
; !D - IO Model Paste
; ------------------------
!d:: {
    try xl := ComObjActive("Excel.Application")
    catch {
        TrayTip "Excel is not running."
        return
    }

    dateStr := FormatTime(, "M/d/yyyy")
    firstName := CONFIG["FirstName"]

    cell := xl.ActiveCell

    ; Row 1
    cell.Offset(0, 0).Value := dateStr              ; Current column
    cell.Offset(0, 2).Value := "Maxilla"
    cell.Offset(0, 3).Value := "Intraoral Scan"
    cell.Offset(0, 4).Value := "N/A"
    cell.Offset(0, 5).Value := "Y"
    cell.Offset(0, 6).Value := firstName

    ; Row 2
    cell.Offset(1, 0).Value := dateStr
    cell.Offset(1, 2).Value := "Mandible"
    cell.Offset(1, 3).Value := "Intraoral Scan"
    cell.Offset(1, 4).Value := "N/A"
    cell.Offset(1, 5).Value := "Y"
    cell.Offset(1, 6).Value := firstName
}

; ------------------------
; !F - Open Case Folder
; ------------------------
!f:: {
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    Sleep(100)

    Send("!t")  ; PowerToys Text Extractor

    if !ClipWait(5) {
        TrayTip("Failed to get text from PowerToys.")
        A_Clipboard := ClipSaved
        return
    }

    userInput := A_Clipboard
    userInput := StrReplace(userInput, " ")
    userInput := StrReplace(userInput, "`r")
    userInput := StrReplace(userInput, "`n")
    cleanNumber := StrReplace(userInput, "c")
    cleanNumber := StrReplace(cleanNumber, "C")

    ; Call function
    OpenIPSGate(cleanNumber, CONFIG, ClipSaved)
}


OpenIPSGate(cleanNumber, CONFIG, ClipSaved) {
    locations := [
        CONFIG["OpenIPSGate"],
        CONFIG["OpenIPSGateDomestic"],
        CONFIG["ClosedIPSGate"]
    ]

    opened := false

    for base in locations {
        if (base = "")
            continue

        link := base . cleanNumber

        try {
            Run(link)
            opened := true
            break
        }
    }

    if (!opened) {
        TrayTip("Unable to open link in any configured location.")
    }

    ; Restore clipboard
    A_Clipboard := ClipSaved
}
; ------------------------
; !Q - Pin Folder to Quick Access
; ------------------------
!q:: { ; Alt+Q
	
	global A_ClipboardAll
	
    OldClip := ClipboardAll()
    A_Clipboard := ""

    Send "^c"
    if !ClipWait(1) {
        TrayTip "No folder selected."
        A_Clipboard := OldClip
        return
    }

    SelectedPath := A_Clipboard
    A_Clipboard := OldClip

    shell := ComObject("Shell.Application")
    folder := shell.Namespace(SelectedPath)

    if !folder {
        SplitPath SelectedPath, , &parentPath
        folder := shell.Namespace(parentPath)
    }

    if folder
        folder.Self.InvokeVerb("pintohome")
}
; ------------------------
; !R - Parse Names from Text
; ------------------------
!r:: {
    ; --- Save clipboard ---
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    Sleep 100

    ; --- Trigger PowerToys Text Extractor ---
    Send("!t")

    ; --- Wait up to 5 seconds for new clipboard content ---
    ClipWaitTimeout := 5000
    StartTime := A_TickCount
    Loop {
        Sleep 100
        if (A_Clipboard != "" && A_Clipboard != ClipSaved)
            break
        if ((A_TickCount - StartTime) > ClipWaitTimeout) {
            TrayTip("Failed to get text from PowerToys.")
            A_Clipboard := ClipSaved
            return
        }
    }

    ; --- Split clipboard into lines ---
    text := StrReplace(A_Clipboard, "`r", " ")
    lines := StrSplit(text, "`n")

    ; --- Count lines safely ---
    lineCount := 0
    for index, _ in lines
        lineCount++

    ; --- Detect name-like lines ---
    nameLikeLines := 0
    cleanedLines := []
    for index, line in lines {
        line := Trim(line)
        line := StripNameTitle(line)
        if RegExMatch(line, "^[A-Z][a-z]+(\s[A-Z][a-z]+)+$") {
            nameLikeLines++
            cleanedLines.Push(line)
        }
    }

    ; --- Build names output ---
    names := ""
    if (lineCount > 0 && nameLikeLines >= lineCount * 0.75) {
        names := StrJoin("; ", cleanedLines*)
    } else {
        selectedLines := []
        for index, line in lines {
            if Mod(index, 2) = 1 {
                line := Trim(line)
                line := StripNameTitle(line)
                if (line != "")
                    selectedLines.Push(line)
            }
        }
        names := StrJoin("; ", selectedLines*)
    }

    ; --- Restore clipboard ---
    A_Clipboard := names
    ClipSaved := ""
}

; --- Strip common titles ---
StripNameTitle(line) {
    titles := ["Sales Associate", "Manager", "Director", "VP", "Vice President", "Executive", "Coordinator", "Assistant", "Engineer", "Specialist", "Senior Sales Consultant", "Senior Executive Sales Consultant", "Sales Consultant", "Mr."]
    for _, title in titles {
        line := RegExReplace(line, "^(?i)\s*" . title . "\s+", "")
        line := RegExReplace(line, "\s+" . title . "\s*$(?i)", "")
    }
    return line
}

; --- Join array elements with a separator ---
StrJoin(sep, arr*) {
    result := ""
    for _, val in arr
        result .= (result = "" ? "" : sep) . val
    return result
}

; ------------------------
; !X - Create Case Folder Script
; ------------------------
!x:: {
    ClipContent := A_Clipboard
    if !RegExMatch(ClipContent, "^C(\d{7})", &Match)
        return

    CaseNumber := Match[1]

    ; ------------------------
    ; Determine Flags
    ; ------------------------
    isDomestic := InStr(ClipContent, "(D)")
    isPOC := InStr(ClipContent, "POC")

    if isDomestic
        FolderPath := CONFIG["OpenIPSGateDomestic"] . CaseNumber
    else
        FolderPath := CONFIG["OpenIPSGate"] . CaseNumber

    Run(CONFIG["SharedTemplates"] "Sample Folder Structure\Case Scripts\Create Case Folder Script_1.0.bat")
    WinWaitActive("ahk_class CASCADIA_HOSTING_WINDOW_CLASS")
    Send("^v")

    ; ------------------------
    ; Open IPS Data Log safely
    ; ------------------------
    dataDir := CONFIG["DataCheck"]
    fileFound := false

    Loop Files dataDir "\*.xltm", "F" {
        if RegExMatch(A_LoopFileName, "IPS Data Log .*\.xltm$") {
            Run(A_LoopFileFullPath)
            fileFound := true
            break
        }
    }

    if !fileFound {
        MsgBox("No IPS Data Log file found in " . dataDir)
        return
    }

    WinWaitActive("ahk_class XLMAIN")

    ; ------------------------
    ; Fill Excel
    ; ------------------------
    FillCell("B2")
    Send(Format("{}/{}/{}", A_MM, A_DD, A_YYYY))
    Send("{Enter}")

    FillCell("B11")
    Send("n{Enter}")

    FillCell("B15")
    Send("n{Enter}")

    FillCell("B16")
    Send("n{Enter}")

    ; ------------------------
    ; C5 depends on POC| Uncomment when full list of POC locations is implemented
    ; ------------------------
    ;FillCell("C5")
    ;Send((isPOC ? "y" : "n") "{Enter}")
    
	FillCell("C5")
    Send("n{Enter}")
	
    ; ------------------------
    ; C7 depends on Domestic
    ; ------------------------
    FillCell("C7")
    Send((isDomestic ? "y" : "n") "{Enter}")

    FillCell("B12")
    Send(GetFullName() "{Enter}")
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Generates fresh datalog
!s:: {

    ; ------------------------
    ; Open IPS Data Log safely
    ; ------------------------
    dataDir := CONFIG["DataCheck"]
    fileFound := false

    Loop Files dataDir "\*.xltm", "F" {
        if RegExMatch(A_LoopFileName, "IPS Data Log .*\.xltm$") {
            Run(A_LoopFileFullPath)
            fileFound := true
            break
        }
    }

    if !fileFound {
        MsgBox("No IPS Data Log file found in " . dataDir)
        return
    }

    WinWaitActive("ahk_class XLMAIN")

    ; ------------------------
    ; Fill Excel
    ; ------------------------
    FillCell("B2")
    Send(Format("{}/{}/{}", A_MM, A_DD, A_YYYY))
    Send("{Enter}")

    FillCell("B11")
    Send("n{Enter}")

    FillCell("B15")
    Send("n{Enter}")

    FillCell("B16")
    Send("n{Enter}")
    
	FillCell("C5")
    Send("n{Enter}")
	
    FillCell("C7")
    Send("n{Enter}")

    FillCell("B12")
    Send(GetFullName() "{Enter}")
}
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;; AutoChecker

Inputf  := A_Desktop "\AutoChecker\Input"
Outputf := A_Desktop "\AutoChecker\Output"

ExeFile := A_Desktop "\AutoChecker\Binaries\dicomcheck_v8.exe"

!a::{
    SelectedItems := GetSelectedExplorerItems()

    if (SelectedItems.Length = 0) {
        TrayTip "No files or folders selected in Explorer."
        return
    }

    ; Empty destination folders
    ClearFolderContents(Inputf)
    ClearFolderContents(Outputf)

    ; Move selected items into Folder1
    for Item in SelectedItems {

        SplitPath(Item, &Name)
        Destination := Inputf "\" Name

        try {
            if DirExist(Item)
                DirMove(Item, Destination)
            else if FileExist(Item)
                FileMove(Item, Destination)
        }
        catch Error as Err {
            TrayTip "Failed to move:`n" Item "`n`n" Err.Message
            return
        }
    }

AutoCheckerDir := A_Desktop "\AutoChecker"
ExeFile := AutoCheckerDir "\Binaries\dicomcheck_v8.exe"

cmd := A_ComSpec
    . ' /k cd /d "' AutoCheckerDir '" && "' ExeFile '"'
    . ' -i "' Inputf '"'
    . ' -o "' Outputf '"'
    . ' --location us'
    . ' --indication reconstruction'
    . ' --direct'

RunWait(cmd)
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Indication Fill
FillCaseType(caseType) {
    Send("{F5}")
    Sleep(50)
    Send("B4")       
    Send("{Enter}")
    Send(caseType)
    Send("{Enter}")
}

FillAnatomy(anatomyKey) {
    Send("{F5}")
    Sleep(50)
    Send("B5")       
    Send("{Enter}")
    Sleep(100)
    Send(anatomyKey)
    Send("{Enter}")
}


!Numpad1::{  
    FillCaseType("Orthog")
    FillAnatomy("m") 
}

!Numpad2::{  
    FillCaseType("Distract")
    FillAnatomy("m")
}

!Numpad3::{  
    FillCaseType("Thorax")
    FillAnatomy("r") 
}

!Numpad4::{  
    FillCaseType("Trauma")
    FillAnatomy("o") 
}

!Numpad5::{  
    FillCaseType("Rush Trauma")
    FillAnatomy("m")
}

!Numpad6::{  
    FillCaseType("Graft Recon")
    FillAnatomy("m")
}

!Numpad7::{  
    FillCaseType("Cr Vault")
    FillAnatomy("c") ; Cranium
}

!Numpad8::{  
    FillCaseType("Craniopl")
    FillAnatomy("c")
}

!Numpad9::{  
    FillCaseType("Transf")
    FillAnatomy("m")
}

!Numpad0::{  
    FillCaseType("IJR")
    FillAnatomy("m")
}

!NumpadDot::{  
    FillCaseType("Prepro")
    FillAnatomy("m")
}

