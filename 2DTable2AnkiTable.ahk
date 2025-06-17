#Requires AutoHotkey v2.0
#SingleInstance Force

; ====================================== VARIABLES ======================================
; Set these to the heading of your last column/row so the counter knows when to stop
;Case sensitive
columnTerminator := "W"
rowTerminator := "W"

; Sets the header for the .txt file
matrixStr := "#separator:tab`n#html:true"

; delay between keyboard actions in milliseconds—increase if your spreadsheet is lagging
timeDelay := 100 

; Tracks size of table
numOfColumns := 0
numOfRows := 0

; Holds the table representation
matrix := []

; Makes a note of whatever your clipboard has, so it can be restored at the end (won't work with images or keep formatting)
clip := A_Clipboard

; ====================================== HOTKEYS ======================================
1::generateFile()

Esc::ExitApp
; ====================================== FUNCTIONS ======================================

; Runs all subfunctions and creates output file
generateFile(){
    global
    scan()
    matrixBuilder()

    for(row in matrix){
        i := A_Index
        ; Skip header row
        if(i == 1){
            continue
        }

        for(cell in row){
            j := A_Index 
            ; Skip header column
            if(j == 1){
                continue
            } 
            
            ; Gets rid of invisible characters so an empty cell reads as an empty string
            cell := Trim(cell, "`n `t `r")
            if(!!cell){
                matrixStr .= "`n" . matrix[1][j] . "-" . matrix[i][1] . A_Tab . cell
            }
        }

    }
    FileAppend(matrixStr, "table2Anki" "_" FormatTime(A_Now, "ddMMyy") "_" FormatTime(A_Now, "HHmmss") ".txt")
    A_Clipboard := clip
    ExitApp
}


matrixBuilder(){
    global
    row := []

    ; Go from the top down until we have enough rows
    while(matrix.length < numOfRows){
        SendInput("^c")
        Sleep(timeDelay/2)
        row.Push(A_Clipboard)
        ; MsgBox(A_Clipboard)

        ; If at the end of the row go back to the start, move down, and start building another row
        if(row.Length == numOfColumns){
            matrix.Push(row)
            row := []
            reverseColumn()
            SendInput("{Down}")
        } else {
            SendInput("{Right}")
        }         
        Sleep(100)
    }
}

; Finds the width and height of table
scan(){
    global
    getNumOfColumns()
    reverseColumn()
    getNumOfRows()
    reverseRow()
}

; Moves the cursor back a number of times equal to scanned table dimensions
reverseColumn(){
    global
    Loop(numOfColumns){
        SendInput("{Left}")
        Sleep(timeDelay)
    }
}

; Moves the cursor up a number of times equal to scanned table dimensions
reverseRow(){
    global
    Loop(numOfRows){
        SendInput("{Up}")
        Sleep(timeDelay)
    }
}

; These count the number of rows and columns so that the full grid can be covered 
; They know to stop based on what the column and row terminator string is

; Sets how many columns in the table
getNumOfColumns(){
    global
    SendInput("^c")
    while(A_Clipboard != columnTerminator){
        SendInput("{Right}")
        SendInput("^c")
        numOfColumns += 1
        ; Sheets seems to get a little fussy without a delay between inputs
        Sleep(timeDelay)
    }
    A_Clipboard := ""
}

; Sets how many rows in the table
getNumOfRows(){
    global
    SendInput("^c")
    while(A_Clipboard != rowTerminator){
        SendInput("{Down}")
        SendInput("^c")
        numOfRows += 1
        Sleep(timeDelay)
    }
    A_Clipboard := ""
}