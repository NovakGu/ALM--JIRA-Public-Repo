'|**********************************************************************;
'* Project           : ALM->JIRA import vba automate
'*
'* Program name      : importV07.vbs
'*
'* Author            : Changheng Gu
'*
'* Editor            : Fabio Choi
'*
'* Date created      : 2015-11-01
'*
'* Purpose           : automate the manual process of import master file and working file
'*
'* Revision History  :  importV01
'*                      importV02
'*                      importV03
'*                      importV04
'*                      importV05
'*                      importV06
'*
'* Date        Author            Ref    Revision (Subtask format)
'* 2015-01-29  Changheng GU             Added extra Parent JIRA ID Field
'*
'|**********************************************************************;


Sub Main()
    preProcess
    vlookUp
    FilterTime
    postProcess
    modify
    useCaseLookUp
End Sub

'#########################
'function PreProcess change Clm Detected by to "saiqa.chaudry" and defect type Clm to "Bug"
'Hides Clm E&H&j&M&I&g
'#########################
Sub preProcess()

Set sh = ThisWorkbook.Sheets("Sheet1")

'init var
Dim name As String
name = "saiqa.chaudry"
Dim defect_type As String
defect_type = "Bug"

'count current #s of rows on the active worksheet
Dim k As Long
Set rn = sh.UsedRange
k = rn.Rows.Count + rn.row - 1

'Set row height of the worksheet as 12.5
Rows("1:" & k).RowHeight = 12.5
        
'Override all defect type as Bug
    Dim row As Integer
        For row = 2 To k
            Cells(row, 7).Value = defect_type
            Next row
            
'Override all deected by as saiqa.chaudry
    Dim row_1 As Integer
        For row_1 = 2 To k
            Cells(row_1, 8).Value = name
        Next row_1
'Hide CLM E H J M I G
    Worksheets("Sheet1").Columns("E").Hidden = True
    Worksheets("Sheet1").Columns("H").Hidden = True
    Worksheets("Sheet1").Columns("J").Hidden = True
    Worksheets("Sheet1").Columns("M").Hidden = True
    Worksheets("Sheet1").Columns("I").Hidden = True
    Worksheets("Sheet1").Columns("G").Hidden = True
End Sub

'#########################
'VLookup perform vLookUp Functions on the current master file and the earlier master file to match up all
'avaliable ALM_ID in both workSheets
'#########################
Sub vlookUp()

Set sh = ThisWorkbook.Sheets("Sheet1")

'Count current #s of rows on the active worksheet
Dim k As Long
Set rn = sh.UsedRange
k = rn.Rows.Count + rn.row - 1

'set up copy_range for table1 & table2 in VLOOKUP function
Dim ALM_Row As Long
Dim ALM_Clm As Long
Dim copyRange_1 As String
Dim copyRange_2 As String
startRow = 2
endRow = k
Let copyRange_1 = "A" & startRow & ":" & "A" & endRow
Let copyRange_2 = "A" & startRow & ":" & "B" & endRow

'set table1&table2 and start row# & clm#
Table1 = Sheet1.Range(copyRange_1)
Table2 = Sheet2.Range(copyRange_2)
ALM_Row = Sheet1.Range("B2").row
ALM_Clm = Sheet1.Range("B2").Column

'loop through all rows in table 1
'if in previous master file, Sheet2.cl == "#N/A", then Sheet1.cl = "#N/A"
'if Sheet.cl.value==null,Sheet1.cl = "#N/A"
'else perform vlookup function
For Each cl In Table1
    If (Sheet2.Cells(ALM_Row, ALM_Clm).Text = "#N/A") Then
        Sheet1.Cells(ALM_Row, ALM_Clm) = "#N/A"
    ElseIf (Sheet2.Cells(ALM_Row, ALM_Clm).Text = "") Then
        Sheet1.Cells(ALM_Row, ALM_Clm) = "#N/A"
    Else
        Sheet1.Cells(ALM_Row, ALM_Clm).Value = Application.WorksheetFunction.vlookUp(cl, Table2, 2, False)
    End If
    
    ALM_Row = ALM_Row + 1
Next cl

End Sub

'#########################
'useCaseLookup perform vlookup functions on the current master file on UseCase Clm and mathch with the
'UseCase List WorkSheet(WorkSheet 3)
'#########################
Sub useCaseLookUp()

Set sh = ThisWorkbook.Sheets("Sheet1")

'Count current #s of rows on the active worksheet
Dim k As Long
Set rn = sh.UsedRange
k = rn.Rows.Count + rn.row - 1

'set up copy_range for table1 & table2 in VLOOKUP function
Dim workingRow As Long
Dim workingClm As Long
Dim copyRange_1 As String
Dim copyRange_2 As String
startRow = 2
endRow = k
Let copyRange_1 = "N" & startRow & ":" & "N" & endRow
Let copyRange_2 = "A" & startRow & ":" & "B74"

'set table1&table2 and start row# & clm#
Table1 = sh.Range(copyRange_1)
Table2 = Sheet3.Range(copyRange_2)
workingRow = sh.Range("R2").row
workingClm = sh.Range("R2").Column
For Each cl In Table1
   On Error Resume Next
        sh.Cells(workingRow, workingClm) = Application.WorksheetFunction.vlookUp(cl, Table2, 2, False)

    workingRow = workingRow + 1
Next cl

End Sub

'#########################
'FilterTime function takes the lastest modified time from Sheet 4, user need to manually find the time from the
'last master file using filter and copy the time to Sheet4 A1
'the function will filter and delete any defect with last modified time earlier than the given time
'#########################
Sub FilterTime()

Set sh = ThisWorkbook.Sheets("Sheet1")

'Count current #s of rows on the active worksheet
Dim k As Long
Set rn = sh.UsedRange
k = rn.Rows.Count + rn.row - 1

Dim dbDate As Double

'check if A1 in Sheet1 is a properly formatted time, set filter if the time is valid
If IsDate(Sheet4.Range("A1")) Then
    dbDate = Sheet4.Range("A1")
    Set MyRange = Range("K1:K" & k)
    MyRange.AutoFilter Field:=1, Criteria1:=">" & dbDate
End If

'delete all filtered rows
Dim i As Integer
For j = 1 To 10
For i = 2 To k
If Rows(i).Hidden = True Then
Rows(i).EntireRow.Delete
End If
Next i
Next j

End Sub

'#########################
'PostProcess add Clm "project key", "resolution", "relates", "epic", "project type"
'filter out defect with status as "Fixed", "Closed", "Dev Assigned", and "Dev Rework" and delete defects with other status on the WorkSheet
'#########################
Sub postProcess()

Set sh = ThisWorkbook.Sheets("Sheet1")

'Count current #s of rows on the active worksheet
Dim i As Long
Set rn = sh.UsedRange
i = rn.Rows.Count + rn.row - 1

Worksheets("Sheet1").Range("Q1").Value = "project key"
Worksheets("Sheet1").Range("R1").Value = "resolution"
Worksheets("Sheet1").Range("S1").Value = "relates"
Worksheets("Sheet1").Range("T1").Value = "epic"
Worksheets("Sheet1").Range("U1").Value = "Project type"

Dim Status_Row As Long
Dim Status_Clm As Long
Dim copyRange_1 As String
startRow = 2
endRow = i
Let copyRange_1 = "C" & startRow & ":" & "C" & endRow

Dim start As Integer
For start = 1 To 10
Table1 = Sheet1.Range(copyRange_1)
Status_Row = Sheet1.Range("C2").row
Status_Clm = Sheet1.Range("C2").Column
'filter out status other than closed or fixed or dev assigned ot dev rework
For Each cl In Table1
    If ((Not Sheet1.Cells(Status_Row, Status_Clm) = "Closed") And (Not Sheet1.Cells(Status_Row, Status_Clm) = "Fixed") And (Not Sheet1.Cells(Status_Row, Status_Clm) = "Dev Assigned") And (Not Sheet1.Cells(Status_Row, Status_Clm) = "Dev Re-work")) Then
         Rows(Status_Row).EntireRow.Delete
    End If
    Status_Row = Status_Row + 1
Next cl
Next start

End Sub

'#########################
'modifySeverity changes severity to according level in JIRA formatted
'#########################
Sub modifySeverity(iParam As Integer)
    
Dim cellString As String
Let cellNum = "L" & iParam
cellString = Worksheets("Sheet1").Range(cellNum).Value

If (cellString = "1-Showstopper") Then
     Worksheets("Sheet1").Range(cellNum).Value = "blocker"
ElseIf (cellString = "2-Critical") Then
     Worksheets("Sheet1").Range(cellNum).Value = "critical"
ElseIf (cellString = "3-Medium") Then
     Worksheets("Sheet1").Range(cellNum).Value = "major"
ElseIf (cellString = "4-Low") Then
     Worksheets("Sheet1").Range(cellNum).Value = "minor"
ElseIf (cellString = "5-Enhancement") Then
     Worksheets("Sheet1").Range(cellNum).Value = "trivial"
Else
End If

End Sub

'#########################
'modifySeverity changes severity to according level in JIRA formatted
'#########################
Sub modifyWorkStreamAndProjectKey(iParam As Integer)

Dim cellString As String
Let cellNum = "O" & iParam
cellString = Worksheets("Sheet1").Range(cellNum).Value

If (cellString = "Optimization") Then
    Worksheets("Sheet1").Range(cellNum).Value = "Core Optimization"
    Worksheets("Sheet1").Range("Q" & iParam).Value = "OPT"
    Worksheets("Sheet1").Range("U" & iParam).Value = "software"
ElseIf (cellString = "Code Merge 4.6") Then
    Worksheets("Sheet1").Range("Q" & iParam).Value = "MRG"
    Worksheets("Sheet1").Range("U" & iParam).Value = "software"
    
Else
    Worksheets("Sheet1").Range(cellNum).Value = "Wealth360 5.1"
    Worksheets("Sheet1").Range("Q" & iParam).Value = "RP"
    Worksheets("Sheet1").Range("U" & iParam).Value = "software"
End If

End Sub


'#########################
'Check to make sure that Parent JIRA ID is empty for everything except code merge
'#########################


Sub filterParent(iParam As Integer)

Dim cellString As String
Let cellNum = "O" & iParam
cellString = Worksheets("Sheet1").Range(cellNum).Value


If (Not cellString = "Code Merge 4.6") Then
     Worksheets("Sheet1").Range("P" & iParam).Clear
    
    
Else
  
End If

End Sub

'#########################
'modifyStatus changes defects according to the defects' status
'For closed tickets, the function will clear the comments cell, and assign the defects to saiqa.chaudry
'if the closed defect does not have a JIRA ID, the entire row will be deleted
'if the defect belongs to project AE50 or AE51, mark defect as Done

'For Fixed tickets, the function will clear the comments for imported tickets which have a JIRA ID
'for tickets have not been imported, function will leave the comment abd clear the #N/A in the JIRA ID cell
'all tickets will be assigned to chetan.puwar

'For Dev Assigned tickets, if the defect belong to the AE project, change the project name to Wealth360 5.1, if the defect belongs to optimiziton, change to Core Optimization
'delete the comment for all dev assigned tickets
'if the dec assigned tickets was not imported before the function will clear the #N/A in the JIRA ID cell
'change the status of all defects with project AE to In Progress
 
'For Dev Rework Tickets, change the tickets' status to Open if the tickets belongs to AE, change status to Tech Analysis if belongs to Optimizaion

'For AE defects without a JIRA ID, add epic link to according Clm as 5.1 Defects
'For AE defects' JIRA ID starts with AE, using wildcard to identify and add epic link to according Clm as 5.0 Defcts Carry over
'#########################
Sub modifyStatus(iParam As Integer)

Dim cellString As Variant
Let cellNum = "B" & iParam
cellString = Worksheets("Sheet1").Range(cellNum).Text
Dim status As String
status = Worksheets("Sheet1").Range("C" & iParam).Text
Dim projectName As String
projectName = Worksheets("Sheet1").Range("O" & iParam).Text
Dim jiraId As String
jiraId = Worksheets("Sheet1").Range("B" & iParam).Text


If (status = "Closed") Then
    If (cellString = "#N/A") Then
        Worksheets("Sheet1").Range("A" & iParam).Clear
    Else
        If (projectName = "Advisor Essential 5.1" Or projectName = "Advisor Essential 5.0") Then
            Worksheets("Sheet1").Range("R" & iParam).Value = "Done"
        End If
        Worksheets("Sheet1").Range("F" & iParam).Clear
        Worksheets("Sheet1").Range("D" & iParam).Value = "saiqa.chaudry"
    End If
ElseIf (status = "Fixed") Then
    If (cellString = "#N/A") Then
        Worksheets("Sheet1").Range("B" & iParam).Clear
    Else
        Worksheets("Sheet1").Range("F" & iParam).Clear
    End If

    Worksheets("Sheet1").Range("D" & iParam).Value = "chetan.puwar"
    Worksheets("Sheet1").Range("C" & iParam).Value = "Resolved"
ElseIf (status = "Dev Assigned") Then
    Dim workStream As String
    workStream = Worksheets("Sheet1").Range("O" & iParam).Value
        If (workStream = "Optimization") Then
            Rows(iParam).EntireRow.Delete
        ElseIf (workStream = "Code Merge 4.6") Then
            If (cellString = "#N/A") Then
                Worksheets("Sheet1").Range("B" & iParam).Clear
                Worksheets("Sheet1").Range("C" & iParam).Value = "Implementation"
            Else
                Worksheets("Sheet1").Range("F" & iParam).Clear
                Worksheets("Sheet1").Range("C" & iParam).Value = "Implementation"
            End If
        Else
            If (cellString = "#N/A") Then
                Worksheets("Sheet1").Range("B" & iParam).Clear
                Worksheets("Sheet1").Range("C" & iParam).Value = "In Progress"
            Else
                Worksheets("Sheet1").Range("F" & iParam).Clear
                Worksheets("Sheet1").Range("C" & iParam).Value = "In Progress"
            End If
        End If
ElseIf (status = "Dev Rework") Then
    Dim workStream_0 As String
    workStream_0 = Worksheets("Sheet1").Range("O" & iParam).Value

        If (workStream_0 = "Optmization") Then
            Worksheets("Sheet1").Range("C" & iParam).Value = "Tech Analysis"
        Else
            Worksheets("Sheet1").Range("C" & iParam).Value = "Open"
        End If

    Worksheets("Sheet1").Range("D" & iParam).Value = "unassigned"
Else
End If

If (jiraId = "") Then
    Worksheets("Sheet1").Range("T" & iParam).Value = "5.1 Defects"
ElseIf (jiraId Like "AE-*") Then
    Worksheets("Sheet1").Range("T" & iParam).Value = "5.0 Defects Carry over"
End If


End Sub

'#########################
'modifyRow wrap all functions that modify the tickets
'#########################
Sub modifyRow(a As Integer)
    filterParent (a)
    modifyStatus (a)
    modifySeverity (a)
    modifyWorkStreamAndProjectKey (a)
    cleanUp (a)
End Sub

'#########################
'modify loops through all tickets and calls modifyRow function to change the defects accordingly
'#########################
Sub modify()

Set sh = ThisWorkbook.Sheets("Sheet1")

Dim i As Long

Set rn = sh.UsedRange
i = rn.Rows.Count + rn.row - 1

Dim l As Integer
For l = 1 To 3
    Dim flag As Integer
    For flag = 2 To i
    modifyRow (flag)
    Next flag
Next l

sh.AutoFilterMode = False

End Sub

'#########################
'cleanup function handles unexpected tickets after the modify process
'#########################
Sub cleanUp(i As Integer)
    Dim alm_id As Variant
    alm_id = Worksheets("sheet1").Range("A" & i).Value
    
    If (alm_id = "") Then
        Rows(i).EntireRow.Delete
    End If
    
End Sub






