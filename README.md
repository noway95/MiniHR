Sub CalnPaste()
    'To do copy and paste
    Sheets("Attendance").Select 'Select the correct sheet
    Range("C7").Select 'Select the cell that starts the day label
    Range(Selection, Selection.End(xlToRight)).Select 'Select and hold the range of cells that has day labels
    Selection.Copy 'Copy the range
    
    Sheets("Timetable").Select 'Select the right sheet to copy range
    Range("E4").Select 'Select the right cell to copy range
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True 'Paste Special to the cell
  
    'To clear previous data
    Range("H4:L34").ClearContents
    
    'To key in OT data
    Dim Day 'Stores the day of the week
    Dim PH 'Stores whether a particular day is a public holiday
    Dim OT_Hours As Double 'Stores the OT Hours on a weekday
    Dim Total_Hrs As Double 'Stores the total number of hours worked in a particular day
    Dim day_count As Integer 'Sets the loop from the row of day 1 to row of day 31
    Dim Rmk 'Stores if LL/OL/CSE/etc
    
    Range("E4").Select 'Select the right cell to start for keying and calculation of OT hours, note its in "Template" sheet
    
    For day_count = 1 To 31 'Loop for one month, maximum is 31 days
 
        Day = ActiveCell.Value 'Current cell gives the value of the day of the week, to be stored as day
        ActiveCell.Offset(0, 1).Select 'Move 1 cell to the right
        PH = ActiveCell.Value 'Current cell states whether its a public holiday, to be stored as PH
        ActiveCell.Offset(0, 1).Select 'Move 1 cell to the right
        Total_Hrs = CDec(ActiveCell.Value) 'Current cell gives the total number of hours worked that day, to be stored as Total
        'All cases that results in the different OT hours
        
        If Total_Hrs = 0 Then
        ElseIf Day = "SA" And PH = "Y" Then 'If Saturday and Public Holiday
            ActiveCell.Offset(0, 4).Select 'Move 4 cells to the right
            ActiveCell.Value = Total_Hrs 'Assign the active cell OT hours (X3)
        ElseIf Day = "SA" Then 'If Saturday and not Public Holiday
            ActiveCell.Offset(0, 2).Select 'Move 2 cells to the right
            ActiveCell.Value = Total_Hrs 'Assign the active cell OT hours (X1.5)
        ElseIf Day = "SU" And PH = "Y" Then 'If Sunday and Public Holiday
            ActiveCell.Offset(0, 5).Select 'Move 5 cells to the right
            ActiveCell.Value = Total_Hrs 'Assign the active cell OT hours (X4)
        ElseIf Day = "SU" Then 'If Sunday and not Public Holiday
            ActiveCell.Offset(0, 3).Select 'Move 3 cells to the right
            ActiveCell.Value = Total_Hrs 'Assign the active cell OT hours (X2)
        ElseIf PH = "Y" Then 'If Public Holiday on weekday
            If Total_Hrs > 8 Then 'If there is OT on Public Holiday
                OT_Hours = Total_Hrs - 8 'Calculate the effective OT hours
                ActiveCell.Offset(0, 1).Select 'Move 1 cell to the right
                ActiveCell.Value = 8  'Assign the active cell OT hours (X1)
                ActiveCell.Offset(0, 1).Select 'Move 1 cell to the right
                ActiveCell.Value = OT_Hours 'Assign the active cell OT hours (X1.5)
            Else
                ActiveCell.Offset(0, 1).Select 'If no OT on Public holiday, just move one cell right and key in
                ActiveCell.Value = Total_Hrs 'Assign active cell OT hours (X1)
            End If
        ElseIf PH = "HD" Then
            If Total_Hrs > 4 Then 'If there is OT on a half-day working day
                OT_Hours = Total_Hrs - 4 'Calculate the effective OT hours
                ActiveCell.Offset(0, 2).Select 'Move 1 cell to the right
                ActiveCell.Value = OT_Hours 'Assign the active cell OT hours (X1.5)
            Else
            End If
        ElseIf Total_Hrs > 8 Then 'If Weekday and OT
            OT_Hours = Total_Hrs - 8 'Calculate the effective OT hours
            ActiveCell.Offset(0, 2).Select 'Move 2 cells to the right
            ActiveCell.Value = OT_Hours 'Assign the active cell OT hours (X1.5)
        End If
        OT_Hours = 0 'Restart OT Hours
        Selection.End(xlToRight).Select 'Move selected cell all the way to the right
        ActiveCell.Offset(1, -8).Select 'Move selected cell to next row and column to start
    Next day_count
    
    'Append the monthly data from each staff to the Raw Data Sheet
    'Declare all required variables to store data
    Dim Name 'Name of personnel
    Dim Month 'Month of work
    Dim SCode 'Staff Code
    Dim NormHr As Double 'Normal Hours, applicable for part-time
    Dim OTHr As Double 'OT Hours, applicable for full-time
    Dim Other As Double
    Dim DOR As Date 'Date of report
    Dim DoneBy 'HR Personnel that prepared the spreadsheet
    Dim Leave As Double 'Days of leave
    Dim MC As Integer 'Days of MC
    Dim TrainingHr As Double 'TrainingHrs in that month
    Dim Days As Integer 'Days in that month
    
    'Store all data as variables
    Month = Range("C2").Value 'Store month
    SCode = Range("G2") 'Store staff code
    NormHr = CDec(Range("G36").Value) 'Store normal hours of work, for part-timers
    OTHr = CDec(Range("G37").Value) 'Store OT hours (effective, e.g. if X1.5, OTHr = 1.5 and not 1)
    Other = CDec(Range("G38").Value) 'Store other reburisements/deductions
    DOR = Range("G41").Value 'Store date of report
    DoneBy = Range("G42").Value 'Store personnel who key in data
    Leave = CDec(Application.WorksheetFunction.CountIf(Range("C4:C34"), "LL")) 'Store total local leave
    Leave = Leave + Application.WorksheetFunction.CountIf(Range("C4:C34"), "OL") 'Store total local leave + overseas value
    Leave = Leave - 0.5 * Range("G39").Value 'Account for half-day leave
    MC = Application.WorksheetFunction.CountIf(Range("C4:C34"), "MC") 'Store total MC days for that month
    TrainingHr = CDec(Range("G40").Value) 'Store total training hours for that month
    Days = 31 - WorksheetFunction.CountA("F4:F34") 'Store total no. of days in that month
    
    Sheets("StaffData").Select 'Move to sheet "StaffData"
    'Select Cell B2 to start
    
    'Declare all required variables to store data
    'Start with variables from spreadsheet
    Dim SCodeSD 'Staff Code from the StaffData Sheet
    Dim Nat 'Nationality
    Dim Status 'Whether personnel is full/part-time
    Dim Basic As Double 'Basic Pay (For full-time)
    Dim DayRate As Double 'Day Rate (For part-time)
    Dim HrRate As Double 'Hourly Rate (For both full and part-time)
    Dim Age As Double 'Age, will affect CPF rate
    Dim PRYears As Double 'No. of years of PR, will affect CPF rate
    Dim LeaveE As Double 'Leave Entitled
    
    'Other intermediate variables required
    Dim Total As Double 'Total amount from basic and over-time
    Dim LeaveL As Double 'Leave Left
    Dim MCT As Integer 'MC Total
    Dim TrainingT As Double 'Training Total
    Dim CPF_ee As Double 'CPF Employee
    Dim CPF_er As Double 'CPF Employer
    Dim CDAC As Double 'CDAC Contributions
   
    ActiveSheet.ListObjects("StaffInfo").Range.AutoFilter Field:=2, _
        Criteria1:=SCode 'Filter by staff code
    'Creating dictionary, looping through and store data is not a good solution because it takes up process time
    'expotentially as data increases and increases interim memory space
    'Better method is to do filter as above
    Range("StaffInfo[[#Headers],[Staff Code]]").Select  'Select Cell B2 (column of staff code to start)
    Selection.End(xlDown).Select 'Select the last row (actually there is only one row of available name after filter)
    Name = ActiveCell.Offset(0, -1).Value 'Store name
    ActiveCell.Offset(0, 2).Select 'Move three columns right
    Nat = ActiveCell.Value 'Stores whether or not he/she is Singaporean
    ActiveCell.Offset(0, 5).Select 'Move five columns right
    Status = ActiveCell.Value 'Stores whether he/she is full/part-time
    ActiveCell.Offset(0, 1).Select 'Move one column right
    Basic = CDec(ActiveCell.Value) 'Stores monthly pay of employee
    ActiveCell.Offset(0, 1).Select 'Move one column right
    DayRate = CDec(ActiveCell.Value) 'Stores day rate of employee
    ActiveCell.Offset(0, 1).Select 'Move one column right
    HrRate = CDec(ActiveCell.Value) 'Stores hour rate of employee
    ActiveCell.Offset(0, 1).Select 'Move one column right
    Age = CDec(ActiveCell.Value) 'Stores age of employee
    ActiveCell.Offset(0, 1).Select 'Move one column right
    
    If ActiveCell.Value = "NA" Then
        PRYears = NA
    Else
        PRYears = CDec(ActiveCell.Value) 'Stores no. of years of PR
    End If
    ActiveCell.Offset(0, 2).Select 'Move two colums right
    LeaveE = ActiveCell.Value 'Stores no. of days of leave entitled
    
    If Status = "F" Then
        Total = CDec(Basic + OTHr * HrRate)
    ElseIf Status = "P" Then
        Total = CDec(NormHr * HrRate + OTHr * HrRate)
    End If
    
    'Don't use nested If. Use dictionary keypad values for CPF or create another sheet
    '(like what I have done)
    'Consider customising cases if possible
    
    'Nationality keypad values
    If Nat = "PR" And PRYears > 3 Then
        Nat = "S"
    ElseIf Nat = "PR" And PRYears > 2 Then
        Nat = "PR2"
    ElseIf Nat = "PR" And PRYears > 1 Then
        Nat = "PR1"
    ElseIf Nat = "S" Then
        Nat = "S"
    Else
        Nat = "N"
    End If
    
    Sheets("CPF").Select 'Move sheet to CPF to key in data to calculate CPF
    
    CDAC = CDec(Range("L1").Value) 'Assign CDAC ammount from cell
    Range("C1") = Total 'Assign cell with total salary
    Range("F1") = Nat 'Assign cell with Nationality
    Range("I1") = Age 'Assign cell with age
    Range("H3").Select 'Select the start row of CPF employee
    Selection.End(xlDown).Select 'Shift cell all the way down
    CPF_ee = CDec(ActiveCell.Value) 'Assign CPF of employee from cell value
    ActiveCell.Offset(0, 1).Select 'Shift cell one column to right
    CPF_er = CDec(ActiveCell.Value) 'Assign CPF of employer from cell value
    
    Sheets("PaymentData").Select 'Move sheet to RawData to key in stuff
    Range("B2").Select 'Select the right cell to start
    
    Selection.End(xlDown).Select 'Move cursor to end row that is filled
    ActiveCell.Offset(1, 0).Select 'Move one row down
    ActiveCell.Value = Name 'Assign name to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = SCode 'Assign staff code to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = Month 'Assign month to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    If Basic = 0 Then
        Basic = Total - OTHr * HrRate
    Else
    End If
    ActiveCell.Value = Basic 'Assign Basic Pay amount to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = OTHr * HrRate 'Assign OT Pay amount to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = CPF_ee 'Assign CPF Employee amount to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = CPF_er 'Assign CPF Employer amount to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = Other
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = Leave 'Assign no. of days of leave on that month to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = MC 'Assign total no. of MC days
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = TrainingHr 'Assign total no. of training hours on that month to cell
    
    'Creating dictionary, looping through and store data is not a good solution because it takes up process time
    'expotentially as data increases and increases interim memory space
    'Copy data from Template and StaffData, verify and count if required
    ActiveSheet.ListObjects("MonthlyData").Range.AutoFilter Field:=2, Criteria1 _
        :=SCode 'Filter by Staff Code
    Range("J2").Select 'Column that stores leave that month
    LeaveL = CDec(LeaveE - WorksheetFunction.SumIf(Range("C:C"), SCode, Range("J:J")))
    'Assign leave left by deducting entitled - sums up all cells in that column
    Range("K2").Select 'Column that stores MC that month
    MCT = CDec(WorksheetFunction.SumIf(Range("C:C"), SCode, Range("K:K"))) 'Total no. of MCs
    Range("L2").Select 'Column that stores training hours that month
    TrainingT = CDec(WorksheetFunction.SumIf(Range("C:C"), SCode, Range("L:L"))) 'Total no. of training hours
    
    Selection.End(xlDown).Select 'Move cursor to end row that is filled
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = LeaveL 'Assign no. of days of leave left to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = MCT 'Assign cumulative no. of MC to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = TrainingT 'Assign total training hours to cell
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = DoneBy 'Assign HR Personnel that prepared the spreadsheet
    ActiveCell.Offset(0, 1).Select 'Move one column right
    ActiveCell.Value = CDec(Total - CPF_ee - CDAC + Other) 'Assign Nett Pay
    ActiveCell.Offset(0, -14).Select 'Move 14 columns to left and end
    
End Sub
