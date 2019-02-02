Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub

Sub Button1_Click()
    
    Dim xSh As Worksheet
    
    Call OptimizeCode_Begin
    
    For Each xSh In Worksheets
        xSh.Select
        
            
            Call TickerVol
            Call YearChange
            Call Format
            Call Percent
            Call Max
            
    Next
    
    Call OptimizeCode_End

End Sub