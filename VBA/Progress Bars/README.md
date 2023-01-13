## Progress Bar via Userform


```vba

Sub ProgressBar_UserForm_Update(pctdone, ProgressDesc_Caption, NumericProgress_Caption, FullDesc_Caption, ProgressBar_Title)
    
    'update bar for visable progress
    ProgressBar.LabelProgress.Width = pctdone * (ProgressBar.FrameProgress.Width)
    
    'set text progress indication label. default is Progress:
    ProgressBar.ProgressDesc.Caption = ProgressDesc_Caption
    
    'update numeric progress to show 1/1
    ProgressBar.NumericProgress.Caption = NumericProgress_Caption
    
    'update desc label to give desc of step
    ProgressBar.FullDesc.Caption = FullDesc_Caption
    
    'update userform title
    ProgressBar.Caption = ProgressBar_Title
    
    'reapply changes
    ProgressBar.Repaint
    
End Sub

Sub progressbar_userform_demo()
    
    ProgressBar.Show vbModeless
        
    For i = 1 To 5
        Call ProgressBar_UserForm_Update((i / end_loop), "Progress:", i & "/" & end_loop, "Loop Progress: " & i, "Title")
        Application.Wait (Now + TimeValue("00:00:01"))
    Next
    
    ProgressBar.Hide
    
End Sub
```


## Progress Bar via Status Bar

### Class Module Named ProgressBar_StatusBar


```vba
Private statusBarState As Boolean
Private enableEventsState As Boolean
Private screenUpdatingState As Boolean

Private BAR_CHAR_SM As String
Private BAR_CHAR_MD As String
Private BAR_CHAR_LG As String
Private BAR_CHAR_FL As String
Private SPACE_CHAR As String

Private Sub Class_Initialize()

    'save the state of the variables to change
    statusBarState = Application.DisplayStatusBar
    enableEventsState = Application.EnableEvents
    screenUpdatingState = Application.ScreenUpdating
    
    'set the progress bar chars (should be equal size)
    BAR_CHAR_SM = ChrW(9602)
    BAR_CHAR_MD = ChrW(9603)
    BAR_CHAR_LG = ChrW(9605)
    BAR_CHAR_FL = ChrW(9607)
    SPACE_CHAR = ChrW(9620)
    
    'set the desired state
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False

End Sub

Private Sub Class_Terminate()

    'restore settings
    Application.DisplayStatusBar = statusBarState
    Application.ScreenUpdating = screenUpdatingState
    Application.EnableEvents = enableEventsState
    Application.StatusBar = False
    
End Sub

Public Sub Update(iter, total_num, Optional desc = "", Optional pct_value = True, Optional num_of_bars = 10)
    
    'set progress bar as empty
    progress_bar = ""
    
    'get % of iteration
    value_pct = iter / total_num
    
    'format as whole numbers
    value_base = value_pct * num_of_bars
    
    'fill progressbar with full bars for every base %
    For j = 1 To value_base
        progress_bar = progress_bar & BAR_CHAR_FL
    Next
    
    'find the fractional percentages for smaller bars
    lRemain = value_base - Fix(value_base)
    
    'when there are 10 bars use smaller bars to represent fractions of percentages. don't add space for 10/10 complete
    If num_of_bars = 10 Then
        If lRemain = 0 And (iter <> total_num) Then
            progress_bar = progress_bar & SPACE_CHAR
        ElseIf lRemain > 0 And lRemain <= 0.35 Then
            progress_bar = progress_bar & BAR_CHAR_SM
        ElseIf lRemain > 0.35 And lRemain < 0.65 Then
            progress_bar = progress_bar & BAR_CHAR_MD
        ElseIf lRemain >= 0.65 Then
            progress_bar = progress_bar & BAR_CHAR_LG
        End If
    End If
    
    'fill in rest of progress bar with space
    Do While Len(progress_bar) < num_of_bars
        progress_bar = progress_bar & SPACE_CHAR
    Loop
    
    'format percentage at the end (1.23%)
    If pct_value Then
        pct_text = "(" & Format(iter / total_num, "0.00%") & ")"
    Else
        pct_text = ""
    End If
    
    'set StatusBar text as the end results
    Application.StatusBar = "|" & progress_bar & "|" & " " & pct_text & " " & desc
    
End Sub
```

### Example code to test Progress Bar

```vba

Sub progressbar_statusbar_demo()
    
    Dim ProgressBar As New ProgressBar_StatusBar
    
    total_num = 10
    num_of_bars = 20
    
    For i = 1 To total_num
        Call ProgressBar.Update(i, total_num, "Loop: " & i, True, num_of_bars)
        Application.Wait (Now + TimeValue("00:00:01"))
    Next
    
End Sub
```
