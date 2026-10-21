Attribute VB_Name = "ExamTimer"
' =====================================================================
'  Exam Countdown Timer for PowerPoint  (free, no trial, works offline)
'
'  HOW TO USE
'    1. Go to the slide where you want the timer (a blank slide is best).
'    2. Press Alt+F8, choose InsertExamTimer, click Run.
'    3. Type the exam length in minutes (e.g. 90).
'    4. Start the slide show (F5 / Shift+F5) and click the Start button.
'
'  Buttons in slide show:  -5  -1  Start/Pause  Reset  +1  +5  Set time
'  Colours: normal -> amber at 10 min left -> red at 5 min left
'           -> flashing "TIME'S UP" at zero.
' =====================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' ---- settings you can change ----
Private Const DEFAULT_MINUTES As Double = 60
Private Const WARN_MINUTES As Double = 10      ' turns amber
Private Const DANGER_MINUTES As Double = 5     ' turns red
Private Const FONT_NAME As String = "Segoe UI" ' digits are fixed-width, no jiggle

Private Const CLR_TEXT As Long = 1315860       ' RGB(20,20,20)
Private Const CLR_OK As Long = 5287936         ' RGB(0,176,80)  green
Private Const CLR_WARN As Long = 41215         ' RGB(255,160,0) amber
Private Const CLR_DANGER As Long = 3289805     ' RGB(205,50,50) red
Private Const CLR_MUTED As Long = 7895160      ' RGB(120,120,120)
Private Const CLR_BTN As Long = 15790320       ' RGB(240,240,240)
Private Const CLR_TRACK As Long = 15132390     ' RGB(230,230,230)

' ---- runtime state ----
Private mSlideID As Long
Private mTotal As Double        ' full duration, seconds
Private mRemaining As Double    ' seconds left
Private mEndAt As Double        ' absolute end time (seconds) while running
Private mRunning As Boolean
Private mLoopActive As Boolean
Private mLastShown As Long

' =====================================================================
'  INSERT
' =====================================================================
Public Sub InsertExamTimer()
    Dim sld As Slide
    On Error Resume Next
    Set sld = ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then
        MsgBox "Click on a slide in Normal view first, then run InsertExamTimer again.", vbExclamation, "Exam Timer"
        Exit Sub
    End If

    Dim mins As Double
    mins = AskMinutes(DEFAULT_MINUTES)
    If mins <= 0 Then Exit Sub

    DeleteTimerShapes sld

    Dim W As Single, H As Single
    W = ActivePresentation.PageSetup.SlideWidth
    H = ActivePresentation.PageSetup.SlideHeight

    Dim s As Shape
    ' heading
    Set s = AddText(sld, "ET_Label", 0, H * 0.06, W, H * 0.08, "TIME REMAINING", H * 0.04, CLR_MUTED, True)

    ' big digits
    Dim fs As Single
    fs = H * 0.36
    If W * 0.19 < fs Then fs = W * 0.19
    Set s = AddText(sld, "ET_Time", 0, H * 0.15, W, H * 0.46, "00:00", fs, CLR_TEXT, True)

    ' status line ("Ends at 11:30 AM", "Paused", ...)
    Set s = AddText(sld, "ET_Status", 0, H * 0.61, W, H * 0.08, "", H * 0.045, CLR_MUTED, False)

    ' progress bar
    Set s = sld.Shapes.AddShape(msoShapeRoundedRectangle, W * 0.1, H * 0.72, W * 0.8, H * 0.025)
    s.Name = "ET_BarBG": s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = CLR_TRACK
    Set s = sld.Shapes.AddShape(msoShapeRoundedRectangle, W * 0.1, H * 0.72, W * 0.8, H * 0.025)
    s.Name = "ET_Bar": s.Line.Visible = msoFalse: s.Fill.ForeColor.RGB = CLR_OK

    ' buttons
    Dim gap As Single, small As Single, x As Single, bt As Single, bh As Single
    gap = W * 0.015: small = W * 0.075: bt = H * 0.83: bh = H * 0.1
    x = (W - (4 * small + W * 0.16 + W * 0.12 + W * 0.13 + 6 * gap)) / 2
    AddButton sld, "ET_BtnMinus5", x, bt, small, bh, ChrW(&H2212) & "5", "ET_Minus5": x = x + small + gap
    AddButton sld, "ET_BtnMinus1", x, bt, small, bh, ChrW(&H2212) & "1", "ET_Minus1": x = x + small + gap
    AddButton sld, "ET_BtnStart", x, bt, W * 0.16, bh, "Start", "ET_StartPause": x = x + W * 0.16 + gap
    AddButton sld, "ET_BtnReset", x, bt, W * 0.12, bh, "Reset", "ET_Reset": x = x + W * 0.12 + gap
    AddButton sld, "ET_BtnPlus1", x, bt, small, bh, "+1", "ET_Plus1": x = x + small + gap
    AddButton sld, "ET_BtnPlus5", x, bt, small, bh, "+5", "ET_Plus5": x = x + small + gap
    AddButton sld, "ET_BtnSet", x, bt, W * 0.13, bh, "Set time", "ET_SetTime"

    mRunning = False
    mSlideID = sld.SlideID
    mTotal = mins * 60
    mRemaining = mTotal
    sld.Tags.Add "ET_TOTAL", CStr(mTotal)
    mLastShown = -1
    Render
End Sub

Public Sub RemoveExamTimer()
    Dim sld As Slide
    On Error Resume Next
    Set sld = ActiveWindow.View.Slide
    On Error GoTo 0
    If sld Is Nothing Then Exit Sub
    mRunning = False
    DeleteTimerShapes sld
End Sub

' =====================================================================
'  BUTTON MACROS (called by clicking shapes during the slide show)
' =====================================================================
Public Sub ET_StartPause(shp As Shape)
    Bind shp
    If mRunning Then
        mRemaining = mEndAt - NowSec()
        If mRemaining < 0 Then mRemaining = 0
        mRunning = False
        mLastShown = -1
        Render
    Else
        If mRemaining <= 0 Then mRemaining = mTotal
        mEndAt = NowSec() + mRemaining
        mRunning = True
        mLastShown = -1
        RunLoop
    End If
End Sub

Public Sub ET_Reset(shp As Shape)
    Bind shp
    mRunning = False
    mRemaining = mTotal
    mLastShown = -1
    Render
End Sub

Public Sub ET_Plus1(shp As Shape)
    Bind shp
    AddTime 60
End Sub

Public Sub ET_Plus5(shp As Shape)
    Bind shp
    AddTime 300
End Sub

Public Sub ET_Minus1(shp As Shape)
    Bind shp
    AddTime -60
End Sub

Public Sub ET_Minus5(shp As Shape)
    Bind shp
    AddTime -300
End Sub

Public Sub ET_SetTime(shp As Shape)
    Bind shp
    Dim mins As Double
    mins = AskMinutes(mTotal / 60)
    If mins <= 0 Then Exit Sub
    mRunning = False
    mTotal = mins * 60
    mRemaining = mTotal
    TimerSlide.Tags.Add "ET_TOTAL", CStr(mTotal)
    mLastShown = -1
    Render
End Sub

' =====================================================================
'  INTERNALS
' =====================================================================
Private Sub Bind(shp As Shape)
    Dim sid As Long
    sid = shp.Parent.SlideID
    If sid <> mSlideID Or mTotal <= 0 Then
        mRunning = False
        mSlideID = sid
        mTotal = Val(shp.Parent.Tags("ET_TOTAL"))
        If mTotal <= 0 Then mTotal = DEFAULT_MINUTES * 60
        mRemaining = mTotal
        mLastShown = -1
    End If
End Sub

Private Sub AddTime(ByVal secs As Double)
    If mRunning Then
        mRemaining = mEndAt - NowSec()
    End If
    mRemaining = mRemaining + secs
    If mRemaining < 0 Then mRemaining = 0
    If mRemaining > mTotal Then mTotal = mRemaining   ' keep bar sensible
    If mRunning Then mEndAt = NowSec() + mRemaining
    mLastShown = -1
    Render
    If mRunning Then RunLoop
End Sub

Private Sub RunLoop()
    If mLoopActive Then Exit Sub
    mLoopActive = True
    On Error GoTo Finish
    Do While mRunning
        mRemaining = mEndAt - NowSec()
        If mRemaining <= 0 Then
            mRemaining = 0
            mRunning = False
            mLastShown = -1
            Render
            FlashTimesUp      ' if Start/+time is clicked meanwhile, the loop carries on
        Else
            Render
        End If
        DoEvents
        Sleep 50
    Loop
Finish:
    mLoopActive = False
End Sub

Private Sub FlashTimesUp()
    ' flash for ~20 seconds, stop early if teacher presses Start/Reset/+time
    Dim i As Long, t As Shape
    Set t = TimerSlide.Shapes("ET_Time")
    For i = 1 To 40
        If mRunning Or mRemaining > 0 Then Exit For
        If i Mod 2 = 0 Then
            t.TextFrame.TextRange.Font.Color.RGB = CLR_DANGER
        Else
            t.TextFrame.TextRange.Font.Color.RGB = CLR_TRACK
        End If
        Wait 0.5
    Next i
    If Not mRunning And mRemaining <= 0 Then t.TextFrame.TextRange.Font.Color.RGB = CLR_DANGER
End Sub

Private Sub Wait(ByVal secs As Double)
    Dim untilT As Double
    untilT = NowSec() + secs
    Do While NowSec() < untilT
        DoEvents
        Sleep 30
    Loop
End Sub

Private Sub Render()
    Dim sld As Slide
    Set sld = TimerSlide
    If sld Is Nothing Then Exit Sub

    Dim shown As Long
    shown = -Int(-mRemaining)          ' round up: 59.2s shows 1:00 -> 0:59
    If shown < 0 Then shown = 0

    Dim clr As Long, barClr As Long
    If shown = 0 Then
        clr = CLR_DANGER: barClr = CLR_DANGER
    ElseIf shown <= DANGER_MINUTES * 60 Then
        clr = CLR_DANGER: barClr = CLR_DANGER
    ElseIf shown <= WARN_MINUTES * 60 Then
        clr = CLR_WARN: barClr = CLR_WARN
    Else
        clr = CLR_TEXT: barClr = CLR_OK
    End If

    ' progress bar (cheap, update every tick for smoothness)
    Dim frac As Double
    If mTotal > 0 Then frac = mRemaining / mTotal Else frac = 0
    If frac > 1 Then frac = 1
    If frac < 0 Then frac = 0
    With sld.Shapes("ET_Bar")
        If frac <= 0 Then
            .Visible = msoFalse
        Else
            .Visible = msoTrue
            .Width = sld.Shapes("ET_BarBG").Width * frac
        End If
        .Fill.ForeColor.RGB = barClr
    End With

    If shown = mLastShown Then Exit Sub
    mLastShown = shown

    With sld.Shapes("ET_Time").TextFrame.TextRange
        .Text = FormatTime(shown)
        .Font.Color.RGB = clr
    End With

    Dim st As String
    If shown = 0 Then
        st = "TIME'S UP " & ChrW(&H2014) & " please stop writing"
    ElseIf mRunning Then
        st = "Ends at " & Format(DateAdd("s", shown, Now), "h:mm AM/PM")
    ElseIf mRemaining >= mTotal Then
        st = "Ready " & ChrW(&H2014) & " " & FormatLength(mTotal)
    Else
        st = "Paused"
    End If
    With sld.Shapes("ET_Status").TextFrame.TextRange
        .Text = st
        If shown = 0 Then .Font.Color.RGB = CLR_DANGER Else .Font.Color.RGB = CLR_MUTED
        .Font.Bold = IIf(shown = 0, msoTrue, msoFalse)
    End With

    If mRunning Then
        sld.Shapes("ET_BtnStart").TextFrame.TextRange.Text = "Pause"
    ElseIf mRemaining < mTotal And mRemaining > 0 Then
        sld.Shapes("ET_BtnStart").TextFrame.TextRange.Text = "Resume"
    Else
        sld.Shapes("ET_BtnStart").TextFrame.TextRange.Text = "Start"
    End If
End Sub

Private Function FormatTime(ByVal secs As Long) As String
    Dim h As Long, m As Long, s As Long
    h = secs \ 3600: m = (secs Mod 3600) \ 60: s = secs Mod 60
    If h > 0 Then
        FormatTime = h & ":" & Format(m, "00") & ":" & Format(s, "00")
    Else
        FormatTime = Format(m, "00") & ":" & Format(s, "00")
    End If
End Function

Private Function FormatLength(ByVal secs As Double) As String
    Dim m As Long
    m = CLng(secs / 60)
    If m >= 60 And m Mod 60 = 0 Then
        FormatLength = (m \ 60) & IIf(m = 60, " hour", " hours")
    ElseIf m >= 60 Then
        FormatLength = (m \ 60) & " h " & (m Mod 60) & " min"
    Else
        FormatLength = m & " min"
    End If
End Function

Private Function TimerSlide() As Slide
    On Error Resume Next
    Set TimerSlide = ActivePresentation.Slides.FindBySlideID(mSlideID)
End Function

Private Function NowSec() As Double
    ' wall-clock seconds; using the clock (not counting ticks) means no drift
    NowSec = CDbl(Date) * 86400# + Timer
End Function

Private Function AskMinutes(ByVal defMins As Double) As Double
    Dim r As String
    r = InputBox("Exam length in minutes (for example 90, or 1:30 for 1 hour 30 min):", "Exam Timer", CStr(defMins))
    r = Trim$(r)
    If r = "" Then AskMinutes = 0: Exit Function
    If InStr(r, ":") > 0 Then
        AskMinutes = Val(Split(r, ":")(0)) * 60 + Val(Split(r, ":")(1))
    Else
        AskMinutes = Val(r)
    End If
End Function

Private Function AddText(sld As Slide, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                         ByVal txt As String, ByVal sz As Single, ByVal clr As Long, ByVal isBold As Boolean) As Shape
    Dim s As Shape
    Set s = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h)
    s.Name = nm
    With s.TextFrame
        .AutoSize = ppAutoSizeNone
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        With .TextRange
            .Text = txt
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Name = FONT_NAME
            .Font.Size = sz
            .Font.Bold = IIf(isBold, msoTrue, msoFalse)
            .Font.Color.RGB = clr
        End With
    End With
    s.Top = t: s.Height = h
    Set AddText = s
End Function

Private Sub AddButton(sld As Slide, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                      ByVal caption As String, ByVal macroName As String)
    Dim s As Shape
    Set s = sld.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    s.Name = nm
    s.Line.Visible = msoFalse
    s.Fill.ForeColor.RGB = CLR_BTN
    s.Shadow.Visible = msoFalse
    With s.TextFrame
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = caption
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Name = FONT_NAME
            .Font.Size = h * 0.38
            .Font.Bold = msoTrue
            .Font.Color.RGB = CLR_TEXT
        End With
    End With
    With s.ActionSettings(ppMouseClick)
        .Action = ppActionRunMacro
        .Run = macroName
    End With
End Sub

Private Sub DeleteTimerShapes(sld As Slide)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If Left$(sld.Shapes(i).Name, 3) = "ET_" Then sld.Shapes(i).Delete
    Next i
End Sub
