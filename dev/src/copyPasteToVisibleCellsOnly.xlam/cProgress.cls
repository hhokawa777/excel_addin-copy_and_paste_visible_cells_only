VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'from http://hp.vector.co.jp/authors/VA033788/kowaza.html#0181

Private maxLen As Long      'string length of progress bar
Private initTime As Date
Private StartTime As Date   'execution start time
Private MinVal As Currency      'min value
Private MaxVal As Currency      'max value
Private myVal As Currency       'current value
Private myView As Long      'Display: 0. only bar, 1. +%, 2. + remained time, 3. +% & remained time
Private myDisplayStatusBar As Boolean   'initial state of status bar
Private myMsg As String     'message display

Private Const msgRemained As String = "Remained:"

Private Const msgElapsed As String = "Elapsed:"


Public Property Let Length(L As Long)
    maxLen = L
End Property
Public Sub Init()
    initTime = Now
End Sub

Public Sub Start()
    StartTime = Now
End Sub

Public Property Let Min(MinValue As Long)
    MinVal = MinValue
End Property
Public Property Get Min() As Long
    Min = MinVal
End Property

Public Property Let Max(MaxValue As Long)
    MaxVal = MaxValue
End Property
Public Property Get Max() As Long
    Max = MaxVal
End Property

Public Property Let View(V As Long)
    If V >= 0 And V <= 3 Then
        myView = V
    End If
End Property

Public Property Let Msg(Message As String)
    myMsg = Message
    If myMsg <> "" Then
        If (Left$(myMsg, 1) <> "Å@") And (Left$(myMsg, 1) <> " ") Then
            myMsg = "Å@" & myMsg
        End If
    End If
End Property

Private Sub Class_Initialize()
    maxLen = 20
    MinVal = 0
    MaxVal = 100
    myVal = 0
    myView = 1
    myDisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    myMsg = ""
    Start
End Sub

Private Sub Class_Terminate()
    Application.StatusBar = False
    Application.DisplayStatusBar = myDisplayStatusBar
End Sub

Public Property Let Value(V As Long)
    Dim P As Double
    Dim Zan As Date
    Dim Bar As String
    Dim Lvl As Long
    Dim Keika As Date
    
    If V < MinVal Then
        myVal = MinVal
    ElseIf V > MaxVal Then
        myVal = MaxVal
    Else
        myVal = V
    End If
    
    P = (myVal - MinVal) / (MaxVal - MinVal)
    If P > 0 Then
        Zan = (Now - StartTime) / P  'estimated total process time
        Zan = Zan * (1 - P)          'remained time
        Keika = Now - initTime
    End If
    
'    Bar = String$(maxLen, "Å†")   '** this character work in different lang environment ?
    Bar = String$(maxLen, "_")   '** this character work in different lang environment ?
    If P > 0 Then
        Lvl = maxLen * P
'        Mid$(Bar, 1, Lvl) = String$(Lvl, "Å°")   '** this character work in different lang environment ?
        Mid$(Bar, 1, Lvl) = String$(Lvl, "*")   '** this character work in different lang environment ?
        Select Case myView
            Case 1
                Bar = Bar & CInt(P * 100) & "%"
            Case 2
                Bar = Bar & msgRemained & Zan
            Case 3
                Bar = Bar & " " & CInt(P * 100) & "% ( " & myVal & " / " & MaxVal & " ), " & msgElapsed & " " & Keika & ", " & msgRemained & " " & Zan
        End Select
    End If
    
    Application.StatusBar = Bar & myMsg
End Property
Public Property Get Value() As Long
    Value = myVal
End Property
