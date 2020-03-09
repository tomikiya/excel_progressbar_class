VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pMax As Long
Private pValue As Long
Private pLength As Long
Private per As Long
Private commentString As String

Public Property Let ProgressMax(ByVal v As Long)
    pMax = v
End Property
Public Property Get ProgressMax() As Long
    ProgressMax = pMax
End Property

Public Property Let ProgressValue(ByVal v As Long)
    pValue = v
    Call DisplayProgressbar
End Property
Public Property Get ProgressValue() As Long
    ProgressValue = pValue
End Property

Public Property Let ProgressLength(ByVal v As Long)
    pLength = v
End Property

Public Property Let Comment(ByVal s As String)
    commentString = s
End Property

Public Sub Init()
    pMax = 100
    pValue = 0
    pLength = 20
    commentString = ""
    Application.StatusBar = False
End Sub

Private Sub DisplayProgressbar()
    Dim preCalcPer As Long
    preCalcPer = CalculatePercent()
    If preCalcPer <> per Or Application.StatusBar = False Then
        per = preCalcPer
        Application.StatusBar = GetProgressString()
        DoEvents
    End If
End Sub

Private Function CalculatePercent()
    If pValue >= pMax Then
        CalculatePercent = 100
    ElseIf pValue < 0 Then
        CalculatePercent = 0
    Else
        CalculatePercent = Int(pValue / pMax * 100) - Int(pValue / pMax * 100) Mod (100 / pLength)
    End If
End Function

Private Function GetProgressString() As String
    Dim progressCount As Long
    progressCount = Round(per / 100 * pLength, 0)
    GetProgressString = String(progressCount, "¡") & String(pLength - progressCount, " ") _
                     & "  " & per & "%:" & commentString
End Function


