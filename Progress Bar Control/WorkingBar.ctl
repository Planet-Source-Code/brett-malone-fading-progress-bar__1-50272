VERSION 5.00
Begin VB.UserControl WorkingBar 
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ForeColor       =   &H0000C000&
   MouseIcon       =   "WorkingBar.ctx":0000
   Picture         =   "WorkingBar.ctx":0442
   ScaleHeight     =   705
   ScaleWidth      =   1695
   Begin VB.Timer tmrWorking 
      Interval        =   40
      Left            =   720
      Top             =   0
   End
   Begin VB.PictureBox shpbar2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000006&
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   270
      Width           =   1635
   End
End
Attribute VB_Name = "WorkingBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------------------
' Module    : frmWorking
' DateTime  : 12/2/2003
' Author    : Brett Malone
' Purpose   : Displays fading progress bar
'---------------------------------------------------------------------------------------
Option Explicit
Private Enum theDirection 'Enum to keep things a little easier
    Forward = True
    reverse = False
End Enum
Dim BarDirection  As theDirection 'To keep track of direction of the bright spot
Dim intBrightnessSpot As Long 'The location of the bright spot
Dim theTailLen As Integer 'For keeping track of the current tail length
Private NumElements As Integer
Private DecayRate As Long
Private BrightColor As OLE_COLOR
Private Const pnBC  As String = "Color"
Private Const pnEN  As String = "Enabled"
Private Const pnSP  As String = "Speed"
Private Const pnDR  As String = "Decay"
Private Const pnNE  As String = "Elements"
Private Const erZR  As String = "Illegal Zero Value"
Private Sub InitBar()
Dim counter As Long

On Error GoTo ErrorTrap

    'set direction and speed
    BarDirection = Forward
    'create the elements
    For counter = 1 To NumElements - 1
        Load shpbar2(counter)
        shpbar2(counter).Left = shpbar2(0).Width * counter
        shpbar2(counter).Visible = True
    Next counter
    shpbar2(0).BackColor = BrightColor
    Exit Sub

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "InitBar" & Err.Description


End Sub

Private Sub tmrworking_Timer()
Dim counter As Integer

Dim intTempTail As Integer
   
'On Error GoTo ErrorTrap

    If BarDirection = Forward And intBrightnessSpot < NumElements - 1 Then
        'set the new brightspot
        intBrightnessSpot = intBrightnessSpot + 1
        shpbar2(intBrightnessSpot).BackColor = BrightColor
    End If

    If BarDirection = reverse And intBrightnessSpot > 0 Then
        'set the new brightspot
        intBrightnessSpot = intBrightnessSpot - 1
        shpbar2(intBrightnessSpot).BackColor = BrightColor
    End If
    
    Do While counter > -1 And counter < NumElements
    
        'incriment the length of the trail
        intTempTail = intTempTail + 1
        
        'temporary tail length for this loop only
        counter = counter + 1
        
        If BarDirection = Forward Then
            'special decay in front (overwrite fading tail)
            If intBrightnessSpot + counter < NumElements Then
                'decrease the brightness by subtracting decay from current brightness level
                shpbar2(intBrightnessSpot + counter).BackColor = DecrimentColor(shpbar2(intBrightnessSpot + counter).BackColor, DecayRate)
            End If
            
            'normal decay in back
            If intBrightnessSpot - counter > -1 Then
                'decrease brightness by the distance from the bright spot
                shpbar2(intBrightnessSpot - counter).BackColor = DecrimentColor(shpbar2(intBrightnessSpot - counter).BackColor, DecayRate)
            End If
        Else
            'normal decay in front
            If intBrightnessSpot + counter < NumElements Then
                'decrease brightness by the distance from the bright spot
                shpbar2(intBrightnessSpot + counter).BackColor = DecrimentColor(shpbar2(intBrightnessSpot + counter).BackColor, DecayRate)
            End If
            
            'special decay in back (overwrite fading tail)
            If intBrightnessSpot - counter > -1 Then
                'decrease the brightness by subtracting decay from current brightness level
                shpbar2(intBrightnessSpot - counter).BackColor = DecrimentColor(shpbar2(intBrightnessSpot - counter).BackColor, DecayRate)
            End If
        End If
    Loop
    
    'update the tail length only if it has actually grown
    If intTempTail > theTailLen Then theTailLen = intTempTail
        
    'if we are at the end of the row, we need to decay the tail
    If intBrightnessSpot = 0 Or intBrightnessSpot = NumElements - 1 Then
        counter = 0
        
        'work the entire length of the tail
        Do While counter < theTailLen And intBrightnessSpot - counter > 0
            counter = counter + 1
            If BarDirection = Forward Then
                'reset brightness at end of tail
                shpbar2(intBrightnessSpot - counter).BackColor = DecrimentColor(shpbar2(intBrightnessSpot - counter).BackColor, DecayRate)
            Else
                'reset brightness at end of tail
                shpbar2(intBrightnessSpot + counter).BackColor = DecrimentColor(shpbar2(intBrightnessSpot + counter).BackColor, DecayRate)
            End If
        Loop
        'we're done decaying the tail, so shorten it
        theTailLen = theTailLen - 1
        
        'reverse direction
        If intBrightnessSpot = 0 Then
            BarDirection = Forward
        Else
            BarDirection = reverse
        End If
    End If

    Exit Sub

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "tmrworking_Timer" & Err.Description

             
End Sub

Public Property Get Color() As OLE_COLOR

On Error GoTo ErrorTrap

    Color = BrightColor

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Color" & Err.Description


End Property
Public Property Let Color(ByVal nuColor As OLE_COLOR)

On Error GoTo ErrorTrap

    BrightColor = nuColor
    PropertyChanged pnBC

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Color" & Err.Description


End Property
Public Property Get Elements() As Integer

On Error GoTo ErrorTrap

    Elements = NumElements

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Elements" & Err.Description


End Property
Public Property Let Elements(ByVal nuElements As Integer)
Dim counter As Integer
On Error GoTo ErrorTrap

    For counter = 1 To NumElements - 1
        Unload shpbar2(counter)
    Next counter
    NumElements = Abs(CInt(nuElements))
    intBrightnessSpot = 0
    theTailLen = 0
    InitBar
    ResizeArray Width, Height
    PropertyChanged pnNE

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Elements" & Err.Description

End Property

Public Property Get Enabled() As Boolean

On Error GoTo ErrorTrap

    Enabled = tmrWorking.Enabled

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Enabled" & Err.Description

End Property


Public Property Let Enabled(ByVal nuEnabled As Boolean)


On Error GoTo ErrorTrap

    tmrWorking.Enabled = (nuEnabled <> False)
    If tmrWorking.Enabled = False Then
        Refresh
    End If
    PropertyChanged pnEN



    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Enabled" & Err.Description

End Property


Public Property Get Speed() As Long

On Error GoTo ErrorTrap

    Speed = 200 - tmrWorking.Interval

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Speed" & Err.Description


End Property

Public Property Let Speed(ByVal nuSpeed As Long)

On Error GoTo ErrorTrap

    tmrWorking.Interval = 200 - nuSpeed
    PropertyChanged pnSP

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Speed" & Err.Description

End Property

Public Property Get Decay() As Long

On Error GoTo ErrorTrap

    Decay = DecayRate

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Decay" & Err.Description


End Property

Public Property Let Decay(ByVal nuDecay As Long)

On Error GoTo ErrorTrap

    DecayRate = Abs(nuDecay)
    Dim counter As Integer
    For counter = 0 To NumElements
'        intBrightness(counter) = 0
    Next counter
    PropertyChanged pnDR

    Exit Property

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "Decay" & Err.Description

End Property

Private Sub UserControl_InitProperties()

On Error GoTo ErrorTrap

    DecayRate = 3
    NumElements = 40
    tmrWorking.Interval = 40
    BrightColor = vbRed
    InitBar

    Exit Sub

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "UserControl_InitProperties" & Err.Description

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error GoTo ErrorTrap

    With PropBag
        BrightColor = .ReadProperty(pnBC, vbRed)
        tmrWorking.Interval = 200 - .ReadProperty(pnSP, 160)
        DecayRate = .ReadProperty(pnDR, 3)
        NumElements = .ReadProperty(pnNE, 35)
        tmrWorking.Enabled = .ReadProperty(pnEN, True)
    End With

    Exit Sub

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "UserControl_ReadProperties" & Err.Description


End Sub

Private Sub UserControl_Resize()
On Error GoTo ErrorTrap

    ResizeArray Width, Height

    Exit Sub

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "UserControl_Resize" & Err.Description

End Sub
Private Sub ResizeArray(lngWidth As Long, lngHeight As Long)
Dim counter As Integer
Dim newWidth As Integer
On Error GoTo ErrorTrap
    newWidth = Width / NumElements
    For counter = 0 To NumElements - 1
        shpbar2(counter).Move newWidth * counter, 0, newWidth + 40, Height
    Next counter

    Exit Sub

ErrorTrap:
    InitBar
    Resume Next

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error GoTo ErrorTrap

    With PropBag
        .WriteProperty pnBC, BrightColor, 0
        .WriteProperty pnEN, tmrWorking.Enabled, True
        .WriteProperty pnSP, 200 - tmrWorking.Interval, 160
        .WriteProperty pnDR, DecayRate, 3
        .WriteProperty pnNE, NumElements, 35
    End With 'PROPBAG

    Exit Sub

ErrorTrap:
    MsgBox "User Control" & " WorkingBar" & "UserControl_WriteProperties" & Err.Description

End Sub
Private Function DecrimentColor(SourceColor As Long, DecayValue As Long) As Long
Dim NumRed As Long
Dim NumGreen As Long
Dim NumBlue As Long

    NumRed = SourceColor And &HFF
    NumGreen = (SourceColor \ &H100) And &HFF
    NumBlue = (SourceColor \ &H10000) And &HFF
    
    NumGreen = NumGreen - DecayValue
    NumRed = NumRed - DecayValue
    NumBlue = NumBlue - DecayValue
    
    If NumGreen < 0 Then NumGreen = 0
    If NumRed < 0 Then NumRed = 0
    If NumBlue < 0 Then NumBlue = 0

    DecrimentColor = RGB(NumRed, NumGreen, NumBlue)
            
End Function
