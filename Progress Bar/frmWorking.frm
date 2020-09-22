VERSION 5.00
Begin VB.Form frmWorking 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2940
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox shpbar2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   7  'Diagonal Cross
      ForeColor       =   &H8000000B&
      Height          =   400
      Index           =   0
      Left            =   100
      ScaleHeight     =   405
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   350
      Width           =   45
   End
   Begin VB.Timer tmrWorking 
      Interval        =   50
      Left            =   90
      Top             =   900
   End
   Begin VB.Label lblwork 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Working, Please Wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   2670
   End
   Begin VB.Shape Shape 
      FillStyle       =   0  'Solid
      Height          =   500
      Left            =   50
      Top             =   300
      Width           =   2835
   End
End
Attribute VB_Name = "frmWorking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Public NumElements As Integer 'Number of elements to display
Public DecayRate As Long 'How fast to decay the primary color
Public BrightColor As OLE_COLOR 'Holds primary color
Public BarSpeed As Integer 'Lower = faster (more of a delay than a speed)

Private Sub Form_Activate()
Dim counter As Long

    BarDirection = Forward

'    create the elements
    For counter = 1 To NumElements - 1
        Load shpbar2(counter)
        shpbar2(counter).Left = (shpbar2(0).Width * counter) + shpbar2(0).Left
        shpbar2(counter).Visible = True
    Next counter
    
'    set the form up - Speed, color, form width
    tmrWorking.Interval = BarSpeed
    shpbar2(0).BackColor = BrightColor
    frmWorking.Width = shpbar2(0).Width * NumElements + shpbar2(0).Left + 200
    Shape.Width = shpbar2(0).Width * NumElements + shpbar2(0).Width + 60
    lblwork.Width = shpbar2(0).Width * NumElements + 100
End Sub

Private Sub tmrworking_Timer()
Dim counter As Integer

Dim intTempTail As Integer
   
On Error GoTo ErrorTrap

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
    MsgBox "User Control" & " WorkingBar: " & "tmrworking_Timer" & Err.Description
             
End Sub

Private Function DecrimentColor(SourceColor As Long, DecayValue As Long) As Long
'Calculates the next lowest color in all three RGB values
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
