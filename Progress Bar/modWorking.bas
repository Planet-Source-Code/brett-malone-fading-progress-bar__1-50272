Attribute VB_Name = "modWorking"
Option Explicit
'This sub used to call the working progress bar, set the caption, color, elements, speed and decayrate
Public Sub Working(ShowForm As Boolean, Optional strCaption As String, Optional ShowColor As OLE_COLOR, Optional NumElements As Integer, Optional BarSpeed As Integer, Optional DecayRate As Integer)

    If ShowColor <> 0 Then
        frmWorking.BrightColor = ShowColor
    Else
        frmWorking.BrightColor = &HFF8080
    End If
    
    If DecayRate <> 0 Then
        frmWorking.DecayRate = DecayRate
    Else
        frmWorking.DecayRate = 6
    End If
    
    If BarSpeed > 0 Then
        frmWorking.BarSpeed = BarSpeed
    Else
        frmWorking.BarSpeed = 30
    End If
    
    If NumElements <> 0 Then
        frmWorking.NumElements = NumElements
    Else
        frmWorking.NumElements = 60
    End If
    
    If strCaption <> "" And Not IsNull(strCaption) Then
        frmWorking.lblwork.Caption = Trim(strCaption)
    End If
    
    If ShowForm = True Then
        frmWorking.Show
        frmWorking.Refresh
    Else
        Unload frmWorking
    End If
    
End Sub
