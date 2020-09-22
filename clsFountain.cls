VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===========================================================
'
' Due to VB's lack of inheritance the classes are alot
' bulkier and use repeating code than they should.
' This could probably be solved if I spent a bit more
' time on this, but I made this in an hour or so.
'
' The only main difference between the classes is
' in the resetParticle subroutine.
'
'===========================================================

Option Explicit

Private pVertex(3) As D3DTLVERTEX
Private Particle() As clsParticle

Private X1 As Single
Private X2 As Single
Private Y1 As Single
Private Y2 As Single

Private ParticleCount As Integer
Public Property Let ParticleCounts(pVal As Integer)
    Dim i As Integer
    
    If pVal > ParticleCount Then
        ReDim Preserve Particle(0 To pVal)
        For i = ParticleCount + 1 To pVal
            Set Particle(i) = New clsParticle
            Call ResetParticle(i)
        Next i
    Else
        If pVal < 0 Then pVal = 0
        ReDim Preserve Particle(0 To pVal)
    End If
    
    g_ParticleCount = g_ParticleCount + (pVal - ParticleCount)
    
    ParticleCount = pVal
End Property
Public Property Get ParticleCounts() As Integer
    ParticleCounts = ParticleCount
End Property

Public Sub SetPosition(pX1 As Single, pX2 As Single, pY1 As Single, pY2 As Single)
    Dim i As Integer
    
    X1 = pX1
    X2 = pX2
    Y1 = pY1
    Y2 = pY2
    
    For i = 0 To ParticleCount
        Call ResetParticle(i)
    Next i
End Sub

Private Sub Class_Initialize()
    Dim i As Integer
    
    X1 = 275
    X2 = 285
    Y1 = 385
    Y2 = 390
    
    ParticleCount = 39
    g_ParticleCount = g_ParticleCount + ParticleCount + 1
    ReDim Particle(0 To ParticleCount)
    
    For i = 0 To ParticleCount
        Set Particle(i) = New clsParticle
        Call ResetParticle(i)
    Next i
End Sub

Public Sub Update(mlngElapsed As Long)
    Dim i As Integer
    
    For i = 0 To ParticleCount
        Call Particle(i).UpdateParticle(mlngElapsed / 100)
        If Particle(i).y > 480 Or Particle(i).x < -Particle(i).Width Or Particle(i).x > 640 Or Particle(i).Alpha <= 0 Then
            Call ResetParticle(i)
        End If
    Next i
End Sub

Public Sub Render()
    Dim i As Integer
    Dim col As Long
    
    For i = 0 To ParticleCount
        col = DX.CreateColorRGBA(Particle(i).Red, Particle(i).Green, Particle(i).Blue, Particle(i).Alpha)
        Call Set_Vertex(pVertex(), Particle(i).x, Particle(i).y, Particle(i).x + Particle(i).Width, Particle(i).y, Particle(i).x, Particle(i).y + Particle(i).Height, Particle(i).x + Particle(i).Width, Particle(i).y + Particle(i).Height, col, col, col, col)
        Call Blt3D(ddsTexture1, pVertex())
    Next i
End Sub

Private Sub ResetParticle(i As Integer)
    Dim x As Single
    Dim y As Single
    
    x = (X2 - X1 + 1) * Rnd + X1
    y = (Y2 - Y1 + 1) * Rnd + Y1
    
    Call Particle(i).ResetParticle(x, y, Rnd - 0.5, (Rnd - 1.1) * 2.5, 0, 0.05, 16, 16)

    Call Particle(i).SetColour(1, 0.5, 0.1, (Rnd / 2) + 0.5, (Rnd * 0.4) + 0.02)
End Sub

' Remember to release all the particles
Private Sub Class_Terminate()
    Dim i As Integer
    
    For i = 0 To ParticleCount
        Set Particle(i) = Nothing
    Next i
End Sub
