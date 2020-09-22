VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===========================================================
'
' This is the class that actually holds the data for
' each individual particle.
'
' You could easily add more functionality to this,
' eg. Wind, Z componants for 3D etc.
'
' Also, I only have VB 5, and therefore couldn't pass
' user defined types to class functions which is a real
' pain the arse.
'
'===========================================================
Option Explicit

Private cX As Single
Private cY As Single

Private cXSpeed As Single
Private cYSpeed As Single

Private cXAcceleration As Single
Private cYAcceleration As Single

Private cWidth As Integer
Private cHeight As Integer

Private cRed As Single
Private cGreen As Single
Private cBlue As Single
Private cAlpha As Single

Private cAlphaDecay As Single

Public Property Get x() As Single 'Syntax: Debug.Print X.ImageID
Attribute x.VB_Description = "A pointer to the surface used for this particle"
    x = cX
End Property
Public Property Get y() As Single 'Syntax: Debug.Print X.ImageID
    y = cY
End Property
Public Property Get Width() As Integer 'Syntax: Debug.Print X.ImageID
    Width = cWidth
End Property
Public Property Get Height() As Integer 'Syntax: Debug.Print X.ImageID
    Height = cHeight
End Property

Public Property Get Red() As Single 'Syntax: Debug.Print X.ImageID
    Red = cRed
End Property
Public Property Get Green() As Single 'Syntax: Debug.Print X.ImageID
    Green = cGreen
End Property
Public Property Get Blue() As Single 'Syntax: Debug.Print X.ImageID
    Blue = cBlue
End Property
Public Property Get Alpha() As Single 'Syntax: Debug.Print X.ImageID
    Alpha = cAlpha
End Property

Public Sub SetColour(pR As Single, pG As Single, pB As Single, pA As Single, pAlphaDecay)
    cRed = pR
    cGreen = pG
    cBlue = pB
    cAlpha = pA
    
    cAlphaDecay = pAlphaDecay
End Sub

Public Sub ResetParticle(x As Single, y As Single, XSpeed As Single, ySpeed As Single, XAcc As Single, YAcc As Single, pWidth As Integer, pHeight As Integer)
    cX = x: cY = y
    cXSpeed = XSpeed: cYSpeed = ySpeed
    cXAcceleration = XAcc: cYAcceleration = YAcc
    cWidth = pWidth: cHeight = pHeight
    
    cRed = 1
    cGreen = 1
    cBlue = 1
    cAlpha = 1
    
    cAlphaDecay = 0.01
End Sub

Public Sub UpdateParticle(pTime As Single)
    cX = cX + cXSpeed * pTime
    cY = cY + cYSpeed * pTime
    
    cXSpeed = cXSpeed + cXAcceleration * pTime
    cYSpeed = cYSpeed + cYAcceleration * pTime
    
    cAlpha = cAlpha - cAlphaDecay * Time
End Sub
