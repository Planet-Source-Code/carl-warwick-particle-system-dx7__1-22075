Attribute VB_Name = "modMain"
'===========================================================
' Title:        Particle System
'
' Name:         Carl Warwick
'
' E-mail:       CC006740@ntu.ac.uk
'
' Company:      Black Edge Designs
'
' Website:      www.freeridedesign.cjb.net
'
' Date:         1st April 2001
'===========================================================

Option Explicit

Private Sub Main()
    frmMain.Show
    Randomize Timer
    
    'Initialize DirectX with screen 640x480x16
    Call Initialize_DX(frmMain, 640, 480, 16)
    
    'Create our two texture surfaces
    'Note...
    'Texture surface dimensions must be to a power or 2 (^2)
    'eg. possible sizes are :- 256x256, 128x128, 128,256, 16x16 etc...
    'Its not advisable to use surface bigger than 256x256 or at the most 512x512
    Set ddsTexture1 = CreateTextureSurface(App.Path & "\Particle.bmp", 32, 32, 0)
    
    'Create a normal DDraw surface
    Call DDCreateSurface(ddsurface, 640, 480, App.Path & "\BG.bmp", 0, False)
    
    'Sets the forecolor, this is only to make the text white when displaying
    'the Frames per second
    Call BackBuffer.SetForeColor(RGB(255, 255, 255))
    
    'Start the main loop
    bRunning = True
    Call Main_Loop
    
    'End the program and return to your normal life
    Endit
End Sub


'The main game loop is in here
Public Sub Main_Loop()
On Error Resume Next

    Dim mlngTimer As Long           '----}
    Dim mlngFrameTimer As Long      '    }  Used to calculate the Frames per second
    Dim mintFPSCounter As Integer   '----}
    Dim FPS As Integer              'Holds the number of frames per second
    Dim mlngElapsed As Long
    
    Dim tmpRect As RECT             'A RECT to be used for bltting
    tmpRect.Left = 0: tmpRect.Top = 0
    tmpRect.Right = 640: tmpRect.Bottom = 480
      
    g_ParticleCount = 0
    
    Set Fire = New clsFire
    Set Fire1 = New clsFire
    Set Smoke = New clsSmoke
    Set Snow = New clsSnow
    
    Call Fire1.SetPosition(530, 540, 340, 340)
    
    'Start our loop, press the 'Esc' key to exit
    Do Until bRunning = False
    
        Call Fire.Update(mlngElapsed)
        Call Fire1.Update(mlngElapsed)
        Call Smoke.Update(mlngElapsed)
        Call Snow.Update(mlngElapsed)
        
        'Clear the device after all the calls to Set_Vertex, and before
        'calling Device.BeginScene and any bltting.
        Call Clear_Device
        
        
        'Use Bltfast on our normal DDraw surface to draw the background
        Call BackBuffer.BltFast(0, 0, ddsurface, tmpRect, DDBLTFAST_WAIT)
        
        
        'Begin the scene, must do this after Calling Clear_Device, and before
        'calling Blt3D()
        Device.BeginScene
            
            Set_Alphablend True
            
                Call Fire.Render
                Call Fire1.Render
                Call Smoke.Render
                Call Snow.Render
            
            Set_Alphablend False
            
            
        'Call EndScene when finished using D3D routines
        Device.EndScene
            
        
        '===============================
        'Calculate the frame rate
        '(Thanks to lucky for this code)
        '===============================
        mlngElapsed = DX.TickCount() - mlngTimer
        mlngTimer = DX.TickCount()
        If DX.TickCount() - mlngFrameTimer >= 1000 Then
            mlngFrameTimer = DX.TickCount()
            FPS = mintFPSCounter
            mintFPSCounter = 0
        Else
            mintFPSCounter = mintFPSCounter + 1
        End If
        
        'Display the FPS
        Call BackBuffer.DrawText(0, 0, "FPS :" & Format(FPS), False)
        Call BackBuffer.DrawText(0, 12, "Time Elapsed :" & Format(mlngElapsed), False)
        Call BackBuffer.DrawText(0, 24, "Number of Particles :" & Format(g_ParticleCount), False)
        Call BackBuffer.DrawText(0, 40, "Press the Up Arrow to Increase the Number of Snow Particles.", False)
        Call BackBuffer.DrawText(0, 52, "Press the Down Arrow to Decrease them.", False)
        
        
        'Flip the primary surface
        Primary.Flip Nothing, DDFLIP_WAIT
        
        'Let windows do its stuff
        DoEvents
    Loop
End Sub


