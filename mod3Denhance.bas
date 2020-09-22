Attribute VB_Name = "mod3Denhance"
'****************************************************************
'   This is an example program to show you how to use           '
'   mod3Denhance.bas, which is a module that allows             '
'   you to use Direct3D for bltting to the screen.              '
'   This is very useful for some features that are not          '
'   available in DDraw, eg. Alpha-Blending.                     '
'   It is also useful for isometrics, as your not restricted    '
'   to rectangles.                                              '
'                                                               '
'   Feel free to use this code for anything you want, but       '
'   if you do use it for anything, then I'd appreciate an       '
'   E-mail telling me what you've done and a Thank-you          '
'   would be nice aswell.                                       '
'   Also, please E-mail me with any comments, suggestions or    '
'   problems with the code.                                     '
'   Or Please just E-mail to say what you think of the code     '
'                                                               '
'List of subs/functions                                         '
'   Initialize_DX           - Initializes Direct X              '
'   CreateTextureSurface    - Creates a texture surface (D3D)   '
'   DDCreateSurface         - Creates a DDraw surface           '
'   Set_Vertex              - Sets the position and color       '
'   Blt3D                   - Blts a 3D surface to the screen   '
'   Set_Alphablend          - Turns Alpha blending on/off       '
'   Set_ColorKey            - Turns Color Keying on/off         '
'   Clear_Device            - Used to clear the D3D Device      '
'   Endit                   - Ends the program                  '
'                                                               '
'   Please look at the subs comments for more info.             '
'                                                               '
'   Carl Warwick     - ( CC006740@ntu.ac.uk )                   '
'   Freeride Designs - ( http://freeridedesign.cjb.net )        '
'   1 - April - 2001                                            '
'****************************************************************
Option Explicit

Public DX As DirectX7                    'Direct X
Public DD As DirectDraw7                 'Direct Draw
Public Primary As DirectDrawSurface7     'Primary surface
Public BackBuffer As DirectDrawSurface7  'Backbuffer
Public Direct3D As Direct3D7             'Direct 3D
Public Device As Direct3DDevice7         'Direct 3D Device

Public ddsTexture1 As DirectDrawSurface7 'Texture surface for logo
Public ddsurface As DirectDrawSurface7   'Normal dd surface for background

Public bRunning As Boolean               'Boolean to check if program is still running

Public Fire As clsFire
Public Fire1 As clsFire
Public Smoke As clsSmoke
Public Snow As clsSnow
Public g_ParticleCount As Integer

'**************************************************************************************
'This sub initializes DDraw and D3D
'If you know what you are doing in DDraw the n this should be fairly easy to
'understand, and change to suit you own needs (eg. If you want windowed mode)
'**************************************************************************************
Public Sub Initialize_DX(mForm As Form, Optional Width As Integer = 640, Optional Height As Integer = 480, Optional BPP As Byte = 16)
    Dim ddsd As DDSURFACEDESC2
    Dim caps As DDSCAPS2
    Dim DEnum As Direct3DEnumDevices
    Dim Guid As String
    
    ' Create the DirectDraw object and set the application
    ' cooperative level.
    Set DX = New DirectX7
    Set DD = DX.DirectDrawCreate("")
    
    DD.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT
    DD.SetDisplayMode Width, Height, BPP, 0, DDSDM_DEFAULT
    
    ' Prepare and create the primary surface.
    ddsd.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_3DDEVICE Or DDSCAPS_PRIMARYSURFACE
    ddsd.lBackBufferCount = 1
    
    Set Primary = DD.CreateSurface(ddsd)
    
    'Attach the backbuffer. the DDSCAPS_3DDEVICE tells the
    'backbuffer that its to be used for 3D stuff
    caps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_3DDEVICE
    Set BackBuffer = Primary.GetAttachedSurface(caps)

    
    Set Direct3D = DD.GetDirect3D
    
    'Get the last device driver (last one is usually the best one)
    Set DEnum = Direct3D.GetDevicesEnum
    Guid = DEnum.GetGuid(DEnum.GetCount)
    
    'Set the Device
    Set Device = Direct3D.CreateDevice(Guid, BackBuffer)
    
    'Our blending states.
    'you can change D3DBLEND_ONE, check the DX7 SDK for other values
    'Device.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_INVSRCALPHA
    'Device.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
    Device.SetRenderState D3DRENDERSTATE_DESTBLEND, D3DBLEND_ONE
    Device.SetRenderState D3DRENDERSTATE_SRCBLEND, D3DBLEND_SRCALPHA
    Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
End Sub


'**************************************************************************************
'Call this sub to create a texture surface (surface to be used by D3D)
'w = Width, h = Height
'ColKey is used for setting the colour key,
'0 = no colour key : 1 = Black colour key : 2 = Magneta (255,0,255) colour key
'
'Note.
'-----
'Texture surface dimensions must be to a power or 2 (^2)
'eg. possible sizes are :- 256x256, 128x128, 128,256, 16x16 etc...
'Its not advisable to use surface bigger than 256x256 or at the most 512x512
'**************************************************************************************
Public Function CreateTextureSurface(sFile As String, w As Long, h As Long, Optional ColKey As Integer = 0) As DirectDrawSurface7

    Dim bOK As Boolean
    Dim enumTex As Direct3DEnumPixelFormats
    Dim sLoadFile As String
    Dim i As Long
    Dim ddsd As DDSURFACEDESC2
    Dim SurfaceObject As DirectDrawSurface7
    Dim Init As Boolean

    ddsd.lFlags = DDSD_CAPS Or DDSD_TEXTURESTAGE Or DDSD_PIXELFORMAT
    If ((h <> 0) And (w <> 0)) Then
        ddsd.lFlags = ddsd.lFlags Or DDSD_HEIGHT Or DDSD_WIDTH
        ddsd.lHeight = h
        ddsd.lWidth = w
    End If
    
    Set enumTex = Device.GetTextureFormatsEnum()

    'check if device supports 16bit surfaces
    For i = 1 To enumTex.GetCount()
        bOK = True
        Call enumTex.GetItem(i, ddsd.ddpfPixelFormat)

        With ddsd.ddpfPixelFormat
            If .lRGBBitCount <> 16 Then bOK = False
        End With
        If bOK = True Then Exit For
    Next

    If bOK = False Then
        Debug.Print "Unable to find 16bit surface support on your hardware - exiting"
        Init = False
    End If

    'set some texture surface flags
    If Device.GetDeviceGuid() = "IID_IDirect3DHALDevice" Then
            ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
            ddsd.ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
            ddsd.lTextureStage = 0
    Else
            ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
            ddsd.ddsCaps.lCaps2 = 0
            ddsd.lTextureStage = 0
    End If

    'If no filename was passed then create a blank surface
    If sFile = "" Then
        Set SurfaceObject = DD.CreateSurface(ddsd)
    Else
        Set SurfaceObject = DD.CreateSurfaceFromFile(sFile, ddsd)
    End If

    Set CreateTextureSurface = SurfaceObject

    'Colour key
    Dim ddckColourKey As DDCOLORKEY
    Dim ddpf As DDPIXELFORMAT
    If ColKey = 1 Then 'Black colorkey
        ddckColourKey.low = 0
        ddckColourKey.high = 0
        CreateTextureSurface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
    ElseIf ColKey = 2 Then 'Magneta Colorkey
        CreateTextureSurface.GetPixelFormat ddpf
        ddckColourKey.low = ddpf.lBBitMask + ddpf.lRBitMask
        ddckColourKey.high = ddckColourKey.low
        CreateTextureSurface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
    End If
End Function

'**************************************************************************************
'Create a normal DirectDraw Surface
'TransCol is used for setting the colour key,
'0 = no colour key : 1 = Black colour key : 2 = Magneta (255,0,255) colour key
'**************************************************************************************
Public Sub DDCreateSurface(surface As DirectDrawSurface7, sWidth, sHeight, sFile As String, Optional TransCol As Integer = 0, Optional UseSysMem As Boolean = False)
    'This sub will load a bitmap from a file
    'into a specified dd surface. Transparent
    'colour is black (0) by default.
    
    Dim tempddsd As DDSURFACEDESC2
    Set surface = Nothing
    
    'Load sprite
    tempddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If UseSysMem = True Then
        tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Else
        tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    End If
    tempddsd.lWidth = sWidth
    tempddsd.lHeight = sHeight
    Set surface = DD.CreateSurfaceFromFile(sFile, tempddsd)
    
    'Colour key
    Dim ddckColourKey As DDCOLORKEY
    Dim ddpf As DDPIXELFORMAT
    If TransCol = 1 Then 'Black colorkey
        ddckColourKey.low = 0
        ddckColourKey.high = 0
        surface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
    ElseIf TransCol = 2 Then 'Magneta Colorkey
        surface.GetPixelFormat ddpf
        ddckColourKey.low = ddpf.lBBitMask + ddpf.lRBitMask
        ddckColourKey.high = ddckColourKey.low
        surface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
    End If
End Sub

'**************************************************************************************
'Call this to set the position and colour of the four corners of a
'texture surface.
'The corners don't need to make a rectangle, but they must be set in a
'clockwise order.
'eg. X1,Y1 = Top Left     :  X2,Y2 = Top Right
'eg. X3,Y3 = Bottom Left  :  X4,Y4 = Bottom Right
'This makes it perfect for drawing isometric tiles!!!
'
'You can also set the colour of each corner to produce fading and light effects,
'if color values are emmited when calling it will be normal (white)
'**************************************************************************************
Public Sub Set_Vertex(Vertex() As D3DTLVERTEX, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, Optional color1 As Long = &HFFFFFF, Optional color2 As Long = &HFFFFFF, Optional color3 As Long = &HFFFFFF, Optional color4 As Long = &HFFFFFF)
    Call DX.CreateD3DTLVertex(X1, Y1, 0, 1, color1, 1, 1, 0, Vertex(0))
    Call DX.CreateD3DTLVertex(X2, Y2, 0, 1, color2, 1, 0, 0, Vertex(1))
    Call DX.CreateD3DTLVertex(X4, Y4, 0, 1, color4, 1, 0, 1, Vertex(3))
    Call DX.CreateD3DTLVertex(X3, Y3, 0, 1, color3, 1, 1, 1, Vertex(2))
End Sub


'**************************************************************************************
'Call this sub to Blt a texture surface to the screen (must be a texture surface),
'This must be called between Device.BeginScene and Device.EndScene
'tmpSurface = a texture surface
'Vertex()   = a vertex list as defined by Set_Vertex()
'**************************************************************************************
Public Sub Blt3D(tmpSurface As DirectDrawSurface7, Vertex() As D3DTLVERTEX)
    Call Device.SetTexture(0, tmpSurface)
    Call Device.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, Vertex(0), 4, D3DDP_DEFAULT)
End Sub


'**************************************************************************************
'Turn alpha blending on or off, you can call this anywhere.
'But this should be called as little as possible, so make sure you do all
'your alpha blended bltting at the same time so you only need to call this once
'to turn on and once to turn off
'**************************************************************************************
Public Sub Set_Alphablend(BlendOn As Boolean)
    Device.SetRenderState D3DRENDERSTATE_ALPHABLENDENABLE, BlendOn
End Sub

'**************************************************************************************
'Turn colour keying on or off, you can call this anywhere.
'But this should be called as little as possible, so make sure you do all
'your color keyed bltting at the same time so you only need to call this once
'to turn on and once to turn off
'**************************************************************************************
Public Sub Set_ColorKey(BlendOn As Boolean)
    Device.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, BlendOn
End Sub


'**************************************************************************************
'Clears the device ready for drawing operations
'This needs to be called after all the calls to Set_Vertex,
'and before Device.BeginScene and any bltting operations
'**************************************************************************************
Public Sub Clear_Device()
    Dim ClearRect(0 To 0) As D3DRECT
    ClearRect(0).X2 = 640
    ClearRect(0).Y2 = 480
    Device.Clear 1, ClearRect, D3DCLEAR_TARGET, 0, 0, 0
End Sub


'**************************************************************************************
'Restore the display mode and return to normal life
'**************************************************************************************
Public Sub Endit()
    ' Clean up DirectX
    Set Fire = Nothing
    
    Call DD.RestoreDisplayMode
    Call DD.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    
    Set Device = Nothing
    Set Direct3D = Nothing
    Set ddsTexture1 = Nothing
    Set ddsurface = Nothing
    Set BackBuffer = Nothing
    Set Primary = Nothing
    Set DD = Nothing
    Set DX = Nothing
    
    End
End Sub

