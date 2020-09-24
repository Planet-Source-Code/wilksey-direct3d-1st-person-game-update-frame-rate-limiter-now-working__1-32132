Attribute VB_Name = "modMain"
' Direct3D 1st Person Game VERSION 1.1 (24/02/2002)
' By Frederico Machado (indiofu@bol.com.br)
' Please vote for me if you like the game.
'
' I'd like to thanks to all DirectX programmers at PSC
' I've downloaded all I found about DirectX, and I learned
' too much from it.
' Special thanks to the guy who made the
' tutorial "Direct3D For Dummiez (In VB6)", that helped me
' a lot, and to Simon Price, the EngineX gave me desire to
' make my own 3D game (The EngineX is in DirectX8, so I
' couldn't understand, I think it's too hard.)
' and to Dave Cline, his Mouse Sub helped me a lot too.
'
' Sorry my English, I'm Brazilian! :)
'
' ************************************************** '
' I need help to use DirectInput to take a look      '
' with the mouse.                                    '
' That sky is really sucks, I thought in something   '
' like the sky of TrueVisionSDK.                     '
' If anyone knows how to do that, please help me.    '
' ************************************************** '

' Variables of DirectX and others

'================================
' Direct Input Added By Wilksey.
'================================

' Main variables
Global DX_Main As New DirectX7  ' The boss, main object
Global DD_Main As DirectDraw4   ' DirectDraw object
Global D3D_Main As Direct3DRM3  ' Direct3D object

' DirectInput components
Global DI_Main As DirectInput   ' Main obj of DirectInput
Global DI_Device As DirectInputDevice ' DirectInput Device
Global DI_State As DIKEYBOARDSTATE ' Hold the state of the keys
Global DI_MouseDevice As DirectInputDevice 'Direct Input Mouse Device *Added By Wilksey*
Global DI_MouseState As DIMOUSESTATE       'Direct Input Mouse State *Added By Wilksey*

' DirectDraw surfaces, where the screen is draw.
Global DS_Front As DirectDrawSurface4 ' The front buffer, what we see on the screen
Global DS_Back As DirectDrawSurface4 ' The back buffer, where everything is draw before it's put on the screen
Global SD_Front As DDSURFACEDESC2 ' The surface description
Global DD_Back As DDSCAPS2 ' General surface info

' ViewPort and Direct3D Device
Global D3D_Device As Direct3DRMDevice3 ' The Main Direct3D Retained Mode Device
Global D3D_ViewPort As Direct3DRMViewport2 ' The Direct3D Retained Mode Viewport (Kinda the camera)

' Frames
Global FR_Root As Direct3DRMFrame3 ' The Main Frame (The other frames are put under this one (Like a tree))
Global FR_Camera As Direct3DRMFrame3 ' We will use this frame as a camera.
Global FR_Light As Direct3DRMFrame3 ' This frame contains our spotlights
Global FR_Building As Direct3DRMFrame3 ' Frame containing all walls, floors and roofs
                                       ' I think it's easyer than use more frames to walls, floors, ...
                                       ' but, if anyone thinks it's dumb or slow, please tell me, I don't know too much about DX.

' Lights
Global LT_Ambient As Direct3DRMLight ' The main (ambient) light, that illuminates everything
Global LT_Spot As Direct3DRMLight ' Our spot light, makes it look more realistic.

Global Path As String
Global ESC As Boolean

' I've used those functions because I couldn't use the DirectInput mouse
' I don't know to use it yet. If anyone knows how to use the DirectInput
' to look around with the mouse, please help me.
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Type POINTAPI
    X As Long
    y As Long
End Type

Global Mpos As POINTAPI ' Contains the mouse position

Global Rotation As Single
Public LastTimeDrawn As Long
'Frame rate limiter by Thomas Sturm
Public Const Delaytime = 1000 / 60 'this delay is in milliseconds
'----------------------------------



' DirectX Initialization
Public Sub DX_Init()

  On Error GoTo InitError

  Set DD_Main = DX_Main.DirectDraw4Create("") ' Create the DirectDraw object

  DD_Main.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE ' Set the screen mode (full screen)
  DD_Main.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT ' Set Resolution and BitDepth (Lets use 32-bit color)

  SD_Front.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  SD_Front.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
  SD_Front.lBackBufferCount = 1 ' Make one backbuffer
  Set DS_Front = DD_Main.CreateSurface(SD_Front) ' Initialize the front buffer (the screen)
  ' The Previous block of code just created the screen and the backbuffer.

  DD_Back.lCaps = DDSCAPS_BACKBUFFER
  Set DS_Back = DS_Front.GetAttachedSurface(DD_Back)
  DS_Back.SetForeColor RGB(255, 255, 255)
  ' The backbuffer was initialized and the DirectDraw text color was set to white.

  Set D3D_Main = DX_Main.Direct3DRMCreate() ' Creates the Direct3D Retained Mode Object

  Set D3D_Device = D3D_Main.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD_Main, DS_Back, D3DRMDEVICE_DEFAULT) ' Tell the Direct3D Device that we are using hardware rendering (HALDevice)

  D3D_Device.SetBufferCount 2 ' Set the number of buffers
  D3D_Device.SetQuality D3DRMRENDER_GOURAUD ' Set the rendering quality. GOURAUD has the best rendering quality.
  D3D_Device.SetTextureQuality D3DRMTEXTURE_LINEAR ' Set the texture quality
  D3D_Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY ' Set the render mode

  Set DI_Main = DX_Main.DirectInputCreate() ' Create the DirectInput Device
  Set DI_Device = DI_Main.CreateDevice("GUID_SysKeyboard") ' Set it to use the keyboard.
  DI_Device.SetCommonDataFormat DIFORMAT_KEYBOARD ' Set the data format to the keyboard format
  DI_Device.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE ' Set Cooperative level
  DI_Device.Acquire
  '*****************************Added By Wilksey************************************
  Set DI_Main = DX_Main.DirectInputCreate()     'Create a new Direct Input Device
  Set DI_MouseDevice = DI_Main.CreateDevice("GUID_SysMouse")   'Set it to use the mouse
  DI_MouseDevice.SetCommonDataFormat DIFORMAT_MOUSE 'Set the data format to mouse
  DI_MouseDevice.SetCooperativeLevel frmMain.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE  'Set Co-Operative Level
  DI_MouseDevice.Acquire 'Assign the device.
  '*********************************************************************************
  ' The above block of code configures the DirectInput Device and starts it.

  Exit Sub

InitError:

  ' Lets restore the display mode to show the msgbox
  DD_Main.RestoreDisplayMode
  DD_Main.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
  DI_Device.Unacquire

  MsgBox "Error: Could not initialize DirectX (See DX_Init).", vbCritical, "Error"
  Unload frmMain
 
End Sub

' Creates the objects, lights, meshes, etc..
Public Sub DX_MakeObjects()

  On Error GoTo ObjectsError
  
  Set FR_Root = D3D_Main.CreateFrame(Nothing) ' This will be the root frame of the "tree"
  Set FR_Camera = D3D_Main.CreateFrame(FR_Root) ' Our Camera's Sub Frame. It goes under FR_Root in the "tree".
  Set FR_Light = D3D_Main.CreateFrame(FR_Root) ' Light sub frame
  Set FR_Building = D3D_Main.CreateFrame(FR_Root) ' Our building (will contains the walls, roofs and floors) sub frame

  FR_Root.SetSceneBackgroundRGB 0, 0, 0 ' Set the background color. Use decimals, not the standerd 255 = max.

  FR_Camera.SetPosition Nothing, 20, 6, 20 ' Set the camera position. X = 20, Y = 6, Z = 20
  Set D3D_ViewPort = D3D_Main.CreateViewport(D3D_Device, FR_Camera, 0, 0, 800, 600) ' Make the viewport and set it to be our camera.

  D3D_ViewPort.SetBack 800 ' How far back it will draw the image. (like a visibility limit)

  Set LT_Ambient = D3D_Main.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.45, 0.45, 0.45) ' Create the ambient light.
  FR_Root.AddLight LT_Ambient ' Add the light to its frame
  
  LoadMap Path & "Maps\test.txt" ' Load the map (it contains the coords of walls, roofs, floors and lights and its properties)

  Exit Sub

ObjectsError:

  ' Lets restore the display mode to show the msgbox
  DD_Main.RestoreDisplayMode
  DD_Main.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
  DI_Device.Unacquire

  MsgBox "Error: Could not create game objects (See DX_MakeOjbects).", vbCritical, "Error"
  Unload frmMain

End Sub

' The Main Loop. Ends when we hit esc
Public Sub DX_Render()

  Do While ESC = False
    If DX_Main.TickCount >= LastTimeDrawn + Delaytime Then  'Frame rate limiter by Thomas Sturm
    LastTimeDrawn = DX_Main.TickCount                       ' ----------------------------------
    On Local Error Resume Next ' Lets go on at all costs
    DoEvents ' think PC, and do what you have to do...
    
    DX_Keyboard ' Call function to verify keys
    DX_Mouse ' Look around with the mouse
    D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER ' Clean the ViewPort
    D3D_Device.Update ' Update the Direct3D device
    D3D_ViewPort.Render FR_Root ' Render our objects (meshes, lights, etc)
    DS_Front.Flip Nothing, DDFLIP_WAIT ' Flip the back buffer with the front buffer.
    End If
Loop
    
End Sub

' Lets verify the keys from keyboard
' Check the collison detection
Public Sub DX_Keyboard()

  Const Sin5 = 8.715574E-02! ' Sin(5°) ' We'd use this if we want to rotate left and right
  Const Cos5 = 0.9961947! ' Cos(5°)    ' but the mouse do that, then we just strafe left and right
  
  Dim OldPos As D3DVECTOR
  FR_Camera.GetPosition Nothing, OldPos
  
  DI_Device.GetDeviceStateKeyboard DI_State ' Get the array of keyboard keys and their current states
  
  If DI_State.Key(DIK_ESCAPE) <> 0 Then Unload frmMain ' If user presses [esc] then exit end the program.
  
  If DI_State.Key(DIK_LEFT) <> 0 Then
    FR_Camera.SetPosition FR_Camera, -1, 0, 0 ' Move the viewport to the left
  End If
  
  If DI_State.Key(DIK_RIGHT) <> 0 Then
    FR_Camera.SetPosition FR_Camera, 1, 0, 0 ' Move the viewport to the right
  End If
  
  
  If DI_State.Key(DIK_UP) <> 0 Then
    If DI_State.Key(DIK_LSHIFT) <> 0 Or DI_State.Key(DIK_RSHIFT) <> 0 Then
      FR_Camera.SetPosition FR_Camera, 0, 0, 2.5 ' (Run) Move the viewport forward
    Else
      FR_Camera.SetPosition FR_Camera, 0, 0, 1.5 ' Move the viewport forward
    End If
  End If
  
  If DI_State.Key(DIK_DOWN) <> 0 Then
    FR_Camera.SetPosition FR_Camera, 0, 0, -1 ' Move the ViewPort back
  End If

  ' Lets check the collision detection
  Dim CVector As D3DVECTOR
  FR_Camera.GetPosition Nothing, CVector
  
  CVector = CheckCollision(OldPos, CVector) ' Colision Detection
  
  FR_Camera.SetPosition Nothing, CVector.X, 6, CVector.Z

End Sub

' I want to thanks to Dave Cline
' I got this sub from his game "ShotEm"
' He has all credits for this!!!
' I really want to use DirectInput to look around with the mouse
' so, if anyone knows, please help me.
' Direct Input Added By Wilksey.
Public Sub DX_Mouse()

  Const RDiv = 100
  Dim MPosNow As POINTAPI
  DI_MouseDevice.GetDeviceStateMouse DI_MouseState
  Mpos.X = DI_MouseState.X * 0.5
  Mpos.y = DI_MouseState.y * 0.5
  If Mpos.X > MPosNow.X Then
    rspeed = Abs(MPosNow.X - Mpos.X) / RDiv
    If rspeed < 2 Then FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, rspeed Else FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
  ElseIf Mpos.X < MPosNow.X Then
    rspeed = Abs(Mpos.X - MPosNow.X) / RDiv
    If rspeed < 2 Then FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -rspeed Else FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
  End If
  
  If Mpos.y > MPosNow.y Then
    rspeed = Abs(MPosNow.y - Mpos.y) / RDiv
    'If InvertMouse = True Then rspeed = -rspeed
    rspeed = rspeed / 2
    If rspeed < 1.8 Then FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, rspeed Else FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
  ElseIf Mpos.y < MPosNow.y Then
    rspeed = Abs(MPosNow.y - Mpos.y) / RDiv
    'If InvertMouse = True Then rspeed = -rspeed
    rspeed = rspeed / 2
    If rspeed < 1.8 Then FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -rspeed Else FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
  End If
  Mpos.X = MPosNow.X
  Mpos.y = MPosNow.y

  'Makes the camera not barrel rolled
  Dim RC As D3DVECTOR
  Dim RCU As D3DVECTOR
  FR_Camera.GetOrientation Nothing, RC, RCU
  FR_Camera.SetOrientation Nothing, RC.X, RC.y, RC.Z, 0, 1, 0

End Sub
