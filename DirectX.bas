Attribute VB_Name = "DirectX"
Option Explicit
'declare objects
Dim g_dx As New DirectX7
Dim m_dd As DirectDraw7
Dim m_ddClipper As DirectDrawClipper
Dim m_rm As Direct3DRM3

'declare devices
Dim m_rmDevice As Direct3DRMDevice3

'declare viewports
Dim m_rmViewport As Direct3DRMViewport2

'declare frames
Dim m_rootFrame As Direct3DRMFrame3
Dim m_lightFrame As Direct3DRMFrame3
Dim m_cameraFrame As Direct3DRMFrame3
Dim m_objectFrame As Direct3DRMFrame3

'declare meshes
Dim m_meshBuilder As Direct3DRMMeshBuilder3

'delare lights
Dim m_light As Direct3DRMLight
Dim m_ambientLight As Direct3DRMLight

'viewport sizes
Dim m_width As Long
Dim m_height As Long

'mousedown positions for rotation
Public m_LastX As Integer
Public m_LastY As Integer


Sub InitRM(PicBox As PictureBox)

    'Create Direct Draw From Current Display Mode
    Set m_dd = g_dx.DirectDrawCreate("")
    
    'Create new clipper object and associate it with a window'
    Set m_ddClipper = m_dd.CreateClipper(0)
    m_ddClipper.SetHWnd PicBox.hWnd
            
    'save the width and height of the picture in pixels
    m_width = PicBox.ScaleWidth
    m_height = PicBox.ScaleHeight
    
    'Create the Retained Mode object
    Set m_rm = g_dx.Direct3DRMCreate()
    
    'Create the Retained Mode device to draw to
    Set m_rmDevice = m_rm.CreateDeviceFromClipper(m_ddClipper, "", m_width, m_height)
    
    'set shading quality
    m_rmDevice.SetQuality D3DRMRENDER_GOURAUD
    
End Sub

Sub CleanUp()
       
    'free memory taken by frames, etc
    Set m_light = Nothing
    Set m_ambientLight = Nothing
    Set m_meshBuilder = Nothing
    Set m_rmViewport = Nothing
    
    Set m_lightFrame = Nothing
    Set m_cameraFrame = Nothing
    Set m_objectFrame = Nothing
    Set m_rootFrame = Nothing

    Set m_rmDevice = Nothing
    Set m_ddClipper = Nothing
    Set m_rm = Nothing
    Set m_dd = Nothing
 
End Sub

Sub InitScene()

    'Setup a scene graph with a camera and lighting
    Set m_rootFrame = m_rm.CreateFrame(Nothing)
    Set m_cameraFrame = m_rm.CreateFrame(m_rootFrame)
    Set m_lightFrame = m_rm.CreateFrame(m_rootFrame)
    Set m_objectFrame = m_rm.CreateFrame(m_rootFrame)
    
    'position the camera and create the Viewport
    'provide the device the viewport uses to render, the frame whose orientation and position
    'is used to determine the camera, and a rectangle describing the extents of the viewport
    m_cameraFrame.SetPosition Nothing, 0, 0, -100
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_cameraFrame, 0, 0, m_width, m_height)
    
    'set the depth of the viewport
    m_rmViewport.SetBack (2000)
    
    'create the white light and hang it off the light frame
    Set m_light = m_rm.CreateLight(D3DRMLIGHT_DIRECTIONAL, &HFFFFFFFF)
    m_lightFrame.AddLight m_light
    
    'create ambient lighting
    Set m_ambientLight = m_rm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)   'Shadow
    m_lightFrame.AddLight m_ambientLight
        
End Sub

Sub AddShape(NumVerts As Integer, VertArray() As D3DVECTOR, SideFaces() As Long)
    
    'create a meshbuilder object
    Set m_meshBuilder = m_rm.CreateMeshBuilder()
    
    Dim NormArray(0) As D3DVECTOR   'normal arrays
    
    'add the shape to the mesh to "create" an object
    m_meshBuilder.AddFaces NumVerts, VertArray, 0, NormArray, SideFaces
    
    'set the color
    m_meshBuilder.SetColorRGB 0, 1, 1
    
    'add the object to the scene
    m_objectFrame.AddVisual m_meshBuilder
    
End Sub

Sub RenderScene()

On Local Error Resume Next
    'clear the rendering surface rectangle described by the viewport
    m_rmViewport.Clear D3DRMCLEAR_ALL
    
    'render to the device
    m_rmViewport.Render m_rootFrame
    
    'blt the image to the screen
    m_rmDevice.Update
    
    'allow other tasks to be processed (be nice)
    DoEvents
End Sub

Sub RotateTrackBall(X As Integer, Y As Integer)
    
    'this function taken from MS engine sample in the VB sample that
    'comes with MS DirectX7 SDK. It works as follows:
        'select point on screen interpret as though selecting
        'point on sphere. as new point is passed when mouse is
        'moved rotate in coresponding direction(s) on sphere.
    
    Dim delta_x As Single, delta_y As Single
    Dim delta_r As Single, radius As Single, denom As Single, Angle As Single
    
    ' rotation axis in camcoords, worldcoords, sframecoords
    Dim axisC As D3DVECTOR
    Dim wc As D3DVECTOR
    Dim axisS As D3DVECTOR
    Dim base As D3DVECTOR
    Dim origin As D3DVECTOR
    
    delta_x = X - m_LastX
    delta_y = Y - m_LastY
    m_LastX = X
    m_LastY = Y

            
     delta_r = Sqr(delta_x * delta_x + delta_y * delta_y)
     radius = 50
     denom = Sqr(radius * radius + delta_r * delta_r)
    
    If (delta_r = 0 Or denom = 0) Then Exit Sub
    Angle = (delta_r / denom)

    axisC.X = (-delta_y / delta_r)
    axisC.Y = (-delta_x / delta_r)
    axisC.Z = 0

    m_cameraFrame.Transform wc, axisC
    m_objectFrame.InverseTransform axisS, wc
        
    m_cameraFrame.Transform wc, origin
    m_objectFrame.InverseTransform base, wc
    
    axisS.X = axisS.X - base.X
    axisS.Y = axisS.Y - base.Y
    axisS.Z = axisS.Z - base.Z
    
    m_objectFrame.AddRotation D3DRMCOMBINE_BEFORE, axisS.X, axisS.Y, axisS.Z, Angle
    
End Sub

Sub RotatePlayer(Angle)
 m_cameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, Angle
End Sub

Sub TiltPlayer(Angle)
 m_cameraFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, Angle
End Sub


Sub MovePlayerForward(Distance)
 m_cameraFrame.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, Distance
End Sub

