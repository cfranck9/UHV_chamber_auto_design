Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst

Module Module1

    Structure stPoint
        Dim X As Double
        Dim Y As Double
        Dim Z As Double
        Public Sub New(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
            Me.X = X
            Me.Y = Y
            Me.Z = Z
        End Sub
    End Structure

    Structure stPortSchematic
        Dim FocalPointType As Integer
        Dim Distance As Double
        Dim theta As Double
        Dim phi As Double
        Dim FlangeType As Integer

        Public Sub New(ByVal FocalPointType As Integer, ByVal Distance As Double, ByVal theta As Double,
                       ByVal phi As Double, ByVal FlangeType As Integer)
            Me.FocalPointType = FocalPointType
            Me.Distance = Distance
            Me.theta = theta
            Me.phi = phi
            Me.FlangeType = FlangeType
        End Sub
    End Structure

    Structure stPortCoordinates
        Dim X1 As Double
        Dim Y1 As Double
        Dim Z1 As Double
        Dim X2 As Double                'For coordinates of the 3D sketch endpoints
        Dim Y2 As Double
        Dim Z2 As Double
        Dim X3 As Double                'For coordinates of the actual termination point of the cylindrical extrusion. Needed for shell operation on offset extrusion.
        Dim Y3 As Double
        Dim Z3 As Double

        'constructor
        Public Sub New(ByVal X1 As Double, ByVal Y1 As Double, ByVal Z1 As Double,
                       ByVal X2 As Double, ByVal Y2 As Double, ByVal Z2 As Double,
                       ByVal X3 As Double, ByVal Y3 As Double, ByVal Z3 As Double)
            Me.X1 = X1
            Me.Y1 = Y1
            Me.Z1 = Z1
            Me.X2 = X2
            Me.Y2 = Y2
            Me.Z2 = Z2
            Me.X3 = X3
            Me.Y3 = Y3
            Me.Z3 = Z3
        End Sub
    End Structure

    Dim intstatus, errors As Integer
    Dim boolstatus As Boolean
    Dim swApp As Object
    Dim swPart As ModelDoc2
    Dim swAssembly As AssemblyDoc
    Dim skSegment As SketchSegment
    Dim myFeature As Feature
    Dim myRefPlane As Object
    Dim myComponent As Component2
    Dim myModelView As Object
    Dim AssemblyTitle As String
    Dim rootFolder As String = "c:\users\(user name)\desktop\"
    Dim finalAssemblyName As String = "my_chamber_assembly"

    Sub Main()
        Dim i As Integer

        'ports geometry
        Dim portschematic() As stPortSchematic = {
            New stPortSchematic(0, 0.25, 0, 0, 4),
            New stPortSchematic(0, 0.25, 180, 0, 7),
            New stPortSchematic(0, 0.2, 67.5, 315, 0),
            New stPortSchematic(0, 0.2, 67.5, 45, 0),
            New stPortSchematic(0, 0.2, 67.5, 135, 0),
            New stPortSchematic(0, 0.2, 67.5, 225, 0),
            New stPortSchematic(0, 0.2, 90, 315, 0),
            New stPortSchematic(0, 0.2, 90, 0, 0),
            New stPortSchematic(0, 0.2, 90, 45, 0),
            New stPortSchematic(0, 0.2, 90, 90, 0),
            New stPortSchematic(0, 0.2, 90, 135, 0),
            New stPortSchematic(0, 0.2, 90, 180, 2),
            New stPortSchematic(0, 0.2, 90, 225, 0),
            New stPortSchematic(1, 0.2, 90, 0, 2),
            New stPortSchematic(1, 0.2, 90, 45, 2),
            New stPortSchematic(1, 0.2, 90, 90, 0),
            New stPortSchematic(1, 0.2, 90, 135, 2),
            New stPortSchematic(1, 0.2, 90, 180, 2)
        }

        'focal points
        Dim FocalPoints() As stPoint = {
            New stPoint(0, 0, 0),
            New stPoint(0, -0.12, 0)
            }

        'Flange Type - 0: CF2750, 1: CF3375, 2: CF4500, 3: CF6000, 4: CF8000, 5: CF10000, 6: CF12000, 7: CF14000
        Dim FlangeTypeString() As String = {"CF2750", "CF3375", "CF4500", "CF6000", "CF8000", "CF10000", "CF12000", "CF14000"}
        Dim FlangeOffsetDistance() As Double = {0.00531, 0.005715, 0.009525, 0.0111, 0.0127, 0.0127, 0.0127, 0.0127}
        Dim PortExtrusionRaduis() As Double = {0.022352, 0.0254, 0.03175, 0.0508, 0.0762, 0.1016, 0.127, 0.15}

        Dim ports() As stPortCoordinates = {
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0),
            New stPortCoordinates(0, 0, 0, 0, 0, 0, 0, 0, 0)
        }

        For i = 0 To ports.Length - 1
            ports(i).X1 = FocalPoints(portschematic(i).FocalPointType).X
            ports(i).Y1 = FocalPoints(portschematic(i).FocalPointType).Y
            ports(i).Z1 = FocalPoints(portschematic(i).FocalPointType).Z
            ports(i).X2 = ports(i).X1 + portschematic(i).Distance * Math.Sin(portschematic(i).theta / 180 * Math.PI) * Math.Sin(portschematic(i).phi / 180 * Math.PI)
            ports(i).Y2 = ports(i).Y1 + portschematic(i).Distance * Math.Cos(portschematic(i).theta / 180 * Math.PI)
            ports(i).Z2 = ports(i).Z1 + portschematic(i).Distance * Math.Sin(portschematic(i).theta / 180 * Math.PI) * Math.Cos(portschematic(i).phi / 180 * Math.PI)
            ports(i).X3 = ports(i).X1 + (portschematic(i).Distance - FlangeOffsetDistance(portschematic(i).FlangeType)) * Math.Sin(portschematic(i).theta / 180 * Math.PI) * Math.Sin(portschematic(i).phi / 180 * Math.PI)
            ports(i).Y3 = ports(i).Y1 + (portschematic(i).Distance - FlangeOffsetDistance(portschematic(i).FlangeType)) * Math.Cos(portschematic(i).theta / 180 * Math.PI)
            ports(i).Z3 = ports(i).Z1 + (portschematic(i).Distance - FlangeOffsetDistance(portschematic(i).FlangeType)) * Math.Sin(portschematic(i).theta / 180 * Math.PI) * Math.Cos(portschematic(i).phi / 180 * Math.PI)
        Next

        '============================================================================================================================================='

        swApp = New SldWorks()
        If swApp IsNot Nothing Then
            swApp.Visible = True
        End If

        swPart = swApp.NewDocument("C:\Program Files\SolidWorks Corp\SolidWorks\lang\english\Tutorial\part.prtdot", 0, 0, 0)
        swPart = swApp.ActivateDoc3("", False, False, intstatus)

        'change display units (not internal units used within the present source code)
        boolstatus = swPart.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, 0, swUnitSystem_e.swUnitSystem_MMGS)
        boolstatus = swPart.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 0, 3)

        'build 3d sketch
        swPart.SketchManager.Insert3DSketch(True)   'opens 3d sketch
        For i = 0 To ports.Length - 1
            skSegment = swPart.SketchManager.CreateLine(ports(i).X1, ports(i).Y1, ports(i).Z1, ports(i).X2, ports(i).Y2, ports(i).Z2)
        Next i
        swPart.SketchManager.InsertSketch(True)     'closes 3d sketch
        swPart.ClearSelection2(True)
        swPart.ViewZoomtofit2()

        'planes
        For i = 0 To ports.Length - 1
            'name field can be left blank, as long as coordinates are precisely specified
            'type of object such as "PLANE" / "FACE" has to be all UPPERCASE. Won't work otherwise.
            boolstatus = swPart.Extension.SelectByID2("", "EXTSKETCHSEGMENT", ports(i).X2, ports(i).Y2, ports(i).Z2, True, 0, Nothing, 0)
            boolstatus = swPart.Extension.SelectByID2("", "EXTSKETCHPOINT", ports(i).X2, ports(i).Y2, ports(i).Z2, True, 1, Nothing, 0)
            myRefPlane = swPart.FeatureManager.InsertRefPlane(2, 0, 4, 0, 0, 0)
        Next i
        swPart.ClearSelection2(True)
        swPart.ViewZoomtofit2()

        'primary cylinder
        Dim bodyradius As Double = 0.15
        Dim bodytopheight As Double = 0.14
        Dim bodybottomheight As Double = 0.2
        swPart.Extension.SelectByID2("top", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        swPart.SketchManager.InsertSketch(True)     'opens (2d) sketch
        skSegment = swPart.SketchManager.CreateCircleByRadius(0.0, 0.0, 0.0, bodyradius)    'unit is by defalut meter
        myFeature = swPart.FeatureManager.FeatureExtrusion2(False, False, False, 0, 0, bodytopheight, bodybottomheight, False, False, False, False, 0.01, 0.01, False, False, False, False, True, True, True, 0, 0, False)
        swPart.ClearSelection2(True)

        'port cylinders (solid)
        For i = 0 To ports.Length - 1
            swPart.Extension.SelectByID2("plane" & CStr(i + 1), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
            swPart.SketchManager.InsertSketch(True)     'opens (2d) sketch
            skSegment = swPart.SketchManager.CreateCircleByRadius(0, 0, 0, PortExtrusionRaduis(portschematic(i).FlangeType))     'center position is relative to the origin of the sketch plane
            'direction to extrude flipped except for i=1 (bottom extrusion): 3d sketch for bottom extrusion must be associated with i=1
            'iif(i = 1, false, true)
            myFeature = swPart.FeatureManager.FeatureExtrusion2(True, False, True, 2, 0, 0, 0, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 3, FlangeOffsetDistance(portschematic(i).FlangeType), True)
        Next i
        swPart.ClearSelection2(True)

        'shell
        'shell with multiple radii doesn't usually work..
        swPart.ShowNamedView2("", 8)
        'the above line is very very very important. it turns out that the selection below doesn't work for face that is along the line of your view
        '(surface normal lying in the monitor plane). So it is important to rotate the view properly such that no face to be selected is as such.
        'no need to be di- or tri-metric. May depend on monitor spec & setting.
        swPart.ViewZoomtofit2()
        For i = 0 To ports.Length - 1
            swPart.Extension.SelectByID2("", "FACE", ports(i).X3, ports(i).Y3, ports(i).Z3, True, 0, Nothing, 0)
        Next i
        swPart.InsertFeatureShell(0.002, False)
        swPart.ClearSelection2(True)

        'clear display
        swPart.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches, False)
        swPart.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplayPlanes, False)

        intstatus = swPart.SaveAs3(rootFolder & "chamber.sldprt", 0, 2)
        swApp.CloseDoc("chamber")

        'swApp.ExitApp()

        '============================================================================================================================================='

        'Assembly
        'Finding out the right combination of the following four lines took me many hours...
        'One needs to "opendoc" a part file to be inserted into an assembly first of all, even if that part document is not "CloseDoc"ed above.
        swPart = swApp.NewDocument("C:\Program Files\SolidWorks Corp\SolidWorks\lang\english\Tutorial\assem.asmdot", 0, 0, 0)
        AssemblyTitle = swPart.GetTitle()

        swPart = swApp.ActivateDoc3(AssemblyTitle, False, False, intstatus)
        swAssembly = swPart

        swPart = swApp.OpenDoc6(rootFolder & "chamber.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, intstatus)
        myComponent = swAssembly.AddComponent5(rootFolder & "chamber.SLDPRT", 0, "", False, "", 0.01, 0.01, 0.01) ' doesn't need to be AddComponent5. AddComponent also works.
        swAssembly.ViewZoomtofit2()

        swPart = swApp.ActivateDoc3(AssemblyTitle, False, False, intstatus)
        'the above is a MUST. It activates the assembly window and pulls it up to front so that components in it can be selected
        'as soon as OpenDoc6 and AddComponent5 executes back-to-back, the active window is that for the just-opened part.
        swPart.Extension.SelectByID2("chamber-1@" & AssemblyTitle, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        swPart.UnFixComponent()
        'Solidworks Macro saves unfixing operation into FixComponent() maybe as a toggle function, but in stand-alone mode UnFixComponent() should be used.
        swPart.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        swPart.Extension.SelectByID2("Front@chamber-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        swAssembly.AddMate4(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, intstatus)
        swPart.ClearSelection2(True)
        swPart.Extension.SelectByID2("Top", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        swPart.Extension.SelectByID2("Top@chamber-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        swAssembly.AddMate4(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, intstatus)
        swPart.ClearSelection2(True)
        swPart.Extension.SelectByID2("Right", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        swPart.Extension.SelectByID2("Right@chamber-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        swAssembly.AddMate4(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, intstatus)
        swPart.ClearSelection2(True)
        swPart.Extension.SelectByID2("chamber-1@" & AssemblyTitle, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        swPart.FixComponent()
        swPart.Extension.SelectByID2("Coincident1", "MATE", 0, 0, 0, False, 0, Nothing, 0)
        swPart.Extension.SelectByID2("Coincident2", "MATE", 0, 0, 0, True, 0, Nothing, 0)
        swPart.Extension.SelectByID2("Coincident3", "MATE", 0, 0, 0, True, 0, Nothing, 0)
        swPart.EditDelete()
        swPart.ClearSelection2(True)

        Dim FlangeUsageCount() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}

        For i = 0 To ports.Length - 1
            FlangeUsageCount(portschematic(i).FlangeType) += 1
            swPart = swApp.OpenDoc6(rootFolder & "Flanges\" & FlangeTypeString(portschematic(i).FlangeType) & "flange.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, intstatus)
            myComponent = swAssembly.AddComponent5(rootFolder & "Flanges\" & FlangeTypeString(portschematic(i).FlangeType) & "flange.SLDPRT", 0, "", False, "", 0.01, 0.01, 0.01)
            swPart = swApp.ActivateDoc3(AssemblyTitle, False, False, intstatus)

            ' Mate #1
            swPart.Extension.SelectByID2("Plane" & CStr(i + 1) & "@chamber-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            swPart.Extension.SelectByID2("Top Plane@" & FlangeTypeString(portschematic(i).FlangeType) & "flange-" & FlangeUsageCount(portschematic(i).FlangeType) & "@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            swAssembly.AddMate4(5, 1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, intstatus)
            swPart.ClearSelection2(True)

            ' Mate #2
            If (i < 2) Then    'for top and bottom ports
                swPart.Extension.SelectByID2("Front@chamber-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            Else
                swPart.Extension.SelectByID2("Line1@3DSketch1@chamber-1@" & AssemblyTitle, "EXTSKETCHSEGMENT", 0, 0, 0, True, 1, Nothing, 0)
            End If
            swPart.Extension.SelectByID2("Front Plane@" & FlangeTypeString(portschematic(i).FlangeType) & "flange-" & FlangeUsageCount(portschematic(i).FlangeType) & "@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            swAssembly.AddMate4(3, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, intstatus)
            swPart.ClearSelection2(True)

            ' Mate #3
            swPart.Extension.SelectByID2("Line" & CStr(i + 1) & "@3DSketch1@chamber-1@" & AssemblyTitle, "EXTSKETCHSEGMENT", 0, 0, 0, True, 1, Nothing, 0)
            swPart.Extension.SelectByID2("Point1@Origin@" & FlangeTypeString(portschematic(i).FlangeType) & "flange-" & FlangeUsageCount(portschematic(i).FlangeType) & "@" & AssemblyTitle, "EXTSKETCHPOINT", 0, 0, 0, True, 1, Nothing, 0)
            swAssembly.AddMate4(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, intstatus)
            swPart.ClearSelection2(True)
        Next i

        swPart.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches, False)
        swPart.EditRebuild3()
        intstatus = swAssembly.SaveAs3(rootFolder & finalAssemblyName & ".sldasm", 0, 2)
        swApp.CloseAllDocuments(True)

        swPart = swApp.OpenDoc6(rootFolder & finalAssemblyName & ".sldasm", swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, intstatus)
        swAssembly = swPart
        swAssembly.ResolveAllLightweight()

        'Maximize Solidworks Window
        swApp.FrameState = swWindowState_e.swWindowMaximized
    End Sub

End Module