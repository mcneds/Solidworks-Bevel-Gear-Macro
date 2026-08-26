Attribute VB_Name = "modBevelGear"
Option Explicit

' ============================================================
' SOLIDWORKS 2025 - Approximate Straight Bevel Gear Generator (v15.1)
' ------------------------------------------------------------
' Geometry:
'   - conical blank made with a loft
'   - straight bevel tooth spaces made with a lofted cut
'   - transverse tooth flanks use a segmented involute
'   - tooth-space cut is circular-patterned
'
' Pair mode:
'   - computes both pitch cone angles from shaft angle + tooth counts
'   - generates both gears as separate solid bodies
'   - when no custom plane/face was selected before starting the macro,
'     gear 2 is moved into the calculated intersecting-axis position.
'
' Units in the form are mm / degrees.
' Internal API geometry uses meters / radians.
' ============================================================

Public swApp As SldWorks.SldWorks
Public swModel As SldWorks.ModelDoc2
Public swPart As SldWorks.PartDoc
Public gBaseRef As Object
Public gUsingDefaultBase As Boolean

Private Const PI As Double = 3.14159265358979
Private Const FLANK_SAMPLES As Long = 16
Private Const ROOT_ARC_SAMPLES As Long = 6

Public Sub main()
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim selType As Long
    Dim swFace As SldWorks.Face2
    Dim swSurf As SldWorks.Surface
    Dim swFeat As SldWorks.Feature

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Open a part before running the bevel gear macro.", vbExclamation, "Bevel Gear"
        Exit Sub
    End If

    If swModel.GetType <> swDocPART Then
        MsgBox "This macro currently generates gears in a part document.", vbExclamation, "Bevel Gear"
        Exit Sub
    End If

    Set swPart = swModel
    Set swSelMgr = swModel.SelectionManager

    gUsingDefaultBase = True
    Set gBaseRef = Nothing

    If swSelMgr.GetSelectedObjectCount2(-1) > 0 Then
        selType = swSelMgr.GetSelectedObjectType3(1, -1)

        If selType = swSelDATUMPLANES Then
            Set gBaseRef = swSelMgr.GetSelectedObject6(1, -1)
            gUsingDefaultBase = False

        ElseIf selType = swSelFACES Then
            Set swFace = swSelMgr.GetSelectedObject6(1, -1)
            Set swSurf = swFace.GetSurface
            If Not swSurf Is Nothing Then
                If swSurf.IsPlane Then
                    Set gBaseRef = swFace
                    gUsingDefaultBase = False
                End If
            End If
        End If
    End If

    If gBaseRef Is Nothing Then
        Set swFeat = GetNthReferencePlane(1)
        If swFeat Is Nothing Then
            MsgBox "Could not find an origin reference plane.", vbCritical, "Bevel Gear"
            Exit Sub
        End If
        Set gBaseRef = swFeat
        gUsingDefaultBase = True
    End If

    UserForm1.Show
End Sub

Public Sub GenerateFromForm(ByVal frm As Object)
    On Error GoTo EH

    Dim pairMode As Boolean
    Dim lockMode As Long

    Dim m As Double
    Dim d1 As Double, d2 As Double
    Dim z1 As Long, z2 As Long
    Dim sigma As Double
    Dim delta1 As Double, delta2 As Double
    Dim pressure As Double
    Dim faceW As Double
    Dim bore1 As Double, bore2 As Double

    Dim rp1 As Double, rp2 As Double
    Dim R1 As Double, R2 As Double
    Dim minR As Double
    Dim msg As String
    Dim ans As VbMsgBoxResult

    Dim body1 As SldWorks.Body2
    Dim body2 As SldWorks.Body2
    Dim buildX As Double

    Dim apex1X As Double, apex1Y As Double, apex1Z As Double
    Dim axis1X As Double, axis1Y As Double, axis1Z As Double
    Dim apex2X As Double, apex2Y As Double, apex2Z As Double
    Dim axis2X As Double, axis2Y As Double, axis2Z As Double

    pairMode = (Ctl(frm, "cmbMode").ListIndex = 1)
    lockMode = Ctl(frm, "cmbLock").ListIndex

    pressure = DegToRad(ParsePositiveDouble(Ctl(frm, "txtPressureAngle").Text, "Pressure angle"))
    faceW = MmToM(ParsePositiveDouble(Ctl(frm, "txtFaceWidth").Text, "Face width"))
    bore1 = MmToM(ParseNonNegativeDouble(Ctl(frm, "txtBore1").Text, "Gear 1 bore"))

    If pressure <= 0 Or pressure >= DegToRad(45#) Then
        Err.Raise vbObjectError + 100, , "Pressure angle must be between 0° and 45°."
    End If

    Select Case lockMode
        Case 0  ' Module + teeth -> diameter
            m = MmToM(ParsePositiveDouble(Ctl(frm, "txtModule").Text, "Module"))
            z1 = ParsePositiveInteger(Ctl(frm, "txtTeeth1").Text, "Gear 1 teeth")
            d1 = m * z1

        Case 1  ' Diameter + teeth -> module
            d1 = MmToM(ParsePositiveDouble(Ctl(frm, "txtPitchDia1").Text, "Gear 1 pitch diameter"))
            z1 = ParsePositiveInteger(Ctl(frm, "txtTeeth1").Text, "Gear 1 teeth")
            m = d1 / z1

        Case 2  ' Module + diameter -> teeth
            m = MmToM(ParsePositiveDouble(Ctl(frm, "txtModule").Text, "Module"))
            d1 = MmToM(ParsePositiveDouble(Ctl(frm, "txtPitchDia1").Text, "Gear 1 pitch diameter"))
            z1 = CLng(d1 / m + 0.5)

            If Abs(d1 - m * z1) > 0.0000005 Then
                Err.Raise vbObjectError + 101, , _
                    "Module + pitch diameter does not produce an integer tooth count." & vbCrLf & _
                    "Calculated teeth = " & Format(d1 / m, "0.000")
            End If
            d1 = m * z1

        Case Else
            Err.Raise vbObjectError + 102, , "Select a parameter lock mode."
    End Select

    If z1 < 5 Then
        Err.Raise vbObjectError + 103, , "Gear 1 must have at least 5 teeth."
    End If

    If pairMode Then
        z2 = ParsePositiveInteger(Ctl(frm, "txtTeeth2").Text, "Gear 2 teeth")
        If z2 < 5 Then
            Err.Raise vbObjectError + 104, , "Gear 2 must have at least 5 teeth."
        End If

        ' Pair mode uses one common bore diameter for both gears.
        bore2 = bore1
        sigma = DegToRad(ParsePositiveDouble(Ctl(frm, "txtShaftAngle").Text, "Shaft angle"))

        If sigma <= 0 Or sigma >= PI Then
            Err.Raise vbObjectError + 105, , "Shaft angle must be greater than 0° and less than 180°."
        End If

        d2 = m * z2

        delta1 = Atan2(Sin(sigma), (CDbl(z2) / CDbl(z1)) + Cos(sigma))
        delta2 = Atan2(Sin(sigma), (CDbl(z1) / CDbl(z2)) + Cos(sigma))

        If delta1 <= 0 Or delta2 <= 0 Then
            Err.Raise vbObjectError + 106, , "The requested shaft-angle/tooth-count combination produced an invalid pitch cone."
        End If
    Else
        bore2 = 0#
        d2 = 0#
        z2 = 0
        sigma = 0#
        delta1 = DegToRad(ParsePositiveDouble(Ctl(frm, "txtPitchAngle1").Text, "Pitch cone angle"))

        If delta1 <= 0 Or delta1 >= (PI / 2#) Then
            Err.Raise vbObjectError + 107, , "Single-gear pitch cone angle must be between 0° and 90°."
        End If
    End If

    rp1 = d1 / 2#
    R1 = rp1 / Sin(delta1)

    If pairMode Then
        rp2 = d2 / 2#
        R2 = rp2 / Sin(delta2)
        minR = R1
        If R2 < minR Then minR = R2
    Else
        minR = R1
    End If

    If faceW >= minR Then
        Err.Raise vbObjectError + 108, , "Face width must be smaller than the cone distance."
    End If

    If faceW > minR / 3# Then
        msg = "The requested face width is " & Format(MToMm(faceW), "0.###") & " mm." & vbCrLf & _
              "A common straight-bevel design guideline is face width <= about one-third of cone distance." & vbCrLf & _
              "Cone distance = " & Format(MToMm(minR), "0.###") & " mm." & vbCrLf & vbCrLf & _
              "Generate it anyway?"
        ans = MsgBox(msg, vbYesNo + vbExclamation, "Large bevel gear face width")
        If ans <> vbYes Then Exit Sub
    End If

    ValidateBore m, z1, delta1, faceW, bore1, "Gear 1"
    If pairMode Then ValidateBore m, z2, delta2, faceW, bore2, "Gear 2"

    ' Reflect calculated values back into the form.
    Ctl(frm, "txtModule").Text = Format(MToMm(m), "0.######")
    Ctl(frm, "txtTeeth1").Text = CStr(z1)
    Ctl(frm, "txtPitchDia1").Text = Format(MToMm(d1), "0.######")

    If pairMode Then
        Ctl(frm, "txtPitchDia2").Text = Format(MToMm(d2), "0.######")
        Ctl(frm, "txtPitchAngle1").Text = Format(RadToDeg(delta1), "0.######")
        Ctl(frm, "txtPitchAngle2").Text = Format(RadToDeg(delta2), "0.######")
    End If

    swModel.ClearSelection2 True

    Set body1 = BuildGear( _
                    gBaseRef, 0#, _
                    m, z1, delta1, pressure, faceW, bore1, _
                    0#, _
                    apex1X, apex1Y, apex1Z, _
                    axis1X, axis1Y, axis1Z, _
                    "BG1")

    If body1 Is Nothing Then
        Err.Raise vbObjectError + 109, , "Gear 1 was not created."
    End If

    If pairMode Then
        ' Build gear 2 well away from gear 1 so its cut features cannot touch gear 1.
        buildX = 3# * ((d1 / 2#) + m + (d2 / 2#) + m)
        ' Build Gear 2 with its tooth-space pattern phase shifted by half
        ' one tooth pitch.  This replaces the old BG2_MeshPhase Move/Copy
        ' operation entirely.
        Set body2 = BuildGear( _
                        gBaseRef, buildX, _
                        m, z2, delta2, pressure, faceW, bore2, _
                        PI / z2, _
                        apex2X, apex2Y, apex2Z, _
                        axis2X, axis2Y, axis2Z, _
                        "BG2")

        If body2 Is Nothing Then
            Err.Raise vbObjectError + 110, , "Gear 2 was not created."
        End If

        If gUsingDefaultBase Then
            PositionMatingGearDefaultBase _
                "BG1_Body", "BG2_Body", sigma, _
                apex1X, apex1Y, apex1Z, _
                axis1X, axis1Y, axis1Z, _
                apex2X, apex2Y, apex2Z, _
                axis2X, axis2Y, axis2Z
        Else
            MsgBox _
                "Both matched gears were generated, but automatic pair positioning was skipped because a custom plane/face was selected." & vbCrLf & vbCrLf & _
                "The tooth geometry and pitch-cone angles are correct. Automatic 3D pair placement in this version is applied when the macro starts with no custom plane/face selected.", _
                vbInformation, "Matched pair generated"
        End If

        ' The Move/Copy Body feature moves only the solid body, not the
        ' reference planes/axes used to construct it. Hide those references
        ' after generation so their original locations do not confuse the result.
        HideRefFeatureByName "BG1_SmallEndPlane"
        HideRefFeatureByName "BG1_Axis"
        HideRefFeatureByName "BG2_SmallEndPlane"
        HideRefFeatureByName "BG2_Axis"
    Else
        HideRefFeatureByName "BG1_SmallEndPlane"
        HideRefFeatureByName "BG1_Axis"
    End If

    swModel.EditRebuild3
    swModel.ViewZoomtofit2

    If pairMode Then
        MsgBox _
            "Bevel gear pair generated." & vbCrLf & _
            "Module: " & Format(MToMm(m), "0.###") & " mm" & vbCrLf & _
            "Gear 1: " & z1 & "T, pitch dia " & Format(MToMm(d1), "0.###") & " mm, pitch cone " & Format(RadToDeg(delta1), "0.###") & "°" & vbCrLf & _
            "Gear 2: " & z2 & "T, pitch dia " & Format(MToMm(d2), "0.###") & " mm, pitch cone " & Format(RadToDeg(delta2), "0.###") & "°" & vbCrLf & _
            "Shaft angle: " & Format(RadToDeg(sigma), "0.###") & "°", _
            vbInformation, "Bevel Gear"
    Else
        MsgBox _
            "Bevel gear generated." & vbCrLf & _
            z1 & "T, module " & Format(MToMm(m), "0.###") & " mm, pitch dia " & Format(MToMm(d1), "0.###") & " mm", _
            vbInformation, "Bevel Gear"
    End If

    Unload frm
    Exit Sub

EH:
    MsgBox Err.Description, vbCritical, "Bevel Gear"
End Sub

Private Function BuildGear( _
    ByVal baseRef As Object, _
    ByVal centerX As Double, _
    ByVal m As Double, _
    ByVal z As Long, _
    ByVal pitchCone As Double, _
    ByVal pressure As Double, _
    ByVal faceW As Double, _
    ByVal bore As Double, _
    ByVal phaseOffset As Double, _
    ByRef blankApexX As Double, _
    ByRef blankApexY As Double, _
    ByRef blankApexZ As Double, _
    ByRef blankAxisX As Double, _
    ByRef blankAxisY As Double, _
    ByRef blankAxisZ As Double, _
    ByVal prefix As String) As SldWorks.Body2

    On Error GoTo EH

    Dim rp As Double
    Dim ra As Double
    Dim R As Double
    Dim scaleSmall As Double
    Dim axialSmall As Double

    Dim innerPlane As SldWorks.Feature
    Dim skLargeBlank As SldWorks.Feature
    Dim skSmallBlank As SldWorks.Feature
    Dim blankFeat As SldWorks.Feature
    Dim axisFeat As SldWorks.Feature
    Dim skLargeBore As SldWorks.Feature
    Dim skSmallBore As SldWorks.Feature
    Dim boreFeat As SldWorks.Feature
    Dim skLargeTooth As SldWorks.Feature
    Dim skSmallTooth As SldWorks.Feature
    Dim cutFeat As SldWorks.Feature
    Dim patFeat As SldWorks.Feature
    Dim swBody As SldWorks.Body2

    rp = m * z / 2#
    ra = rp + m
    R = rp / Sin(pitchCone)
    scaleSmall = (R - faceW) / R
    axialSmall = faceW * Cos(pitchCone)

    If scaleSmall <= 0# Then
        Err.Raise vbObjectError + 200, , prefix & ": face width collapses the small end of the gear."
    End If

    Set innerPlane = CreateOffsetPlane(baseRef, axialSmall, prefix & "_SmallEndPlane")
    If innerPlane Is Nothing Then
        Err.Raise vbObjectError + 201, , prefix & ": failed to create the small-end plane."
    End If

    Set skLargeBlank = CreateCircleSketch(baseRef, centerX, ra, prefix & "_BlankLarge")
    Set skSmallBlank = CreateCircleSketch(innerPlane, centerX, ra * scaleSmall, prefix & "_BlankSmall")

    SelectSketchProfile skLargeBlank, False
    SelectSketchProfile skSmallBlank, True

    Set blankFeat = swModel.FeatureManager.InsertProtrusionBlend2( _
                        False, False, False, 1#, _
                        0, 0, 0#, 0#, False, False, _
                        False, 0#, 0#, 0, _
                        False, False, True, 0)

    If blankFeat Is Nothing Then
        Err.Raise vbObjectError + 202, , prefix & ": blank loft failed."
    End If
    blankFeat.Name = prefix & "_Blank"

    ' Capture the gear-axis line and common bevel apex NOW, while the blank
    ' still has one clean, unambiguous outer conical face.  Do not try to
    ' rediscover this from the finished tooth-cut body later.
    If Not GetCleanBlankConeApexAxis( _
                blankFeat, _
                blankApexX, blankApexY, blankApexZ, _
                blankAxisX, blankAxisY, blankAxisZ) Then

        Err.Raise vbObjectError + 214, , _
            prefix & ": could not read the clean blank's conical axis/apex."
    End If

    Set swBody = GetBodyFromFeature(blankFeat)
    If swBody Is Nothing Then
        Err.Raise vbObjectError + 203, , prefix & ": could not obtain the solid body from the blank."
    End If

    Dim bodyName As String
    bodyName = prefix & "_Body"

    If swModel.FeatureManager.IsNameUsed(swBodyName, bodyName) Then
        Err.Raise vbObjectError + 211, , _
            prefix & ": solid-body name '" & bodyName & "' is already in use. Use a fresh part for this test."
    End If

    swBody.Name = bodyName

    If StrComp(swBody.Name, bodyName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 212, , prefix & ": SOLIDWORKS did not retain the requested body name."
    End If

    swModel.FeatureManager.UpdateFeatureTree

    Set axisFeat = CreateAxisFromConicalFace(blankFeat, prefix & "_Axis")
    If axisFeat Is Nothing Then
        Err.Raise vbObjectError + 204, , prefix & ": failed to create gear axis."
    End If

    If bore > 0# Then
        Set skLargeBore = CreateCircleSketch(baseRef, centerX, bore / 2#, prefix & "_BoreLarge")
        Set skSmallBore = CreateCircleSketch(innerPlane, centerX, bore / 2#, prefix & "_BoreSmall")

        SelectSketchProfile skLargeBore, False
        SelectSketchProfile skSmallBore, True

        Set boreFeat = swModel.FeatureManager.InsertCutBlend( _
                        False, False, False, 1#, _
                        0, 0, False, 0#, 0#, 0, _
                        False, True)

        If boreFeat Is Nothing Then
            Err.Raise vbObjectError + 207, , prefix & ": bore lofted cut failed."
        End If
        boreFeat.Name = prefix & "_Bore"
    End If

    Set skLargeTooth = CreateToothSpaceSketch( _
                            baseRef, centerX, m, z, pressure, _
                            1#, phaseOffset, _
                            prefix & "_SpaceLarge")

    Set skSmallTooth = CreateToothSpaceSketch( _
                            innerPlane, centerX, m, z, pressure, _
                            scaleSmall, phaseOffset, _
                            prefix & "_SpaceSmall")

    SelectSketchProfile skLargeTooth, False
    SelectSketchProfile skSmallTooth, True

    Set cutFeat = swModel.FeatureManager.InsertCutBlend( _
                    False, False, False, 1#, _
                    0, 0, False, 0#, 0#, 0, _
                    False, True)

    If cutFeat Is Nothing Then
        Err.Raise vbObjectError + 205, , prefix & ": seed lofted tooth-space cut failed."
    End If
    cutFeat.Name = prefix & "_ToothSpaceCut"

    ' SOLIDWORKS 2025 documented VBA circular-pattern sequence:
    ' seed feature = mark 4
    ' direction axis = mark 1
    ' then FeatureCircularPattern4.
    swModel.ClearSelection2 True

    SelectFeatureWithMark cutFeat, False, 4, prefix & " tooth-space cut"
    SelectFeatureWithMark axisFeat, True, 1, prefix & " pattern axis"

    Set patFeat = swModel.FeatureManager.FeatureCircularPattern4( _
                    z, 2# * PI, False, "NULL", _
                    False, True, False)

    If patFeat Is Nothing Then
        Err.Raise vbObjectError + 206, , _
            prefix & ": SOLIDWORKS failed to create the circular tooth pattern."
    End If

    patFeat.Name = prefix & "_ToothSpaces"

    swModel.EditRebuild3

    ' Do not return the pre-cut Body2 pointer captured from the blank.
    ' Reacquire the final body after the bore, tooth cut and pattern.
    Set swBody = GetSolidBodyByName(bodyName)

    If swBody Is Nothing Then
        Err.Raise vbObjectError + 213, , _
            prefix & ": could not reacquire final solid body '" & bodyName & "'."
    End If

    Set BuildGear = swBody
    Exit Function

EH:
    MsgBox Err.Description, vbCritical, prefix & " generation error"
    Set BuildGear = Nothing
End Function

Private Function CreateCircleSketch( _
    ByVal refObj As Object, _
    ByVal centerX As Double, _
    ByVal radius As Double, _
    ByVal featureName As String) As SldWorks.Feature

    Dim swSketch As SldWorks.Sketch
    Dim swFeat As SldWorks.Feature
    Dim swSeg As SldWorks.SketchSegment

    StartSketchOn refObj

    swModel.SketchManager.AddToDB = True
    swModel.SketchManager.DisplayWhenAdded = False

    Set swSeg = swModel.SketchManager.CreateCircleByRadius(centerX, 0#, 0#, radius)
    If swSeg Is Nothing Then
        swModel.SketchManager.AddToDB = False
        swModel.SketchManager.DisplayWhenAdded = True
        Err.Raise vbObjectError + 310, , "Failed to create circle in " & featureName & "."
    End If

    Set swSketch = swModel.GetActiveSketch2
    If swSketch Is Nothing Then
        swModel.SketchManager.AddToDB = False
        swModel.SketchManager.DisplayWhenAdded = True
        Err.Raise vbObjectError + 311, , "No active sketch while creating " & featureName & "."
    End If

    Set swFeat = swSketch
    swFeat.Name = featureName

    swModel.SketchManager.AddToDB = False
    swModel.SketchManager.DisplayWhenAdded = True
    swModel.SketchManager.InsertSketch True

    Set CreateCircleSketch = swFeat
End Function

Private Function CreateToothSpaceSketch( _
    ByVal refObj As Object, _
    ByVal centerX As Double, _
    ByVal m As Double, _
    ByVal z As Long, _
    ByVal pressure As Double, _
    ByVal radialScale As Double, _
    ByVal phaseOffset As Double, _
    ByVal featureName As String) As SldWorks.Feature

    Dim xs() As Double
    Dim ys() As Double
    Dim n As Long
    Dim i As Long, j As Long
    Dim swSketch As SldWorks.Sketch
    Dim swFeat As SldWorks.Feature

    BuildToothSpacePoints m, z, pressure, radialScale, centerX, phaseOffset, xs, ys, n

    StartSketchOn refObj
    swModel.SketchManager.AddToDB = True
    swModel.SketchManager.DisplayWhenAdded = False

    For i = 0 To n - 1
        j = i + 1
        If j >= n Then j = 0
        swModel.SketchManager.CreateLine xs(i), ys(i), 0#, xs(j), ys(j), 0#
    Next i

    Set swSketch = swModel.GetActiveSketch2
    If swSketch Is Nothing Then
        swModel.SketchManager.AddToDB = False
        swModel.SketchManager.DisplayWhenAdded = True
        Err.Raise vbObjectError + 312, , "No active sketch while creating " & featureName & "."
    End If
    Set swFeat = swSketch
    swFeat.Name = featureName

    swModel.SketchManager.AddToDB = False
    swModel.SketchManager.DisplayWhenAdded = True
    swModel.SketchManager.InsertSketch True

    Set CreateToothSpaceSketch = swFeat
End Function

Private Sub BuildToothSpacePoints( _
    ByVal m As Double, _
    ByVal z As Long, _
    ByVal pressure As Double, _
    ByVal radialScale As Double, _
    ByVal centerX As Double, _
    ByVal phaseOffset As Double, _
    ByRef xs() As Double, _
    ByRef ys() As Double, _
    ByRef n As Long)

    Dim rp As Double, ra As Double, rf As Double, rb As Double
    Dim invPitch As Double, tPitch As Double
    Dim spaceCenter As Double
    Dim rStart As Double
    Dim spaceRoot As Double, spaceTip As Double
    Dim r As Double, spaceHalf As Double
    Dim rOut As Double
    Dim i As Long
    Dim a As Double

    ReDim xs(0 To 100)
    ReDim ys(0 To 100)
    n = 0

    rp = m * z / 2#
    ra = rp + m
    rf = rp - 1.25 * m
    If rf < 0.05 * m Then rf = 0.05 * m

    rb = rp * Cos(pressure)

    tPitch = Tan(pressure)
    invPitch = tPitch - pressure

    ' Tooth centres repeat every 2*pi/z.
    ' The space between tooth 0 and tooth 1 is centred halfway between them.
    spaceCenter = (PI / z) + phaseOffset

    If rf > rb Then
        rStart = rf
        spaceRoot = SpaceHalfAtRadius(rf, rb, invPitch, z)
    Else
        rStart = rb
        spaceRoot = SpaceHalfAtRadius(rb, rb, invPitch, z)
    End If

    ' One root-side boundary of the tooth SPACE.
    AddPolarPoint xs, ys, n, centerX, _
                  rf * radialScale, _
                  spaceCenter + spaceRoot

    ' Below the base circle there is no true involute.
    ' Continue the base-circle flank angle radially down to the root circle.
    If rf < rb Then
        AddPolarPoint xs, ys, n, centerX, _
                      rb * radialScale, _
                      spaceCenter + spaceRoot
    End If

    ' First involute boundary of the tooth space.
    For i = 0 To FLANK_SAMPLES
        r = rStart + (ra - rStart) * (CDbl(i) / CDbl(FLANK_SAMPLES))
        spaceHalf = SpaceHalfAtRadius(r, rb, invPitch, z)

        AddPolarPoint xs, ys, n, centerX, _
                      r * radialScale, _
                      spaceCenter + spaceHalf
    Next i

    spaceTip = SpaceHalfAtRadius(ra, rb, invPitch, z)

    ' Put the closing chord outside the addendum circle so the cut opens
    ' completely through the outside of the blank.
    rOut = (ra * 1.025) / Cos(spaceTip)

    AddPolarPoint xs, ys, n, centerX, _
                  rOut * radialScale, _
                  spaceCenter + spaceTip

    AddPolarPoint xs, ys, n, centerX, _
                  rOut * radialScale, _
                  spaceCenter - spaceTip

    ' Opposite involute boundary, returning toward the root.
    For i = FLANK_SAMPLES To 0 Step -1
        r = rStart + (ra - rStart) * (CDbl(i) / CDbl(FLANK_SAMPLES))
        spaceHalf = SpaceHalfAtRadius(r, rb, invPitch, z)

        AddPolarPoint xs, ys, n, centerX, _
                      r * radialScale, _
                      spaceCenter - spaceHalf
    Next i

    If rf < rb Then
        AddPolarPoint xs, ys, n, centerX, _
                      rb * radialScale, _
                      spaceCenter - spaceRoot
    End If

    AddPolarPoint xs, ys, n, centerX, _
                  rf * radialScale, _
                  spaceCenter - spaceRoot

    ' Root-circle arc across the bottom of the tooth space.
    For i = 1 To ROOT_ARC_SAMPLES - 1
        a = (spaceCenter - spaceRoot) + _
            (2# * spaceRoot) * (CDbl(i) / CDbl(ROOT_ARC_SAMPLES))

        AddPolarPoint xs, ys, n, centerX, _
                      rf * radialScale, _
                      a
    Next i

    ReDim Preserve xs(0 To n - 1)
    ReDim Preserve ys(0 To n - 1)
End Sub

Private Function SpaceHalfAtRadius( _
    ByVal r As Double, _
    ByVal rb As Double, _
    ByVal invPitch As Double, _
    ByVal z As Long) As Double

    Dim t As Double
    Dim invR As Double
    Dim toothHalf As Double
    Dim spaceHalf As Double
    Dim minAngle As Double
    Dim maxAngle As Double

    If r <= rb Then
        invR = 0#
    Else
        t = Sqr((r / rb) * (r / rb) - 1#)
        invR = t - Atn(t)
    End If

    ' Standard involute angular half-thickness of ONE TOOTH:
    '   toothHalf = pi/(2z) + inv(alpha_pitch) - inv(alpha_r)
    toothHalf = (PI / (2# * z)) + invPitch - invR

    ' Tooth centre to neighbouring tooth centre is 2*pi/z.
    ' Half the intervening SPACE is therefore:
    spaceHalf = (PI / z) - toothHalf

    minAngle = DegToRad(0.15)
    maxAngle = (PI / z) - minAngle

    If spaceHalf < minAngle Then spaceHalf = minAngle
    If spaceHalf > maxAngle Then spaceHalf = maxAngle

    SpaceHalfAtRadius = spaceHalf
End Function

Private Sub AddPolarPoint( _
    ByRef xs() As Double, _
    ByRef ys() As Double, _
    ByRef n As Long, _
    ByVal centerX As Double, _
    ByVal r As Double, _
    ByVal a As Double)

    xs(n) = centerX + r * Cos(a)
    ys(n) = r * Sin(a)
    n = n + 1
End Sub

Private Sub StartSketchOn(ByVal refObj As Object)
    Dim swSketch As SldWorks.Sketch

    swModel.ClearSelection2 True
    SelectReference refObj, False, 0

    swModel.SketchManager.InsertSketch True
    Set swSketch = swModel.GetActiveSketch2

    If swSketch Is Nothing Then
        Err.Raise vbObjectError + 300, , "Could not start sketch on selected plane/face."
    End If
End Sub

Private Function CreateOffsetPlane( _
    ByVal refObj As Object, _
    ByVal offsetDistance As Double, _
    ByVal newName As String) As SldWorks.Feature

    Dim swRefPlane As SldWorks.RefPlane
    Dim swFeat As SldWorks.Feature

    swModel.ClearSelection2 True
    SelectReference refObj, False, 0

    Set swRefPlane = swModel.FeatureManager.InsertRefPlane(8, offsetDistance, 0, 0#, 0, 0#)

    If swRefPlane Is Nothing Then
        Set CreateOffsetPlane = Nothing
        Exit Function
    End If

    Set swFeat = swRefPlane
    swFeat.Name = newName
    swModel.ClearSelection2 True

    Set CreateOffsetPlane = swFeat
End Function

Private Function CreateAxisFromConicalFace( _
    ByVal sourceFeat As SldWorks.Feature, _
    ByVal newName As String) As SldWorks.Feature

    Dim vFaces As Variant
    Dim i As Long
    Dim swFace As SldWorks.Face2
    Dim swSurf As SldWorks.Surface
    Dim swAxisFeat As SldWorks.Feature
    Dim ok As Boolean

    vFaces = sourceFeat.GetFaces
    If IsEmpty(vFaces) Then
        Set CreateAxisFromConicalFace = Nothing
        Exit Function
    End If

    For i = LBound(vFaces) To UBound(vFaces)
        Set swFace = vFaces(i)
        Set swSurf = swFace.GetSurface

        If Not swSurf Is Nothing Then
            If swSurf.IsCone Then
                swModel.ClearSelection2 True
                ok = swFace.Select4(False, Nothing)
                If ok Then
                    ok = swModel.InsertAxis2(True)
                    If ok Then
                        Set swAxisFeat = FindNewestFeatureByType("RefAxis")
                        If Not swAxisFeat Is Nothing Then swAxisFeat.Name = newName
                        Set CreateAxisFromConicalFace = swAxisFeat
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    Set CreateAxisFromConicalFace = Nothing
End Function

Private Function GetBodyFromFeature(ByVal swFeat As SldWorks.Feature) As SldWorks.Body2
    Dim vFaces As Variant
    Dim swFace As SldWorks.Face2

    vFaces = swFeat.GetFaces
    If IsEmpty(vFaces) Then
        Set GetBodyFromFeature = Nothing
        Exit Function
    End If

    Set swFace = vFaces(LBound(vFaces))
    Set GetBodyFromFeature = swFace.GetBody
End Function

Private Sub PositionMatingGearDefaultBase( _
    ByVal body1Name As String, _
    ByVal body2Name As String, _
    ByVal sigma As Double, _
    ByVal apex1X As Double, _
    ByVal apex1Y As Double, _
    ByVal apex1Z As Double, _
    ByVal axis1X As Double, _
    ByVal axis1Y As Double, _
    ByVal axis1Z As Double, _
    ByVal apex2X As Double, _
    ByVal apex2Y As Double, _
    ByVal apex2Z As Double, _
    ByVal axis2X As Double, _
    ByVal axis2Y As Double, _
    ByVal axis2Z As Double)

    Dim dx As Double, dy As Double, dz As Double
    Dim shaftRX As Double, shaftRY As Double, shaftRZ As Double
    Dim actualX As Double, actualY As Double, actualZ As Double
    Dim measuredAngle As Double
    Dim moveFeat As SldWorks.Feature

    ' Both blank axes came from the untouched conical blanks before any
    ' tooth geometry existed. They should therefore be parallel.
    If AxisLineAngle( _
            axis1X, axis1Y, axis1Z, _
            axis2X, axis2Y, axis2Z) > DegToRad(0.05) Then

        Err.Raise vbObjectError + 392, , _
            "The two clean bevel-gear blank axes are not parallel before pair placement." & vbCrLf & _
            "Gear 1 blank axis = (" & Format(axis1X, "0.000000") & ", " & Format(axis1Y, "0.000000") & ", " & Format(axis1Z, "0.000000") & ")" & vbCrLf & _
            "Gear 2 blank axis = (" & Format(axis2X, "0.000000") & ", " & Format(axis2Y, "0.000000") & ", " & Format(axis2Z, "0.000000") & ")"
    End If

    ' ------------------------------------------------------------
    ' 1) Shaft-angle rotation only.
    '
    ' Gear 2 was ALREADY generated with the required half-tooth mesh
    ' phase in its sketches, so no phase Move/Copy body feature exists.
    '
    ' Rotate about Gear 2's clean-blank apex. Because the starting axis
    ' is aligned to one of the standard origin-plane normals, choose one
    ' perpendicular global principal axis.
    ' ------------------------------------------------------------
    PerpendicularShaftRotationComponents _
        axis2X, axis2Y, axis2Z, _
        sigma, _
        shaftRX, shaftRY, shaftRZ

    Set moveFeat = CreateMoveBodyRotationFeature( _
                        body2Name, _
                        apex2X, apex2Y, apex2Z, _
                        shaftRX, shaftRY, shaftRZ, _
                        "BG2_ShaftAngle")

    If moveFeat Is Nothing Then
        Err.Raise vbObjectError + 400, , _
            "Could not rotate mating gear 2 to the requested shaft angle."
    End If

    ' Verify the ACTUAL solid-body shaft direction using a planar end face.
    ' This avoids using tooth-generated conical surfaces as an axis detector.
    If Not GetBodyAxisFromLargestPlanarFace( _
                body2Name, actualX, actualY, actualZ) Then

        Err.Raise vbObjectError + 394, , _
            "Could not measure Gear 2's shaft axis from its planar end face after rotation."
    End If

    measuredAngle = AxisLineAngle( _
                        axis1X, axis1Y, axis1Z, _
                        actualX, actualY, actualZ)

    If Abs(measuredAngle - sigma) > DegToRad(0.05) Then
        Err.Raise vbObjectError + 395, , _
            "Gear 2 rotation feature was created, but the measured shaft-axis angle is not the requested value." & vbCrLf & _
            "Requested: " & Format(RadToDeg(sigma), "0.000") & " deg" & vbCrLf & _
            "Measured: " & Format(RadToDeg(measuredAngle), "0.000") & " deg" & vbCrLf & _
            "Gear 1 clean-blank axis = (" & Format(axis1X, "0.000000") & ", " & Format(axis1Y, "0.000000") & ", " & Format(axis1Z, "0.000000") & ")" & vbCrLf & _
            "Gear 2 measured axis = (" & Format(actualX, "0.000000") & ", " & Format(actualY, "0.000000") & ", " & Format(actualZ, "0.000000") & ")" & vbCrLf & _
            MoveCopyFeatureDiagnostic(moveFeat)
    End If

    ' ------------------------------------------------------------
    ' 2) Apex-to-apex translation.
    '
    ' The shaft rotation was ABOUT apex2, so apex2 did not move during
    ' the rotation. Move it directly onto Gear 1's saved clean-blank apex.
    ' ------------------------------------------------------------
    dx = apex1X - apex2X
    dy = apex1Y - apex2Y
    dz = apex1Z - apex2Z

    Set moveFeat = CreateMoveBodyTranslationFeature( _
                        body2Name, _
                        dx, dy, dz, _
                        "BG2_MeshPosition")

    If moveFeat Is Nothing Then
        Err.Raise vbObjectError + 401, , _
            "Could not translate Gear 2's bevel apex onto Gear 1's bevel apex."
    End If

    ' Translation cannot change the shaft angle, but verify the actual
    ' body orientation one final time.
    If Not GetBodyAxisFromLargestPlanarFace( _
                body2Name, actualX, actualY, actualZ) Then

        Err.Raise vbObjectError + 396, , _
            "Could not verify Gear 2's final shaft axis."
    End If

    measuredAngle = AxisLineAngle( _
                        axis1X, axis1Y, axis1Z, _
                        actualX, actualY, actualZ)

    If Abs(measuredAngle - sigma) > DegToRad(0.05) Then
        Err.Raise vbObjectError + 397, , _
            "Final pair shaft-angle verification failed." & vbCrLf & _
            "Requested: " & Format(RadToDeg(sigma), "0.000") & " deg" & vbCrLf & _
            "Measured: " & Format(RadToDeg(measuredAngle), "0.000") & " deg"
    End If
End Sub

Private Function GetCleanBlankConeApexAxis( _
    ByVal blankFeat As SldWorks.Feature, _
    ByRef apexX As Double, _
    ByRef apexY As Double, _
    ByRef apexZ As Double, _
    ByRef axisX As Double, _
    ByRef axisY As Double, _
    ByRef axisZ As Double) As Boolean

    Dim vFaces As Variant
    Dim vCone As Variant
    Dim swFace As SldWorks.Face2
    Dim swSurf As SldWorks.Surface

    Dim i As Long
    Dim bestArea As Double
    Dim faceArea As Double
    Dim swBestSurf As SldWorks.Surface

    Dim axisNorm As Double
    Dim halfAngle As Double
    Dim apexDistance As Double

    Dim bestPlanarArea As Double
    Dim planarArea As Double
    Dim swBestPlanarFace As SldWorks.Face2
    Dim vPlanarNormal As Variant
    Dim pnX As Double, pnY As Double, pnZ As Double
    Dim pnNorm As Double

    vFaces = blankFeat.GetFaces
    If IsEmpty(vFaces) Then Exit Function

    bestArea = -1#
    bestPlanarArea = -1#

    ' The clean blank has one outer conical side face and two planar end
    ' faces. Capture BOTH independently, then verify that the cone axis is
    ' parallel to the planar-face normal. This prevents any ambiguous
    ' conical-surface interpretation from becoming the gear shaft axis.
    For i = LBound(vFaces) To UBound(vFaces)
        Set swFace = vFaces(i)
        Set swSurf = swFace.GetSurface

        If Not swSurf Is Nothing Then
            If swSurf.IsCone Then
                faceArea = swFace.GetArea

                If faceArea > bestArea Then
                    bestArea = faceArea
                    Set swBestSurf = swSurf
                End If
            ElseIf swSurf.IsPlane Then
                planarArea = swFace.GetArea

                If planarArea > bestPlanarArea Then
                    bestPlanarArea = planarArea
                    Set swBestPlanarFace = swFace
                End If
            End If
        End If
    Next i

    If swBestSurf Is Nothing Then Exit Function

    ' Same SOLIDWORKS cone-parameter interpretation used by the official
    ' "Locate Apex of Conical Face" VBA example.
    vCone = swBestSurf.ConeParams

    axisX = CDbl(vCone(3))
    axisY = CDbl(vCone(4))
    axisZ = CDbl(vCone(5))

    axisNorm = Sqr(axisX * axisX + _
                   axisY * axisY + _
                   axisZ * axisZ)

    If axisNorm <= 0# Then Exit Function

    axisX = axisX / axisNorm
    axisY = axisY / axisNorm
    axisZ = axisZ / axisNorm

    halfAngle = CDbl(vCone(7))
    If Abs(Tan(halfAngle)) < 0.000000001 Then Exit Function

    apexDistance = CDbl(vCone(6)) / Tan(halfAngle)

    apexX = CDbl(vCone(0)) + axisX * apexDistance
    apexY = CDbl(vCone(1)) + axisY * apexDistance
    apexZ = CDbl(vCone(2)) + axisZ * apexDistance

    If swBestPlanarFace Is Nothing Then Exit Function

    vPlanarNormal = swBestPlanarFace.Normal
    pnX = CDbl(vPlanarNormal(0))
    pnY = CDbl(vPlanarNormal(1))
    pnZ = CDbl(vPlanarNormal(2))

    pnNorm = Sqr(pnX * pnX + pnY * pnY + pnZ * pnZ)
    If pnNorm <= 0# Then Exit Function

    pnX = pnX / pnNorm
    pnY = pnY / pnNorm
    pnZ = pnZ / pnNorm

    If AxisLineAngle(axisX, axisY, axisZ, pnX, pnY, pnZ) > DegToRad(0.05) Then
        Err.Raise vbObjectError + 215, , _
            "Clean blank cone axis does not agree with its planar end-face normal." & vbCrLf & _
            "Cone axis = (" & Format(axisX, "0.000000") & ", " & Format(axisY, "0.000000") & ", " & Format(axisZ, "0.000000") & ")" & vbCrLf & _
            "End-face normal = (" & Format(pnX, "0.000000") & ", " & Format(pnY, "0.000000") & ", " & Format(pnZ, "0.000000") & ")"
    End If

    ' Use the planar face normal as the canonical shaft direction. The
    ' cone is retained only for the apex calculation.
    axisX = pnX
    axisY = pnY
    axisZ = pnZ

    GetCleanBlankConeApexAxis = True
End Function

Private Function GetBodyAxisFromLargestPlanarFace( _
    ByVal bodyName As String, _
    ByRef axisX As Double, _
    ByRef axisY As Double, _
    ByRef axisZ As Double) As Boolean

    Dim swBody As SldWorks.Body2
    Dim swFace As SldWorks.Face2
    Dim swSurf As SldWorks.Surface
    Dim vFaces As Variant
    Dim vNormal As Variant

    Dim i As Long
    Dim faceArea As Double
    Dim bestArea As Double
    Dim swBestFace As SldWorks.Face2
    Dim n As Double

    Set swBody = GetSolidBodyByName(bodyName)
    If swBody Is Nothing Then Exit Function

    vFaces = swBody.GetFaces
    If IsEmpty(vFaces) Then Exit Function

    bestArea = -1#

    ' The large and small end faces remain planar after the tooth-space
    ' cuts. The largest planar face is a stable representation of the
    ' gear's physical shaft-normal direction.
    For i = LBound(vFaces) To UBound(vFaces)
        Set swFace = vFaces(i)
        Set swSurf = swFace.GetSurface

        If Not swSurf Is Nothing Then
            If swSurf.IsPlane Then
                faceArea = swFace.GetArea

                If faceArea > bestArea Then
                    bestArea = faceArea
                    Set swBestFace = swFace
                End If
            End If
        End If
    Next i

    If swBestFace Is Nothing Then Exit Function

    vNormal = swBestFace.Normal

    axisX = CDbl(vNormal(0))
    axisY = CDbl(vNormal(1))
    axisZ = CDbl(vNormal(2))

    n = Sqr(axisX * axisX + axisY * axisY + axisZ * axisZ)
    If n <= 0# Then Exit Function

    axisX = axisX / n
    axisY = axisY / n
    axisZ = axisZ / n

    GetBodyAxisFromLargestPlanarFace = True
End Function

Private Sub PerpendicularShaftRotationComponents( _
    ByVal ux As Double, _
    ByVal uy As Double, _
    ByVal uz As Double, _
    ByVal angle As Double, _
    ByRef rotX As Double, _
    ByRef rotY As Double, _
    ByRef rotZ As Double)

    Dim tol As Double
    tol = 0.9999

    rotX = 0#
    rotY = 0#
    rotZ = 0#

    ' The clean blank is generated normal to one of SOLIDWORKS'
    ' standard origin planes, so its initial axis is a global
    ' principal axis. Choose a different principal axis for the
    ' shaft-angle rotation.
    If Abs(ux) >= tol And Abs(uy) < 0.01 And Abs(uz) < 0.01 Then
        ' Initial shaft along X -> swing it about global Z.
        rotZ = angle
        Exit Sub
    End If

    If Abs(uy) >= tol And Abs(ux) < 0.01 And Abs(uz) < 0.01 Then
        ' Initial shaft along Y -> swing it about global X.
        rotX = angle
        Exit Sub
    End If

    If Abs(uz) >= tol And Abs(ux) < 0.01 And Abs(uy) < 0.01 Then
        ' Initial shaft along Z -> swing it about global X.
        rotX = angle
        Exit Sub
    End If

    Err.Raise vbObjectError + 403, , _
        "Could not choose a principal-axis shaft rotation for measured blank axis (" & _
        Format(ux, "0.000000") & ", " & _
        Format(uy, "0.000000") & ", " & _
        Format(uz, "0.000000") & ")."
End Sub

Private Function AxisLineAngle( _
    ByVal ax As Double, _
    ByVal ay As Double, _
    ByVal az As Double, _
    ByVal bx As Double, _
    ByVal by As Double, _
    ByVal bz As Double) As Double

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim crossNorm As Double
    Dim dotVal As Double

    cx = ay * bz - az * by
    cy = az * bx - ax * bz
    cz = ax * by - ay * bx

    crossNorm = Sqr(cx * cx + cy * cy + cz * cz)
    dotVal = ax * bx + ay * by + az * bz

    ' A shaft axis is an undirected line:
    ' +axis and -axis describe the same physical axis.
    AxisLineAngle = Atan2(crossNorm, Abs(dotVal))
End Function

Private Function VectorAngle( _
    ByVal ax As Double, _
    ByVal ay As Double, _
    ByVal az As Double, _
    ByVal bx As Double, _
    ByVal by As Double, _
    ByVal bz As Double) As Double

    Dim cx As Double, cy As Double, cz As Double
    Dim crossNorm As Double
    Dim dotVal As Double

    cx = ay * bz - az * by
    cy = az * bx - ax * bz
    cz = ax * by - ay * bx

    crossNorm = Sqr(cx * cx + cy * cy + cz * cz)
    dotVal = ax * bx + ay * by + az * bz

    VectorAngle = Atan2(crossNorm, dotVal)
End Function

Private Sub HideRefFeatureByName(ByVal featureName As String)
    Dim swFeat As SldWorks.Feature

    Set swFeat = swModel.FeatureByName(featureName)
    If swFeat Is Nothing Then Exit Sub

    swModel.ClearSelection2 True

    ' Hiding is cosmetic only; do not fail gear generation if SOLIDWORKS
    ' chooses not to hide a particular reference feature.
    On Error Resume Next
    swFeat.Select2 False, 0
    swModel.BlankRefGeom
    swModel.ClearSelection2 True
    On Error GoTo 0
End Sub

Private Sub ValidateBore( _
    ByVal m As Double, _
    ByVal z As Long, _
    ByVal delta As Double, _
    ByVal faceW As Double, _
    ByVal bore As Double, _
    ByVal gearName As String)

    Dim rp As Double, rf As Double, R As Double
    Dim scaleSmall As Double
    Dim smallRootDia As Double

    If bore <= 0# Then Exit Sub

    rp = m * z / 2#
    rf = rp - 1.25 * m
    R = rp / Sin(delta)
    scaleSmall = (R - faceW) / R
    smallRootDia = 2# * rf * scaleSmall

    If bore >= smallRootDia Then
        Err.Raise vbObjectError + 500, , _
            gearName & " bore (" & Format(MToMm(bore), "0.###") & " mm) is too large." & vbCrLf & _
            "Approximate small-end root diameter is " & Format(MToMm(smallRootDia), "0.###") & " mm."
    End If
End Sub

Private Sub SelectFeatureWithMark( _
    ByVal swFeat As SldWorks.Feature, _
    ByVal append As Boolean, _
    ByVal mark As Long, _
    ByVal description As String)

    Dim swSelMgr As SldWorks.SelectionMgr
    Dim beforeCount As Long
    Dim afterCount As Long
    Dim selectName As String
    Dim selectType As String
    Dim byIdResult As Boolean
    Dim directResult As Boolean

    If swFeat Is Nothing Then
        Err.Raise vbObjectError + 620, , _
            "Cannot select " & description & ": feature reference is Nothing."
    End If

    Set swSelMgr = swModel.SelectionManager

    If Not append Then
        swModel.ClearSelection2 True
    End If

    beforeCount = swSelMgr.GetSelectedObjectCount2(mark)

    ' Preferred named selection path.
    selectName = ""
    selectType = ""
    selectName = swFeat.GetNameForSelection(selectType)

    If Len(selectName) > 0 And Len(selectType) > 0 Then
        byIdResult = swModel.Extension.SelectByID2( _
                        selectName, selectType, _
                        0#, 0#, 0#, _
                        append, mark, Nothing, swSelectOptionDefault)
    End If

    afterCount = swSelMgr.GetSelectedObjectCount2(mark)
    If afterCount > beforeCount Then Exit Sub

    ' SOLIDWORKS also documents IFeature.Select2 for directly selecting
    ' a feature or reference axis with the ordered-selection mark.
    directResult = swFeat.Select2(append, mark)

    afterCount = swSelMgr.GetSelectedObjectCount2(mark)
    If afterCount > beforeCount Then Exit Sub

    Err.Raise vbObjectError + 621, , _
        "Could not add " & description & " to the SOLIDWORKS selection set." & vbCrLf & _
        "Feature: " & swFeat.Name & vbCrLf & _
        "Selection name: " & selectName & vbCrLf & _
        "Selection type: " & selectType & vbCrLf & _
        "Required mark: " & CStr(mark) & vbCrLf & _
        "SelectByID2 returned: " & CStr(byIdResult) & vbCrLf & _
        "Feature.Select2 returned: " & CStr(directResult) & vbCrLf & _
        "Marked count before: " & CStr(beforeCount) & vbCrLf & _
        "Marked count after: " & CStr(afterCount)
End Sub

Private Sub SelectSketchProfile(ByVal swFeat As SldWorks.Feature, ByVal append As Boolean)
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim beforeCount As Long
    Dim afterCount As Long
    Dim selectName As String
    Dim selectType As String
    Dim apiResult As Boolean
    Dim fallbackResult As Boolean

    If swFeat Is Nothing Then
        Err.Raise vbObjectError + 610, , "A loft-profile feature reference is Nothing."
    End If

    Set swSelMgr = swModel.SelectionManager

    If Not append Then
        swModel.ClearSelection2 True
    End If

    ' What matters to a loft is the actual selection set and selection mark.
    ' Record how many mark-1 profile objects exist before adding this sketch.
    beforeCount = swSelMgr.GetSelectedObjectCount2(1)

    ' Use SOLIDWORKS' selectable name/type rather than IFeature.Name directly.
    selectName = ""
    selectType = ""
    selectName = swFeat.GetNameForSelection(selectType)

    If Len(selectName) > 0 And Len(selectType) > 0 Then
        apiResult = swModel.Extension.SelectByID2( _
                        selectName, selectType, _
                        0#, 0#, 0#, _
                        append, 1, Nothing, swSelectOptionDefault)
    Else
        apiResult = False
    End If

    afterCount = swSelMgr.GetSelectedObjectCount2(1)

    ' IMPORTANT:
    ' Do not reject a valid selection merely because SelectByID2 returned False.
    ' If the mark-1 selection count increased, SOLIDWORKS actually selected it.
    If afterCount > beforeCount Then Exit Sub

    ' Fallback: select the Feature object directly with loft-profile mark 1,
    ' then verify the actual SelectionManager state again.
    fallbackResult = swFeat.Select2(append, 1)
    afterCount = swSelMgr.GetSelectedObjectCount2(1)

    If afterCount > beforeCount Then Exit Sub

    Err.Raise vbObjectError + 612, , _
        "Could not add loft profile to the actual SOLIDWORKS selection set." & vbCrLf & _
        "Feature tree name: " & swFeat.Name & vbCrLf & _
        "GetNameForSelection name: " & selectName & vbCrLf & _
        "GetNameForSelection type: " & selectType & vbCrLf & _
        "SelectByID2 returned: " & CStr(apiResult) & vbCrLf & _
        "Feature.Select2 returned: " & CStr(fallbackResult) & vbCrLf & _
        "Mark-1 count before: " & CStr(beforeCount) & vbCrLf & _
        "Mark-1 count after: " & CStr(afterCount)
End Sub

Private Function GetSolidBodyByName(ByVal bodyName As String) As SldWorks.Body2
    Dim vBodies As Variant
    Dim i As Long
    Dim swBody As SldWorks.Body2

    vBodies = swPart.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then
        Set GetSolidBodyByName = Nothing
        Exit Function
    End If

    For i = LBound(vBodies) To UBound(vBodies)
        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then
            If StrComp(swBody.Name, bodyName, vbTextCompare) = 0 Then
                Set GetSolidBodyByName = swBody
                Exit Function
            End If
        End If
    Next i

    Set GetSolidBodyByName = Nothing
End Function

Private Function SolidBodyNames() As String
    Dim vBodies As Variant
    Dim i As Long
    Dim swBody As SldWorks.Body2
    Dim s As String

    vBodies = swPart.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then
        SolidBodyNames = "<none>"
        Exit Function
    End If

    For i = LBound(vBodies) To UBound(vBodies)
        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then
            If Len(s) > 0 Then s = s & ", "
            s = s & swBody.Name
        End If
    Next i

    SolidBodyNames = s
End Function

Private Function CreateMoveBodyRotationFeature( _
    ByVal bodyName As String, _
    ByVal originX As Double, _
    ByVal originY As Double, _
    ByVal originZ As Double, _
    ByVal rotX As Double, _
    ByVal rotY As Double, _
    ByVal rotZ As Double, _
    ByVal featureName As String) As SldWorks.Feature

    On Error GoTo EH

    Dim swFeat As SldWorks.Feature
    Dim moveData As SldWorks.MoveCopyBodyFeatureData
    Dim ok As Boolean
    Dim stage As String

    ' The body selection already defines the body to move.
    stage = "selecting body"
    SelectBodyForMove bodyName

    ' Insert the feature using the documented rotation-only argument set.
    stage = "inserting Move/Copy Body rotation feature"
    Set swFeat = swModel.FeatureManager.InsertMoveCopyBody2( _
                    0#, 0#, 0#, 0#, _
                    originX, originY, originZ, _
                    rotX, rotY, rotZ, _
                    False, 1)

    If swFeat Is Nothing Then
        Err.Raise vbObjectError + 630, , _
            "InsertMoveCopyBody2 returned Nothing for the rotation feature."
    End If

    ' SOLIDWORKS documented edit sequence:
    ' GetDefinition -> AccessSelections -> set transform fields -> ModifyDefinition.
    stage = "getting Move/Copy feature definition"
    Set moveData = swFeat.GetDefinition

    If moveData Is Nothing Then
        Err.Raise vbObjectError + 631, , _
            "SOLIDWORKS created the Move/Copy feature but returned no feature definition."
    End If

    stage = "accessing Move/Copy selections"
    ok = moveData.AccessSelections(swModel, Nothing)

    If Not ok Then
        Err.Raise vbObjectError + 632, , _
            "AccessSelections returned False for the Move/Copy rotation feature."
    End If

    ' Do NOT reset Bodies, Copy, TransformReferenceEntity, or TranslateToVertex.
    ' The inserted feature already has the correct body selection and copy state.
    ' The official 'setting transforms' workflow only changes the transform data.
    stage = "setting rotation transform type"
    moveData.TransformType = _
        swMoveCopyBodyFeatureTransformType_e.swTransformType_Rotation

    stage = "setting rotation origin"
    moveData.RotationOriginX = originX
    moveData.RotationOriginY = originY
    moveData.RotationOriginZ = originZ

    stage = "setting rotation angles"
    moveData.TransformX = rotX
    moveData.TransformY = rotY
    moveData.TransformZ = rotZ

    stage = "committing rotation definition"
    ok = swFeat.ModifyDefinition(moveData, swModel, Nothing)

    If Not ok Then
        On Error Resume Next
        moveData.ReleaseSelectionAccess
        On Error GoTo EH

        Err.Raise vbObjectError + 633, , _
            "ModifyDefinition returned False for the explicit rotation definition."
    End If

    ' ModifyDefinition restores the rollback state. ReleaseSelectionAccess is
    ' harmless in the SOLIDWORKS examples, but do not let it hide a successful move.
    On Error Resume Next
    moveData.ReleaseSelectionAccess
    On Error GoTo EH

    swFeat.Name = featureName
    swModel.EditRebuild3

    Set CreateMoveBodyRotationFeature = swFeat
    Exit Function

EH:
    Err.Raise vbObjectError + 634, "CreateMoveBodyRotationFeature", _
        "Move/Copy rotation failed while " & stage & "." & vbCrLf & _
        "VBA error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CreateMoveBodyTranslationFeature( _
    ByVal bodyName As String, _
    ByVal dx As Double, _
    ByVal dy As Double, _
    ByVal dz As Double, _
    ByVal featureName As String) As SldWorks.Feature

    On Error GoTo EH

    Dim swFeat As SldWorks.Feature
    Dim moveData As SldWorks.MoveCopyBodyFeatureData
    Dim ok As Boolean
    Dim stage As String

    stage = "selecting body"
    SelectBodyForMove bodyName

    stage = "inserting Move/Copy Body translation feature"
    Set swFeat = swModel.FeatureManager.InsertMoveCopyBody2( _
                    dx, dy, dz, 0#, _
                    0#, 0#, 0#, _
                    0#, 0#, 0#, _
                    False, 1)

    If swFeat Is Nothing Then
        Err.Raise vbObjectError + 640, , _
            "InsertMoveCopyBody2 returned Nothing for the translation feature."
    End If

    stage = "getting Move/Copy feature definition"
    Set moveData = swFeat.GetDefinition

    If moveData Is Nothing Then
        Err.Raise vbObjectError + 641, , _
            "SOLIDWORKS created the translation feature but returned no feature definition."
    End If

    stage = "accessing Move/Copy selections"
    ok = moveData.AccessSelections(swModel, Nothing)

    If Not ok Then
        Err.Raise vbObjectError + 642, , _
            "AccessSelections returned False for the Move/Copy translation feature."
    End If

    stage = "setting translation transform type"
    moveData.TransformType = _
        swMoveCopyBodyFeatureTransformType_e.swTransformType_Translation

    stage = "setting translation vector"
    moveData.TransformX = dx
    moveData.TransformY = dy
    moveData.TransformZ = dz

    stage = "committing translation definition"
    ok = swFeat.ModifyDefinition(moveData, swModel, Nothing)

    If Not ok Then
        On Error Resume Next
        moveData.ReleaseSelectionAccess
        On Error GoTo EH

        Err.Raise vbObjectError + 643, , _
            "ModifyDefinition returned False for the explicit translation definition."
    End If

    On Error Resume Next
    moveData.ReleaseSelectionAccess
    On Error GoTo EH

    swFeat.Name = featureName
    swModel.EditRebuild3

    Set CreateMoveBodyTranslationFeature = swFeat
    Exit Function

EH:
    Err.Raise vbObjectError + 644, "CreateMoveBodyTranslationFeature", _
        "Move/Copy translation failed while " & stage & "." & vbCrLf & _
        "VBA error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MoveCopyFeatureDiagnostic( _
    ByVal swFeat As SldWorks.Feature) As String

    Dim moveData As SldWorks.MoveCopyBodyFeatureData

    On Error GoTo EH

    If swFeat Is Nothing Then
        MoveCopyFeatureDiagnostic = "Move/Copy feature data: <feature is Nothing>"
        Exit Function
    End If

    Set moveData = swFeat.GetDefinition

    If moveData Is Nothing Then
        MoveCopyFeatureDiagnostic = "Move/Copy feature data: <definition unavailable>"
        Exit Function
    End If

    MoveCopyFeatureDiagnostic = _
        "Move/Copy definition readback:" & vbCrLf & _
        "  TransformType = " & CStr(moveData.TransformType) & _
        " (0=None, 1=Translation, 2=Rotation)" & vbCrLf & _
        "  TransformXYZ = (" & _
            Format(moveData.TransformX, "0.000000") & ", " & _
            Format(moveData.TransformY, "0.000000") & ", " & _
            Format(moveData.TransformZ, "0.000000") & ")" & vbCrLf & _
        "  RotationOrigin = (" & _
            Format(moveData.RotationOriginX, "0.000000") & ", " & _
            Format(moveData.RotationOriginY, "0.000000") & ", " & _
            Format(moveData.RotationOriginZ, "0.000000") & ")" & vbCrLf & _
        "  Bodies count = " & CStr(moveData.GetBodiesCount)

    Exit Function

EH:
    MoveCopyFeatureDiagnostic = _
        "Move/Copy definition readback failed: " & Err.Description
End Function

Private Sub SelectBodyForMove(ByVal bodyName As String)
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim selData As SldWorks.SelectData
    Dim swBody As SldWorks.Body2
    Dim beforeCount As Long
    Dim afterCount As Long
    Dim byIdResult As Boolean
    Dim directResult As Boolean

    Set swSelMgr = swModel.SelectionManager

    swModel.ClearSelection2 True
    beforeCount = swSelMgr.GetSelectedObjectCount2(1)

    ' This follows the SOLIDWORKS 2025 Move/Copy Body examples:
    ' select a SOLIDBODY using mark 1 before InsertMoveCopyBody2.
    byIdResult = swModel.Extension.SelectByID2( _
                    bodyName, "SOLIDBODY", _
                    0#, 0#, 0#, _
                    False, 1, Nothing, swSelectOptionDefault)

    afterCount = swSelMgr.GetSelectedObjectCount2(1)
    If afterCount > beforeCount Then Exit Sub

    ' Reacquire the current body after preceding cut/pattern/move features.
    ' Do not depend on a stale Body2 pointer captured before those features.
    Set swBody = GetSolidBodyByName(bodyName)

    If Not swBody Is Nothing Then
        Set selData = swSelMgr.CreateSelectData
        selData.Mark = 1

        directResult = swBody.Select2(False, selData)

        afterCount = swSelMgr.GetSelectedObjectCount2(1)
        If afterCount > beforeCount Then Exit Sub
    End If

    Err.Raise vbObjectError + 611, , _
        "Could not select solid body '" & bodyName & "' for Move/Copy Body." & vbCrLf & _
        "SelectByID2 returned: " & CStr(byIdResult) & vbCrLf & _
        "Body2.Select2 returned: " & CStr(directResult) & vbCrLf & _
        "Mark-1 count before: " & CStr(beforeCount) & vbCrLf & _
        "Mark-1 count after: " & CStr(afterCount) & vbCrLf & _
        "Current solid bodies: " & SolidBodyNames()
End Sub

Private Sub SelectReference(ByVal refObj As Object, ByVal append As Boolean, ByVal mark As Long)
    ' SOLIDWORKS can leave the requested object correctly selected even when
    ' IFeature::Select2 / SelectByID2 reports False.  Do not use that Boolean
    ' alone as the success test.  The selection is cleared before this routine,
    ' so SelectionManager is the authoritative check.

    Dim ok As Boolean
    Dim swFeat As SldWorks.Feature
    Dim swFace As SldWorks.Face2
    Dim selData As SldWorks.SelectData
    Dim swSelMgr As SldWorks.SelectionMgr
    Dim featName As String
    Dim featType As String
    Dim errDetail As String
    Dim selCount As Long
    Dim selObjType As Long

    Set swSelMgr = swModel.SelectionManager

    ' ----- Reference-plane / feature path -----
    On Error Resume Next
    Set swFeat = Nothing
    Set swFeat = refObj
    On Error GoTo 0

    If Not swFeat Is Nothing Then
        ' This is the direct pattern from the official SOLIDWORKS
        ' "Select Plane Example (VBA)".
        On Error Resume Next
        ok = swFeat.Select2(append, mark)
        On Error GoTo 0

        ' IMPORTANT: trust the actual selection state, not only 'ok'.
        selCount = swSelMgr.GetSelectedObjectCount2(-1)
        If selCount > 0 Then
            selObjType = swSelMgr.GetSelectedObjectType3(1, -1)
            If selObjType = swSelDATUMPLANES Then Exit Sub
        End If

        ' Official SOLIDWORKS fallback: get the selection name/type from the
        ' feature, then select it through IModelDocExtension::SelectByID2.
        featName = ""
        featType = ""

        On Error Resume Next
        featName = swFeat.GetNameForSelection(featType)
        On Error GoTo 0

        If Len(featName) > 0 And Len(featType) > 0 Then
            On Error Resume Next
            ok = swModel.Extension.SelectByID2( _
                    featName, featType, _
                    0#, 0#, 0#, _
                    append, mark, Nothing, swSelectOptionDefault)
            On Error GoTo 0

            selCount = swSelMgr.GetSelectedObjectCount2(-1)
            If selCount > 0 Then
                selObjType = swSelMgr.GetSelectedObjectType3(1, -1)
                If selObjType = swSelDATUMPLANES Then Exit Sub
            End If
        End If

        errDetail = "Feature name: " & swFeat.Name & _
                    ", feature type: " & swFeat.GetTypeName & _
                    ", GetNameForSelection name/type: " & featName & "/" & featType & _
                    ", selected count: " & CStr(selCount)

    Else
        ' ----- Planar-face path -----
        On Error Resume Next
        Set swFace = Nothing
        Set swFace = refObj
        On Error GoTo 0

        If Not swFace Is Nothing Then
            Set selData = swSelMgr.CreateSelectData
            selData.Mark = mark

            On Error Resume Next
            ok = swFace.Select4(append, selData)
            On Error GoTo 0

            selCount = swSelMgr.GetSelectedObjectCount2(-1)
            If selCount > 0 Then
                selObjType = swSelMgr.GetSelectedObjectType3(1, -1)
                If selObjType = swSelFACES Then Exit Sub
            End If

            errDetail = "Reference object is a Face2; selected count: " & CStr(selCount)
        Else
            errDetail = "Object supports neither IFeature nor IFace2. VBA TypeName=" & TypeName(refObj)
        End If
    End If

    Err.Raise vbObjectError + 600, , _
        "SOLIDWORKS did not leave the requested reference selected." & vbCrLf & errDetail
End Sub

Private Function GetNthReferencePlane(ByVal reqPlane As Long) As SldWorks.Feature
    Dim swFeat As SldWorks.Feature
    Dim planeCount As Long

    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "RefPlane" Or swFeat.GetTypeName = "RefPlane" Then
            planeCount = planeCount + 1
            If planeCount = reqPlane Then
                Set GetNthReferencePlane = swFeat
                Exit Function
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop

    Set GetNthReferencePlane = Nothing
End Function

Private Function FindNewestFeatureByType(ByVal typeName As String) As SldWorks.Feature
    Dim i As Long
    Dim maxI As Long
    Dim swFeat As SldWorks.Feature
    Dim t As String

    maxI = swModel.GetFeatureCount + 10

    For i = 0 To maxI
        Set swFeat = swModel.FeatureByPositionReverse(i)
        If swFeat Is Nothing Then Exit For

        t = swFeat.GetTypeName2
        If t = "" Then t = swFeat.GetTypeName

        If StrComp(t, typeName, vbTextCompare) = 0 Then
            Set FindNewestFeatureByType = swFeat
            Exit Function
        End If
    Next i

    Set FindNewestFeatureByType = Nothing
End Function

Private Function Ctl(ByVal frm As Object, ByVal controlName As String) As Object
    Set Ctl = frm.Controls(controlName)
End Function

Private Function ParsePositiveDouble(ByVal s As String, ByVal fieldName As String) As Double
    Dim v As Double

    If Not IsNumeric(s) Then
        Err.Raise vbObjectError + 700, , fieldName & " must be numeric."
    End If

    v = CDbl(s)
    If v <= 0# Then
        Err.Raise vbObjectError + 701, , fieldName & " must be greater than zero."
    End If

    ParsePositiveDouble = v
End Function

Private Function ParseNonNegativeDouble(ByVal s As String, ByVal fieldName As String) As Double
    Dim v As Double

    If Trim$(s) = "" Then
        ParseNonNegativeDouble = 0#
        Exit Function
    End If

    If Not IsNumeric(s) Then
        Err.Raise vbObjectError + 702, , fieldName & " must be numeric."
    End If

    v = CDbl(s)
    If v < 0# Then
        Err.Raise vbObjectError + 703, , fieldName & " cannot be negative."
    End If

    ParseNonNegativeDouble = v
End Function

Private Function ParsePositiveInteger(ByVal s As String, ByVal fieldName As String) As Long
    Dim v As Double
    Dim n As Long

    If Not IsNumeric(s) Then
        Err.Raise vbObjectError + 704, , fieldName & " must be numeric."
    End If

    v = CDbl(s)
    n = CLng(v)

    If n <= 0 Or Abs(v - n) > 0.000001 Then
        Err.Raise vbObjectError + 705, , fieldName & " must be a positive whole number."
    End If

    ParsePositiveInteger = n
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        Atan2 = Atn(y / x)
    ElseIf x < 0# And y >= 0# Then
        Atan2 = Atn(y / x) + PI
    ElseIf x < 0# And y < 0# Then
        Atan2 = Atn(y / x) - PI
    ElseIf x = 0# And y > 0# Then
        Atan2 = PI / 2#
    ElseIf x = 0# And y < 0# Then
        Atan2 = -PI / 2#
    Else
        Atan2 = 0#
    End If
End Function

Private Function MmToM(ByVal mm As Double) As Double
    MmToM = mm / 1000#
End Function

Private Function MToMm(ByVal m As Double) As Double
    MToMm = m * 1000#
End Function

Private Function DegToRad(ByVal deg As Double) As Double
    DegToRad = deg * PI / 180#
End Function

Private Function RadToDeg(ByVal rad As Double) As Double
    RadToDeg = rad * 180# / PI
End Function
