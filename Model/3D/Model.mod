'# MWS Version: Version 2019.0 - Sep 20 2018 - ACIS 28.0.2 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 1 fmax = 5
'# created = '[VERSION]2019.0|28.0.2|20180920[/VERSION]


'@ use template: Planar Filter_37.cfg

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
'set the units
With Units
    .Geometry "mm"
    .Frequency "GHz"
    .Voltage "V"
    .Resistance "Ohm"
    .Inductance "H"
    .TemperatureUnit  "Kelvin"
    .Time "ns"
    .Current "A"
    .Conductance "Siemens"
    .Capacitance "F"
End With
'----------------------------------------------------------------------------
'set the frequency range
Solver.FrequencyRange "1", "5"
'----------------------------------------------------------------------------
With Background
     .Type "Normal"
     .Epsilon "1.0"
     .Mu "1.0"
     .XminSpace "0.0"
     .XmaxSpace "0.0"
     .YminSpace "0.0"
     .YmaxSpace "0.0"
     .ZminSpace "0.0"
     .ZmaxSpace "0.0"
End With
With Boundary
     .Xmin "electric"
     .Xmax "electric"
     .Ymin "electric"
     .Ymax "electric"
     .Zmin "electric"
     .Zmax "electric"
End With
' mesh - Tetrahedral
With Mesh
     .MeshType "Tetrahedral"
     .SetCreator "High Frequency"
End With
With MeshSettings
     .SetMeshType "Tet"
     .Set "Version", 1%
     .Set "StepsPerWaveNear", "6"
     .Set "StepsPerBoxNear", "10"
     .Set "CellsPerWavelengthPolicy", "automatic"
     .Set "CurvatureOrder", "2"
     .Set "CurvatureOrderPolicy", "automatic"
     .Set "CurvRefinementControl", "NormalTolerance"
     .Set "NormalTolerance", "60"
     .Set "SrfMeshGradation", "1.5"
     .Set "UseAnisoCurveRefinement", "1"
     .Set "UseSameSrfAndVolMeshGradation", "1"
     .Set "VolMeshGradation", "1.5"
End With
With MeshSettings
     .SetMeshType "Unstr"
     .Set "MoveMesh", "1"
     .Set "OptimizeForPlanarStructures", "1"
End With
With Mesh
     .MeshType "PBA"
     .SetCreator "High Frequency"
     .AutomeshRefineAtPecLines "True", "4"
     .UseRatioLimit "True"
     .RatioLimit "50"
     .LinesPerWavelength "20"
     .MinimumStepNumber "10"
     .Automesh "True"
End With
With MeshSettings
     .SetMeshType "Hex"
     .Set "RatioLimitGeometry", "50"
     .Set "StepsPerWaveNear", "20"
     .Set "EdgeRefinementOn", "1"
     .Set "EdgeRefinementRatio", "4"
End With
' mesh - Multilayer (Preview)
' default
' solver - FD settings
With FDSolver
     .Reset
     .Method "Tetrahedral Mesh" ' i.e. general purpose
     .AccuracyHex "1e-6"
     .AccuracyTet "1e-5"
     .AccuracySrf "1e-3"
     .SetUseFastResonantForSweepTet "False"
     .Type "Direct"
     .MeshAdaptionHex "False"
     .MeshAdaptionTet "True"
     .InterpolationSamples "5001"
End With
With MeshAdaption3D
    .SetType "HighFrequencyTet"
    .SetAdaptionStrategy "Energy"
    .MinPasses "3"
    .MaxPasses "10"
End With
FDSolver.SetShieldAllPorts "True"
With FDSolver
     .Method "Tetrahedral Mesh (MOR)"
     .HexMORSettings "", "5001"
End With
FDSolver.Method "Tetrahedral Mesh" ' i.e. general purpose
' solver - TD settings
With MeshAdaption3D
    .SetType "Time"
    .SetAdaptionStrategy "Energy"
    .CellIncreaseFactor "0.5"
    .AddSParameterStopCriterion "True", "0.0", "10", "0.01", "1", "True"
End With
With Solver
     .Method "Hexahedral"
     .SteadyStateLimit "-40"
     .MeshAdaption "True"
     .NumberOfPulseWidths "50"
     .FrequencySamples "5001"
     .UseArfilter "True"
End With
' solver - M settings
'----------------------------------------------------------------------------
Dim sDefineAt As String
sDefineAt = "1;3;5"
Dim sDefineAtName As String
sDefineAtName = "1;3;5"
Dim sDefineAtToken As String
sDefineAtToken = "f="
Dim aFreq() As String
aFreq = Split(sDefineAt, ";")
Dim aNames() As String
aNames = Split(sDefineAtName, ";")
Dim nIndex As Integer
For nIndex = LBound(aFreq) To UBound(aFreq)
Dim zz_val As String
zz_val = aFreq (nIndex)
Dim zz_name As String
zz_name = sDefineAtToken & aNames (nIndex)
' Define E-Field Monitors
With Monitor
    .Reset
    .Name "e-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Efield"
    .MonitorValue  zz_val
    .Create
End With
Next
'----------------------------------------------------------------------------
With MeshSettings
     .SetMeshType "Tet"
     .Set "Version", 1%
End With
With Mesh
     .MeshType "Tetrahedral"
End With
'set the solver type
ChangeSolverType("HF Frequency Domain")
'----------------------------------------------------------------------------

'@ new component: component1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Component.New "component1"

'@ define brick: component1:solid1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-4", "4" 
     .Yrange "0.5", "4" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-3", "3" 
     .Yrange "0.9", "03.6" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2.5", "2.5" 
     .Yrange "1.6", "3.1" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid4

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2", "02" 
     .Yrange "01.8", "2.9" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid2"

'@ boolean subtract shapes: component1:solid3, component1:solid4

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid3", "component1:solid4"

'@ define brick: component1:solid4

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.2", "0.2" 
     .Yrange "0.9", "0.5" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid5

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid5" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.1", "0.1" 
     .Yrange "3.1", "2.9" 
     .Zrange "1.6", "01.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid1, component1:solid4

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid1", "component1:solid4"

'@ boolean subtract shapes: component1:solid3, component1:solid5

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid3", "component1:solid5"

'@ define brick: component1:solid4

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2.5", "02.5" 
     .Yrange "-0.5", "0.5" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid5

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid5" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2.3", "02.3" 
     .Yrange "-0.3", "0.3" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid4, component1:solid5

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid4", "component1:solid5"

'@ define brick: component1:solid5

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid5" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-4", "4" 
     .Yrange "-0.5", "-4" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid6

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid6" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-3", "3" 
     .Yrange "-0.9", "-3.6" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid5, component1:solid6

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid5", "component1:solid6"

'@ define brick: component1:solid6

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid6" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2", "2" 
     .Yrange "-1.8", "-2.9" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid7

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid7" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2.5", "2.5" 
     .Yrange "-1.6", "-3.1" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid7, component1:solid6

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid7", "component1:solid6"

'@ define brick: component1:solid8

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid8" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.2", "0.2" 
     .Yrange "-0.9", "0-0.5" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid9

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid9" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.1", "0.1" 
     .Yrange "3.1", "2.9" 
     .Zrange "1.6", "01.635" 
     .Create
End With

'@ delete shape: component1:solid9

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid9"

'@ define brick: component1:solid9

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid9" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.1", "0.1" 
     .Yrange "-3.1", "-2.9" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid5, component1:solid8

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid5", "component1:solid8"

'@ boolean subtract shapes: component1:solid7, component1:solid9

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid7", "component1:solid9"

'@ define brick: component1:solid8

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid8" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "4", "7.5" 
     .Yrange "-0.03", "0.03" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid9

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid9" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "7.5", "17.5" 
     .Yrange "-2", "2" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid10

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid10" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-4", "-7.5" 
     .Yrange "-0.03", "0.03" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid11

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid11" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-7.5", "-17.5" 
     .Yrange "-2", "2" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid12

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid12" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-17.5", "17.5" 
     .Yrange "-10", "10" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ delete shape: component1:solid12

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid12"

'@ define material: FR-4 (lossy)

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Material
     .Reset
     .Name "FR-4 (lossy)"
     .Folder ""
.FrqType "all"
.Type "Normal"
.SetMaterialUnit "GHz", "mm"
.Epsilon "4.3"
.Mu "1.0"
.Kappa "0.0"
.TanD "0.025"
.TanDFreq "10.0"
.TanDGiven "True"
.TanDModel "ConstTanD"
.KappaM "0.0"
.TanDM "0.0"
.TanDMFreq "0.0"
.TanDMGiven "False"
.TanDMModel "ConstKappa"
.DispModelEps "None"
.DispModelMu "None"
.DispersiveFittingSchemeEps "General 1st"
.DispersiveFittingSchemeMu "General 1st"
.UseGeneralDispersionEps "False"
.UseGeneralDispersionMu "False"
.Rho "0.0"
.ThermalType "Normal"
.ThermalConductivity "0.3"
.SetActiveMaterial "all"
.Colour "0.94", "0.82", "0.76"
.Wireframe "False"
.Transparency "0"
.Create
End With

'@ define brick: component1:solid12

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid12" 
     .Component "component1" 
     .Material "FR-4 (lossy)" 
     .Xrange "-17.5", "017.5" 
     .Yrange "-10", "10" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ delete shape: component1:solid12

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid12"

'@ define brick: component1:solid12

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid12" 
     .Component "component1" 
     .Material "FR-4 (lossy)" 
     .Xrange "-17.5", "017.5" 
     .Yrange "-10", "10" 
     .Zrange "0", "1.6" 
     .Create
End With

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid12", "2"

'@ define brick: component1:solid13

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid13" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-17.5", "17.5" 
     .Yrange "-10", "10" 
     .Zrange "0", "-0.035" 
     .Create
End With

'@ clear picks

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.ClearAllPicks

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid9", "6"

'@ define port:1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient
With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0", "0"
  .YrangeAdd "1.6*6.7", "1.6*6.7"
  .ZrangeAdd "1.6", "1.6*6.7"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid11", "4"

'@ define port:2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient
With Port
  .Reset
  .PortNumber "2"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0", "0"
  .YrangeAdd "1.6*6.7", "1.6*6.7"
  .ZrangeAdd "1.6", "1.6*6.7"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define frequency domain solver parameters

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Mesh.SetCreator "High Frequency" 
With FDSolver
     .Reset 
     .SetMethod "Tetrahedral", "General purpose" 
     .OrderTet "Second" 
     .OrderSrf "First" 
     .Stimulation "All", "All" 
     .ResetExcitationList 
     .AutoNormImpedance "False" 
     .NormingImpedance "50" 
     .ModesOnly "False" 
     .ConsiderPortLossesTet "True" 
     .SetShieldAllPorts "True" 
     .AccuracyHex "1e-6" 
     .AccuracyTet "1e-5" 
     .AccuracySrf "1e-3" 
     .LimitIterations "False" 
     .MaxIterations "0" 
     .SetCalcBlockExcitationsInParallel "True", "True", "" 
     .StoreAllResults "False" 
     .StoreResultsInCache "False" 
     .UseHelmholtzEquation "True" 
     .LowFrequencyStabilization "True" 
     .Type "Direct" 
     .MeshAdaptionHex "False" 
     .MeshAdaptionTet "True" 
     .AcceleratedRestart "True" 
     .FreqDistAdaptMode "Distributed" 
     .NewIterativeSolver "True" 
     .TDCompatibleMaterials "False" 
     .ExtrudeOpenBC "False" 
     .SetOpenBCTypeHex "Default" 
     .SetOpenBCTypeTet "Default" 
     .AddMonitorSamples "True" 
     .CalcPowerLoss "True" 
     .CalcPowerLossPerComponent "False" 
     .StoreSolutionCoefficients "True" 
     .UseDoublePrecision "False" 
     .UseDoublePrecision_ML "True" 
     .MixedOrderSrf "False" 
     .MixedOrderTet "False" 
     .PreconditionerAccuracyIntEq "0.15" 
     .MLFMMAccuracy "Default" 
     .MinMLFMMBoxSize "0.3" 
     .UseCFIEForCPECIntEq "True" 
     .UseFastRCSSweepIntEq "false" 
     .UseSensitivityAnalysis "False" 
     .RemoveAllStopCriteria "Hex"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
     .RemoveAllStopCriteria "Tet"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
     .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
     .RemoveAllStopCriteria "Srf"
     .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
     .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
     .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
     .SweepMinimumSamples "3" 
     .SetNumberOfResultDataSamples "5001" 
     .SetResultDataSamplingMode "Automatic" 
     .SweepWeightEvanescent "1.0" 
     .AccuracyROM "1e-4" 
     .AddSampleInterval "", "", "1", "Automatic", "True" 
     .AddSampleInterval "", "", "", "Automatic", "False" 
     .MPIParallelization "False"
     .UseDistributedComputing "False"
     .NetworkComputingStrategy "RunRemote"
     .NetworkComputingJobCount "3"
     .UseParallelization "True"
     .MaxCPUs "128"
     .MaximumNumberOfCPUDevices "2"
End With
With IESolver
     .Reset 
     .UseFastFrequencySweep "True" 
     .UseIEGroundPlane "False" 
     .SetRealGroundMaterialName "" 
     .CalcFarFieldInRealGround "False" 
     .RealGroundModelType "Auto" 
     .PreconditionerType "Auto" 
     .ExtendThinWireModelByWireNubs "False" 
End With
With IESolver
     .SetFMMFFCalcStopLevel "0" 
     .SetFMMFFCalcNumInterpPoints "6" 
     .UseFMMFarfieldCalc "True" 
     .SetCFIEAlpha "0.500000" 
     .LowFrequencyStabilization "False" 
     .LowFrequencyStabilizationML "True" 
     .Multilayer "False" 
     .SetiMoMACC_I "0.0001" 
     .SetiMoMACC_M "0.0001" 
     .DeembedExternalPorts "True" 
     .SetOpenBC_XY "True" 
     .OldRCSSweepDefintion "False" 
     .SetAccuracySetting "Custom" 
     .CalculateSParaforFieldsources "True" 
     .ModeTrackingCMA "True" 
     .NumberOfModesCMA "3" 
     .StartFrequencyCMA "-1.0" 
     .SetAccuracySettingCMA "Default" 
     .FrequencySamplesCMA "0" 
     .SetMemSettingCMA "Auto" 
End With

'@ define brick: component1:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid14" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-17.5", "17.5" 
     .Yrange "-0.5", "0.5" 
     .Zrange "0", "0.035" 
     .Create
End With

'@ define brick: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-0.5", "0.5" 
     .Yrange "-10", "10" 
     .Zrange "0", "0.035" 
     .Create
End With

'@ delete shape: component1:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid14"

'@ delete shape: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid15"

'@ define brick: component1:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid14" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-1.9", "1.9" 
     .Yrange "-0.2", "0.2" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-1.5", "1.5" 
     .Yrange "-0.1", "0.1" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid14, component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid14", "component1:solid15"

'@ delete shape: component1:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid14"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid3", "39"

'@ unpick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.UnpickEndpointFromId "component1:solid3", "39"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid3", "38"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid3", "35"

'@ define distance dimension

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .Reset
    .CreationType "picks"
    .SetType "Distance"
    .SetID "0"
    .SetOrientation "Smart Mode"
    .SetDistance "0.434750"
    .SetViewVector "-0.115179", "0.060427", "-0.991505"
    .SetConnectedElement1 "component1:solid3"
    .SetConnectedElement2 "component1:solid3"
    .Create
End With
Pick.ClearAllPicks

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid1", "35"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid1", "38"

'@ define distance dimension

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .Reset
    .CreationType "picks"
    .SetType "Distance"
    .SetID "1"
    .SetOrientation "Smart Mode"
    .SetDistance "0.397132"
    .SetViewVector "-0.115179", "0.060427", "-0.991505"
    .SetConnectedElement1 "component1:solid1"
    .SetConnectedElement2 "component1:solid1"
    .Create
End With
Pick.ClearAllPicks

'@ delete dimension 0

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .RemoveDimension "0"
End With

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid3", "35"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid3", "38"

'@ define distance dimension

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .Reset
    .CreationType "picks"
    .SetType "Distance"
    .SetID "2"
    .SetOrientation "Smart Mode"
    .SetDistance "0.315052"
    .SetViewVector "0.100388", "0.014109", "-0.994848"
    .SetConnectedElement1 "component1:solid3"
    .SetConnectedElement2 "component1:solid3"
    .Create
End With
Pick.ClearAllPicks

'@ delete dimension 1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .RemoveDimension "1"
End With

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid1", "35"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid1", "38"

'@ define distance dimension

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .Reset
    .CreationType "picks"
    .SetType "Distance"
    .SetID "3"
    .SetOrientation "Smart Mode"
    .SetDistance "0.343939"
    .SetViewVector "0.100388", "0.014102", "-0.994848"
    .SetConnectedElement1 "component1:solid1"
    .SetConnectedElement2 "component1:solid1"
    .Create
End With
Pick.ClearAllPicks

'@ delete dimension 2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .RemoveDimension "2"
End With

'@ delete dimension 3

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .RemoveDimension "3"
End With

'@ delete port: port1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Port.Delete "1"

'@ delete port: port2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Port.Delete "2"

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid9", "6"

'@ define port:1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient
With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0", "0"
  .YrangeAdd "1.6*6.7", "1.6*6.7"
  .ZrangeAdd "1.6", "1.6*6.7"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid11", "4"

'@ define port:2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
' Port constructed by macro Solver -> Ports -> Calculate port extension coefficient
With Port
  .Reset
  .PortNumber "2"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0", "0"
  .YrangeAdd "1.6*6.7", "1.6*6.7"
  .ZrangeAdd "1.6", "1.6*6.7"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define lumped element: element1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With LumpedElement
     .Reset 
     .SetName "element1" 
     .Folder "" 
     .SetType "RLCSerial" 
     .SetR  "100" 
     .SetL  "0.01" 
     .SetC  "0.001" 
     .SetGs "0" 
     .SetI0 "1e-14" 
     .SetT  "300" 
     .SetP1 "False", "0", "0.1", "0" 
     .SetP2 "False", "0", "-0.1", "0" 
     .SetInvert "False" 
     .SetMonitor "True" 
     .SetRadius "0.05" 
     .Wire "" 
     .Position "end1" 
     .CircuitFileName "" 
     .CircuitId "1" 
     .UseCopyOnly "True" 
     .UseRelativePath "False" 
     .Create
End With

'@ delete lumped element: element1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
LumpedElement.Delete "element1"

'@ delete shape: component1:solid4

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid4"

'@ new component: component2

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Component.New "component2"

'@ define brick: component2:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid14" 
     .Component "component2" 
     .Material "PEC" 
     .Xrange "-2.3", "2.3" 
     .Yrange "-0.2", "0.2" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ delete shape: component2:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component2:solid14"

'@ define brick: component1:solid14

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid14" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-4", "4" 
     .Yrange "-0.5", "0.5" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define lumped element: Folder1:element1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With LumpedElement
     .Reset 
     .SetName "element1" 
     .Folder "Folder1" 
     .SetType "RLCSerial" 
     .SetR  "1000" 
     .SetL  "0.05" 
     .SetC  "0.02" 
     .SetGs "0" 
     .SetI0 "1e-14" 
     .SetT  "300" 
     .SetP1 "False", "-1", "-0.2", "1.6" 
     .SetP2 "False", "1", "0.2", "1.635" 
     .SetInvert "False" 
     .SetMonitor "True" 
     .SetRadius "0.04" 
     .Wire "" 
     .Position "end1" 
     .CircuitFileName "" 
     .CircuitId "1" 
     .UseCopyOnly "True" 
     .UseRelativePath "False" 
     .Create
End With

'@ delete lumped folder: Folder1

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
LumpedElement.DeleteFolder "Folder1"

'@ define brick: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-3", "3" 
     .Yrange "-0.3", "0.3" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ boolean subtract shapes: component1:solid14, component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Subtract "component1:solid14", "component1:solid15"

'@ define brick: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "2.5", "3" 
     .Yrange "-0.2", "0.2" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid16

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid16" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2.5", "-3" 
     .Yrange "-0.2", "0.2" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ delete shape: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid15"

'@ delete shape: component1:solid16

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid16"

'@ define brick: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "2", "2.5" 
     .Yrange "-0.05", "0.05" 
     .Zrange "1.6", "1.635" 
     .Create
End With

'@ define brick: component1:solid16

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid16" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-2", "-2.5" 
     .Yrange "-0.05", "0.05" 
     .Zrange "1.6", "1.65" 
     .Create
End With

'@ delete shape: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid15"

'@ delete shape: component1:solid16

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid16"

'@ define sphere: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Sphere 
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Axis "z" 
     .CenterRadius "4" 
     .TopRadius "2" 
     .BottomRadius "3" 
     .Center "2", "1.5", "2" 
     .Segments "0" 
     .Create 
End With

'@ delete shape: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid15"

'@ clear picks

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.ClearAllPicks

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid13", "2"

'@ define brick: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid15" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-1.5", "17.5" 
     .Yrange "-10", "10" 
     .Zrange "-0.035", "-0.035" 
     .Create
End With

'@ define brick: component1:solid16

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid16" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-17.5", "-1.5" 
     .Yrange "-10", "10" 
     .Zrange "-0.035", "-0.035" 
     .Create
End With

'@ clear picks

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.ClearAllPicks

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid15", "1"

'@ new component: component3

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Component.New "component3"

'@ define brick: component3:solid17

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid17" 
     .Component "component3" 
     .Material "PEC" 
     .Xrange "-1.5", "8" 
     .Yrange "-10", "10" 
     .Zrange "-0.035", "-0.035" 
     .Create
End With

'@ clear picks

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.ClearAllPicks

'@ delete shape: component3:solid17

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component3:solid17"

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid15", "1"

'@ define brick: component1:solid17

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid17" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-1.5", "8" 
     .Yrange "-10", "10" 
     .Zrange "-0.035", "-0.035" 
     .Create
End With

'@ clear picks

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.ClearAllPicks

'@ pick face

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickFaceFromId "component1:solid16", "1"

'@ define brick: component1:solid18

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Brick
     .Reset 
     .Name "solid18" 
     .Component "component1" 
     .Material "PEC" 
     .Xrange "-10", "-1.5" 
     .Yrange "-10", "10" 
     .Zrange "-0.035", "-0.035" 
     .Create
End With

'@ clear picks

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.ClearAllPicks

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid12", "4"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid12", "1"

'@ define distance dimension

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .Reset
    .CreationType "picks"
    .SetType "Distance"
    .SetID "0"
    .SetOrientation "Smart Mode"
    .SetDistance "4.113060"
    .SetViewVector "-0.679621", "-0.248544", "-0.690175"
    .SetConnectedElement1 "component1:solid12"
    .SetConnectedElement2 "component1:solid12"
    .Create
End With
Pick.ClearAllPicks

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid12", "1"

'@ pick end point

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Pick.PickEndpointFromId "component1:solid12", "2"

'@ define distance dimension

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
With Dimension
    .Reset
    .CreationType "picks"
    .SetType "Distance"
    .SetID "1"
    .SetOrientation "Smart Mode"
    .SetDistance "8.075624"
    .SetViewVector "-0.679621", "-0.248544", "-0.690175"
    .SetConnectedElement1 "component1:solid12"
    .SetConnectedElement2 "component1:solid12"
    .Create
End With
Pick.ClearAllPicks

'@ delete shape: component1:solid17

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid17"

'@ delete shape: component1:solid18

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid18"

'@ delete shape: component1:solid15

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid15"

'@ delete shape: component1:solid16

'[VERSION]2019.0|28.0.2|20180920[/VERSION]
Solid.Delete "component1:solid16" 

