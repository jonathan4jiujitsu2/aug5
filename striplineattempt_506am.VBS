' Hairpin Filter: Fully Filled Geometry as Single Closed Polyline + Thicken

Dim Wio, L1, L2, L3, L4, L5, L6, L7, L9, L10, L11, L12, L13, L14, L15
Dim oAnsoftApp, oDesktop, oProject, oDesign, oEditor

Wio = 7.0
L1 = 50
L2 = 51
L3 = 26.28
L4 = 20.48
L5 = 103.36
L6 = 100.26
L7 = 14.3378
L9 = 2.65
L10 = 8.58
L11 = 45.118
L12 = 28
L13 = 55
L14 = 133
L15 = 108

Set oAnsoftApp = CreateObject("AnsoftHfss.HfssScriptInterface")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow
Set oProject = oDesktop.NewProject
Set oDesign = oProject.InsertDesign("HFSS", "HairpinFilterFilled", "DrivenModal", "")
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

Dim pts
pts = Array( _
  Array("NAME:PLPoint", "X:=" , "0mil", "Y:=" , "0mil", "Z:=" , "0mil"), _                     ' Start (Wio base)
  Array("NAME:PLPoint", "X:=" , "0mil", "Y:=" , CStr(Wio) & "mil", "Z:=" , "0mil"), _              ' Top of Wio
  Array("NAME:PLPoint", "X:=" , CStr(L1) & "mil", "Y:=" , CStr(Wio) & "mil", "Z:=" , "0mil"), _    ' End of Line1
  Array("NAME:PLPoint", "X:=" , CStr(L1) & "mil", "Y:=" , "0mil", "Z:=" , "0mil"), _              ' End of Line2
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071) & "mil", "Y:=" , CStr(L3*0.7071) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071) & "mil", "Y:=" , CStr(L3*0.7071 + L5) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071 + L7*0.7071) & "mil", "Y:=" , CStr(L3*0.7071 + L5 + L7*0.7071) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071 + L7*0.7071) & "mil", "Y:=" , CStr(L3*0.7071 + L5 + L7*0.7071 + L9) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071 + L7*0.7071 + L11) & "mil", "Y:=" , CStr(L3*0.7071 + L5 + L7*0.7071 + L9) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071 + L7*0.7071 + L11) & "mil", "Y:=" , CStr(L3*0.7071 + L5 + L7*0.7071 + L9 - L15) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071 + L7*0.7071 + L11 + L13) & "mil", "Y:=" , CStr(L3*0.7071 + L5 + L7*0.7071 + L9 - L15) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L3*0.7071 + L7*0.7071 + L11 + L13) & "mil", "Y:=" , CStr(L3*0.7071 + L5 + L7*0.7071 + L9 - L15 + L14) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L4*0.7071 + L13) & "mil", "Y:=" , CStr(Wio + L4*0.7071 + L6 + L12*0.7071) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L4*0.7071) & "mil", "Y:=" , CStr(Wio + L4*0.7071 + L6) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L4*0.7071 - L7*0.7071) & "mil", "Y:=" , CStr(Wio + L4*0.7071 + L6 + L7*0.7071) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1 + L4*0.7071) & "mil", "Y:=" , CStr(Wio + L4*0.7071) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , CStr(L1) & "mil", "Y:=" , CStr(Wio) & "mil", "Z:=" , "0mil"), _
  Array("NAME:PLPoint", "X:=" , "0mil", "Y:=" , CStr(Wio) & "mil", "Z:=" , "0mil") _
)

Call oEditor.CreatePolyline( _
  Array("NAME:PolylineParameters", _
    "IsPolylineCovered:=" , true, _
    "IsPolylineClosed:=" , true, _
    Array("NAME:PolylinePoints", pts), _
    Array("NAME:PolylineSegments", _
      Array("NAME:PLSegment", "SegmentType:=" , "Line", "StartIndex:=" , 0, "NoOfPoints:=" , UBound(pts)+1) _
    ), _
    Array("NAME:PolylineXSection", _
      "XSectionType:=" , "None", _
      "XSectionOrient:=" , "Auto", _
      "XSectionWidth:=" , "0mil", _
      "XSectionTopWidth:=" , "0mil", _
      "XSectionHeight:=" , "0mil", _
      "XSectionNumSegments:=" , "0", _
      "XSectionBendType:=" , "Corner") _
  ), _
  Array("NAME:Attributes", _
    "Name:=" , "HairpinClosedLoop", _
    "Flags:=" , "", _
    "Color:=" , "(255 255 0)", _
    "Transparency:=" , 0, _
    "PartCoordinateSystem:=" , "Global", _
    "MaterialValue:=" , Chr(34) & "copper" & Chr(34), _
    "SolveInside:=" , false) _
)

Call oEditor.ThickenSheet( _
  Array("NAME:Selections", "Selections:=" , "HairpinClosedLoop"), _
  Array("NAME:ThickenSheetParameters", _
    "Thickness:=" , "0.65mil", _
    "BothSides:=" , false, _
    "ReplaceSheet:=" , true) _
)
