Attribute VB_Name = "Train_de_navettes"
Sub Validation_des_Hypothèses_Train_de_Navettes()

    Dim Resultat_E As Variant
    Dim Resultat_M As Variant
    Dim Resultat_P As Variant
    Dim Resultat_I As Variant
    Dim I As String
    Dim Nom As String
    Dim C_navette_X As Variant
    Dim C_navette_Y As Variant
    Dim C_navette As Variant
    Dim Validation As Variant
    Dim N_Objet As Byte
    Dim N_Cycle_Eff As Byte
    Dim N_Cycle_X_Y As Byte
    Dim M_Limite As Variant
    Dim P_Limite As Variant
    Dim N_Navette_2 As Variant

If Resultat_N <> "" Then GoTo Nombre_de_Navette

Saisie_N:
    Resultat_N = Application.InputBox(Prompt:="Please enter the conveyor plan name or number.", Title:="Conveyor name", Type:=2)
    If InStr(1, Resultat_N, "/", 0) > 0 Then GoTo Avertissement_N
    If InStr(1, Resultat_N, "\", 0) > 0 Then GoTo Avertissement_N
    If InStr(1, Resultat_N, ":", 0) > 0 Then GoTo Avertissement_N
    If InStr(1, Resultat_N, "*", 0) > 0 Then GoTo Avertissement_N
    If InStr(1, Resultat_N, "?", 0) > 0 Then GoTo Avertissement_N
    If Resultat_N = "Faux" Then GoTo Annuler
    If Resultat_N = "" Then GoTo Saisie_N

Message_1_1:
     MsgBox "                        You will check the section :" & Chr(10) & "1. Limit of the load and the position of Gc per shuttle"

Nombre_de_Navette:
    N_Navette = Application.InputBox(Prompt:="Please enter the number of shuttles, in the shuttle train.", Title:="Shuttles Number", Default:=3, Type:=1)
    If N_Navette = "Faux" Then GoTo Annuler
    If N_Navette < 1 Then GoTo Nombre_de_Navette
    If N_Navette <> Round(N_Navette, 0) Then GoTo Nombre_de_Navette

Nombre_dOutillage:
    N_Outillage = Application.InputBox(Prompt:="Please enter the number of toolings, in the shuttle train.", Title:="Toolings Number", Default:=2, Type:=1)
    If N_Outillage = "Faux" Then GoTo Annuler
    If N_Outillage < 1 Then GoTo Nombre_dOutillage
    If N_Outillage > N_Navette Then
    Avertissement = MsgBox("The number of toolings is higher than the number of shuttles." & Chr(10) & "Please increase the quantity of shuttles or reduce the quantity of toolings.", vbInformation + vbRetryCancel, "Result")
    If Avertissement = vbRetry Then GoTo Nombre_de_Navette
    If Avertissement = vbCancel Then GoTo Annuler
    End If
    If N_Outillage = False Then GoTo Annuler
    If N_Outillage <> Round(N_Outillage, 0) Then GoTo Nombre_dOutillage
    If N_Navette = 1 Then Call Validation_des_Hypothèses_Standard

N_Cycle_X_Y:
    N_Cycle_X_Y = 1
    Do While N_Cycle_X_Y < N_Outillage + 1

Force_sur_la_Navette_n:
    F = Application.InputBox(Prompt:="Please enter the load F" & N_Cycle_X_Y & ", in daN, exerted at point Gc" & N_Cycle_X_Y & " on the shuttle.", Title:="F" & N_Cycle_X_Y & " Value", Default:=75, Type:=1)
    If F = "Faux" Then GoTo Annuler
    If F <= 0 Then GoTo Force_sur_la_Navette_n
    F = (F * 10) / 9.81

Position_sur_la_Navette_en_x:
    x = Application.InputBox(Prompt:="Please enter the x position of F" & N_Cycle_X_Y & ", on the shuttle, in mm.", Title:="Value x of F" & N_Cycle_X_Y, Default:=0, Type:=1)
    If x = "Faux" Then GoTo Annuler
    If x < -60 Or x > 60 Then GoTo Position_sur_la_Navette_en_x

Position_sur_la_Navette_en_y:
    y = Application.InputBox(Prompt:="Please enter the y position of F" & N_Cycle_X_Y & ", on the shuttle, in mm.", Title:="Value y of F" & N_Cycle_X_Y, Default:=0, Type:=1)
    If y = "Faux" Then GoTo Annuler
    If y < -100 Or y > 100 Then GoTo Position_sur_la_Navette_en_y

Calcul_de_ma_Limite_de_Charge_sur_la_Navette:
    If x > 0 And y > 0 Then
    Resultat_C = (1920000) / (((120 / 2) + x) * ((200 / 2) + y))
    ElseIf x <= 0 And y > 0 Then
    Resultat_C = (1920000) / (((120 / 2) - x) * ((200 / 2) + y))
    ElseIf x <= 0 And y <= 0 Then
    Resultat_C = (1920000) / (((120 / 2) - x) * ((200 / 2) - y))
    ElseIf x > 0 And y <= 0 Then
    Resultat_C = (1920000) / (((120 / 2) + x) * ((200 / 2) - y))
    End If
    If Resultat_C < F Or 150 < F Then
    Resultat_C = MsgBox("The load exerted on the shuttle isn't conform." & Chr(10) & "Please recenter the position of the COG on the shuttle" & Chr(10) & "and/or reduce the stress exerted (tooling + part(s))." & Chr(10) & "Refer to the section dealing with limit of the load per shuttle." & Chr(10) & Chr(10) & "Click on Ignore to make a waiver request.", vbAbortRetryIgnore + vbCritical + vbDefaultButton1, "Result")
    If Resultat_C = vbAbort Then GoTo Annuler
    If Resultat_C = vbRetry Then GoTo Force_sur_la_Navette_n
    Else: Resultat_C = MsgBox("The configuration on the load F" & N_Cycle_X_Y & " and the position of Gc" & N_Cycle_X_Y & " is validated.", vbOKOnly, "Result")
    End If
    If Resultat_C = vbIgnore Then C_navette = "Non validées"
    F = WorksheetFunction.RoundUp(F, 1)
    Resultat_P0 = Resultat_P0 + F
    If N_Cycle_X_Y = 1 Then Fn = F
    If N_Cycle_X_Y > 1 Then Fn = Fn & " + " & F
    If F < 80 Then F = 80
    If F > 80 And F < 150 Then F = 150
    If F > 150 Then F = "-"
    If N_Cycle_X_Y = 1 Then nI = F
    If F = "-" And nI = 80 Or F = "-" And nI = 150 Then nI = F
    If F = 150 And nI = 80 Then nI = F
    x = WorksheetFunction.RoundUp(x, 1)
    y = WorksheetFunction.RoundUp(y, 1)
    If N_Cycle_X_Y = 1 Then P_navette_X = x
    If N_Cycle_X_Y > 1 Then P_navette_X = P_navette_X & " I " & x
    If N_Cycle_X_Y = 1 Then P_navette_Y = y
    If N_Cycle_X_Y > 1 Then P_navette_Y = P_navette_Y & " I " & y
    N_Cycle_X_Y = N_Cycle_X_Y + 1
    Loop
    If C_navette <> "Non validées" Then C_navette = "Validées"

Message_2:
     MsgBox "    You will check the section :" & Chr(10) & "     2. Load limit in the bends"

N_Cycle_M_E:
    N_Cycle_M_E = 1
    Do While N_Cycle_M_E < N_Outillage + 1

Saisie_Mn:
    Resultat_Mn = Application.InputBox(Prompt:="Please enter the mass of the tooling " & N_Cycle_M_E & ", in kg.", Title:="Mass Tooling " & N_Cycle_M_E, Default:=10, Type:=1)
    If Resultat_Mn = "Faux" Then GoTo Annuler
    If Resultat_Mn <= 0 Then GoTo Saisie_Mn

Saisie_E:
    Resultat_En = Application.InputBox(Prompt:="Please enter the distance E" & N_Cycle_M_E & " of the tooling, in mm.", Title:="E" & N_Cycle_M_E & " Value", Default:=100, Type:=1)
    If Resultat_En = "Faux" Then GoTo Annuler
    If Resultat_En <= 0 Then GoTo Saisie_E

Calcul_de_la_Limite_de_Charge_dans_les_Courbes:
    Cmax = (Range("B16") * Range("B17") * 9.81)
    Resultat = (((Resultat_Mn * ((Resultat_En) + Range("B17")))) * 9.81)
    If Resultat > Cmax Then
    Resultat = MsgBox("The load limit in the bends isn't conform." & Chr(10) & "Please reduce tooling " & N_Cycle_M_E & " weight and/or E" & N_Cycle_M_E & " distance." & Chr(10) & "Refer to the chart : Load limit in the bends." & Chr(10) & Chr(10) & "Click on Ignore to make a waiver request.", vbAbortRetryIgnore + vbCritical + vbDefaultButton1, "Result")
    If Resultat = vbAbort Then GoTo Annuler
    If Resultat = vbRetry Then GoTo Saisie_Mn
    Else: Resultat = MsgBox("The configuration " & N_Cycle_M_E & " in the bends is validated.", vbOKOnly, "Result")
    End If
    If Resultat = vbIgnore Then Hypothèses_Courbe = "Non validées"
    Resultat_Mn = WorksheetFunction.RoundUp(Resultat_Mn, 1)
    If N_Cycle_M_E = 1 Then Resultat_M = Resultat_Mn
    If N_Cycle_M_E > 1 Then Resultat_M = Resultat_M & " + " & Resultat_Mn
    M_Limite = M_Limite + Resultat_Mn
    Resultat_En = WorksheetFunction.RoundUp(Resultat_En, 1)
    If N_Cycle_M_E = 1 Then Resultat_E = Resultat_En
    If N_Cycle_M_E > 1 Then Resultat_E = Resultat_E & " I " & Resultat_En
    N_Cycle_M_E = N_Cycle_M_E + 1
    Loop
    If Hypothèses_Courbe <> "Non validées" Then Hypothèses_Courbe = "Validées"

Calcul_de_P_Par_Outillage:
    Do While N_Cycle_P_Long < N_Outillage
    N_Cycle_P = N_Outillage + N_Cycle_P_Long
    Resultat_Pn = Fn
    Resultat_Mn = Resultat_M
    If N_Cycle_P_Long <> 0 Then
    Do While N_Cycle_P > N_Outillage
    Resultat_Pn = Mid(Resultat_Pn, InStr(Resultat_Pn, "+") + 1)
    Resultat_Mn = Mid(Resultat_Mn, InStr(Resultat_Mn, "+") + 1)
    N_Cycle_P = N_Cycle_P - 1
    Loop
    End If
    N_Cycle_P_Long = N_Cycle_P_Long + 1
    If N_Cycle_P_Long < N_Outillage Then
    Resultat_Pn = StrReverse(Mid(StrReverse(Resultat_Pn), Len(Resultat_Pn) - InStr(Resultat_Pn, "+") + 2))
    Resultat_Mn = StrReverse(Mid(StrReverse(Resultat_Mn), Len(Resultat_Mn) - InStr(Resultat_Mn, "+") + 2))
    End If
    Resultat_Pn = Resultat_Pn - Resultat_Mn
    Resultat_Pn = WorksheetFunction.Round(Resultat_Pn, 1)
    If N_Cycle_P_Long = 1 Then Resultat_P = Resultat_Pn
    If N_Cycle_P_Long > 1 Then Resultat_P = Resultat_P & " + " & Resultat_Pn
    P_Limite = P_Limite + Resultat_Pn
    Loop

Validation:
    If C_navette <> "Validées" Then Validation = "The load per shuttle is invalid," & Chr(10) & "waiver request have been made."
    If Hypothèses_Courbe <> "Validées" Then Validation = "The load in the bends is invalid," & Chr(10) & "a waiver request have been made."
    If C_navette <> "Validées" And Hypothèses_Courbe <> "Validées" Then Validation = "The load per shuttle and in the bends are" & Chr(10) & "invalids, waiver requests have been made."
    If C_navette = "Validées" And Hypothèses_Courbe = "Validées" Then Validation = "The configurations are validated."
    Validation = Validation & " Made the "

Donné_le_Chemin_Accès:
    Nom = ThisWorkbook.Name
    Nom = Left(Nom, (InStrRev(ThisWorkbook.Name, ".") - 1)) & " - " & Resultat_N
    Nom = "\" & Nom
    Set Accès = Application.FileDialog(msoFileDialogFolderPicker)
    With Accès
        .Title = "Please enter the path of the pdf file, which will be generated."
        .AllowMultiSelect = False
        .InitialFileName = Workbooks(ActiveWorkbook.Name).Path & "\"
        If .Show = 0 Then GoTo Annuler
        Accès = .SelectedItems(1)
    End With

Générer_PDF:
    Worksheets("Shuttle Train").Unprotect Password:="Idra01*"
    Worksheets("Shuttle Train").Select
    ActiveWindow.Zoom = 100
    With Range("K1:R55")
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .ShrinkToFit = True
    End With
    Range("K2:R4").MergeCells = True
    
    Range("K2:R4") = "VALIDATION SHEET" & Chr(10) & "IDRAPAL: " & Resultat_N
    With Range("K2:R4").Font
        .Size = 14
        .Underline = xlUnderlineStyleSingle
        .Bold = True
    End With
    Range("L6", "N6").MergeCells = True
    Range("L6") = "Type of IDRAPAL"
    If N_Outillage > 1 Then Range("L6") = "Types of IDRAPAL"
    Range("O6", "O6").MergeCells = True
    Range("O6") = ":"
    Range("P6", "Q6").MergeCells = True
    If N_Outillage > 3 Then Range("P6", "R6").MergeCells = True
    Range("P6") = nI
    Range("L8", "N8").MergeCells = True
    Range("L8") = "Mass of the part(s)"
    Range("O8", "O8").MergeCells = True
    Range("O8") = ":"
    Range("P8", "Q8").MergeCells = True
    If N_Outillage > 3 Then Range("P8", "R8").MergeCells = True
    Range("P8") = Resultat_P & " kg"
    If N_Outillage > 1 Then Range("P8") = Resultat_P & " = " & P_Limite & " kg"
    Range("L10", "N10").MergeCells = True
    Range("L10") = "Mass of the tooling"
    If N_Outillage > 1 Then Range("L10") = "Mass of the toolings"
    Range("O10", "O10").MergeCells = True
    Range("O10") = ":"
    Range("P10", "Q10").MergeCells = True
    If N_Outillage > 3 Then Range("P10", "R10").MergeCells = True
    Range("P10") = Resultat_M & " kg"
    If N_Outillage > 1 Then Range("P10") = Resultat_M & " = " & M_Limite & " kg"
    Range("L12", "N12").MergeCells = True
    Range("L12") = "Position on the shuttle in x"
    If N_Outillage > 1 Then Range("L12") = "Positions on the shuttles in x"
    Range("O12", "O12").MergeCells = True
    Range("O12") = ":"
    Range("P12", "Q12").MergeCells = True
    If N_Outillage > 3 Then Range("P12", "R12").MergeCells = True
    Range("P12") = P_navette_X & " mm"
    Range("L14", "N14").MergeCells = True
    Range("L14") = "Position on the shuttle in y"
    If N_Outillage > 1 Then Range("L14") = "Positions on the shuttles in y"
    Range("O14", "O14").MergeCells = True
    Range("O14") = ":"
    Range("P14", "Q14").MergeCells = True
    If N_Outillage > 3 Then Range("P14", "R14").MergeCells = True
    Range("P14") = P_navette_Y & " mm"
    If C_navette = "Non validées" Then Range("P8", "P14").Font.Color = RGB(210, 125, 0)
    Range("L29", "N29").MergeCells = True
    Range("L29") = "Number of Shuttles"
    Range("O29", "O29").MergeCells = True
    Range("O29") = ":"
    Range("P29", "Q29").MergeCells = True
    If N_Outillage > 3 Then Range("P29", "R29").MergeCells = True
    Range("P29") = N_Navette
    Range("L31", "N31").MergeCells = True
    Range("L31") = "Distance E"
    Range("O31", "O31").MergeCells = True
    Range("O31") = ":"
    Range("P31", "Q31").MergeCells = True
    If N_Outillage > 3 Then Range("P31", "R31").MergeCells = True
    Range("P31") = Resultat_E & " mm"
    If N_Outillage > 1 Then
    Range("L31") = "Distances E"
    End If
    If Hypothèses_Courbe = "Non validées" Then Range("P10, P31").Font.Color = RGB(210, 125, 0)
    Range("L49", "Q49").MergeCells = True
    Range("L49") = Validation & Date & "."
    If Validation <> "The configurations are validated. Made the " Then
    Range("L49").Interior.Color = RGB(255, 194, 105)
    Range("L49", "Q50").MergeCells = True
    Else: Range("L49").Interior.Color = RGB(183, 216, 160)
    End If
    Range("L49").Font.Bold = True
    
    ActiveSheet.Shapes("Image 87").Copy
    Range("K53").Select
    ActiveSheet.Paste
    With Selection.ShapeRange
        .Left = 485
    End With
    N_Objet = N_Objet + 1
    
    ActiveSheet.Shapes("Image 89").Copy
    Range("P53").Select
    ActiveSheet.Paste
    With Selection.ShapeRange
        .Left = 760
        .Top = 795
    End With
    N_Objet = N_Objet + 1
    
    ActiveSheet.Shapes("Image 84").Copy
    Range("L33").Select
    ActiveSheet.Paste
    With Selection.ShapeRange
        .Left = 526
        .Top = 486.2
        .ScaleHeight 1.5, mostrue
        .Line.Weight = 1.5
        .Line.ForeColor.RGB = RGB(10, 10, 10)
    End With
    N_Objet = N_Objet + 1
    
    Do While N_Cycle_M_E_Long < N_Outillage
    N_Cycle_M_E = N_Outillage + N_Cycle_M_E_Long
    Resultat_En = Resultat_E
    Resultat_Mn = Resultat_M
    If N_Cycle_M_E_Long <> 0 Then
    Do While N_Cycle_M_E > N_Outillage
    Resultat_En = Mid(Resultat_En, InStr(Resultat_En, "I") + 1)
    Resultat_Mn = Mid(Resultat_Mn, InStr(Resultat_Mn, "+") + 1)
    N_Cycle_M_E = N_Cycle_M_E - 1
    Loop
    End If
    N_Cycle_M_E_Long = N_Cycle_M_E_Long + 1
    If N_Cycle_M_E_Long < N_Outillage Then
    Resultat_En = StrReverse(Mid(StrReverse(Resultat_En), Len(Resultat_En) - InStr(Resultat_En, "I") + 2))
    Resultat_Mn = StrReverse(Mid(StrReverse(Resultat_Mn), Len(Resultat_Mn) - InStr(Resultat_Mn, "+") + 2))
    End If
    If Resultat_En > 500 Then Resultat_En = 500.1
    If Resultat_Mn > 70 Then Resultat_Mn = 70.1
    ActiveSheet.Shapes.AddShape(msoShapeOval, (560 + (Resultat_En * 0.5845)), (635.2 - (Resultat_Mn * 1.9155)), 5, 5).Select
    With Selection.ShapeRange
        .Fill.ForeColor.RGB = RGB(165, 42, 42)
        .Line.Visible = msoFalse
    End With
    N_Objet = N_Objet + 1
    If Resultat_En < 500 Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, (562.5 + (Resultat_En * 0.5845)), (637.7 - (Resultat_Mn * 1.9155)), (562.5 + (Resultat_En * 0.5845)), 637.7).Select
    With Selection.ShapeRange.Line
        .ForeColor.RGB = RGB(165, 42, 42)
        .Weight = 2
    End With
    N_Objet = N_Objet + 1
    End If
    If Resultat_Mn < 70 Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, (562.1 + (Resultat_En * 0.5845)), (637.7 - (Resultat_Mn * 1.9155)), 562.1, (637.7 - (Resultat_Mn * 1.9155))).Select
    With Selection.ShapeRange.Line
        .ForeColor.RGB = RGB(165, 42, 42)
        .Weight = 2
    End With
    N_Objet = N_Objet + 1
    End If
    Loop
    
    ActiveSheet.Shapes("Image 93").Copy
    Range("M16").Select
    ActiveSheet.Paste
    With Selection.ShapeRange
        .Left = 609
        .Top = 219.5
        .ScaleHeight 1.5, mostrue
    End With
    N_Objet = N_Objet + 1
    
    Do While N_Cycle_X_Y_Long < N_Outillage
    N_Cycle_X_Y = N_Outillage + N_Cycle_X_Y_Long
    x = P_navette_X
    y = P_navette_Y
    If N_Cycle_X_Y_Long <> 0 Then
    Do While N_Cycle_X_Y > N_Outillage
    x = Mid(x, InStr(x, "I") + 1)
    y = Mid(y, InStr(y, "I") + 1)
    N_Cycle_X_Y = N_Cycle_X_Y - 1
    Loop
    End If
    N_Cycle_X_Y_Long = N_Cycle_X_Y_Long + 1
    If N_Cycle_X_Y_Long < N_Outillage Then
    x = StrReverse(Mid(StrReverse(x), Len(x) - InStr(x, "I") + 2))
    y = StrReverse(Mid(StrReverse(y), Len(y) - InStr(y, "I") + 2))
    End If
    ActiveSheet.Shapes.AddShape(msoShapeOval, (686.5 + (x * 0.435)), (309.6 - (y * 0.45)), 5, 5).Select
    With Selection.ShapeRange
        .Fill.ForeColor.RGB = RGB(165, 42, 42)
        .Line.Visible = msoFalse
    End With
    N_Objet = N_Objet + 1
    Loop
    
    ActiveSheet.Protect Password:="Idra01*"

Exporter_PDF:
    CreateObject("WScript.Shell").Run "taskkill.exe /IM AcroRd32.exe /T /F", 0
    CreateObject("WScript.Shell").Run "taskkill.exe /IM Acrobat.exe /T /F", 0
    CreateObject("WScript.Shell").Run "taskkill.exe /IM msedge.exe /T /F", 0
    CreateObject("WScript.Shell").Run "taskkill.exe /IM Chrome.exe /T /F", 0
    PauseTime = 0.5
    Start = Timer
    Do While Timer < Start + PauseTime
    Loop
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Accès & Nom, From:=2, to:=2, Quality:=xlQualityStandard, IncludeDocProperties:=True, OpenAfterPublish:=True

Nettoyer_P2:
    Resultat_N = Null
    ActiveSheet.Unprotect Password:="Idra01*"
    With ActiveSheet.Range("K1:R56")
        .Clear
    End With

    Do While N_Cycle_Eff < N_Objet
    With ActiveSheet
        .Shapes.Range(Array(.Shapes(.Shapes.Count).Name)).Delete
    End With
    N_Cycle_Eff = N_Cycle_Eff + 1
    Loop
    ActiveSheet.Range("K1:R55").Locked = True
    ActiveSheet.Protect Password:="Idra01*"
    Range("C1").Select
    End


Avertissement_p_x:
    Resultat_X = MsgBox("The position, in x, must be between -60mm and 60mm." & Chr(10) & "See diagram.", vbOKOnly + vbCritical + vbDefaultButton1, "x Postion")
    GoTo Position_sur_la_Navette_en_x

Avertissement_p_y:
    Resultat_Y = MsgBox("The position, in y, must be between -100mm and 100mm." & Chr(10) & "See diagram.", vbOKOnly + vbCritical + vbDefaultButton1, "y Postion")
    GoTo Position_sur_la_Navette_en_y

Avertissement_N:
    Resultat_N = MsgBox("The following symbols are not recognized:" & Chr(10) & "                 /        \        :        *        ?", vbOKOnly + vbExclamation + vbDefaultButton1, "Unrecognized Symbols")
    GoTo Saisie_N

Annuler:
    MsgBox "Data verification has been interrupted."
    Resultat_N = Null
    End
    End Sub
