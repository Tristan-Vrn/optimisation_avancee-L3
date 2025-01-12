Attribute VB_Name = "ac_Optimisation"
Option Explicit
Option Base 1
'Procedure permettant d'executer la totalite du code
Sub Principale()

    Call Inconditionnels
    
    Call Conditionnels
    
    Call Comparaison

End Sub
'Procedure calculant les 3 portefeuilles inconditonnels
Sub Inconditionnels()

Dim ws As Variant, wsOpt As Worksheet, x(1 To 3) As Worksheet
Dim i As Integer, j As Variant
Dim y() As Variant
Dim k As Integer
Dim nbSec As Integer
Dim observ As Long
Dim AR(1 To 4) As Double
Dim cellule As Range
Dim r As Range

'Attribution de la feuille de rendement de chaque indice au vecteur x()
Set x(1) = ThisWorkbook.Worksheets("Rendements_MSCI_W")
Set x(2) = ThisWorkbook.Worksheets("Rendements_S&P500")
Set x(3) = ThisWorkbook.Worksheets("Rendements_Stoxx6")

'Attribution de chaque coefficient d'aversion au risque au vecteur AR()
AR(1) = 1
AR(2) = 2
AR(3) = 4
AR(4) = 20

'Definition de la feuille "Optimisation"
Set wsOpt = ThisWorkbook.Worksheets("Optimisation")

'Mise en place du titre principal
With wsOpt.Cells(5, 1)
    .Value = "PORTEFEUILLES INCONDITIONNELS"
    .Font.Bold = True
    .Interior.Color = RGB(200, 200, 200)
End With


'Boucle sur les indices
For Each ws In x
    
    'Nombre de secteurs de l'indice
    nbSec = ws.Cells(1, Columns.Count).End(xlToLeft).column - 1
    
    'Nombre d'obersations maximum de l'indice
    observ = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1
    
    'Mise en place des intitules des indices sur la feuille "Optimisation"
    wsOpt.Cells(8, 1 + k).Resize(nbSec, 1).Value = WorksheetFunction.Transpose(ws.Cells(1, 2).Resize(1, nbSec))
    
    'Mise en place du titre pour l'indice
    wsOpt.Cells(7, 1 + k).Value = "Secteurs du " & Mid(ws.Name, 11, 7)
    
    'Mise en place de l'intitule rdmt moyen
    wsOpt.Cells(7, 2 + k).Value = "Rdmt moyen"
    wsOpt.Rows(7).Font.Bold = True

    
    'CALCUL RDMT MOYEN : boucle sur chaque secteur de l'indice
    For i = 1 To nbSec
        
        'recuperation de la serie du secteur
        y = ws.Cells(2, 1 + i).Resize(observ, 1).Value
        
        'calcul du rendement moyen
        wsOpt.Cells(7 + i, 2 + k).Value = WorksheetFunction.Average(y)
    
    'Prochain secteur
    Next i
    
    'Mise en forme et attribution des noms aux celules informatives sur les performances du portefeuille
    With wsOpt
        .Cells(8, 2 + k).Resize(nbSec, 1).Name = "er_" & Mid(ws.Name, 15, 3)
        .Cells(11 + nbSec, k + 1).Value = "EC "
        .Cells(11 + nbSec, k + 1).Font.Bold = True
        .Cells(12 + nbSec, k + 1).Value = "Somme des parts "
        .Cells(12 + nbSec, k + 1).Font.Bold = True
        .Cells(10 + nbSec, k + 1).Value = "Variance ptf  "
        .Cells(10 + nbSec, k + 1).Font.Bold = True
        .Cells(9 + nbSec, k + 1).Value = "Rdmt Ptf "
        .Cells(9 + nbSec, k + 1).Font.Bold = True
        .Cells(8 + nbSec, k + 1).Interior.Color = vbBlack
    End With
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%   CALCUL PARTS OPTIMALES : boucle pour chaque degre d'aversion  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    For j = 1 To 4
    
    With wsOpt
    
        'Mise en place de l'intitule des degre d'aversion au risque
        .Cells(7, 2 + k + j).Value = "Parts avec AR=" & AR(j)
        
        'Attribution d'un nom UNIQUE a chque cellule et range contenant les valeurs des rendements, de variance et d'EC
        .Cells(11 + nbSec, k + 2 + j).Name = "EC_opt" & Mid(ws.Name, 15, 3) & AR(j)
        .Cells(10 + nbSec, k + 2 + j).Name = "Volat_opt" & Mid(ws.Name, 15, 3) & AR(j)
        .Cells(8, 2 + k + j).Resize(nbSec, 1).Name = "parts_opt" & Mid(ws.Name, 15, 3) & AR(j)
        
        'Initialisation des parts en equipondere
        .Range("parts_opt" & Mid(ws.Name, 15, 3) & AR(j)).Value = 1 / nbSec
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Mise en place des formules%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        'Formulation et attribution d'un nom unique a la cellule sommant les parts
        .Cells(12 + nbSec, k + 2 + j).FormulaR1C1 = "=SUM(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & ")"
        .Cells(12 + nbSec, k + 2 + j).Name = "SumParts_" & Mid(ws.Name, 15, 3) & AR(j)
        
        'Formulation et attribution d'un nom unique a la cellule indiquant le rendement du portefeuille
        .Cells(9 + nbSec, k + 2 + j).FormulaR1C1 = "=SUMPRODUCT(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & ",er_" & Mid(ws.Name, 15, 3) & ")"
        .Cells(9 + nbSec, k + 2 + j).Name = "PtfEr_opt" & Mid(ws.Name, 15, 3) & AR(j)
        
        'Formulation de la cellule indiquant la variance du portefeuille (en se servant des noms des differentes matrices de covriances definies dan le module setup)
        .Range("Volat_opt" & Mid(ws.Name, 15, 3) & AR(j)).FormulaArray = "=MMult(Transpose(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "), MMult(cov_" & Mid(ws.Name, 15, 3) & ", parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "))"
        
          'Formulation de la cellule indiquant la valeur de l'EC
        .Range("EC_opt" & Mid(ws.Name, 15, 3) & AR(j)).FormulaR1C1 = "=PtfEr_opt" & Mid(ws.Name, 15, 3) & AR(j) & " - (" & AR(j) & " / 2) * Volat_opt" & Mid(ws.Name, 15, 3) & AR(j)
    
        'Masquage de la ligne vide
        .Cells(8 + nbSec, k + 2 + j).Interior.Color = vbBlack
        .Cells(8 + nbSec, k + 2).Interior.Color = vbBlack
    
    End With


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% SOLVEUR %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    wsOpt.Activate
    
    'Renitialisation du solveur
    SolverReset
    
    '1er modele
    SolverOk SetCell:="EC_opt" & Mid(ws.Name, 15, 3) & AR(j), MaxMinVal:=1, ByChange:="parts_opt" & Mid(ws.Name, 15, 3) & AR(j)
    
    'ajout de la contrainte budgetaire
    SolverAdd CellRef:="SumParts_" & Mid(ws.Name, 15, 3) & AR(j), Relation:=2, FormulaText:=1
    
    'Interdiction de la vente a decouvert
    solveroptions assumenonneg:=True
    
    'lancement du solver
    SolverSolve userfinish:=True
  '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    'Surlignage en jaune des parts calculees non nulles du portefeuille
    For Each cellule In wsOpt.Range("parts_opt" & Mid(ws.Name, 15, 3) & AR(j))
        If cellule.Value <> 0 Then
            cellule.Interior.Color = vbYellow
        Else
            cellule.Interior.Color = vbWhite
        End If
    Next cellule
    
    'Prochain degre d'aversion
    Next j
    
  
   'Definition du tableau des portefeuilles optimaux de l'indice pour chaque degre d'aversion au risque
   Set r = wsOpt.Cells(7, 1 + k).Resize(6 + nbSec, 6)
   
    'Mise en forme generale
   With r
        .HorizontalAlignment = xlCenter
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    End With
    
    'Incrementation de k avant de passer au prochain indice pour decaler les colonnes
     k = k + 7

'Prochain indice
Next ws

'Mise en forme generale
wsOpt.Columns.AutoFit
ActiveWindow.DisplayGridlines = False

End Sub
'Procedure calculant les portfeuilles optimaux pour chaque indice ET chaque SOUS PERIODE
Sub Conditionnels()

Dim ws As Variant, wsOpt As Worksheet, z(1 To 3) As Worksheet
Dim i As Integer, j As Variant, t As Long, a As Integer
Dim x() As Variant, y(1 To 3, 1 To 6) As Variant
Dim k As Integer, c As Long
Dim nbSec As Integer
Dim observ As Long
Dim AR(1 To 4) As Double
Dim adresse1 As Variant
Dim adresse2 As Variant
Dim nbR As Long
Dim nbRD As Long
Dim cellule As Range
Dim r As Range

'Attribution de la feuille de rendement de chaque indice au vecteur z()
Set z(1) = ThisWorkbook.Worksheets("Rendements_MSCI_W")
Set z(2) = ThisWorkbook.Worksheets("Rendements_S&P500")
Set z(3) = ThisWorkbook.Worksheets("Rendements_Stoxx6")

'Attribution de chaque coefficient d'aversion au risque au vecteur AR()
AR(1) = 1
AR(2) = 2
AR(3) = 4
AR(4) = 20

'Attribution a wsOpt de la feuille "Optimisation"
Set wsOpt = ThisWorkbook.Worksheets("Optimisation")

'Mise en place du titre
With wsOpt.Cells(42, 1)
    .Value = "PORTEFEUILLES CONDITIONNELS"
    .Font.Bold = True
    .Interior.Color = RGB(200, 200, 200)
End With

'On definit les dates entre chaque periode pour chaque indice :
' - MSCI WORLD : 31/03/2000, 31/03/2003, 31/10/2007, 27/02/2009
' - S&P 500 : 31/08/2000,28/02/2003, 31/10/2007, 27/02/2009
'- STOXX600 : 31/03/2000, 31/03/2003, 31/05/2007, 27/02/2009

'MSCI
y(1, 1) = "28/02/1995"
y(1, 2) = "31/08/2000"
y(1, 3) = "31/03/2003"
y(1, 4) = "31/10/2007"
y(1, 5) = "27/02/2009"
y(1, 6) = "28/02/2020"

'S&P
y(2, 1) = "31/10/1989"
y(2, 2) = "31/08/2000"
y(2, 3) = "28/02/2003"
y(2, 4) = "31/10/2007"
y(2, 5) = "27/02/2009"
y(2, 6) = "28/02/2020"

'STOXX
y(3, 1) = "30/01/1987"
y(3, 2) = "31/03/2000"
y(3, 3) = "31/03/2003"
y(3, 4) = "31/05/2007"
y(3, 5) = "27/02/2009"
y(3, 6) = "28/02/2020"

'Initialisation de c a la ligne 40
c = 40

'Boucle sur les 3 indices
For i = 1 To 3
    
    'Initialisaton de k a la colonne 0
    k = 0
    'Attribution de ws a la feuille de rendement de l'indice z(i)
    Set ws = z(i)
    'Nombre de secteurs dans l'indice i
    nbSec = ws.Cells(1, Columns.Count).End(xlToLeft).column - 1
    
    'Mise en forme de l'intitule de l'indice
    With wsOpt.Cells(c + 5, 1 + k)
        .Value = Mid(ws.Name, 11, 7)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    'Boucle sur les periodes
    For t = 1 To 5
        
        'Mise en place des intitules des secteurs et des periodes
        wsOpt.Cells(c + 8, 1 + k).Resize(nbSec, 1).Value = WorksheetFunction.Transpose(ws.Cells(1, 2).Resize(1, nbSec))
        wsOpt.Cells(c + 7, 1 + k).Value = "Du " & y(i, t) & " au " & y(i, t + 1)
        wsOpt.Cells(c + 7, 1 + k).Interior.Color = RGB(220, 220, 220)
        wsOpt.Cells(c + 7, 2 + k).Value = "Rdmt moyen"
        wsOpt.Rows(c + 7).Font.Bold = True
    
        'Recherche de la cellule contenant la date y(i,t) dans la feuille des rendements grace a la fonction FIND
        Set adresse1 = ws.Columns(1).Find(What:=y(i, t), LookIn:=xlValues)
        Set adresse2 = ws.Columns(1).Find(What:=y(i, t + 1), LookIn:=xlValues)
        
        'Pour eviter des bugs
        If adresse1 Is Nothing Or adresse2 Is Nothing Then
            Set adresse1 = ws.Columns(1).Find(What:=y(i, t), LookIn:=xlValues)
            Set adresse2 = ws.Columns(1).Find(What:=y(i, t + 1), LookIn:=xlValues)
        End If
        
        'Attricution a nbRD de la ligne de la date de Depart et a nbR la ligne de la date de fin
        nbRD = adresse1.Row
        nbR = adresse2.Row
        
        If Not adresse1 Is Nothing And Not adresse2 Is Nothing Then
            
            For a = 1 To nbSec
    
                'r_cup_ration de la s_rie du secteur
                x = ws.Cells(nbRD, 1 + a).Resize(nbR - nbRD, 1).Value
                
                'calcul du rendement moyen
                wsOpt.Cells(7 + a + c, 2 + k).Value = WorksheetFunction.Average(x)
    
            Next a
     
             'Mise en forme et attribution des noms aux celules informatives sur les performances du portefeuille
             With wsOpt
                .Cells(c + 8, 2 + k).Resize(nbSec, 1).Name = "er_" & Mid(ws.Name, 15, 3) & t
                .Cells(c + 11 + nbSec, k + 1).Value = "EC "
                .Cells(c + 11 + nbSec, k + 1).Font.Bold = True
                .Cells(c + 12 + nbSec, k + 1).Value = "Somme des parts "
                .Cells(c + 12 + nbSec, k + 1).Font.Bold = True
                .Cells(c + 10 + nbSec, k + 1).Value = "Variance ptf : "
                .Cells(c + 10 + nbSec, k + 1).Font.Bold = True
                .Cells(c + 9 + nbSec, k + 1).Value = "Rdmt Ptf "
                .Cells(c + 9 + nbSec, k + 1).Font.Bold = True
                .Cells(c + 8 + nbSec, k + 1).Interior.Color = vbBlack
            End With
                    
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%  CALCUL parts optimales : boucle sur chaque degre d'aversion ET chaque periode %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
             For j = 1 To 4
                
                With wsOpt
                
                    'Mise en place de l'intitule des degre d'aversion au risque
                    .Cells(c + 7, 2 + k + j).Value = "Parts avec AR=" & AR(j)
                    
                    'Attribution d'un nom UNIQUE a chque cellule et range contenant les valeurs des rendements, de variance et d'EC
                    .Cells(c + 11 + nbSec, k + 2 + j).Name = "EC_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
                    .Cells(c + 10 + nbSec, k + 2 + j).Name = "Volat_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
                    .Cells(c + 8, 2 + k + j).Resize(nbSec, 1).Name = "parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
                    
                    'Initialisation des parts en equipondere
                    .Range("parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t).Value = 1 / nbSec
                    
                    'Formulation et attribution d'un nom unique a la cellule sommant les parts
                    .Cells(c + 12 + nbSec, k + 2 + j).FormulaR1C1 = "=SUM(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t & ")"
                    .Cells(c + 12 + nbSec, k + 2 + j).Name = "SumParts_" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
                    
                    'Formulation et attribution d'un nom unique a la cellule indiquant le rendement du portefeuille
                    .Cells(c + 9 + nbSec, k + 2 + j).FormulaR1C1 = "=SUMPRODUCT(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t & ",er_" & Mid(ws.Name, 15, 3) & t & ")"
                    .Cells(c + 9 + nbSec, k + 2 + j).Name = "PtfEr_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
                    
                    'Formulation de la cellule indiquant la variance du portefeuille (en se servant des noms des differentes matrices de covriances definies dan le module setup)
                    .Range("Volat_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t).FormulaArray = "=MMult(Transpose(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t & "), MMult(cov_" & Mid(ws.Name, 15, 3) & "_periode_" & t & ", parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t & "))"
                    
                    'Formulation de la cellule indiquant la valeur de l'EC
                    .Range("EC_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t).FormulaR1C1 = "=PtfEr_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t & " - (" & AR(j) & " / 2) * Volat_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
                
                    'Masquage de la ligne vide
                    .Cells(c + 8 + nbSec, k + 2 + j).Interior.Color = vbBlack
                
                End With
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%% SOLVEUR %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        wsOpt.Activate
        
        'Renitialisation du solveur
        SolverReset
        
        '1er modlee
        SolverOk SetCell:="EC_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t, MaxMinVal:=1, ByChange:="parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t
        
        'ajout de la contrainte budg_taire
        SolverAdd CellRef:="SumParts_" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t, Relation:=2, FormulaText:=1
        
        'Interdiciton de la vente a decouvert
        solveroptions assumenonneg:=True
        
        'lancement du solver
        SolverSolve userfinish:=True
            
       '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        'Surlignage des parts non nulles du portefeuille
        For Each cellule In wsOpt.Range("parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t)
            If cellule.Value <> 0 Then
                cellule.Interior.Color = vbYellow
            Else
                cellule.Interior.Color = vbWhite
            End If
        Next cellule
        
                    
        'Prochain degre d'aversion
        Next j
                  
                  
    End If
    
    'Mise en forme generale
    Set r = wsOpt.Cells(c + 7, 1 + k).Resize(6 + nbSec, 6)
    With r
        .HorizontalAlignment = xlCenter
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    End With
    
    'Decalage des colonnes avant de passer a la periode suivante
    k = k + 7
        
    'Prochaine periode
    Next t
    
    'Decalage des lignes avant de passer a l'indice suivant
    c = c + nbSec + 12

Next i


End Sub
'Procedure qui va comparer les performances des portefeuilles incontionnels avec celles des portefeuilles incondtionnel au sein de chaque sous periode
Sub Comparaison()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsC As Worksheet
    Dim wsOpt As Worksheet
    Dim x(1 To 3) As Worksheet
    Dim y(1 To 3, 1 To 6) As Variant
    Dim AR(1 To 4) As Double
    Dim r As Range
    
    Dim t As Integer
    Dim i As Long
    Dim j As Integer
    Dim k As Long
    Dim c As Integer
    Dim p As Integer
    Dim q As Integer
    
 'MSCI
y(1, 1) = "28/02/1995"
y(1, 2) = "31/08/2000"
y(1, 3) = "31/03/2003"
y(1, 4) = "31/10/2007"
y(1, 5) = "27/02/2009"
y(1, 6) = "28/02/2020"

'S&P
y(2, 1) = "31/10/1989"
y(2, 2) = "31/08/2000"
y(2, 3) = "28/02/2003"
y(2, 4) = "31/10/2007"
y(2, 5) = "27/02/2009"
y(2, 6) = "28/02/2020"

'STOXX
y(3, 1) = "30/01/1987"
y(3, 2) = "31/03/2000"
y(3, 3) = "31/03/2003"
y(3, 4) = "31/05/2007"
y(3, 5) = "27/02/2009"
y(3, 6) = "28/02/2020"

'Attribution de chaque coefficient d'aversion au risque au vecteur AR()
AR(1) = 1
AR(2) = 2
AR(3) = 4
AR(4) = 20


    'Definition du classeur actif comme wb
    Set wb = ThisWorkbook
    'Definition de la feuille de calcul "Optimisation" comme wsOpt
    Set wsOpt = wb.Worksheets("Optimisation")
    
    'Definition de la feuille "comparaison" selon que l'on ai deja executer la procedure ou non
    If wb.Worksheets("Calcul") Is Nothing Then
        'Ajout d'une nouvelle feuille de calcul
        Set wsC = wb.Worksheets.Add
        'Renommer la nouvelle feuille de calcul "Comparaison"
        wsC.Name = "Comparaison"
        'D_placement de la feuille "Comparaison" aprs la feuille "Optimisation""
        wsC.Move After:=wsOpt
    Else
        Set wsC = wb.Worksheets("Calcul")
    End If
    
    Set x(1) = ThisWorkbook.Worksheets("Rendements_MSCI_W")
    Set x(2) = ThisWorkbook.Worksheets("Rendements_S&P500")
    Set x(3) = ThisWorkbook.Worksheets("Rendements_Stoxx6")
    
    
    'Boucle sur les indices
    For i = 1 To 3
        
        'Initialisation a 0 de toutes les variables de mise en page (permettant de decaler ligne/colonne)
        p = 0
        c = 0
        q = 0
        
        'Defintion de ws selon la feuille de l'indice
        Set ws = x(i)
        
        'Mise en places des intitules de chaque indice
        With wsC.Cells(1 + k, 1)
            .Value = Mid(ws.Name, 11, 7)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        'Mise en place des intitules des indicateurs de performances des portefeuilles
        wsC.Cells(5 + k, 1).Resize(3, 1).Value = wsOpt.Cells(34, 1).Resize(3, 1).Value
        wsC.Cells(4 + k, 1).Value = "Type de portefeuille"
        wsC.Cells(3 + k, 1).Value = "Niveau d'aversion au risque"
        
        'Boucle sur les sous periodes
        For t = 1 To 5
            
            'Mise en place des intitules des periodes
            With wsC.Cells(2 + k, 8 + q)
                .Value = "Du " & y(i, t) & " au " & y(i, t + 1)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            wsC.Cells(2 + k, 3 + q).Resize(1, 11).Interior.Color = RGB(220, 220, 22)
        
            'Boucle sur les degre d'aversion
            For j = 1 To 4
            
                'Mise en place de l'intitules sur le degre d'aversion au risque
               Application.DisplayAlerts = False
                With wsC.Cells(3 + k, 3 + p + c)
                    .Resize(1, 2).Merge
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                End With
                Application.DisplayAlerts = True
                If j = 1 Then
                    wsC.Cells(3 + k, 3 + p + c).Value = "Offensif"
                ElseIf j = 2 Then
                    wsC.Cells(3 + k, 3 + p + c).Value = "Equilibre"
                ElseIf j = 3 Then
                    wsC.Cells(3 + k, 3 + p + c).Value = "Conservateur"
                Else
                    wsC.Cells(3 + k, 3 + p + c).Value = "Prudent"
                End If
                            
                'Mise en place des intitules sur le type de portefeuille conditionnel/incontionnel
                With wsC.Cells(4 + k, 3 + c + p)
                    .Value = "CONDITIONNEL"
                    .HorizontalAlignment = xlCenter
                End With
                
                With wsC.Cells(4 + k, 4 + c + p)
                    .Value = "INCONDITIONNEL"
                    .HorizontalAlignment = xlCenter
                End With
                
                
                
                'Report du rendement de chaque ptf CONDITIONNEL depuis la feuille "optimisation"
                wsC.Cells(5 + k, 3 + p + c).Value = wsOpt.Range("PtfEr_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t)
                
                'Report de la variance de chaque ptf CONDITIONNEL depuis la feuille "optimisation"
                wsC.Cells(6 + k, 3 + p + c).Value = wsOpt.Range("Volat_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t)
                
                'Report de l'EC de chaque ptf CONDITIONNEL depuis la feuille "optimisation"
                wsC.Cells(7 + k, 3 + p + c).Value = wsOpt.Range("EC_opt" & Mid(ws.Name, 15, 3) & AR(j) & "_" & t)
                
                
                'CALCUL du rendement de chaque ptf INCONDITIONNEL dans la periode concernee depuis la feuille "optimisation" (parts du ptf inconditionnel * rdmt moyen de la periode t)
                wsC.Cells(5 + k, 4 + p + c).FormulaR1C1 = "=SUMPRODUCT(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & ",er_" & Mid(ws.Name, 15, 3) & t & ")"
                
                'CALCUL de la variance de chaque ptf INCONDITIONNEL
                wsC.Cells(6 + k, 4 + p + c).FormulaArray = "=MMult(Transpose(parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "), MMult(cov_" & Mid(ws.Name, 15, 3) & "_periode_" & t & ", parts_opt" & Mid(ws.Name, 15, 3) & AR(j) & "))"

                 'CALUL de l'EC de chaque ptf INCONDITIONNEL
                wsC.Cells(7 + k, 4 + p + c).FormulaR1C1 = "=R[-2]C - " & AR(j) & "/2 * R[-1]C"
                
                
                '%%%%%%%%%%%%%%%%%%%% Determination pour chaque cas des meilleures performance entre les 2 portefeuilles (d'un point de vue averse au risque) %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                
                'Pour le rendement
                If wsC.Cells(5 + k, 3 + p + c).Value > wsC.Cells(5 + k, 4 + p + c).Value Then
                    wsC.Cells(5 + k, 3 + p + c).Interior.Color = RGB(140, 220, 140)
                    wsC.Cells(5 + k, 4 + p + c).Interior.Color = RGB(255, 140, 140)
                Else
                    wsC.Cells(5 + k, 3 + p + c).Interior.Color = RGB(255, 140, 140)
                    wsC.Cells(5 + k, 4 + p + c).Interior.Color = RGB(140, 220, 140)
                End If
                
                'Pour la variance
                 If wsC.Cells(6 + k, 3 + p + c).Value > wsC.Cells(6 + k, 4 + p + c).Value Then
                    wsC.Cells(6 + k, 3 + p + c).Interior.Color = RGB(255, 140, 140)
                    wsC.Cells(6 + k, 4 + p + c).Interior.Color = RGB(140, 220, 140)
                Else
                    wsC.Cells(6 + k, 3 + p + c).Interior.Color = RGB(140, 220, 140)
                    wsC.Cells(6 + k, 4 + p + c).Interior.Color = RGB(255, 140, 140)
                End If
                
                'Pour l'EC
                 If wsC.Cells(7 + k, 3 + p + c).Value > wsC.Cells(7 + k, 4 + p + c).Value Then
                    wsC.Cells(7 + k, 3 + p + c).Interior.Color = RGB(140, 220, 140)
                    wsC.Cells(7 + k, 4 + p + c).Interior.Color = RGB(255, 140, 140)
                Else
                    wsC.Cells(7 + k, 3 + p + c).Interior.Color = RGB(255, 140, 140)
                    wsC.Cells(7 + k, 4 + p + c).Interior.Color = RGB(140, 220, 140)
                End If
                
                'Mise en page
                Set r = wsC.Cells(3 + k, 3 + p + c).Resize(5, 2)
                r.Borders.Weight = xlThin
                
                c = c + 3
            
            'Prochain degre d'aversion
            Next j
                                        'Incrementation de c,p,q et k (servant a bien disposer les donnees sur la feuille)
            
            p = p + 1
            q = q + 13
        
        'Prochaine sous periode
        Next t
        
        k = k + 10
    
    'Prochain indice
    Next i
    
    
    'Mise en forme generale
    wsC.Columns.AutoFit
    ActiveWindow.DisplayGridlines = False
    
    
End Sub

