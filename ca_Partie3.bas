Attribute VB_Name = "ca_Partie3"
Option Explicit
Option Base 1
'Cette procedure reprend le code de la partie 2 en ajoutant une boucle sur les valeurs de m a la place de celle sur les indices
Sub Evaluation2()

Dim wsEV As Worksheet, wb As Workbook, ws As Worksheet, wsCA As Worksheet, wsG As Worksheet
Dim i As Integer, j As Integer, t As Integer, m As Double
Dim x() As Variant, y As Variant
Dim k As Integer, c As Long
Dim nbSec As Integer
Dim nbD As Integer
Dim observ As Long

Dim AR(1 To 4) As Double
Dim adresse As Variant
Dim nbR As Long
Dim nbRD As Long

Dim cellule As Range
Dim r As Range
Dim r1 As Range
Dim r2 As Range



'Attribution de chaque coefficient d'aversion au risque au vecteur AR()
AR(1) = 1
AR(2) = 2
AR(3) = 4
AR(4) = 20

'Definiton de la date de DEPART : on reprend celle de la partie 2 en ajoutant 72 mois
 
 'MSCI
y = "28/02/2001"


'Definition du classeur actif comme wb
    Set wb = ThisWorkbook
    
    'Definition de la feuille de calcul "Optimisation" comme wsC
    Set wsCA = wb.Worksheets("Calcul")

    'Ajout d'une nouvelle feuille de calcul
    Set wsEV = wb.Worksheets.Add
    
    'On renomme la nouvelle feuille de calcul "evaluation"
    wsEV.Name = "Evaluation_2.0"
    
    'Deplacement de la feuille "Evaluation" aprs la feuille "Comparaison""
    wsEV.Move before:=wsCA
  
      

 Set ws = ThisWorkbook.Worksheets("Rendements_MSCI_W")
 
 'Nombre de secteurs dans l'indice
 nbSec = ws.Cells(1, Columns.Count).End(xlToLeft).column - 1

 'Nombre de dates dans l'indice
 nbD = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1

 'Boucle sur chaque vaaleur de m (que l'on divisera ensuite par 2 pour eviter des bugs)
 For m = 0 To 20 Step 1
    
    'Initiliasation de k a 0
    k = 0

     'Mise en forme de l'intitule de la valeur de m
    With wsEV.Cells(c + 2, 1)
        .Value = "M = " & m / 2
        .Interior.Color = vbBlack
        .Font.Color = vbWhite
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    

    'Recherche de la date de depart
     Set adresse = ws.Columns(1).Find(What:=y, LookIn:=xlValues)
    
        'Pour eviter des bugs
        If adresse Is Nothing Then
            Set adresse = ws.Columns(1).Find(What:=CDate(y), LookIn:=xlValues)
        End If
    
    'Mise en place des intitules des indicateurs qui serviront au solveur
    wsEV.Cells(1, 3).Resize(1, 7).Value = Array("rendement", "variance", "variance opti", "EC opti", "", "H", "m")
    'Mise en forme de la case vide
    wsEV.Cells(1, 7).Resize(2, 1).Interior.Color = vbBlack
    
    'Mise en pagedu range qui servira au solveur
    With wsEV.Cells(1, 3).Resize(2, 7)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    
    'Boucle pour reporter les dates (on investit tous les 6 mois et on s'arrete avant les 36 derniers mois)
    For t = adresse.Row To nbD - 36 Step 6
        
        k = 0

        'Boucle sur les degres d'averson
        For j = 1 To 4
             

             'Mise en place de l'intitule sur le degre d'aversion au risque
                With wsEV.Cells(4 + c, 2 + k)
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                    .Interior.Color = RGB(22, 202, 240)
                End With
                If j = 1 Then
                    wsEV.Cells(4 + c, 2 + k).Value = "Offensif"
                ElseIf j = 2 Then
                    wsEV.Cells(4 + c, 2 + k).Value = "Equilibre"
                ElseIf j = 3 Then
                    wsEV.Cells(4 + c, 2 + k).Value = "Conservateur"
                Else
                    wsEV.Cells(4 + c, 2 + k).Value = "Prudent"
                End If
                
            'Report des dates
            wsEV.Cells((t - adresse.Row) / 6 + c + 5, k + 2).Value = ws.Cells(3 + t, 1).Value
            
            'Report des secteurs
            wsEV.Cells(c + 4, k + 3).Resize(1, nbSec).Value = ws.Cells(1, 2).Resize(1, nbSec).Value
            
           With wsCA
                
                'Calcul de la matrice des covariances sur la feuille wsCA grace a la fonciton cov_flexible
                .Cells(1, 1).Resize(nbSec, nbSec).Value = cov_flexible(ws.Cells(3 + t, 1).Value, ws, 72)
                'Attribution d 'un nom a la matrice
                .Cells(1, 1).Resize(nbSec, nbSec).Name = "Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
                
                'Calcul des rendements moyen sur la feuille wsCA grace a la fonctionRdmt
                .Cells(1, 30).Resize(nbSec, 1).Value = Rdmt(ws.Cells(3 + t, 1).Value, 72, ws)
                'Attribution d'un nom
                .Cells(1, 30).Resize(nbSec, 1).Name = "Rdmt_moyen_eval" & "_" & t & "_" & AR(j)
                
           End With
           
           
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Calcul des indicateurs qui serviront au solveur en R1C1 dans le range en haut de la feuille %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       
           With wsEV
           
            'Initialisation des parts du portefeuille en equipondere
            .Cells((t - adresse.Row) / 6 + c + 5, k + 3).Resize(1, nbSec).Value = 1 / nbSec
             'Attribtution d'un nom pour chaque range de parts
            .Cells((t - adresse.Row) / 6 + c + 5, k + 3).Resize(1, nbSec).Name = "Parts_eval" & "_" & t & "_" & AR(j)
           
           
           'RENDEMENT
            .Cells(2, 3).FormulaArray = "=SUMPRODUCT(Parts_eval" & "_" & t & "_" & AR(j) & ", TRANSPOSE(Rdmt_moyen_eval" & "_" & t & "_" & AR(j) & "))"
            .Cells(2, 3).Name = "rdmt_ptf_eval" & "_" & t & "_" & AR(j)
            
            'VARIANCE
            .Cells(2, 4).FormulaArray = "=MMULT(Parts_eval" & "_" & t & "_" & AR(j) & ", MMULT(Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", TRANSPOSE(Parts_eval" & "_" & t & "_" & AR(j) & ")))"
            .Cells(2, 4).Name = "volat_eval" & "_" & t & "_" & AR(j)
        
            'Somme des parts au carre
            .Cells(2, 8).FormulaR1C1 = "=SUMSQ(Parts_eval_" & t & "_" & AR(j) & ")"
           .Cells(c + 4, k + 5 + nbSec).Value = "Somme parts carré"
           
           'Valeur de m
            .Cells(2, 9).Value = m / 2
            
            'VARIANCE OPTIMISEE
            .Cells(2, 5).FormulaR1C1 = "=volat_eval" & "_" & t & "_" & AR(j) & "*(1+" & "RC[4]*RC[3])"
            .Cells(2, 5).Name = "RisqueTotal" & "_" & t & "_" & AR(j)
            
            'EC
            .Cells(2, 6).FormulaR1C1 = "=rdmt_ptf_eval" & "_" & t & "_" & AR(j) & " - (" & AR(j) & " / 2) *RisqueTotal" & "_" & t & "_" & AR(j)
            .Cells(2, 6).Name = "EC_eval" & "_" & t & "_" & AR(j)
            

 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            
            'Formule pour la somme des parts
            .Cells(c + 4, k + 4 + nbSec).Value = "Somme parts"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 4 + nbSec).FormulaR1C1 = "=SUM(Parts_eval" & "_" & t & "_" & AR(j) & ")"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 4 + nbSec).Name = "SommeParts_eval" & "_" & t & "_" & AR(j)
            
        
        
            .Cells(c + 4, k + 6 + nbSec).Value = "Rendement prevu"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 6 + nbSec).Value = wsEV.Range("rdmt_ptf_eval" & "_" & t & "_" & AR(j))
            .Cells((t - adresse.Row) / 6 + c + 5, k + 6 + nbSec).Name = "rdmt_prev_" & "_" & t & "_" & AR(j)
            
            .Cells(c + 4, k + 7 + nbSec).Value = "Variance prevue"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 7 + nbSec).Value = wsEV.Range("RisqueTotal" & "_" & t & "_" & AR(j))
            .Cells((t - adresse.Row) / 6 + c + 5, k + 7 + nbSec).Name = "volat_prev_" & "_" & t & "_" & AR(j)
            
            .Cells(c + 4, k + 8 + nbSec).Value = "EC prevu"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 8 + nbSec).Value = wsEV.Range("EC_eval" & "_" & t & "_" & AR(j))
            .Cells((t - adresse.Row) / 6 + c + 5, k + 8 + nbSec).Name = "EC_prev_" & "_" & t & "_" & AR(j)
        
    
             
         'Activation de la feuille sur laquelle le solveur va optimiser
         .Activate
        
        End With

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% SOLVEUR %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


        'Renitialisation du solveur
        SolverReset
        
        '1er modele
        SolverOk SetCell:=wsEV.Range("EC_eval" & "_" & t & "_" & AR(j)), MaxMinVal:=1, ByChange:=wsEV.Range("Parts_eval" & "_" & t & "_" & AR(j))
    
        'ajout de la contrainte budgetaire
        SolverAdd CellRef:=wsEV.Range("SommeParts_eval" & "_" & t & "_" & AR(j)), Relation:=2, FormulaText:=1
    
        'Interdiction de la vente a decouvert
        solveroptions assumenonneg:=True
        
        'lancement du solver
        SolverSolve userfinish:=True
                 
   '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% REPORT de la valeur des indicateurs depuis le range optimisation en haut de la feuille sur la feuille pour chaque m, date et coefficient d'aversion au risque %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       
       'Report de l'IHH
        wsEV.Cells((t - adresse.Row) / 6 + c + 5, k + 5 + nbSec).FormulaR1C1 = wsEV.Cells(2, 8).Value
        
        'Calcul du rendement effectif
        wsEV.Cells(c + 4, k + 9 + nbSec).Value = "Rendement effectif"
        wsCA.Cells(1, 30).Resize(nbSec, 1).Value = Rdmt(ws.Cells(3 + t, 1).Value, 36, ws, True)
        wsCA.Cells(1, 30).Resize(nbSec, 1).Name = "Rdmt_moyen_eval" & "_" & t & "_" & AR(j)
        wsEV.Cells((t - adresse.Row) / 6 + c + 5, k + 9 + nbSec).Value = wsEV.Range("rdmt_ptf_eval" & "_" & t & "_" & AR(j)).Value
        wsEV.Cells((t - adresse.Row) / 6 + c + 5, k + 9 + nbSec).Name = "rdmt_eff_" & "_" & t & "_" & AR(j)
        
        'Calcul de la variance effectice
        wsEV.Cells(c + 4, k + 10 + nbSec).Value = "Variance effective"
        wsCA.Cells(1, 1).Resize(nbSec, nbSec).Value = cov_flexible(ws.Cells(3 + t, 1).Value, ws, 36, True)
        wsCA.Cells(1, 1).Resize(nbSec, nbSec).Name = "Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
        
        With wsEV
        
            'Report de la variance
            .Cells((t - adresse.Row) / 6 + c + 5, k + 10 + nbSec).Value = wsEV.Range("volat_eval" & "_" & t & "_" & AR(j)).Value
            .Cells((t - adresse.Row) / 6 + c + 5, k + 10 + nbSec).Name = "volat_eff_" & "_" & t & "_" & AR(j)
            
            'Report de l'EC
            .Cells(c + 4, k + 11 + nbSec).Value = "EC effectif"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 11 + nbSec).FormulaR1C1 = "=RC[-4] - " & AR(j) & "/2 * RC[-2]"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 11 + nbSec).Name = "EC_eff_" & "_" & t & "_" & AR(j)
            
            'Report de la variance optimisee
            .Cells(c + 4, k + 12 + nbSec).Value = "Variance opti"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 12 + nbSec).FormulaR1C1 = "=RC[-2]*(1+ " & m & "*RC[-7])"
            
            'Report du ratio de Sharpe
            .Cells(c + 4, k + 13 + nbSec).Value = "Ratio Sharpe"
            .Cells((t - adresse.Row) / 6 + c + 5, k + 13 + nbSec).FormulaR1C1 = "=(RC[-4]-0.0151)/SQRT(RC[-1])"
            
            
            'Mise en page de la case vide
            .Cells(c + 4, k + 3 + nbSec).Interior.Color = vbBlack
            .Cells((t - adresse.Row) / 6 + c + 5, k + 3 + nbSec).Interior.Color = vbBlack
            
            
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Calcul des MOYENNES historiques des indicateurs de performance  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            'Intitule
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + 2).Value = "Moyenne"
            
            'Ratio de Sharpe
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 13).Value = Application.WorksheetFunction.Average(wsEV.Cells(c + 5, k + nbSec + 13).Resize(c + (nbD / 6) - 35, 1).Value)
            
            'Variance optimisee
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 12).Value = Application.WorksheetFunction.Average(wsEV.Cells(c + 5, k + nbSec + 12).Resize(c + (nbD / 6) - 35, 1).Value)
            
            'EC effectif
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 11).Value = Application.WorksheetFunction.Average(wsEV.Cells(c + 5, k + nbSec + 11).Resize(c + (nbD / 6) - 35, 1).Value)
            
            'Variance effective
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 10).Value = Application.WorksheetFunction.Average(wsEV.Cells(c + 5, k + nbSec + 10).Resize(c + (nbD / 6) - 35, 1).Value)
            
            'Rendement effectif
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 9).Value = Application.WorksheetFunction.Average(wsEV.Cells(c + 5, k + nbSec + 9).Resize(c + (nbD / 6) - 35, 1).Value)
            
            'IHH
            .Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 5).Value = Application.WorksheetFunction.Average(wsEV.Cells(c + 5, k + nbSec + 5).Resize(c + (nbD / 6) - 35, 1).Value)
            
        
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
            'Report des valeurs necessaire a la creation des graphiques en bas de la feuille
            'Ratio de Sharpe
            .Cells(1000 + m, 1 + j).Value = wsEV.Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 13).Value
            .Cells(1000 + m, 1).Value = m / 2
            
            'EC
            .Cells(1022 + m, 1 + j).Value = wsEV.Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 11).Value
            .Cells(1022 + m, 1).Value = m / 2
            
            'IHH
            .Cells(1044 + m, 1 + j).Value = wsEV.Cells(Int((nbD / 6)) + c + 4 - 35, k + nbSec + 5).Value
            .Cells(1044 + m, 1).Value = m / 2
        
        End With
        
    
      'Mise en page de chaque "tableau"
      Set r = wsEV.Cells(c + 4, 2 + k).Resize(Int((nbD / 6)) - 34, nbSec + 12)
      With r
      
          .HorizontalAlignment = xlCenter
          .Borders(xlInsideHorizontal).Weight = xlThin
          .Borders(xlInsideVertical).Weight = xlThin
          .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
          
      End With
      
         'Mise en page du range des parts
          Set r1 = wsEV.Cells(c + 5, 3 + k).Resize(Int((nbD / 6)) - 34, nbSec)
          For Each cellule In r1
            If cellule.Value <> 0 Then
                cellule.NumberFormat = "0.00%"
                cellule.Interior.Color = vbYellow
            End If
          Next cellule
        
         'Mise en page du range des indicateurs de performance
          Set r2 = wsEV.Cells(c + 5, 4 + k + nbSec).Resize(Int((nbD / 6)) - 34, 10)
          For Each cellule In r2
              If cellule.Value <> 0 Then
                  cellule.NumberFormat = "0.0000%"
              End If
          Next cellule
    
        'Surlignage en gris des valeurs des indicateurs de performance effective
        With wsEV.Cells(c + 4, k + 9 + nbSec).Resize((nbD / 6) - 35, 3)
            .Interior.Color = RGB(220, 220, 220)
            .Font.Bold = True
        End With
        
        k = k + nbSec + 13
      
      'Prochain degre d'aversion au risque
      Next j
        
    
    'Prochaine date
    Next t
    
     
    c = c + (nbD - adresse.Row) / 6 + 2

'Incrementation de la valeur de m
Next m



'Mise en forme generale
 wsEV.Columns.AutoFit
 ActiveWindow.DisplayGridlines = False

'Appel de la procedure pour la creation des graphiques
 Call Graph(wsEV, 1000, "Ratio de Sharpe")
 Call Graph(wsEV, 1022, "EC")
 Call Graph(wsEV, 1044, "IHH")
    
    
End Sub
'Procedure generant des graphiques de l'indicateur indic selon les donnees a partir de la r-eme ligne de la feuille ws
Sub Graph(ws As Worksheet, r As Integer, indic As String)
    Dim m As Integer
    Dim j As Integer
    Dim chartObj As ChartObject
   
    Dim newWs As Worksheet
    Dim seriesIndex As Integer
    Dim dataRange As Range
    Dim chartSeries As Series
    
    Dim AR(1 To 4) As Integer

AR(1) = 1
AR(2) = 2
AR(3) = 4
AR(4) = 20
    
    ' Creation  d'une nouvelle feuille de calcul pour le graphique
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "Graphique " & indic
    
    ' Creation d'un nouvel objet Chart sur la nouvelle feuille
    Set chartObj = newWs.ChartObjects.Add(Left:=100, Width:=750, Top:=75, Height:=450)
    
    ' Initialisation l'index de série
    seriesIndex = 1
    
    ' Defintion des séries de données pour chaque degre d'aversion au risque
    For j = 1 To 4
        
        
        ' Determinaton du nombre de lignes à considérer pour les données
        m = 21
        
        ' Defition de la plage de données pour chaque valeur de j
        Set dataRange = ws.Range(ws.Cells(r, j + 1), ws.Cells(r + m, j + 1))
        
        ' Ajout de la série au graphique
        Set chartSeries = chartObj.Chart.SeriesCollection.NewSeries
        With chartSeries
            .Values = dataRange
            .XValues = ws.Range(ws.Cells(r, 1), ws.Cells(r + m, 1))
        End With
        
         If j = 1 Then
               chartSeries.Name = "Portefeuille Offensif"
            ElseIf j = 2 Then
                chartSeries.Name = "Portefeuille Equilibre"
            ElseIf j = 3 Then
                chartSeries.Name = "Portefeuille Conservateur"
            Else
                chartSeries.Name = "Portefeuille Prudent"
        End If
                
        
        ' Incrementation l'index de série
        seriesIndex = seriesIndex + 1
    Next j
    
    ' Definition du type de graphique
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = indic & " moyen par niveau d'averson selon m"
    End With
    
End Sub
'Fonction permettant de calculer la matrice des covariances sur un historique de rendements entre une date choisie et un certain nombre (p) de periodes
Function cov_flexible(date_ As String, ws As Worksheet, p As Integer, Optional futur As Boolean)
    
       Dim i As Long
    Dim j As Long
    Dim largeur As Long
    Dim plage As Range
    Dim plage_1 As Range
    Dim plage_2 As Range
    Dim Result()
    Dim adresse As Variant

'Recherche de la cellule contenant la date
Set adresse = ws.Columns(1).Find(What:=date_, LookIn:=xlValues)
    
        'Pour eviter des bugs
        If adresse Is Nothing Then
            Set adresse = ws.Columns(1).Find(What:=CDate(date_), LookIn:=xlValues)
        End If

    'Nombre de secteurs
    largeur = ws.Cells(1, Columns.Count).End(xlToLeft).column - 1
    
    
    'Rediemension du range de la matrice des covariances
    ReDim Result(1 To largeur, 1 To largeur)
    
    'Double boucle sur les secteurs
    For i = 1 To largeur
        For j = 1 To largeur
        
        'Specification si l'on souhaite considerer les p rendements AVANT ou APRES la date entree en argument
        If futur = False Then
            
            Set plage_1 = ws.Cells(adresse.Row - p, 1 + i).Resize(p, 1)
            Set plage_2 = ws.Cells(adresse.Row - p, 1 + j).Resize(p, 1)
            
       Else
                Set plage_1 = ws.Cells(adresse.Row, 1 + i).Resize(p, 1)
                Set plage_2 = ws.Cells(adresse.Row, 1 + j).Resize(p, 1)
       End If
            Result(i, j) = Application.WorksheetFunction.Covariance_S(plage_1, plage_2)
        
        Next j
    Next i
    
    
    cov_flexible = Result
    End Function
    
'Fonction permettant de calculer un range de rendements moyen sur un historique de rendements entre une date choisie et un certain nombre (p) de periodes
Function Rdmt(date_ As String, p As Integer, ws As Worksheet, Optional futur As Boolean)

    
    Dim adresse As Variant
    Dim nbSec As Integer
    Dim i As Integer
    Dim y As Variant, r As Variant

    Set adresse = ws.Columns(1).Find(What:=date_, LookIn:=xlValues)
    
    ' Pour eviter des bugs
    If adresse Is Nothing Then
        Set adresse = ws.Columns(1).Find(What:=CDate(date_), LookIn:=xlValues)
    End If

    'Nombre de secteurs
    nbSec = ws.Cells(1, Columns.Count).End(xlToLeft).column - 1
    
    ReDim r(1 To nbSec)
    For i = 1 To nbSec
        
        'Specification si l'on souhaite considerer les p rendements AVANT ou APRES la date entree en argument
        If futur = True Then
             y = ws.Cells(adresse.Row, 1 + i).Resize(p, 1).Value
        
        Else
        
            ' Recuperation de la serie du secteur
            y = ws.Cells(adresse.Row - p, 1 + i).Resize(p, 1).Value
        
        End If
        
        ' Calcul du rendement moyen
        r(i) = WorksheetFunction.Average(y)
    
    Next i
    
    ' Retourne le rendement moyen
    Rdmt = Application.WorksheetFunction.Transpose(r)
    
End Function
