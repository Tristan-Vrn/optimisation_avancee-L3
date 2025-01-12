Attribute VB_Name = "ba_Partie2"
Option Explicit
Option Base 1
'Procedure calculant pour chaque indice les portefuilles optimaux tous les 6 mois et comparant ensuite les performances prevues et effectives
Sub Evaluation()

Dim wsC As Worksheet, wsE As Worksheet, z(1 To 3) As Worksheet, wb As Workbook, ws As Worksheet, wsCA As Worksheet
Dim i As Integer, j As Integer, t As Integer
Dim x() As Variant, y(1 To 3) As Variant
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

'Attribution de la feuille de rendement de chaque indice au vecteur z()
Set z(1) = ThisWorkbook.Worksheets("Rendements_MSCI_W")
Set z(2) = ThisWorkbook.Worksheets("Rendements_S&P500")
Set z(3) = ThisWorkbook.Worksheets("Rendements_Stoxx6")

'Attribution de chaque coefficient d'aversion au risque au vecteur AR()
AR(1) = 1
AR(2) = 2
AR(3) = 4
AR(4) = 20

'Definiton des dates de depart : on reprend celle de la partie 1 en ajoutant 72 mois
 'MSCI
y(1) = "28/02/2001"

'S&P
y(2) = "31/10/1995"

'STOXX
y(3) = "29/01/1993"

'Definition du classeur actif comme wb
    Set wb = ThisWorkbook
    'Definition de la feuille de calcul "Optimisation" comme wsC
    Set wsC = wb.Worksheets("Comparaison")
    
   
        'Ajout d'une nouvelle feuille de calcul
        Set wsE = wb.Worksheets.Add
        'Renommer la nouvelle feuille de calcul "evaluation"
        wsE.Name = "Evaluation"
        'D_placement de la feuille "Evaluation" aprs la feuille "Comparaison""
        wsE.Move After:=wsC
  
    
    'Definition de la feuille "calcul"

        'Ajout d'une nouvelle feuille de calcul
        Set wsCA = wb.Worksheets.Add
        'Renommer la nouvelle feuille de calcul "Calcul"
        wsCA.Name = "Calcul"
        'D_placement de la feuille "Calcul" aprs la feuille "Evaluation""
        wsCA.Move After:=wsE

'Boucle sur les indices
For i = 1 To 3
    k = 0

    Set ws = z(i)
     
     'Mise en forme de l'intitule de l'indice
    With wsE.Cells(c + 2, 1)
        .Value = Mid(ws.Name, 11, 7)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    'Nombre de secteurs dans l'indice i
    nbSec = ws.Cells(1, Columns.Count).End(xlToLeft).column - 1
    
    'Nombre de dates dans l'indice i
    nbD = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1
   
    
     Set adresse = ws.Columns(1).Find(What:=y(i), LookIn:=xlValues)
    
        'Pour eviter des bugs
        If adresse Is Nothing Then
            Set adresse = ws.Columns(1).Find(What:=CDate(y(i)), LookIn:=xlValues)
        End If
    
    'Boucle pour reporter les dates (on investit tous les 6 mois et on s'arrete avant les 36 derniers mois)
    For t = adresse.Row To nbD - 36 Step 6
        
        'Initialisation a 0 de k (servant a decaler les lignes entre chaque indice
        k = 0
    
        'Boucle sur les degres d'aversion
        For j = 1 To 4
             
              'Nombre de dates dans l'indice i
            nbD = ws.Cells(Rows.Count, 2).End(xlUp).Row - 1
             
             'Mise en place de l'intitule sur le degre d'aversion au risque
                With wsE.Cells(3 + c, 2 + k)
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                    .Interior.Color = RGB(22, 202, 240)
                End With
                If j = 1 Then
                    wsE.Cells(3 + c, 2 + k).Value = "Offensif"
                ElseIf j = 2 Then
                    wsE.Cells(3 + c, 2 + k).Value = "Equilibre"
                ElseIf j = 3 Then
                    wsE.Cells(3 + c, 2 + k).Value = "Conservateur"
                Else
                    wsE.Cells(3 + c, 2 + k).Value = "Prudent"
                End If
                
            'Report des dates
            wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 2).Value = ws.Cells(3 + t, 1).Value
            
            'Report des secteurs
            wsE.Cells(c + 3, k + 3).Resize(1, nbSec).Value = ws.Cells(1, 2).Resize(1, nbSec).Value
            
           With wsCA
                
                'Calcul de la matrice des covariances sur la feuille wsCA grace a la fonciton cov_flexible
                .Cells(1, 1).Resize(nbSec, nbSec).Value = cov_flexible(ws.Cells(3 + t, 1).Value, ws, 72)
                'Attribution d'un nom a la matrice
                .Cells(1, 1).Resize(nbSec, nbSec).Name = "Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
                
                'Calcul des rendements moyen sur la feuille wsCA grace a la fonctionRdmt
                .Cells(1, 40).Resize(nbSec, 1).Value = Rdmt(ws.Cells(3 + t, 1).Value, 72, ws)
                'Attribution d'un nom
                 .Cells(1, 40).Resize(nbSec, 1).Name = "Rdmt_moyen_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
                 
            End With
           
           
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% CALCUL des indicateurs de performance PREVUE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       
           With wsE
           
           'Initialisation des parts du portefeuille en equipondere
            .Cells((t - adresse.Row) / 6 + c + 4, k + 3).Resize(1, nbSec).Value = 1 / nbSec
            'Attribtution d'un nom pour chaque range de parts
            .Cells((t - adresse.Row) / 6 + c + 4, k + 3).Resize(1, nbSec).Name = "Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
           
           'Calcul des indicateurs en R1C1 en bas de la feuille
           
           'Rendement
            .Cells(150, 1).FormulaArray = "=SUMPRODUCT(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", TRANSPOSE(Rdmt_moyen_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & "))"
            .Cells(150, 1).Name = "rdmt_ptf_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            
            'Variance
            .Cells(151, 1).FormulaArray = "=MMULT(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", MMULT(Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", TRANSPOSE(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ")))"
            .Cells(151, 1).Name = "volat_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            
            'EC
            .Cells(152, 1).FormulaR1C1 = "=rdmt_ptf_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & " - (" & AR(j) & " / 2) * volat_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            .Cells(152, 1).Name = "EC_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            
            'Formule pour la somme des parts
            .Cells((t - adresse.Row) / 6 + c + 4, k + 4 + nbSec).FormulaR1C1 = "=SUM(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ")"
            .Cells((t - adresse.Row) / 6 + c + 4, k + 4 + nbSec).Name = "SommeParts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            Cells(c + 3, k + 4 + nbSec).Value = "Somme parts"
        
       
     'Report des indicateurs calcules precedemment sur le haut de la feuille pour chaque date, indice et degre d'aversion
        
            .Cells(c + 3, k + 5 + nbSec).Value = "Rendement prevu"
            .Cells((t - adresse.Row) / 6 + c + 4, k + 5 + nbSec).Value = wsE.Range("rdmt_ptf_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j))
            .Cells((t - adresse.Row) / 6 + c + 4, k + 5 + nbSec).Name = "rdmt_prev_" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            
            .Cells(c + 3, k + 7 + nbSec).Value = "Variance prevue"
            .Cells((t - adresse.Row) / 6 + c + 4, k + 7 + nbSec).Value = wsE.Range("volat_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j))
            .Cells((t - adresse.Row) / 6 + c + 4, k + 7 + nbSec).Name = "volat_prev_" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
            
            .Cells(c + 3, k + 9 + nbSec).Value = "EC prevu"
            .Cells((t - adresse.Row) / 6 + c + 4, k + 9 + nbSec).Value = wsE.Range("EC_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j))
            .Cells((t - adresse.Row) / 6 + c + 4, k + 9 + nbSec).Name = "EC_prev_" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
        
    
         'Activation de la feuille sur laquelle le solveur va optimiser
         .Activate
        
        End With

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% SOLVEUR %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        'Renitialisation du solveur
        SolverReset
        
        '1er modele
        SolverOk SetCell:=wsE.Range("EC_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)), MaxMinVal:=1, ByChange:=wsE.Range("Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j))
    
        'ajout de la contrainte budgetaire
        SolverAdd CellRef:=wsE.Range("SommeParts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)), Relation:=2, FormulaText:=1
    
        'Interdiction de la vente a decouvert
        solveroptions assumenonneg:=True
        
        'lancement du solver
        SolverSolve userfinish:=True
             
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% CALCUL des indicateurs de performance EFFECTIVE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
           
        '                                                                                                      RENDEMENT
        'Intitule
        wsE.Cells(c + 3, k + 6 + nbSec).Value = "Rendement effectif"
        
         'Calcul des rendements moyen grace a la fonction Rdmt
        wsCA.Cells(1, 30).Resize(nbSec, 1).Value = Rdmt(ws.Cells(3 + t, 1).Value, 36, ws, True)
        
        'Attribution d'un nom unique
        wsCA.Cells(1, 30).Resize(nbSec, 1).Name = "Rdmt_moyen_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
        
        'Formule
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 6 + nbSec).FormulaArray = "=SUMPRODUCT(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", TRANSPOSE(Rdmt_moyen_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & "))"
        
        'Attribution d'un nom unique pour chaque valeur calculee
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 6 + nbSec).Name = "rdmt_eff_" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
    
    
        '                                                                                                        VARIANCE
        'Intitule
         wsE.Cells(c + 3, k + 8 + nbSec).Value = "Variance effective"
         
         'Calcul de la matrice des covariances grace a la fonction cov_flexible
        wsCA.Cells(1, 1).Resize(nbSec, nbSec).Value = cov_flexible(ws.Cells(3 + t, 1).Value, ws, 36, True)
        
        'Attribtution d'un nom a la matrice nouvelleemnt creee
        wsCA.Cells(1, 1).Resize(nbSec, nbSec).Name = "Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
        
        'Formule
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 8 + nbSec).FormulaArray = "=MMULT(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", MMULT(Matcov" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ", TRANSPOSE(Parts_eval" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j) & ")))"
        
         'Attribution d'un nom unique pour chaque valeur calculee
         wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 8 + nbSec).Name = "volat_eff_" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
         
         
         '                                                                                                      EC
        'Intitule
        wsE.Cells(c + 3, k + 10 + nbSec).Value = "EC effectif"
        
        'Formule
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 10 + nbSec).FormulaR1C1 = "=RC[-4] - " & AR(j) & "/2 * RC[-2]"
        
        'Attribution d'un nom unique pour chaque valeur calculee
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 10 + nbSec).Name = "EC_eff_" & Mid(ws.Name, 15, 3) & "_" & t & "_" & AR(j)
        
         
         '                                                                                                      RATIO DE SHARPE
        'Intitule
         wsE.Cells(c + 3, k + 11 + nbSec).Value = "Ratio Sharpe"
         
         'Formule
         wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 11 + nbSec).FormulaR1C1 = "=(RC[-5]-0.0151)/SQRT(RC[-3])"
    
    'Mise en page de la colonne vide
    wsE.Cells(c + 3, k + 3 + nbSec).Interior.Color = vbBlack
    wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 3 + nbSec).Interior.Color = vbBlack
    
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% COMPARAISON entre les valeurs des indicateurs de performance prevue et effective %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    'Rendement
    If wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 5 + nbSec).Value > wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 6 + nbSec).Value Then
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 5 + nbSec).Interior.Color = RGB(140, 220, 140)
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 6 + nbSec).Interior.Color = RGB(255, 140, 140)
    Else
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 5 + nbSec).Interior.Color = RGB(255, 140, 140)
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 6 + nbSec).Interior.Color = RGB(140, 220, 140)
    End If
        
    'Variance
    If wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 7 + nbSec).Value > wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 8 + nbSec).Value Then
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 7 + nbSec).Interior.Color = RGB(255, 140, 140)
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 8 + nbSec).Interior.Color = RGB(140, 220, 140)
    Else
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 7 + nbSec).Interior.Color = RGB(140, 220, 140)
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 8 + nbSec).Interior.Color = RGB(255, 140, 140)
    End If
        
    'EC
    If wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 9 + nbSec).Value > wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 10 + nbSec).Value Then
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 9 + nbSec).Interior.Color = RGB(140, 220, 140)
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 10 + nbSec).Interior.Color = RGB(255, 140, 140)
    Else
        wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 9 + nbSec).Interior.Color = RGB(255, 140, 140)
         wsE.Cells((t - adresse.Row) / 6 + c + 4, k + 10 + nbSec).Interior.Color = RGB(140, 220, 140)
    End If
    
    
    'Calcul des moyennes des indicateurs de performances
    With wsE
        
        'Intitule
        .Cells(Int((nbD / 6)) + c + 3 - 35, k + 2).Value = "Moyenne"
        
        'Ratio de Sharpe
        .Cells(Int((nbD / 6)) + c + 3 - 35, k + nbSec + 10).Value = Application.WorksheetFunction.Average(wsE.Cells(c + 4, k + 10).Resize(c + (nbD / 6) - 35).Value)
        
        'EC effectif
        .Cells(Int((nbD / 6)) + c + 3 - 35, k + nbSec + 9).Value = Application.WorksheetFunction.Average(wsE.Cells(c + 4, k + 9).Resize(c + (nbD / 6) - 35).Value)
        
        'EC prevu
        .Cells(Int((nbD / 6)) + c + 3 - 35, k + nbSec + 8).Value = Application.WorksheetFunction.Average(wsE.Cells(c + 4, k + 8).Resize(c + (nbD / 6) - 35).Value)
        
        'Variance effective
        .Cells(Int((nbD / 6)) + c + 3 - 35, k + nbSec + 7).Value = Application.WorksheetFunction.Average(wsE.Cells(c + 4, k + 7).Resize(c + (nbD / 6) - 35).Value)
        
        'Variance prevue
        .Cells(Int((nbD / 6)) + c + 3 - 35, k + nbSec + 6).Value = Application.WorksheetFunction.Average(wsE.Cells(c + 4, k + 6).Resize(c + (nbD / 6) - 35).Value)
    
    End With
      
    'Centrage des donnees
    wsE.UsedRange.HorizontalAlignment = xlCenter
   

        'Decalage des colonnes
        k = k + nbSec + 11
        
        'Prochain degre d'aversion au risque
        Next j
        
    
    'Prochaine periode
    Next t
    
     
     'Decalage des lignes
    c = c + (nbD - adresse.Row) / 6 + 5

'Prochain indice
Next i

 
'Mise en forme generale
    wsE.Columns.AutoFit
    ActiveWindow.DisplayGridlines = False
    
    
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
