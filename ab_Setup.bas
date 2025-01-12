Attribute VB_Name = "ab_Setup"
Option Explicit
Option Base 1
'Procedure principale qui va permettre de creer les feuilles de rendements et les matrices de covariances pour chaque indice et chaque sous periode en appelant notamment d'autres procedures
Sub Rendement()

Dim x(1 To 3) As Workbook
Dim ws As Worksheet
Dim wb As Variant

Dim wbST As Workbook
Dim wbMS As Workbook
Dim wbSP As Workbook
Dim wsR As Worksheet

Dim i As Integer
Dim j As Long
Dim nbSec As Integer
Dim nbRow As Long

'Nettoyage du workbook
Call sup

'Appel de procedure preliminaire de conversion du taux de change
Call tx_change

'Attriubtion de chaque classeur d'indice a une variable wb
Set wbST = Workbooks.Open("/Users/tristan/Desktop/base de donn_es/donn_es indices/Stoxx600_dec86_fev20.xlsx")
Set wbSP = Workbooks.Open("/Users/tristan/Desktop/base de donn_es/donn_es indices/S&P500_dec86_fev20.xlsx")
Set wbMS = Workbooks.Open("/Users/tristan/Desktop/base de donn_es/donn_es indices/MSCI_World_secteurs_dec85_fev20.xlsx")

'Creation d'un vecteur de classeur pour pouvoir faire une boucle
Set x(1) = wbST
Set x(2) = wbSP
Set x(3) = wbMS

'Boucle sur les classeurs
For Each wb In x
    
    'Creation d'une feuille de rendement pour chaque indice
    Set wsR = ThisWorkbook.Worksheets.Add
    
    'Set ws = wb.Worksheets(wb.Worksheets.Count)
    Set ws = wb.Worksheets(27)
    
    'Attribution a la feuille de rendements de chaque indice d'un nom
    wsR.Name = "Rendements_" & Mid(wb.Name, 1, 6)
    
    'Calcul du nombre de secteurs et de date
    nbSec = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column - 1
    nbRow = ws.Cells(ws.Rows.Count, 2).End(xlUp) - 1
    
    'Report des intitules
    wsR.Cells(1, 2).Resize(1, nbSec).Value = ws.Cells(1, 2).Resize(1, nbSec).Value
    'Report des dates
    wsR.Cells(1, 1).Resize(nbRow + 1, 1).Value = ws.Cells(1, 1).Resize(nbRow + 1, 1).Value
    
    'Boucle sur chaque secteur
    For i = 1 To nbSec
    
        'Boucle a chaque date pour le i-eme secteur
        For j = 2 To nbRow
        
        'Condition pour eviter les cellules vides
        If ws.Cells(j, i + 1).Value = "" Or ws.Cells(j + 1, i + 1).Value = "" Then
        Else
            wsR.Cells(j + 1, i + 1).Value = ws.Cells(j + 1, i + 1).Value / ws.Cells(j, i + 1).Value - 1
        End If
        
        'Prochaine date
        Next j
        
    'Prochain secteur
    Next i

    'Fermeture du classeur de l'indice
    wb.Close

'Prochain classeur
Next wb

'Appel de la procedure de tri des secteurs
Call Tri(ThisWorkbook)

'Appel de la procedure pour calculer les matrices covariance
Call MatVar(ThisWorkbook)

End Sub
'Procedure qui va permettre de selectionner uniquement les indices ayant un historique convenable
Sub Tri(wb As Workbook)

Dim ws As Variant
Dim wsR As Worksheet
Dim x(1 To 3) As Worksheet
Dim nbEmpty() As Long
Dim nbCol As Integer
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim column As Range

'Creation d'une nouvelle feuille qui nous servira par la suite et sur laquelle on va reporter au prealable les secteurs qui ont ete supprime pour chaque indice
Set wsR = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
wsR.Name = "Optimisation"

'Attribution des feuilles de rendements de chque indice dans un vecteur de worksheet
Set x(1) = ThisWorkbook.Worksheets("Rendements_MSCI_W")
Set x(2) = ThisWorkbook.Worksheets("Rendements_S&P500")
Set x(3) = ThisWorkbook.Worksheets("Rendements_Stoxx6")

'Boucle sur les feuilles du classeur ayant ete mis en argument dans la sub
For Each ws In x
    
    'Calcul du nombre de secteurs
    nbCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column - 1
    
    'Vecteur qui compte le nombre de cellule vides entre la ligne des noms des secteurs et la cellule non vide la plus ancienne pour chaque secteurs
    ReDim nbEmpty(1 To nbCol)
    
    'Boucle sur les secteurs qui remplit le vecteur nbEmpty
    For i = 1 To nbCol
        nbEmpty(i) = ws.Cells(2, 1 + i).End(xlDown).Row
    Next i
    
    'Application de la regle de decision a partir de nbEmpty pour chaque secteur
    k = 0
    For i = 1 To nbCol
        
        If nbEmpty(i) > Application.WorksheetFunction.Mode(nbEmpty) + 1 Then
            
            'Report du nom du secteur supprime sur la feuille optimisaton
            wsR.Cells(1 + j, k + 2).Value = ws.Cells(1, i - k + 1).Value
            'Suppresion de la colonne
            ws.Columns(i - k + 1).Delete
            
            'k permet de compenser le decalage indiciel des colonnes survenant apres l'execution de la ligne precedente
            k = k + 1
        
        End If
        
    'Prochaine colonne
    Next i

'Mise en place des intitules pour chaque indice de quels secteurs ont ete supprime
wsR.Cells(j + 1, 1).Value = "Secteurs du " & Mid(ws.Name, 11, 7) & " exclus :"
wsR.Cells(j + 1, 1).Font.Bold = True

'Permet de reiterer le report des secteurs supprimes pour le prochain indice une ligne en dessous du precedent (sur la feuille optimsation)
j = j + 1

'Prochaine feuille
Next ws

End Sub

Sub MatVar(wb As Workbook)

Dim x(1 To 3) As Worksheet
Dim y(1 To 3, 1 To 6) As Variant
Dim ws As Variant
Dim Mat As Range
Dim wsN As Worksheet
Dim Secteurs As Range
Dim adresse1 As Variant
Dim adresse2 As Variant

Dim nbC As Integer
Dim nbR As Long
Dim nbRD As Long

Dim i As Integer
Dim j As Integer
Dim k As Long

Set x(1) = ThisWorkbook.Worksheets("Rendements_MSCI_W")
Set x(2) = ThisWorkbook.Worksheets("Rendements_S&P500")
Set x(3) = ThisWorkbook.Worksheets("Rendements_Stoxx6")

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


'Boucle sur la feuille de rendement de chaque indice
For i = 1 To 3

    'Attribution a la variable wsde la feuille de rendemnt de l'indice i
    Set ws = x(i)
    
    'Creation d'une nouvelle feuille pour chaque indice servant a contenir les differentes matrices de covariances
    Set wsN = ThisWorkbook.Worksheets.Add
    
    'Attribtution d'un nom propre a chaque indice de cette feuille
    wsN.Name = "CoVar_" & Mid(ws.Name, 11, 7)
    
    'Nombre de secteurs (colonnes)
    nbC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column - 1
    
    '%%%%%%%%%%%%%%%%%%%%   Procedure pour la matrice de covariances pour la periode TOTALE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    
    'Recherche de la localisation des cellules contenant les deux dates dans la feuille de rendement grace a la fonction FIND
    Set adresse1 = ws.Columns(1).Find(What:=y(i, 1), LookIn:=xlValues)
    Set adresse2 = ws.Columns(1).Find(What:=y(i, 6), LookIn:=xlValues)
    
        If Not adresse1 Is Nothing And Not adresse2 Is Nothing Then
        
        'Report du numero de ligne  des deux dates dans un integer
        
        nbRD = adresse1.Row                 'nbRD = ligne pour la date de depart
        nbR = adresse2.Row                    'nbR = ligne pour la date finale
        
        'Definiton du range des rendements sur lesquels la fonction cov va se baser
        Set Mat = ws.Cells(nbRD, 2).Resize(nbR - nbRD, nbC)
        
        'Incorporation des noms des secteurs dans un range
        Set Secteurs = ws.Cells(1, 2).Resize(1, nbC)
        
        'Report des intitules
        wsN.Cells(1, 1).Value = "Var " & y(i, 1) & " - " & y(i, 6)
        wsN.Cells(1, 1).Font.Bold = True
         wsN.Cells(1, 2).Resize(1, nbC).Value = Secteurs.Value
         wsN.Cells(2, 1).Resize(nbC, 1).Value = Application.WorksheetFunction.Transpose(Secteurs)
         
         'Creation de la matrice des covariances
         wsN.Cells(2, 2).Resize(nbC, nbC) = cov(Mat)
         
         'Attribution d'un nom au range de la matrice
         wsN.Activate
        Cells(2, 2).Resize(nbC, nbC).Name = "cov_" & Mid(ws.Name, 15, 3)
        
        'Mise en page
        wsN.Cells(1, 1).Resize(nbC + 1, nbC + 1).Borders.LineStyle = xlContinuous

        
        End If
    
        
        '%%%%%%%%%%%%%%%%%%%%%%%%% Procedure pour la matrice de covariances pour CHAQUE SOUS PERIODE %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    
    'On reste sur la meme feuille mais avec un saut de 30 lignes a chaque nouvelle iteration
    k = 30
    
    'Boucle sur les 5 sous periodes
    For j = 1 To 5
        
        
    'Recherche de la localisation des cellules contenant les deux dates des sous periodes CONSECUTIVES dans la feuille de rendement grace a la fonction FIND

    Set adresse1 = ws.Columns(1).Find(What:=y(i, j), LookIn:=xlValues)
    Set adresse2 = ws.Columns(1).Find(What:=y(i, j + 1), LookIn:=xlValues)
    
        If Not adresse1 Is Nothing And Not adresse2 Is Nothing Then
        
        
       'Report du numero de ligne  des deux dates dans un integer
        nbRD = adresse1.Row
        nbR = adresse2.Row
        
        'Definiton du range des rendements sur lesquels la fonction cov va se baser
        Set Mat = ws.Cells(nbRD, 2).Resize(nbR - nbRD, nbC)
        
        'Incorporation des noms des secteurs dans un range
        Set Secteurs = ws.Cells(1, 2).Resize(1, nbC)
        
        'Report des intitules de chaque matrice covariance (selon chaque sous periode)
        wsN.Cells(k, 1).Value = "Var " & y(i, j) & " - " & y(i, j + 1)
        'Mise en gras
        wsN.Cells(k, 1).Font.Bold = True
        
        'Report des secteurs
         wsN.Cells(k, 2).Resize(1, nbC).Value = Secteurs.Value
         wsN.Cells(k + 1, 1).Resize(nbC, 1).Value = Application.WorksheetFunction.Transpose(Secteurs)
         
         'Creation de la matrice des covariances
         wsN.Cells(k + 1, 2).Resize(nbC, nbC) = cov(Mat)
         
         'Attribution d'un nom au range de la matrice
         wsN.Activate
        Cells(k + 1, 2).Resize(nbC, nbC).Name = "cov_" & Mid(ws.Name, 15, 3) & "_periode_" & j
        
        'Mise en forme
        wsN.Cells(k, 1).Resize(nbC + 1, nbC + 1).Borders.LineStyle = xlContinuous
        
        End If
        
        'Saut de 30 lignes entre chaque matrice de sous periode pour un indice
         k = k + 30
         
    'Pochaine sous periode
    Next j
    
    'Mise en page generale
    wsN.Columns.AutoFit
    ActiveWindow.DisplayGridlines = False
    
'Prochain indice
Next i

End Sub
'Fonction classique de calcul de covariances
Function cov(plage As Range)
    
    Dim i As Long
    Dim j As Long
    Dim largeur As Long
    Dim plage_1 As Range
    Dim plage_2 As Range
    Dim Result()

    largeur = plage.Columns.Count
    ReDim Result(1 To largeur, 1 To largeur)
    
    For i = 1 To largeur
        For j = 1 To largeur
            Set plage_1 = plage.Cells(1, i).Resize(plage.Rows.Count, 1)
            Set plage_2 = plage.Cells(1, j).Resize(plage.Rows.Count, 1)
            Result(i, j) = Application.WorksheetFunction.Covariance_S(plage_1, plage_2)
        Next j
    Next i

    cov = Result
    
End Function
'Procedure servant a nettoyer complement le workbook pour reexecuter la sub principale
Sub sup()
    
    Dim i As Integer
    Dim k As Integer
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    For i = 1 To wb.Worksheets.Count - 1
    Application.DisplayAlerts = False
        wb.Worksheets(i - k).Delete
        k = k + 1
        Application.DisplayAlerts = True
    Next i
    wb.Worksheets(1).Name = "a"
    Range("A1:ZZ100").Clear
    
End Sub
