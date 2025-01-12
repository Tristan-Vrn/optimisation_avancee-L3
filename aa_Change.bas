Attribute VB_Name = "aa_Change"
Option Explicit
Option Base 1

Sub tx_change()

Dim wsST As Worksheet
Dim wsCH As Worksheet
Dim wbSource1 As Workbook
Dim wbSource2 As Workbook
Dim adresse As String
Dim ws As Worksheet

Dim nbCol As Long
Dim k As Integer
Dim j As Integer

'adresse = Application.GetOpenFilename
'Set wbSource1 = Workbooks.Open(adresse)

'On attribue le classeur contenant la feuille de change a notre variable wbSource1
Set wbSource1 = Workbooks.Open("/Users/tristan/Desktop/base de données/données économiques/EUR_USD_SPOT.xlsb")

If wbSource1 Is Nothing Then
        MsgBox "Impossible d'ouvrir le classeur ou le classeur est vide.", vbExclamation
        Exit Sub
End If


'On attribue ensuite la feuille change a wsCH
Set wsCH = wbSource1.Sheets(1)

'On attribue le classeur contenant STOCXX600 _ wbSource2
Set wbSource2 = Workbooks.Open("/Users/tristan/Desktop/base de données/données indices/Stoxx600_dec86_fev20.xlsx")
wbSource2.Activate

'Conditon If pour eviter demultiplier une nouvelle fois les valeurs par le tx de changesi on reexecute la sub par erreur
If wbSource1.Worksheets(1).Name <> "tx_change" Then
    
    'Boucle sur toutes les feuilles du classeur STOXX600
    For Each ws In wbSource2.Worksheets
     
        'Calcul du nombre de colonnes pour chaque feuille
         nbCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column - 1
        
        'Condition pour ne pas executer le programme sur les ratios
        If InStr(1, ws.Name, "_To_", vbTextCompare) = 0 And InStr(1, ws.Name, "pe", vbTextCompare) = 0 And _
        InStr(1, ws.Name, "aggte", vbTextCompare) = 0 And InStr(1, ws.Name, "yield", vbTextCompare) = 0 And InStr(1, ws.Name, "margin", vbTextCompare) = 0 Then
         
            'Boucle sur les colonnes de ws
             For j = 1 To nbCol
                    
                    
                    'Boucle sur les lignes de la colonne k+1
                    For k = 1 To 399
                        
                        'Condition pour eviter les cellules vides
                        If ws.Cells(k + 1, j + 1).Value <> "" Then
                            ws.Cells(k + 1, j + 1).Value = wsCH.Cells(k + 1, 2).Value * ws.Cells(k + 1, j + 1).Value
                        
                        End If
           
                    Next k
            Next j
        
        End If
        
    'Formatage des valeurs des feuilles modifi_es a 2 decimales
    ws.Cells(2, 2).Resize(nbCol, 1000).NumberFormat = "0.00"
    
    Next ws

    'Permet de cr_er la condition pour ne pas executer plusieurs fois la sub par erreur
    wbSource1.Worksheets(1).Name = "tx_change"
    wbSource1.Save

End If

'Fermeture du classeur contenant le taux de change
wbSource1.Close

'Enregistrement puis fermeture du classeur STOXX600
wbSource2.Save
wbSource2.Close

End Sub
