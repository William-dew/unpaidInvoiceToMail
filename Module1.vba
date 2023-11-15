Attribute VB_Name = "Module1"
Option Explicit
Dim wsTmp As Worksheet

' Initialiser la variable wsTmp avec la feuille nomm�e "tmp"
Sub InitailizewsTmp()
    Set wsTmp = ActiveWorkbook.Sheets("tmp")
End Sub

' Proc�dure principale pour d�marrer le processus
Sub start()
    Application.ScreenUpdating = False
    Call cleartmp
    InitailizewsTmp
    copySelection
    Call CreerEmailDepuisFeuilleActive
    Application.ScreenUpdating = True
End Sub

' Copier le contenu de la feuille "Version beta" vers la feuille "tmp"
Sub copySelection()
    Sheets("Version beta").Select
    Selection.Copy
    Sheets("tmp").Select
    ActiveSheet.Paste
    Sheets("Version beta").Select
    Application.CutCopyMode = False
    Range("A1").Select
End Sub

' Nettoyer le contenu de la feuille "tmp" et redimensionner le tableau
Sub cleartmp()
    Sheets("tmp").Select
    Range("A2:F100").Select
    Selection.Clear
    ActiveSheet.ListObjects("Tableau1").Resize Range("$A$1:$F$2")
    Range("A2").Select
End Sub

' Fonction pour obtenir les adresses e-mail de la feuille "tmp"
Function getMailAdress()
    Dim EmailAddr As String
    Dim i As Integer

    i = 2 ' Commencer � la ligne 2

    ' Continuer jusqu'� trouver une cellule vide
    Do While wsTmp.Cells(i, 1).Value <> ""
        EmailAddr = EmailAddr & wsTmp.Cells(i, 1).Value & ";"
        i = i + 1
    Loop
    getMailAdress = EmailAddr
End Function

' Fonction pour obtenir les num�ros de facture de la feuille "tmp"
Function getBillNumber() As Variant
    Dim wsTmp As Worksheet
    Set wsTmp = Sheets("tmp")
    Dim billNumber() As Variant
    Dim i As Integer

    i = 2
    ReDim billNumber(1 To 1) ' Initialiser le tableau

    ' Parcourir les cellules jusqu'� trouver une cellule vide
    Do While wsTmp.Cells(i, 3).Value <> ""
        If i > 2 Then
            ReDim Preserve billNumber(1 To UBound(billNumber) + 1)
        End If
        billNumber(UBound(billNumber)) = wsTmp.Cells(i, 3).Value
        i = i + 1
    Loop

    If UBound(billNumber) > 1 Then
        ReDim Preserve billNumber(1 To UBound(billNumber))
    End If

    getBillNumber = billNumber
End Function

' Cr�er et configurer un e-mail � partir des donn�es recueillies
Sub CreerEmailDepuisFeuilleActive()
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim Subj As String
    ' Dim CcAddress As String - si besoin plus tard
    Dim Body As String
    Dim ListeFactures As String
    Dim LastRow As Long
    Dim Plage As Range

    ' Construction de la liste des factures
    ListeFactures = CreerListeAPuces(getBillNumber())
    
    ' Trouver la derni�re ligne remplie dans la colonne F de la feuille "tmp"
    LastRow = ThisWorkbook.Sheets("tmp").Cells(ThisWorkbook.Sheets("tmp").Rows.Count, "F").End(xlUp).Row

    ' S�lectionner la plage de B1 � F et la derni�re ligne remplie
    Set Plage = ThisWorkbook.Sheets("tmp").Range("B1:F" & LastRow)
    
    ' Initialiser l'application Outlook et cr�er un nouvel e-mail
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MItem = OutlookApp.CreateItem(0)

    ' D�finir le sujet, le destinataire, le corps du message et l'afficher
    Subj = " Rappel de paiement pour: " & TableauEnChaine(getBillNumber, " - ")
    ' ... [Votre code existant] ...

    ' Construire le corps du message en HTML
    Body = "<html><body>"
    Body = Body & "<p>Bonjour,</p>"
    Body = Body & "<p>Je vous �cris pour vous rappeler le paiement en attente des factures suivantes :</p>"
    Body = Body & RangetoHTML(Plage)
    Body = Body & "<p>Nous vous serions reconnaissants de bien vouloir v�rifier ces points et de proc�der au paiement d�s que possible.</p>"
    Body = Body & "<p>Pour tout renseignement compl�mentaire, n'h�sitez pas � me contacter.</p>"
    Body = Body & "<p>Merci pour votre attention.</p>"
    Body = Body & "<p>Bien � vous,</p>"
    Body = Body & "<p>Signature</p>"
    Body = Body & "</body></html>"

 
    With MItem
        .Subject = Subj
        .To = getMailAdress()
        .BCC = "r.valton@tecnomaster.fr"
        .HTMLBody = Body
        .Display
    End With

    ' Nettoyage des objets Outlook
    Set MItem = Nothing
    Set OutlookApp = Nothing
End Sub

