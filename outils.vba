Attribute VB_Name = "outils"
Option Explicit

Function TableauEnChaine(tableau As Variant, Optional separateur As String = ", ") As String
    Dim i As Integer
    Dim resultat As String

    ' Vérifier si le tableau est vide
    If IsEmpty(tableau) Then
        TableauEnChaine = ""
        Exit Function
    End If

    ' Initialiser la chaîne de caractères avec le premier élément du tableau
    resultat = tableau(LBound(tableau))

    ' Parcourir le tableau à partir du deuxième élément
    For i = LBound(tableau) + 1 To UBound(tableau)
        ' Ajouter le séparateur suivi de l'élément
        resultat = resultat & separateur & tableau(i)
    Next i

    ' Assigner le résultat à la fonction
    TableauEnChaine = resultat
End Function


Function CreerListeAPuces(tableau As Variant) As String
    Dim resultat As String
    Dim i As Long

    resultat = ""
    For i = LBound(tableau) To UBound(tableau)
        resultat = resultat & "     • " & tableau(i) & vbCrLf
    Next i

    CreerListeAPuces = resultat
End Function

Function RangetoHTML(ByVal rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
 
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
 
        .Cells(1).PasteSpecial Paste:=12
        .Cells(1).PasteSpecial Paste:=-4122
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
            .DrawingObjects.Visible = True
            .DrawingObjects.Delete
            .Columns.AutoFit
             Dim col As Range
            For Each col In .UsedRange.Columns
                col.ColumnWidth = col.ColumnWidth + 5 ' Ajuster la largeur de chaque colonne (ajouter 5 par exemple)
            Next col
            .Rows.AutoFit
        On Error GoTo 0
    End With
 
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
 
    TempWB.Close savechanges:=False
    Kill TempFile
 
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
 
End Function

