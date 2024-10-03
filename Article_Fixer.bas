Sub ModificarDocumentoWord()
    ' Chama as funções para cada tipo de modificação
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Call InserirContent
    Call ModificarEtiquetasDeContenido
    Call RemoverEtiquetaP
    Call ModificarEtiquetasDeImagenDeBannerActual
    Call ConverterListasEmBulletPoints
    Call ConverterTitulosEmHeadings
    Call AplicarNegritoEConverterTexto
    Call RemoverTextoEspecifico
    Call RemoverTextoEspecifico2

    
    ' Salva o documento
    ActiveDocument.Save
    MsgBox "Documento modificado com sucesso!", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Sub InserirContent()
    Dim wordDoc As Document
    Dim rng As Range
    Dim contentInserted As Boolean
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    contentInserted = False
    
    With rng.Find
        .text = "Nombre de la imagen: *.[jJ][pP][gG]"
        .MatchWildcards = True
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            If Not contentInserted Then
                rng.Collapse Direction:=wdCollapseEnd
                
                ' Insere "CONTENT:" sem herdar estilos
                rng.InsertParagraphAfter
                rng.InsertAfter "CONTENT:" & vbCrLf
                rng.Style = wdStyleNormal
                
                contentInserted = True
            End If
            
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub


Sub ModificarEtiquetasDeContenido()
    Dim wordDoc As Document
    Dim rng As Range
    Dim startRange As Range
    Dim endRange As Range
    Dim idiomas As Variant
    Dim i As Integer
    
    idiomas = Array("ETIQUETAS DE CONTENIDO:", "SEO:", _
                    "ETIQUETAS DE CONTEÚDO:", "SEO:", _
                    "ETIQUETAS DE CONTENIDO:", "SEO:")
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    
    For i = LBound(idiomas) To UBound(idiomas) Step 2
        With rng.Find
            .text = idiomas(i)
            .Replacement.text = idiomas(i + 1)
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                rng.text = idiomas(i + 1)
                rng.Collapse Direction:=wdCollapseEnd
                
                Set startRange = rng.Duplicate
                startRange.Collapse Direction:=wdCollapseStart
                rng.Collapse Direction:=wdCollapseEnd
                
                With rng.Find
                    .text = "URL SUGERIDA: *^13"
                    .MatchWildcards = True
                    .Forward = True
                    .Wrap = wdFindStop
                    If .Execute Then
                        Set endRange = rng.Duplicate
                        endRange.Collapse Direction:=wdCollapseEnd
                        
                        endRange.InsertParagraphAfter
                        endRange.InsertAfter "FIN DE SEO" & vbCrLf
                        endRange.Style = wdStyleNormal
                    End If
                End With
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next i
End Sub

Sub RemoverEtiquetaP()
    Dim wordDoc As Document
    Dim rng As Range
    Dim idiomas As Variant
    Dim i As Integer
    
    idiomas = Array("Etiqueta P: ", "Etiqueta P: ", "Etiqueta P: ")
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    
    For i = LBound(idiomas) To UBound(idiomas)
        With rng.Find
            .text = idiomas(i)
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                rng.text = ""
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next i
End Sub

Sub ModificarEtiquetasDeImagenDeBannerActual()
    Dim wordDoc As Document
    Dim rng As Range
    Dim startRange As Range
    Dim idiomas As Variant
    Dim i As Integer
    
    idiomas = Array("ETIQUETAS DE IMAGEN DE BANNER ACTUAL:", "ETIQUETAS DE IMAGEN:", _
                    "ETIQUETAS DE IMAGEM DO BANNER ATUAL:", "ETIQUETAS DE IMAGEM:", _
                    "ETIQUETAS DE IMAGEN DE BANNER ACTUAL:", "ETIQUETAS DE IMAGEN:")
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    
    For i = LBound(idiomas) To UBound(idiomas) Step 2
        With rng.Find
            .text = idiomas(i)
            .Replacement.text = idiomas(i + 1)
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                rng.text = idiomas(i + 1)
                rng.Collapse Direction:=wdCollapseEnd
                
                Set startRange = rng.Duplicate
                startRange.Collapse Direction:=wdCollapseStart
                rng.Collapse Direction:=wdCollapseEnd
                
                With rng.Find
                    .text = "Nombre de la imagen: *^13"
                    .MatchWildcards = True
                    .Forward = True
                    .Wrap = wdFindStop
                    Do While .Execute
                        rng.Collapse Direction:=wdCollapseEnd
                        rng.InsertParagraphAfter
                        rng.InsertAfter "FIN DE ETIQUETAS" & vbCrLf
                        rng.Style = wdStyleNormal
                        rng.Collapse Direction:=wdCollapseEnd
                    Loop
                End With
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next i
End Sub

Sub ConverterListasEmBulletPoints()
    Dim wordDoc As Document
    Dim para As Paragraph
    Dim listRange As Range
    Dim inEtiquetasSection As Boolean
    Dim itemCount As Integer
    
    Set wordDoc = ActiveDocument
    inEtiquetasSection = False
    
    For Each para In wordDoc.Paragraphs
        If InStr(para.Range.text, "ETIQUETAS DE IMAGEN:") > 0 Or _
           InStr(para.Range.text, "FIN DE ETIQUETAS") > 0 Then
            inEtiquetasSection = Not inEtiquetasSection
        End If
        
        If Not inEtiquetasSection And Left(para.Range.text, 2) = "- " Then
            itemCount = 1
            Set listRange = para.Range.Duplicate
            
            Do While Not para.Next Is Nothing And Left(para.Next.Range.text, 2) = "- "
                Set para = para.Next
                itemCount = itemCount + 1
                listRange.End = para.Range.End
            Loop
            
            If itemCount > 1 Then
                listRange.ListFormat.ApplyBulletDefault
            End If
        End If
    Next para
End Sub

Sub ConverterTitulosEmHeadings()
    Dim wordDoc As Document
    Dim rng As Range
    Dim i As Integer
    Dim estilo As Style
    Dim estiloNames As Variant
    Dim estiloName As Variant
    Dim estiloPrefix As String

    Set wordDoc = ActiveDocument
    estiloNames = Array("Heading ", "Título ", "Encabezado ")

    ' Loop to convert H1 to H5
    For i = 1 To 5
        estiloPrefix = ""
        Set estilo = Nothing
        
        ' Find the correct style based on the language
        For Each estiloName In estiloNames
            On Error Resume Next
            Set estilo = wordDoc.Styles(estiloName & i)
            On Error GoTo 0
            If Not estilo Is Nothing Then
                estiloPrefix = estiloName
                Exit For
            End If
        Next estiloName
        
        ' If the style doesn't exist, create it
        If estilo Is Nothing Then
            estiloPrefix = "Heading "
            Set estilo = wordDoc.Styles.Add(Name:=estiloPrefix & i, Type:=wdStyleTypeParagraph)
        End If
        
        ' Apply the found style to headings
        Set rng = wordDoc.Content
        
        With rng.Find
            .text = "H" & i & ": "
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                rng.text = Replace(rng.text, "H" & i & ": ", "")
                rng.Style = estilo
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next i
End Sub
Sub AplicarNegritoEConverterTexto()
    Dim wordDoc As Document
    Dim rng As Range
    Dim found As Boolean
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    
    ' Aplica negrito e converte os textos específicos
    found = True
    Do While found
        found = False
        
        ' Check and change "Text Alt:"
        With rng.Find
            .text = "Text Alt:"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            
            If .Execute Then
                rng.Font.Bold = True
                rng.text = "Alt text:"
                found = True
            End If
        End With
        
        ' Check and change "Title de la Imagen:" (with variations)
        With rng.Find
            .text = "Title de la Imagen:*" ' Use wildcard to match any variations
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            
            If .Execute Then
                rng.Font.Bold = True
                rng.text = "Title:" ' Change the text
                found = True
            End If
        End With
        
        ' Check and change "Nombre de la imagen:" (with variations)
        With rng.Find
            .text = "Nombre de la imagen:*" ' Use wildcard to match any variations
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            
            If .Execute Then
                rng.Font.Bold = True
                rng.text = "Nombre de la imagen:" ' Ensure correct formatting
                found = True
            End If
        End With
        
        ' Move to the end of the document for the next iteration
        rng.Collapse Direction:=wdCollapseEnd
    Loop
End Sub

Sub RemoverTextoEspecifico()
    ' Remove specific unwanted text with special handling for spaces and accented characters
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    ' Remove the specific instructional text
    With rng.Find
        ' Use wildcards to account for spaces or formatting changes
        .text = "Se debe*copiar*el*código*que*encuentra*dentro*del*recuadro*y*pegarlo*en*sección*<head>*del*documento*HTML*del*sitio*web.*Es*importante*que*no*modifique*el*contenido*del*mismo."
        .MatchWildcards = True
        .Replacement.text = ""
        .Wrap = wdFindStop
        Do While .Execute
            rng.text = ""  ' Clear the found text
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub
Sub RemoverTextoEspecifico2()
    ' Remove specific unwanted text with special handling for spaces and accented characters
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    ' Remove the specific instructional text
    With rng.Find
        ' Use wildcards to account for spaces or formatting changes
        .text = "Recomendación:*"
        .MatchWildcards = True
        .Replacement.text = ""
        .Wrap = wdFindStop
        Do While .Execute
            rng.text = ""  ' Clear the found text
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub



