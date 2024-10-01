Sub ModificarDocumentoWord()
    ' Chama as funções para cada tipo de modificação
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Call InserirContent
    Call ModificarEtiquetasDeContenido
    Call RemoverEtiquetaP
    Call ModificarEtiquetasDeImagenDeBannerAtual
    Call ConverterListasEmBulletPoints
    Call ConverterTitulosEmHeadings

    ' Modificar "Resume" e configurar como Heading 5
    Call ModificarResume

    ' Remover o texto da seção Schema
    Call RemoverTextoSchema

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
    
    ' Define o documento atual
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    contentInserted = False
    
    With rng.Find
        .text = "Nombre de la imagen: *.[jJ][pP][gG]"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            If Not contentInserted Then
                rng.Collapse Direction:=wdCollapseEnd
            
                ' Insere "CONTENT:" sem herdar estilos
                rng.InsertParagraphAfter
                rng.InsertAfter "CONTENT:" & vbCrLf
                rng.Style = wdStyleNormal
                rng.ParagraphFormat.Alignment = wdAlignParagraphLeft ' Align left
                
                contentInserted = True ' Ensure "CONTENT:" is only added once
            End If
            
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub

Sub ModificarEtiquetasDeContenido()
    Dim wordDoc As Document
    Dim rng As Range
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

                ' Search only for "URL SUGERIDA:" and remove the entire line containing it
                With rng.Find
                    .text = "URL SUGERIDA:*^13"
                    .MatchWildcards = True
                    .Forward = True
                    .Wrap = wdFindStop
                    If .Execute Then
                        ' Only delete the "URL SUGERIDA:..." line
                        rng.Delete
                    End If
                End With

                ' Insert "FIN DE SEO" after removing "URL SUGERIDA"
                rng.InsertParagraphAfter
                rng.InsertAfter "FIN DE SEO" & vbCrLf
                rng.Style = wdStyleNormal
                rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
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

Sub ModificarEtiquetasDeImagenDeBannerAtual()
    Dim wordDoc As Document
    Dim rng As Range
    Dim startRange As Range
    Dim idiomas As Variant
    Dim i As Integer
    
    idiomas = Array("ETIQUETAS DE IMAGEM DE BANNER ATUAL:", "ETIQUETAS DE IMAGEM:", _
                    "ETIQUETAS DE IMAGEM DO BANNER ATUAL:", "ETIQUETAS DE IMAGEM:", _
                    "ETIQUETAS DE IMAGEM DE BANNER ATUAL:", "ETIQUETAS DE IMAGEM:")
    
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
                rng.Font.Bold = True ' Make the title bold
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
        If InStr(para.Range.text, "ETIQUETAS DE IMAGEM:") > 0 Or _
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
    Dim headingText As String
    
    Set wordDoc = ActiveDocument
    estiloNames = Array("Heading ", "Título ", "Encabezado ")
    
    For i = 1 To 5
        estiloPrefix = ""
        Set estilo = Nothing
        
        For Each estiloName In estiloNames
            On Error Resume Next
            Set estilo = wordDoc.Styles(estiloName & i)
            On Error GoTo 0
            If Not estilo Is Nothing Then
                estiloPrefix = estiloName
                Exit For
            End If
        Next estiloName
        
        If estilo Is Nothing Then
            estiloPrefix = "Heading "
            Set estilo = wordDoc.Styles.Add(Name:=estiloPrefix & i, Type:=wdStyleTypeParagraph)
        End If
        
        Set rng = wordDoc.Content
        
        ' Handle cases with <hX> format and remove HTML tags
        With rng.Find
            .text = "<h" & i & ">*</h" & i & ">"
            .MatchWildcards = True
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                headingText = Trim(Replace(Replace(rng.text, "<h" & i & ">", ""), "</h" & i & ">", ""))
                rng.text = headingText ' Remove HTML tags
                rng.Style = estilo
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
        
        ' Normal detection for headings like "H" & i & ": "
        With rng.Find
            .text = "H" & i & ": *"
            .MatchWildcards = True
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindStop
            
            Do While .Execute
                headingText = Trim(Replace(rng.text, "H" & i & ": ", ""))
                rng.text = headingText
                rng.Style = estilo
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End With
    Next i
End Sub

Sub ModificarResume()
    Dim wordDoc As Document
    Dim rng As Range
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    
    With rng.Find
        .text = "Resume"
        .Forward = True
        .Wrap = wdFindStop
        
        If .Execute Then
            rng.text = "" ' Remove "Resume"
            rng.Collapse Direction:=wdCollapseEnd
            
            ' Set the rest of the text as H5 heading
            Do While rng.Paragraphs(1).Range.text <> ""
                rng.Style = "Heading 5" ' Change to Heading 5 style
                rng.Collapse Direction:=wdCollapseEnd
            Loop
        End If
    End With
End Sub

Sub RemoverTextoSchema()
    Dim wordDoc As Document
    Dim rng As Range
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content

    With rng.Find
        .text = "Recomendación: " & vbCrLf & "Se debe copiar el código que se encuentra dentro del recuadro y pegarlo en la sección <head> del documento HTML del sitio web. Es importante que no se modifique el contenido del mismo."
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            rng.text = ""
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub


