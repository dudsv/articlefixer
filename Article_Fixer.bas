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
    Call RemoverTextoEspecifico
    Call EliminarSoloURLSugerida
    Call AplicarConversionTexto

    
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

    Set wordDoc = ActiveDocument

    ' Depuraci?n: Mensaje al comenzar la subrutina
    MsgBox "Iniciando la subrutina para aplicar negrita.", vbInformation

    ' Aplicar negrita a "Alt text:"
    Set rng = wordDoc.Content ' Reiniciar el rango al inicio del documento
    With rng.Find
        .text = "Alt text:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        Do While .Execute
            MsgBox "'Alt text:' encontrado y negrita aplicada.", vbInformation ' Depuraci?n
            rng.Font.Bold = True
            rng.Collapse Direction:=wdCollapseEnd ' Mover el rango al final del texto encontrado
        Loop
    End With

    ' Aplicar negrita a "Title:"
    Set rng = wordDoc.Content ' Reiniciar el rango al inicio del documento
    With rng.Find
        .text = "Title:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        Do While .Execute
            MsgBox "'Title:' encontrado y negrita aplicada.", vbInformation ' Depuraci?n
            rng.Font.Bold = True
            rng.Collapse Direction:=wdCollapseEnd ' Mover el rango al final del texto encontrado
        Loop
    End With

    ' Aplicar negrita a "Nombre de la imagen:"
    Set rng = wordDoc.Content ' Reiniciar el rango al inicio del documento
    With rng.Find
        .text = "Nombre de la imagen:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        Do While .Execute
            MsgBox "'Nombre de la imagen:' encontrado y negrita aplicada.", vbInformation ' Depuraci?n
            rng.Font.Bold = True
            rng.Collapse Direction:=wdCollapseEnd ' Mover el rango al final del texto encontrado
        Loop
    End With

End Sub

Sub RemoverTextoEspecifico()
    Dim wordDoc As Document
    Dim rng As Range
    
    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content
    
    ' Remove "Recomendaci�n:"
    With rng.Find
        .Text = "Recomendaci�n:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            rng.Text = ""
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' Remove the specific long line
    With rng.Find
        .Text = "Se debe copiar el c�digo que se encuentra dentro del recuadro y pegarlo en la secci�n <head> del documento HTML del sitio web. Es importante que no se modifique el contenido del mismo."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            rng.Text = ""
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
    
    ' Clear residual formatting
    rng.Font.Reset
    rng.ParagraphFormat.Reset
End Sub

Sub EliminarSoloURLSugerida()
    Dim wordDoc As Document
    Dim rng As Range

    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content

    ' Inicializar la b�squeda de "URL Sugerida:"
    With rng.Find
        .Text = "URL Sugerida:*.[jJ][pP][gG]" ' Buscar cualquier URL que termine con .jpg
        .MatchWildcards = True ' Usar comodines para encontrar la URL completa
        .Forward = True
        .Wrap = wdFindContinue ' Continuar buscando en todo el documento
        Do While .Execute
            ' Eliminar solo la URL sugerida
            rng.Delete ' Eliminar la URL encontrada sin tocar el resto del contenido
        Loop
    End With

End Sub

Sub AplicarConversionTexto()
    Dim wordDoc As Document
    Dim rng As Range

    Set wordDoc = ActiveDocument
    Set rng = wordDoc.Content

    ' Convertir "Text Alt:" a "Alt text:"
    With rng.Find
        .Text = "Text Alt:"
        .Replacement.Text = "Alt text:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False ' Desactivar el formato para asegurarse de que busca en todo el texto
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Convertir "Title de la Imagen:" a "Title:"
    With rng.Find
        .Text = "Title de la Imagen:"
        .Replacement.Text = "Title:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False ' Asegurarse de que no est� buscando solo en textos con formato espec�fico
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
