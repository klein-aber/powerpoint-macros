Sub KeepMyNames()
    'On slides that use slide layouts, apply the placeholder names of the master layout to the placeholders on the slide

    Dim phIndex As Integer
    Dim sldLayout
    Dim sld As Slide


    For Each sld In ActivePresentation.Slides

        Debug.Print "===Slide@Index" + CStr(sld.SlideIndex) + "==="
        Set sldLayout = sld.CustomLayout
        
        ' Applies the layout to the slide again
        ' Caution: This would overwrite changes you made to the layout placeholders made on the slide
        sld.CustomLayout = sldLayout

        
        phIndex = 1

        For Each ph In sldLayout.Shapes.Placeholders
            Debug.Print ph.Name
            sld.Shapes.Placeholders(phIndex).Name = ph.Name

            ' Adjust order of elements. This is mainly required for the cover because it contains an image overlay 
            ' Which has to mix placeholder and non-placeholders, which results in a wrong order of elements
            If sld.Shapes.Placeholders(phIndex).Name = "cover_image" Then
                sld.Shapes.Placeholders(phIndex).ZOrder msoSendToBack
            ElseIf sld.Shapes.Placeholders(phIndex).Type = msoTextBox Then
                sld.Shapes.Placeholders(phIndex).ZOrder msoBringToFront
            End If
            
            phIndex = phIndex + 1
        Next ph

    Next sld



End Sub
