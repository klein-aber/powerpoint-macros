Sub KeepMyNames()
    'On slides that use slide layouts, apply the placeholder names of the master layout to the placeholders on the slide

    Dim phIndex As Integer
    Dim sldLayout
    Dim sld As Slide


    For Each sld In ActivePresentation.Slides

        Debug.Print "===Slide@Index" + CStr(sld.SlideIndex) + "==="
        Set sldLayout = sld.CustomLayout
        
        ' Applies the layout to the slide again
        ' Caution: This would overwrite changes you made to placeholders from the layout
        sld.CustomLayout = sldLayout

        
        phIndex = 1

        For Each ph In sldLayout.Shapes.Placeholders
            Debug.Print ph.Name
            sld.Shapes.Placeholders(phIndex).Name = ph.Name
            phIndex = phIndex + 1
        Next ph

    Next sld



End Sub
