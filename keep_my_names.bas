Sub KeepMyNames()

Dim phIndex As Integer
Dim sldLayout
Dim sld As Slide


For Each sld In ActivePresentation.Slides

    Debug.Print "===Slide@Index" + CStr(sld.SlideIndex) + "==="
    Set sldLayout = sld.CustomLayout
    
    phIndex = 1
    
    For Each ph In sldLayout.Shapes.Placeholders
        Debug.Print ph.Name
        sld.Shapes.Placeholders(phIndex).Name = ph.Name
        phIndex = phIndex + 1
    Next ph

Next sld



End Sub