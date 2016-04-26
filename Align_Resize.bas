Sub Align_Resize()
    
Dim osld As Slide
Dim oshp As Shape
Dim x As Integer
Dim y As Integer

With ActivePresentation.PageSetup
x = .SlideWidth / 2
y = .SlideHeight / 2
End With


' Schleife über jedes Slide
For Each osld In ActivePresentation.Slides

' Schleife über jedes Shape
For Each oshp In osld.Shapes

' Änderungen am Bild vornehmen.
' Größe definiert durch Höhe, alignment genau mittig

If oshp.Type = msoPicture Then
oshp.Height = 300
oshp.Left = x - (oshp.Width / 2)
oshp.Top = y - (oshp.Height / 2)
End If

'Ende Shape-Schleife
Next

'Ende Slide-Schleife
Next

End Sub
