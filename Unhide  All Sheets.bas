Sub Unhide()
  
  Dim Sh As Worksheet
  
  For Each Sh in Activeworkbook.Worksheets
  
    Sh.visible
    
 Next Sh
 
End Sub
