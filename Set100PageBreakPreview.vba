Sub GotoHomeCell()
'
' GotoHomeCell Macro
'
'
    
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For I = 1 To WS_Count

            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            
            If Sheets(I).Visible Then
            
                Sheets(I).Activate
                Range("A1").Select
                
                Sleep (10)
                
                ActiveWindow.Zoom = 100
                ActiveWindow.View = xlPageBreakPreview
            End If
            
            'MsgBox ActiveWorkbook.Worksheets(I).Name, vbExclamation, "Check Name"
            
    Next I
    
    Sheets(1).Activate
    ActiveWorkbook.Save
         
    'MsgBox "Done!", vbExclamation, "Set Pages"
    
End Sub
