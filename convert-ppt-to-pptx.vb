' convert-ppt-to-pptx
'
' VBA macro to batch convert a folder of *.ppt files to Office Open XML format (*.pptx)
' 
' Create a PowerPoint presentation that is macro-enabled. In the Developer tab of the ribbon, click
' Visual Basic. Right click on the project and choose Insert > Module.
'
' (If unable to access the Developer tab, go to File > Options > Customize Ribbon, tick the Developer tab to add it.)
'
' Paste this into the module, altering "Basedir" below and you can use the Run button inside the 
' VBA project window to perform the conversion on a whole folder of .ppt.
'
' Based on Syon DP's answer to <https://stackoverflow.com/questions/46198205/vba-macro-batch-saveas-pptm-to-pptx#46198973>
' <https://stackoverflow.com/users/2398630/siyon-dp>.
'
' Therefore, this file is licensed under CC BY-SA 4.0. <https://creativecommons.org/licenses/by-sa/4.0/>


Sub ProcessFiles()

    Dim Filename, FileFormat As String
    Dim initialDisplayAlerts As Boolean
    initialDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Dim Basedir As String
    
    Basedir = "C:\users\someone\Desktop\convert\"
    
    Filename = Dir(Basedir & "*.ppt")
    
    Debug.Print (Basedir & Filename)
    
    Do While Filename <> ""
    Presentations.Open Filename:=Basedir & Filename
    With ActivePresentation
            .SaveCopyAs _
            Filename:=.Path & "\" & Left(.Name, InStrRev(.Name, ".")) & "pptx", _
            FileFormat:=ppSaveAsOpenXMLPresentation
            
            Filename = Dir()
            .Close
            End With
            
        Loop
    
    
    Application.DisplayAlerts = initialDisplayAlerts

End Sub

