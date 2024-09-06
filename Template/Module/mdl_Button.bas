Attribute VB_Name = "mdl_Button"
Sub BUTTON_Process1()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayScrollBars = False

    Call Proses1


    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayScrollBars = True
    Application.CutCopyMode = False
End Sub



Sub BUTTON_Process2()
    

End Sub
