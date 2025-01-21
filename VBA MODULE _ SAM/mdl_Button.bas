Attribute VB_Name = "mdl_Button"

'=====================================================================================================
  '-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
  ''# BUTTON SHEET -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'+----------------------------------------------------------------------------------------------------+

Sub BUTTON_PROSES1()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Sub BUTTON_PROSES2()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False


    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
