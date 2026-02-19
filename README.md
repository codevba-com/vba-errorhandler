# VBA ErrorHandler - the HandleError function
This small repository contains an ErrorHandler module with a central HandleError function featuring user message dialog with optional Cancel, logging and reporting by email to the developer/adminstrator.
<i>To use the ErrorHandler requires module MailToProxy.</i>

Below is an example if using the HandleError function - at the end of the procedure. The first known case (5152) gives the user a standard informative message. The case Else apparently was not anticipated and the developer wants this situation to be reported to him by email.    
```vba
HandleExit:
    'whatever ...
    Exit Sub
HandleError:
    Select Case Err.Number
    Case 5152 'This is not a valid file name
        HandleError Err, Feedback:=eftSimpleMessage, Procedure:=cstrProcedure, ExtraInfo:=strFullFileName
    Case Else
        HandleError Err, Feedback:=eftReportableMessage, Procedure:=cstrProcedure,  ExtraInfo:=strFullFileName
    End Select
    Resume HandleExit
End Function
```
