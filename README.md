# VBA ErrorHandler - the HandleError function
This small repository contains an ErrorHandler module with a central HandleError function featuring user message dialog with optional Cancel, logging and reporting by email to the developer/adminstrator.
<i>To use the ErrorHandler requires module MailToProxy.</i>

Below is an example if using the HandleError function - at the end of the procedure. The first known case (5152) gives the user a standard informative message. The case Else apparently was not anticipated and the developer wants this situation to be reported to him by email.    

## Example use of the HandleError function
![Error message featuring sending an email with details to support](path-or-url "Optional title")

```vba
Function Divide2Byte(numerator As Integer, denominator As Integer) As Byte
    On Error GoTo HandleError
    Const cstrProcedure As String = "Divide2Byte"
    Divide2Byte = numerator / denominator
HandleExit:
    Exit Function
HandleError:
    Select Case Err.Number
    Case 11 'Division by zero
        HandleError Err, Feedback:=eftSimpleMessage
    Case Else 'unknown, for fixing we need more details!
        HandleError Err, Feedback:=eftReportableMessage, Procedure:=cstrProcedure, ExtraInfo:=numerator & "/" & denominator
    End Select
    Resume HandleExit
End Function
```
In the Immediate window, the first test will give the message dialog with the simple informative message. 
The second test triggers an unanticipated error. The message dialog opens allowing the user to report this by email with all available info included.
```vba
?Divide2Byte(numerator:=2, denominator:=0)
?Divide2Byte(numerator:=300, denominator:=1)
```
