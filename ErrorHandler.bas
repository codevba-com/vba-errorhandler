Attribute VB_Name = "ErrorHandler"
Option Explicit
' =============================================================================
' Module:        ErrorHandler
' Author:        Mark Uildriks, codevba.com
' Description:   Centralized error handling function VBA projects.
' Comment:       Involves user interaction, so primarily to be used in top level procedures
' Office version 2016 and higher
' Dependencies:  MailToProxy module
' License:       MIT License
' Version        1.0
' Repository:    https://github.com/codevba-com/vba-errorhandler
' =============================================================================

Private Const mcMailAddressTo As String = "support@codevba.com" 'replace by your preferred support email
Private Const mcErrorTitle As String = "Error" 'title of error dialog and email, you can make this more informative

Public Enum ErrorFeedbackType
    eftReportableMessage = 0
    eftSimpleMessage = 1
    eftnone = 2 'user does not notice things have gone wrong, use sparingly!
    eftDefault = 3
End Enum
Public Enum ErrorLoggingType
    elNone = 0
    elImmediateWindow = 1
    elErrorLogFile = 2
End Enum

Private Const mceftDefaultErrorFeedbackType As Long = eftReportableMessage 'change to eftSimpleMessage if that suits you better

Private eftErrorFeedbackType As ErrorFeedbackType
Private eltErrorLoggingType  As ErrorLoggingType
Private strErrorLogFile As String
Private strErrorTitle As String

Public Function HandleError(Err As ErrObject, Optional Feedback As ErrorFeedbackType = eftDefault, _
    Optional Module As String, Optional Procedure As String, _
    Optional ExtraInfo As String, Optional ErrLine As Long, Optional AddCancelButton = False) As Boolean
'If the user presses Cancel HandleError returns False, meaning 'don't continue'
    HandleError = True
    Dim Message As String
    Dim strSource As String
    Dim lngErrNumber As Long
    With Err
        'looses Err object somewhere here
        strSource = .Source
        lngErrNumber = .Number
        
        Message = "Error " & lngErrNumber & ": " & .Description
        If Len(strSource) > 0 Then Message = Message & " in " & DocumentName & strSource 'source generally returns VBAProject
        If Len(Module) > 0 Then Message = Message & " " & Module
    End With
    If Len(Procedure) > 0 Then Message = Message & " " & Procedure
    If Len(ExtraInfo) > 0 Then Message = Message & vbNewLine & ExtraInfo
    If ErrLine > 0 Then Message = Message & " line " & ErrLine
    If Feedback = eftDefault Then Feedback = ErrorFeedbackType
    Select Case lngErrNumber
    Case 2424 'risky to have MsgBox may reset form state
        GoTo ErrorLogging
    End Select
    Select Case Feedback
    Case eftReportableMessage, eftDefault
        Select Case MsgBox(Message & vbCrLf & vbCrLf & _
            DoYouWantToReportTheProblem, IIf(AddCancelButton, vbYesNoCancel, vbYesNo) + vbCritical + vbDefaultButton2, ErrorTitle)
        Case vbYes
            MailToProxy.CreateEmail mcMailAddressTo, ErrorTitle, Message
        Case vbCancel
            HandleError = False
        End Select
    Case eftSimpleMessage
        If AddCancelButton Then
            If vbCancel = MsgBox(Message, vbInformation + vbCancel, ErrorTitle) Then HandleError = False
        Else
            MsgBox Message, vbInformation, ErrorTitle
        End If
    End Select
ErrorLogging:
    Select Case ErrorLoggingType
    Case elImmediateWindow
        Debug.Print Message
    Case elErrorLogFile
        Dim iFile As Integer: iFile = FreeFile
        Open ErrorLogFile For Append As #iFile
        Print #iFile, FormatDateTime(Now) & " " & Replace(Message, vbNewLine, " ")
        Close #iFile
    End Select
End Function

'Use this to specify default FeedbackType
Property Let ErrorFeedbackType(Value As ErrorFeedbackType)
    eftErrorFeedbackType = Value
End Property
Property Get ErrorFeedbackType() As ErrorFeedbackType
    If eftErrorFeedbackType = eftDefault Then
        ErrorFeedbackType = mceftDefaultErrorFeedbackType
    Else
        ErrorFeedbackType = eftErrorFeedbackType
    End If
End Property
Property Get ErrorLoggingType() As ErrorLoggingType
    ErrorLoggingType = eltErrorLoggingType
End Property
Property Let ErrorLoggingType(Value As ErrorLoggingType)
    eltErrorLoggingType = Value
End Property
Property Get ErrorLogFile() As String
    If (Len(strErrorLogFile) > 0) Then
        ErrorLogFile = strErrorLogFile
    Else
        Dim strDocumentFolder As String
        strDocumentFolder = DocumentFolder
        strErrorLogFile = strDocumentFolder & "\" & "ErrorLog.txt"
    End If
    ErrorLogFile = strErrorLogFile
End Property
Property Let ErrorLogFile(Value As String)
    strErrorLogFile = Value
End Property
Property Get ErrorTitle() As String
    If (Len(strErrorTitle) > 0) Then
        ErrorTitle = strErrorTitle
    Else
        ErrorTitle = mcErrorTitle
    End If
End Property
Property Let ErrorTitle(Value As String)
    strErrorTitle = Value
End Property

Private Function DoYouWantToReportTheProblem() As String
    DoYouWantToReportTheProblem = "Do you want to report the problem?"
End Function

Private Function DocumentFolder() As String
'returns without \
    Dim objApplication As Object: Set objApplication = Application
    On Error Resume Next
    Select Case Right$(Application.Name, Len(Application.Name) - 10)
    Case "Access"
        DocumentFolder = objApplication.CurrentProject.Path
    Case "Excel"
        DocumentFolder = objApplication.ActiveWorkbook.Path
    Case "Word"
        DocumentFolder = objApplication.ActiveDocument.Path
    Case "PowerPoint"
        DocumentFolder = objApplication.ActivePresentation.Path
    End Select
End Function
Private Function DocumentName() As String
    Dim objApplication As Object: Set objApplication = Application
    On Error Resume Next
    Select Case Right$(Application.Name, Len(Application.Name) - 10)
    Case "Access"
        Dim strCurrentDbName As String: strCurrentDbName = objApplication.CurrentDb.Name
        DocumentName = Right$(strCurrentDbName, Len(strCurrentDbName) - InStrRev(strCurrentDbName, "\"))
    Case "Excel"
        DocumentName = objApplication.ActiveWorkbook.Name
    Case "Word"
        DocumentName = objApplication.ActiveDocument.Name
    Case "PowerPoint"
        DocumentName = objApplication.ActivePresentation.Name
    End Select
End Function

