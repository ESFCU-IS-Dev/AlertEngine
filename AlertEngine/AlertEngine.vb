Imports System.Data
Imports Hyland.Unity.Application
Imports System.Threading.Timer

Public Class AlertEngine
    Private moAlert As Threading.Timer
    Private moPersist As ERMS.OBJ.Persist
    Private moDatabase As ERMS.OBJ.Database
    Private moData As New DataSet

    'local variable(s) to hold property value(s)
    Private miPollInterval As Integer
    Private mbAlertOffHours As Boolean
    Private msApplicationServer As String
    Private msTestApplicationServer As String
    Private msSQLUser As String
    Private msSQLPassword As String
    Private msOnBaseUser As String
    Private msOnBasePassword As String
    Private moAlerts As New Collection
    Private msSMTPHost As String
    Private msSMTPUser As String
    Private msSMTPPassword As String
    Private msSourceMailbox As String
    Private msRecipientMailbox As String
    Private msSubject As String

#Region "Configuration"
    Public Property PollInterval() As Integer
        Get
            PollInterval = miPollInterval
        End Get
        Set(ByVal viValue As Integer)
            miPollInterval = viValue
        End Set
    End Property
    Public Property AlertOffHours() As Boolean
        Get
            AlertOffHours = mbAlertOffHours
        End Get
        Set(ByVal vbValue As Boolean)
            mbAlertOffHours = vbValue
        End Set
    End Property
    Public Property ApplicationServer() As String
        Get
            ApplicationServer = msApplicationServer
        End Get
        Set(ByVal Value As String)
            msApplicationServer = Value
        End Set
    End Property
    Public Property TestApplicationServer() As String
        Get
            TestApplicationServer = msTestApplicationServer
        End Get
        Set(ByVal Value As String)
            msTestApplicationServer = Value
        End Set
    End Property
    Public Property SQLUser() As String
        Get
            SQLUser = msSQLUser
        End Get
        Set(ByVal Value As String)
            msSQLUser = Value
        End Set
    End Property
    Public Property SQLPassword() As String
        Get
            SQLPassword = msSQLPassword
        End Get
        Set(ByVal Value As String)
            msSQLPassword = Value
        End Set
    End Property
    Public Property OnBaseUser() As String
        Get
            OnBaseUser = msOnBaseUser
        End Get
        Set(ByVal Value As String)
            msOnBaseUser = Value
        End Set
    End Property
    Public Property OnBasePassword() As String
        Get
            OnBasePassword = msOnBasePassword
        End Get
        Set(ByVal Value As String)
            msOnBasePassword = Value
        End Set
    End Property

    Public Property SMTPHost() As String
        Get
            SMTPHost = msSMTPHost
        End Get
        Set(ByVal vsValue As String)
            msSMTPHost = vsValue
        End Set
    End Property
    Public Property SMTPUser() As String
        Get
            SMTPUser = msSMTPUser
        End Get
        Set(ByVal vsValue As String)
            msSMTPUser = vsValue
        End Set
    End Property
    Public Property SMTPPassword() As String
        Get
            SMTPPassword = msSMTPPassword
        End Get
        Set(ByVal vsValue As String)
            msSMTPPassword = vsValue
        End Set
    End Property
    Public Property SourceMailbox() As String
        Get
            SourceMailbox = msSourceMailbox
        End Get
        Set(ByVal vsValue As String)
            msSourceMailbox = vsValue
        End Set
    End Property
    Public Property RecipientMailbox() As String
        Get
            RecipientMailbox = msRecipientMailbox
        End Get
        Set(ByVal vsValue As String)
            msRecipientMailbox = vsValue
        End Set
    End Property
    Public Property Subject() As String
        Get
            Subject = msSubject
        End Get
        Set(ByVal vsValue As String)
            msSubject = vsValue
        End Set
    End Property

    Public ReadOnly Property Alerts() As Collection
        Get
            Alerts = moAlerts
        End Get
    End Property
    Public Function AddAlert(ByVal vsAlert As String) As AlertEngine.Alert
        Dim oAlert As New AlertEngine.Alert
        moAlerts.Add(oAlert, vsAlert)
        AddAlert = moAlerts(vsAlert)
    End Function
    Public NotInheritable Class Alert
        'local variable(s) to hold property value(s)
        Private msName As String
        Private msDBQuery As String
        Private dtLastUpdated As DateTime = #1/1/1900#
        Private iMessagesSent As Integer
        Private mbFixed As Boolean = True
        Private mbAlertOffHours As Boolean
        Private mbFollowUp As Boolean = True
        Private miPriority As Integer = 1
        Private moRecipients As New Collection

        Public Property Name() As String
            Get
                Name = msName
            End Get
            Set(ByVal Value As String)
                msName = Value
            End Set
        End Property
        Public Property DBQuery() As String
            Get
                DBQuery = msDBQuery
            End Get
            Set(ByVal Value As String)
                msDBQuery = Value
            End Set
        End Property
        Public Property LastUpdated() As DateTime
            Get
                LastUpdated = dtLastUpdated
            End Get
            Set(ByVal Value As DateTime)
                dtLastUpdated = Value
            End Set
        End Property
        Public Property CurrentMessagesSent() As Integer
            Get
                CurrentMessagesSent = iMessagesSent
            End Get
            Set(ByVal Value As Integer)
                iMessagesSent = Value
            End Set
        End Property
        Public Property Fixed() As Boolean
            Get
                Fixed = mbFixed
            End Get
            Set(ByVal Value As Boolean)
                mbFixed = Value
            End Set
        End Property
        Public Property AlertOffHours() As Boolean
            Get
                AlertOffHours = mbAlertOffHours
            End Get
            Set(ByVal vbValue As Boolean)
                mbAlertOffHours = vbValue
            End Set
        End Property
        Public Property SendFollowUp() As Boolean
            Get
                SendFollowUp = mbFollowUp
            End Get
            Set(ByVal Value As Boolean)
                mbFollowUp = Value
            End Set
        End Property
        Public Property Priority() As Integer
            Get
                Priority = miPriority
            End Get
            Set(ByVal vbValue As Integer)
                miPriority = vbValue
            End Set
        End Property

        Public ReadOnly Property Recipients() As Collection
            Get
                Recipients = moRecipients
            End Get
        End Property
        Public Function AddRecipient(ByVal vsRecipient As String) As AlertEngine.Alert.Recipient
            Dim oRecipient As New Alert.Recipient
            moRecipients.Add(oRecipient, vsRecipient)
            AddRecipient = moRecipients(vsRecipient)
        End Function
        Public NotInheritable Class Recipient
            'local variable(s) to hold property value(s)
            Private msEmailAddress As String
            Private mbAlertOffHours As Boolean

            Public Property eMailAddress() As String
                Get
                    eMailAddress = msEmailAddress
                End Get
                Set(ByVal Value As String)
                    msEmailAddress = Value
                End Set
            End Property
            Public Property AlertOffHours() As Boolean
                Get
                    AlertOffHours = mbAlertOffHours
                End Get
                Set(ByVal vbValue As Boolean)
                    mbAlertOffHours = vbValue
                End Set
            End Property
        End Class
    End Class
#End Region

#Region "Service Wrapper"
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Dim oError As New ERMS.OBJ.ErrorLog
        Dim oCallBack As System.Threading.TimerCallback

        '***************************************
        ' Load settings from configuration file
        '***************************************
        moPersist = New ERMS.OBJ.Persist
        moDatabase = New ERMS.OBJ.Database

        If moPersist.LoadConfig(Me, "C:\zSupport\Master.xml", "Alerts") = True Then
            ' Load Alerts from configuration file
            If moPersist.LoadConfig(moDatabase, "C:\zSupport\Master.xml", "AlertsDB") = True Then
                'Startup Alert Service
                oCallBack = AddressOf DoWork
                miPollInterval = miPollInterval * 1000
                moAlert = New Threading.Timer(oCallBack, Nothing, 0, miPollInterval)
            Else
                'Log Error
                Dim oE As New System.Exception("Error starting up Alert Engine -- Alerts Database Object Config Error")
                oE.Source = "Alert Engine"
                oError.LogError(oE, "OnStart")
                Me.Stop()
            End If
        Else
            'Log Error
            Dim oE As New System.Exception("Error starting up Alert Engine -- Alerts Object Config Error")
            oE.Source = "Alert Engine"
            oError.LogError(oE, "OnStart")
            Me.Stop()
        End If
    End Sub
    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        moAlert.Dispose()
        moData.Dispose()
    End Sub
#End Region

#Region "Core Logic"
    Public Sub DoWork()
        Dim oError As New ERMS.OBJ.ErrorLog

        Dim bIsOffHours As Boolean
        Dim bSuccess As Boolean
        Dim iData As Integer
        Dim sTrace As String
        Dim sStatus As String
        Dim sMessage As String

        Try
            moData.Dispose()
            moData = New DataSet

            '*******************************************************
            ' Load settings from configuration file (Debug Version)
            '*******************************************************
            If moPersist Is Nothing Then
                moPersist = New ERMS.OBJ.Persist
                moDatabase = New ERMS.OBJ.Database

                If moPersist.LoadConfig(Me, "C:\zSupport\Master.xml", "Alerts") <> True Then
                    MsgBox("Error loading application configuration")
                    Exit Sub
                Else
                    MsgBox("Polling interval is " + CStr(miPollInterval) + " seconds!")
                End If
                If moPersist.LoadConfig(moDatabase, "C:\zSupport\Master.xml", "AlertsDB") <> True Then
                    MsgBox("Error loading database configuration")
                    Exit Sub
                End If
            End If

            If (Format(Now, "Short Time") < #5:00:00 AM# Or Format(Now, "Short Time") > #6:00:00 PM#) Or IsWeekday(Now) = False Then
                bIsOffHours = True
            Else
                bIsOffHours = False
            End If

            'PERFORM Daily cleanup services off hours
            'If bIsOffHours = True And Hour(Now) = 2 And Minute(Now) <= 4 Then
            'CleanupDEVTempFiles()
            'CleanupPRODTempFiles()
            'End If

            'Engine can turn off all alerts off hours (GLOBAL CONFIGURATION PARAMETER)
            If (bIsOffHours = False) Or (bIsOffHours = True And mbAlertOffHours = True) Then
                'Process Alerts
                For Each oAlert As Alert In moAlerts
                    'IF off hours global config parameter is ON, then each alert can be configured for off hours notification
                    If (bIsOffHours = False) Or (bIsOffHours = True And oAlert.AlertOffHours = True) Then
                        bSuccess = moDatabase.RunQuery("Alerts", oAlert.DBQuery, , moData)
                        If bSuccess = True Then
                            For iData = 1 To moData.Tables(0).Rows.Count

                                'Agreed upon message contract (Trace/Status/Message)
                                sTrace = UCase(moData.Tables(0).Rows(iData - 1).Item("Trace")) 'ON or OFF
                                sStatus = UCase(moData.Tables(0).Rows(iData - 1).Item("Status")) 'INFORMATIONAL, WARNING, CRITICAL
                                sMessage = moData.Tables(0).Rows(iData - 1).Item("Message")

                                If sTrace = "ON" OrElse sStatus <> "INFORMATIONAL" OrElse oAlert.Fixed = False Then
                                    If oAlert.Fixed = False AndAlso sStatus = "INFORMATIONAL" AndAlso sMessage = "No Error" Then
                                        'Alert has been addressed/fixed
                                        oAlert.Fixed = True
                                        If oAlert.SendFollowUp = True Then
                                            If SendMail(oAlert, oAlert.Name & " has been resolved!", bIsOffHours, sStatus) = False Then
                                                Dim oE As New System.Exception("Error sending follow up email, alert: " & oAlert.Name & "!")
                                                oE.Source = "Alert Engine"
                                                oError.LogError(oE, "DoWork")
                                            End If
                                        End If
                                        'Reset Alert Counters
                                        oAlert.CurrentMessagesSent = 0
                                        oAlert.LastUpdated = "1/1/1900"
                                    Else
                                        'Alert is active
                                        If sStatus <> "INFORMATIONAL" Then
                                            oAlert.Fixed = False
                                            'Only send a max of 3 messages (Once an hour)
                                            If oAlert.CurrentMessagesSent < 3 Then
                                                If DateDiff(DateInterval.Minute, oAlert.LastUpdated, Now) > 60 Then
                                                    If SendMail(oAlert, sMessage, bIsOffHours, sStatus) = False Then
                                                        Dim oE As New System.Exception("Error sending actual alert: " & oAlert.Name & "!")
                                                        oE.Source = "Alert Engine"
                                                        oError.LogError(oE, "DoWork")
                                                    End If

                                                    'Update Counters
                                                    oAlert.CurrentMessagesSent = oAlert.CurrentMessagesSent + 1
                                                    oAlert.LastUpdated = Now
                                                End If
                                            End If
                                        Else
                                            'Tracing is turned on to test email functionality
                                            If sTrace = "ON" Then
                                                If SendMail(oAlert, sMessage, bIsOffHours, sStatus) = False Then
                                                    Dim oE As New System.Exception("Error sending trace for alert: " & oAlert.Name & "!")
                                                    oE.Source = "Alert Engine"
                                                    oError.LogError(oE, "DoWork")
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'Log Error
                            Dim oE As New System.Exception("Error running DB Query associated with alert: " & oAlert.DBQuery & "!")
                            oE.Source = "Alert Engine"
                            oError.LogError(oE, "DoWork")
                        End If

                        moData = New DataSet
                    End If
                Next
            End If
        Catch ex As System.Exception
            oError.LogError(ex, "DoWork")
        Finally
            moData.Dispose()
        End Try
    End Sub
    Private Function SendMail(ByRef xoAlert As Alert, ByVal vsMessage As String, ByVal vbOffHours As Boolean, ByVal vsStatus As String) As Boolean
        SendMail = True

        '1/18/2018 Hard Coded
        'msSMTPHost = "mail.usecology.com"
        msSMTPHost = "mail.esfcu.org"


        'Instantiate Objects
        Dim oSMTP As New SmtpClient(msSMTPHost, 25) '587
        Dim oMail As New MailMessage()
        Dim sSubject As String = msSubject


        Try
            'Send(authentication)
            'oSMTP.Credentials = New NetworkCredential(msSMTPUser, msSMTPPassword)
            'oSMTP.EnableSsl = True

            'Set mail properties
            With oMail
                .From = New MailAddress(msSourceMailbox)

                'Outlook message priority
                If xoAlert.Priority = 0 Then
                    .Priority = MailPriority.Low
                ElseIf xoAlert.Priority > 1 OrElse vsStatus.ToUpper = "CRITICAL" Then
                    .Priority = MailPriority.High
                Else
                    .Priority = MailPriority.Normal
                End If

                'If xoAlert.Name = "DI MT Refresh" Then
                '    'PACMS Request
                '    'Modified 3/26/2014  --> Bret McHenry
                '    Select Case True
                '        Case UCase(vsMessage) Like "*DEV*"
                '            sSubject = sSubject & " --> DEV Environment"
                '        Case UCase(vsMessage) Like "*PROD*"
                '            sSubject = sSubject & " --> PROD Environment"
                '    End Select

                '    Select Case True
                '        Case UCase(vsMessage) Like "*PACMS*"
                '            sSubject = sSubject & ":PACMS"
                '        Case UCase(vsMessage) Like "*CPCMS*"
                '            sSubject = sSubject & ":CPCMS"
                '        Case UCase(vsMessage) Like "*MDJ*"
                '            sSubject = sSubject & ":MDJ"
                '    End Select
                'End If

                'If xoAlert.Name = "InVoice Overwrite Alert" Then
                '    sSubject = "ASAP Alert"
                'End If

                .Subject = sSubject
                .Body = vsMessage


                'Send Alert (to all recipients)
                For Each oRecip As AlertEngine.Alert.Recipient In xoAlert.Recipients
                    If (vbOffHours = False) Or (vbOffHours = True And xoAlert.AlertOffHours = True) Then
                        .To.Add(oRecip.eMailAddress)
                    End If
                Next
            End With

            'Send the message
            oSMTP.Send(oMail)

            'Clear prior values
            oMail.Attachments.Clear()
            oMail.To.Clear()
            oMail.CC.Clear()
        Catch e As Exception
            'Log Error
            Dim oError As New ERMS.OBJ.ErrorLog
            oError.LogError(e, "SendMail")

            SendMail = False
        Finally
            oMail.Dispose()
            oSMTP.Dispose()
        End Try

    End Function
#End Region

    Private Sub CleanupDEVTempFiles()
        Dim oError As New ERMS.OBJ.ErrorLog
        Dim oApplication As Hyland.Unity.Application = Nothing
        Dim oConnection As New System.Data.Odbc.OdbcConnection
        Dim oDS As New System.Data.DataSet
        Dim oRS As System.Data.Odbc.OdbcDataAdapter = Nothing
        Dim oDocument As Hyland.Unity.Document
        Dim iDocuments As Integer
        Dim iDocHandle As Long

        Try
            oApplication = Hyland.Unity.Application.Connect(Hyland.Unity.Application.CreateOnBaseAuthenticationProperties(msTestApplicationServer, msOnBaseUser, msOnBasePassword, "OBPACMSDEV"))

            oConnection.ConnectionString = "DRIVER=SQL Server Native Client 10.0;UID=" & msSQLUser & ";PWD=" & msSQLPassword & _
                                            ";DATABASE=OnBaseDev;SERVER=ITSPSERMSSQLDV1"
            oConnection.Open()
            oRS = New System.Data.Odbc.OdbcDataAdapter("Select  ID.itemnum " & _
                                                        "From   hsi.itemdata ID " & _
                                                        "Left   Join hsi.keyitem116 DN On ID.itemnum = DN.itemnum " & _
                                                        "Where  (DN.itemnum is null and ID.itemtypegroupnum Not In (1,122,146) and ID.status = 0 and ID.datestored < GetDate()-1)" & _
                                                        "       or " & _
                                                        "       (DN.itemnum is null and ID.itemtypenum =444 and ID.status = 0 and ID.datestored < GetDate()-1)", oConnection)
            oRS.Fill(oDS, "TempFiles")

            For iDocuments = 1 To oDS.Tables(0).Rows.Count
                iDocHandle = oDS.Tables(0).Rows(iDocuments - 1).Item("itemnum")

                oDocument = oApplication.Core.GetDocumentByID(iDocHandle, Hyland.Unity.DocumentRetrievalOptions.LoadKeywords)
                If Not oDocument Is Nothing Then
                    If oDocument.KeywordRecords(0).Keywords.Count = 0 Then
                        'The document does not have any associated keywords (DELETE)
                        oApplication.Core.Storage.DeleteDocument(oDocument)
                    End If
                End If
            Next
        Catch ex As Exception
            oError.LogError(ex, "DoWork - Daily Cleanup")
        Finally
            oApplication.Disconnect()
            oRS.Dispose()
            oDS.Dispose()
            oConnection.Close()
        End Try
    End Sub
    Private Sub CleanupPRODTempFiles()
        Dim oError As New ERMS.OBJ.ErrorLog
        Dim oApplication As Hyland.Unity.Application = Nothing
        Dim oConnection As New System.Data.Odbc.OdbcConnection
        Dim oDS As New System.Data.DataSet
        Dim oRS As System.Data.Odbc.OdbcDataAdapter = Nothing
        Dim oDocument As Hyland.Unity.Document
        Dim iDocuments As Integer
        Dim iDocHandle As Long

        Try
            oConnection = New System.Data.Odbc.OdbcConnection
            oDS = New System.Data.DataSet
            oApplication = Hyland.Unity.Application.Connect(Hyland.Unity.Application.CreateOnBaseAuthenticationProperties(msApplicationServer, msOnBaseUser, msOnBasePassword, "OBPACMSPROD"))
            oConnection.ConnectionString = "DRIVER=SQL Server Native Client 10.0;UID=" & msSQLUser & ";PWD=" & msSQLPassword & _
                                            ";DATABASE=observer;SERVER=ermsmssql"
            oConnection.Open()
            oRS = New System.Data.Odbc.OdbcDataAdapter("Select  ID.itemnum " & _
                                                        "From   hsi.itemdata ID " & _
                                                        "Left   Join hsi.keyitem116 DN On ID.itemnum = DN.itemnum " & _
                                                        "Where  (DN.itemnum is null and ID.itemtypegroupnum Not In (1,122,145) and ID.status = 0 and ID.datestored < GetDate()-1)" & _
                                                        "       or " & _
                                                        "       (DN.itemnum is null and ID.itemtypenum In (455,456) and ID.status = 0 and ID.datestored < GetDate()-1)", oConnection)
            oRS.Fill(oDS, "TempFiles")

            For iDocuments = 1 To oDS.Tables(0).Rows.Count
                iDocHandle = oDS.Tables(0).Rows(iDocuments - 1).Item("itemnum")

                oDocument = oApplication.Core.GetDocumentByID(iDocHandle, Hyland.Unity.DocumentRetrievalOptions.LoadKeywords)
                If Not oDocument Is Nothing Then
                    If oDocument.KeywordRecords(0).Keywords.Count = 0 Then
                        'The document does not have any associated keywords (DELETE)
                        oApplication.Core.Storage.DeleteDocument(oDocument)
                    End If
                End If
            Next
        Catch ex As Exception
            oError.LogError(ex, "DoWork - Daily Cleanup")
        Finally
            oApplication.Disconnect()
            oRS.Dispose()
            oDS.Dispose()
            oConnection.Close()
        End Try
    End Sub
    Private Function IsWeekday(vdtSomeDate As Date) As Boolean
        IsWeekday = False

        Select Case Weekday(vdtSomeDate)
            Case 1 : IsWeekday = False 'Sunday
            Case 2 : IsWeekday = True
            Case 3 : IsWeekday = True
            Case 4 : IsWeekday = True
            Case 5 : IsWeekday = True
            Case 6 : IsWeekday = True
            Case 7 : IsWeekday = False 'Saturday 
        End Select
    End Function
End Class

