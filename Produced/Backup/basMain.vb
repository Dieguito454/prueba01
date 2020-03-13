Option Strict Off
Option Explicit On
Imports FirstEnergy.Email
Imports Microsoft.Practices.EnterpriseLibrary.Logging
Imports System.Configuration
Imports System.Data
Imports System.Data.OracleClient
Imports System.IO
Module basMain
    'DEV
    'Dim server As String = ConfigurationManager.AppSettings("serverDEV").ToString
    'Dim connStr As String = ConfigurationManager.ConnectionStrings("connStrDEV").ToString
    'Dim emailClient As String = ConfigurationManager.AppSettings("emailDEV").ToString
    'DEV
    'PROD
    Dim server As String = ConfigurationManager.AppSettings("serverPROD").ToString
    Dim connStr As String = ConfigurationManager.ConnectionStrings("connStrPROD").ToString
    Dim emailClient As String = ConfigurationManager.AppSettings("emailClient").ToString
    'PROD
    Dim xlsFile As String = ConfigurationManager.AppSettings("xlsFile").ToString
    Dim destPath As String = server & ConfigurationManager.AppSettings("destPath").ToString
    Dim producedPath As String = ConfigurationManager.AppSettings("producedPath").ToString
    Dim strFilePath As String = server & ConfigurationManager.AppSettings("destPath").ToString & "RunRevision.txt"
    Dim emailFrom As String = ConfigurationManager.AppSettings("emailFrom").ToString

    Public Sub Main()

        Dim strRunRevision As String
        Try
            If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length <= 1 Then

                WriteToEventLog(My.Application.Info.ProductName & " started", False)
                SendEmail(emailClient, "Produced MW Load Started", "Produced MW load was started!")

                Dim iFile As New StreamReader(strFilePath)
                strRunRevision = iFile.ReadLine
                iFile.Close()

                File.Copy(producedPath & xlsFile, destPath & xlsFile, True)

                InsertProduced(strRunRevision)
            Else
                End
            End If

            WriteToEventLog(My.Application.Info.ProductName & " ended", False)
        Catch ex As Exception
            WriteToEventLog("Exception raised in sub Main(). Exception: " & ex.Message & vbCrLf & ex.StackTrace, True)
            SendEmail(emailClient, "Exception raised in sub Main()", "Exception raised in sub Main() of Produced on " & Now & vbCrLf & "Exception: " & ex.Message & vbCrLf & ex.StackTrace)
            End
        End Try
 
    End Sub

    Public Sub InsertProduced(ByRef runRevision As String)
        Dim xlApplication As New Excel.Application
        Dim xlWorkbook As Excel.Workbook
        Dim strBudgetCycleYearMonth As String
        Dim strBudgetYear As String
        Dim strRun As String
        Dim strRevision As String
        Dim strSheet As String
        Dim strYear As String
        Dim strBudgeted_In_Year_Month As String
        Dim blnFlag As Boolean
        Dim Gregorian_Date As Date
        Dim Hour_Ending As Short
        Dim Plant As String
        Dim Unit As String
        Dim MW As Double
        Dim Row As Short
        Dim dSQL As String
        Dim SQL As String
        Dim intIndex As Short
        Dim rA As Integer
        Dim connDW As New OracleConnection(connStr)
        Dim oCommand As New OracleCommand
        Dim processSheet As String
        Dim nextPlant As Boolean
        Dim colLimit As Integer = 100  'in case of typo, loops looking for user entered control values will stop after this many columns processed
        xlWorkbook = xlApplication.Workbooks.Open(destPath & xlsFile)
        connDW.Open()
        oCommand.Connection = connDW

        Try

            strBudgetCycleYearMonth = Trim(Left(CStr(runrevision), 6))
            strBudgetYear = Trim(Mid(CStr(runrevision), 10, 4))
            strRun = Trim(Mid(CStr(runRevision), 17, InStr(17, runRevision, "-") - 17))
            strRevision = Trim(Right(CStr(runrevision), Len(runrevision) - InStrRev(runrevision, "-")))

            strSheet = CStr("ALL UNITS")
            xlWorkbook.Worksheets(strSheet).Select()

            strYear = Trim(CStr(xlWorkbook.ActiveSheet.Cells(1, 1).Value))
            strBudgeted_In_Year_Month = Trim(CStr(xlWorkbook.ActiveSheet.Cells(1, 2).Value))
            processSheet = Trim(CStr(xlWorkbook.ActiveSheet.Cells(1, 4).Value)).ToUpper

            blnFlag = False
            If strBudgetCycleYearMonth <> strBudgeted_In_Year_Month Then
                SendEmail(emailClient, "Error in InsertProduced", "Budget Cycle Year Month does not match what is showing on the spreadsheet (Cell B1), please make sure you are using the correct Run/Revision!")
                blnFlag = True
            Else
                If strBudgetYear <> strYear Then
                    SendEmail(emailClient, "Error in InsertProduced", "Budget Year does not match what is showing on the spreadsheet (Cell A1), please make sure you are using the correct Run/Revision!")
                    blnFlag = True
                Else

                    If processSheet = "X" Then
                        Row = 3

                        WriteToEventLog("Delete Produced Data for Run/Revision", False)
                        dSQL = "DELETE FROM C_BUD_PRODUCED_FEED WHERE BUDGETED_IN_YEAR_MONTH = " & CInt(strBudgeted_In_Year_Month) & " "
                        dSQL += "AND RUN = '" & strRun & "' "
                        dSQL += "AND REVISION = '" & strRevision & "' "
                        dSQL += "AND TO_CHAR(EST_GREGORIAN_DATE, 'YYYY') = '" & CStr(strYear) & "'"
                        oCommand.CommandText = dSQL
                        rA = oCommand.ExecuteNonQuery()
                        Do Until Trim(UCase(xlWorkbook.ActiveSheet.Cells(Row, 1).Value)) = "END"
                            If CStr(xlWorkbook.ActiveSheet.Cells(Row, 1).Value) = "" Or IsDBNull(xlWorkbook.ActiveSheet.Cells(Row, 1).Value) Then
                                'nothing
                            Else
                                Gregorian_Date = CDate(CShort(xlWorkbook.ActiveSheet.Cells(Row, 2).Value) & "/" & CShort(xlWorkbook.ActiveSheet.Cells(Row, 3).Value) & "/" & CShort(xlWorkbook.ActiveSheet.Cells(Row, 1).Value))
                                Hour_Ending = CShort(xlWorkbook.ActiveSheet.Cells(Row, 5).Value)

                                nextPlant = True
                                intIndex = 0
                                Do While nextPlant
                                    'For intIndex = 1 To 59
                                    intIndex = intIndex + 1

                                    Plant = UCase(Trim(CStr(xlWorkbook.ActiveSheet.Cells(1, intIndex + 7).Value)))
                                    Unit = UCase(Trim(CStr(xlWorkbook.ActiveSheet.Cells(2, intIndex + 7).Value)))

                                    If Plant <> "END" Then
                                        If Trim(xlWorkbook.ActiveSheet.Cells(Row, intIndex + 7).Value) = "" Or IsDBNull(xlWorkbook.ActiveSheet.Cells(Row, intIndex + 7).Value) = True Then
                                            MW = 0
                                        Else
                                            MW = CDbl(Trim(xlWorkbook.ActiveSheet.Cells(Row, intIndex + 7).Value))
                                        End If

                                        SQL = "INSERT INTO C_BUD_PRODUCED_FEED (BUDGETED_IN_YEAR_MONTH, RUN, REVISION, EST_GREGORIAN_DATE, EST_HOUR_ENDING, "
                                        SQL += "PLANT, UNIT, PRODUCED_MW) VALUES(" & CInt(strBudgeted_In_Year_Month) & ",'"
                                        SQL += strRun & "','" & strRevision & "',to_date('" & Gregorian_Date & "','mm/dd/yyyy'),"
                                        SQL += CShort(Hour_Ending) & ",'" & Plant & "','" & Unit & "'," & MW & ")"
                                        oCommand.CommandText = SQL
                                        rA = oCommand.ExecuteNonQuery()
                                        If intIndex >= colLimit Then
                                            nextPlant = False
                                        End If
                                    Else
                                        nextPlant = False
                                    End If
                                    'Next intIndex
                                Loop
                            End If
                            Row = Row + 1
                        Loop
                        WriteToEventLog("Inserted Produced Data for Run/Revision", False)
                    Else

                        WriteToEventLog("Skipped Produced Data for Run/Revision", False)
                    End If

                    strSheet = CStr("BASE LOAD UNAVAIL")

                    xlWorkbook.Worksheets(strSheet).Select()
                    processSheet = Trim(CStr(xlWorkbook.ActiveSheet.Cells(1, 4).Value)).ToUpper

                    If processSheet = "X" Then

                        Row = 3

                        Do Until Trim(UCase(xlWorkbook.ActiveSheet.Cells(Row, 1).Value)) = "END"
                            If CStr(xlWorkbook.ActiveSheet.Cells(Row, 1).Value) = "" Or IsDBNull(xlWorkbook.ActiveSheet.Cells(Row, 1).Value) Then
                                'nothing
                            Else
                                Gregorian_Date = CDate(CShort(xlWorkbook.ActiveSheet.Cells(Row, 2).Value) & "/" & CShort(xlWorkbook.ActiveSheet.Cells(Row, 3).Value) & "/" & CShort(xlWorkbook.ActiveSheet.Cells(Row, 1).Value))
                                Hour_Ending = CShort(xlWorkbook.ActiveSheet.Cells(Row, 5).Value)

                                nextPlant = True
                                intIndex = 0
                                Do While nextPlant
                                    intIndex = intIndex + 1
                                    'For intIndex = 1 To 7
                                    Plant = UCase(Trim(CStr(xlWorkbook.ActiveSheet.Cells(1, intIndex + 7).Value)))
                                    Unit = UCase(Trim(CStr(xlWorkbook.ActiveSheet.Cells(2, intIndex + 7).Value)))

                                    If Plant <> "END" Then
                                        If Trim(xlWorkbook.ActiveSheet.Cells(Row, intIndex + 7).Value) = "" Or IsDBNull(xlWorkbook.ActiveSheet.Cells(Row, intIndex + 7).Value) = True Then
                                            MW = 0
                                        Else
                                            MW = CDbl(Trim(xlWorkbook.ActiveSheet.Cells(Row, intIndex + 7).Value))
                                        End If

                                        SQL = "UPDATE C_BUD_PRODUCED_FEED SET UNAVAILABLE_MW = " & MW & " "
                                        SQL += "WHERE BUDGETED_IN_YEAR_MONTH = " & CInt(strBudgeted_In_Year_Month) & " "
                                        SQL += "AND RUN = '" & strRun & "' AND REVISION = '" & strRevision & "' "
                                        SQL += "AND EST_GREGORIAN_DATE = to_date('" & Gregorian_Date & "','mm/dd/yyyy') "
                                        SQL += "AND EST_HOUR_ENDING = " & CShort(Hour_Ending) & " AND PLANT = '" & Plant & "' "
                                        SQL += "AND UNIT = '" & Unit & "'"
                                        oCommand.CommandText = SQL
                                        rA = oCommand.ExecuteNonQuery()
                                        'Next intIndex
                                        If intIndex >= colLimit Then
                                            nextPlant = False
                                        End If
                                    Else
                                        nextPlant = False
                                    End If
                                Loop
                            End If
                            Row = Row + 1
                        Loop
                        WriteToEventLog("Inserted Unavailability Data for Run/Revision", False)
                    Else
                        WriteToEventLog("Skipped Unavailability Data for Run/Revision", False)
                    End If
                End If
            End If

            oCommand.Dispose()
            connDW.Close()
            connDW.Dispose()
            connDW.Close()
            xlApplication.DisplayAlerts = False
            xlWorkbook.Close()
            xlWorkbook = Nothing
            xlApplication = Nothing

            If blnFlag = False Then
                SendEmail(emailClient, "Produced MW loaded successfully!", "Produced MW loaded successfully!")
            Else
                SendEmail(emailClient, "Produced MW NOT loaded successfully!", "Produced MW NOT loaded successfully!")
            End If

        Catch ex As Exception
            If Not xlApplication Is Nothing Then
                xlApplication.DisplayAlerts = False
                xlApplication.Quit()
                xlApplication = Nothing
            End If
            WriteToEventLog("Exception raised in sub InsertProduced(). Exception: " & ex.Message & vbCrLf & ex.StackTrace, True)
            SendEmail(emailClient, "Exception raised in sub InsertProduced()", "Exception raised in sub InsertProduced() on " & Now & vbCrLf & "Exception: " & ex.Message & vbCrLf & ex.StackTrace)
            End
        End Try
    End Sub
    Public Sub WriteToEventLog(ByRef strMesg As String, Optional ByVal isError As Boolean = False)
        Try
            Dim logEntry As LogEntry = New LogEntry
            logEntry.Message = strMesg
            logEntry.TimeStamp = Now
            If isError Then
                logEntry.Severity = TraceEventType.Error
            Else
                logEntry.Severity = TraceEventType.Information
            End If
            Logger.Write(logEntry)

        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            End
        End Try
    End Sub
    Public Sub SendEmail(ByRef addressTo As String, ByRef subject As String, ByRef body As String)
        Try
            Dim msg As New FirstEnergy.Email.Message
            With msg
                .Subject = subject
                .Body = body
                .To = addressTo
                .From = emailFrom
                .Send()
            End With
            msg = Nothing
        Catch ex As Exception
            WriteToEventLog("Exception raised in SendEmail: " & ex.Message, True)
        End Try
    End Sub
End Module