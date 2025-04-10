
Imports Microsoft.MetadirectoryServices

Public Class MVExtensionObject
    Implements IMVSynchronization

    Public Sub Initialize() Implements IMvSynchronization.Initialize
        ' TODO: Add initialization code here
    End Sub

    Public Sub Terminate() Implements IMvSynchronization.Terminate
        ' TODO: Add termination code here
    End Sub

    Public Sub Provision(ByVal mventry As MVEntry) Implements IMVSynchronization.Provision
        If mventry.ObjectType.Equals("absence") Then
            Dim KHQA As ConnectedMA = mventry.ConnectedMAs("KH-QAbsence")
            Dim KHSQL As ConnectedMA = mventry.ConnectedMAs("KH-SQL")
            Dim KHSQLcsentry As CSEntry
            Dim DN As ReferenceValue
            Dim rdn As String = mventry("absenceID").Value
            DN = KHSQL.EscapeDNComponent(rdn)
            If KHQA.Connectors.Count >= 1 Then
                If KHSQL.Connectors.Count = 0 Then
                    If mventry("SendDate").IsPresent Then
                        'Dim currentDate As DateTime = DateTime.Now
                        ''Dim currentDate As DateTime = New DateTime(2024, 7, 5)
                        ''Dim endDate As DateTime = New DateTime(2024, 6, 6)
                        'currentDate = New DateTime(currentDate.Year, currentDate.Month, 5)
                        'Dim thresholdDate As Date = Convert.ToDateTime(mventry("SendDate").Value)
                        'If thresholdDate < currentDate Then
                        Dim currentDate As DateTime = DateTime.Now
                        currentDate = New DateTime(currentDate.Year, currentDate.Month, 5)
                        Dim thresholdDate As Date = Convert.ToDateTime(mventry("SendDate").Value)
                        If thresholdDate > currentDate Then
                            'do nothing
                        Else
                            Try
                                KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
                                KHSQLcsentry.DN = DN
                                KHSQLcsentry("AbsenceID").Value = mventry("absenceID").Value
                                KHSQLcsentry("objectType").Value = "absence"
                                KHSQLcsentry("fromBadgeNo").Value = mventry("fromBadgeNo").Value
                                KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue

                                KHSQLcsentry("toDate").Value = mventry("toDate").Value
                                KHSQLcsentry("fromDate").Value = mventry("fromDate").Value
                                If mventry("name").IsPresent Then
                                    KHSQLcsentry("name").Value = mventry("name").Value
                                End If
                                KHSQLcsentry("timeStamp").Value = mventry("fromDate").Value
                                If mventry("sickLevel").IsPresent Then
                                    KHSQLcsentry("sickLevel").Value = mventry("sickLevel").Value
                                End If

                                'KHSQLcsentry("approvedByManager").Value = mventry("approvedByManager").Value
                                KHSQLcsentry.CommitNewConnector()

                            Catch ex As ObjectAlreadyExistsException
                            Catch ex As Exception
                                Throw
                                'IMAExtensible2CallExport 'bla
                            End Try


                            'do nothing
                        End If
                    Else
                        'do nothing
                    End If
                Else
                    'do nothing
                End If
            End If
        End If
        ''--Fix removal 05122024
        ''End If
        'If mventry.ObjectType.Equals("person") Then
        '    Dim KHQA As ConnectedMA = mventry.ConnectedMAs("KH-QAbsence")
        '    Dim KHSQL As ConnectedMA = mventry.ConnectedMAs("KH-SQL")
        '    Dim KHSQLcsentry As CSEntry
        '    Dim DN As ReferenceValue
        '    'Dim timeStamp As DateTime = New DateTime(2024, 7, 5)
        '    Dim timeStamp As String = DateAndTime.Now.ToString
        '    If KHQA.Connectors.Count >= 1 Then
        '        'If KHSQL.Connectors.Count > 0 Then -- 05.11.2021
        '        'do nothing -- 05.11.2021
        '        'Else
        '        If mventry("Q_disabled").IsPresent Then
        '            If mventry("Q_disabled").BooleanValue = True Then
        '                Dim currentDate As DateTime = DateTime.Now
        '                Dim endDate As DateTime = Convert.ToDateTime(mventry("Qend").Value)
        '                Dim endMonth As Integer = Month(endDate)
        '                Dim currentMonth As Integer = Month(currentDate)
        '                Dim endYear As Integer = Year(endDate)
        '                Dim currentYear As Integer = Year(currentDate)
        '                If endYear.Equals(currentYear) Then
        '                    Dim result1 As Integer
        '                    result1 = currentMonth - endMonth
        '                    If result1 = 1 Then
        '                        If mventry("SQLaccountName").Value.Length > 2 Then
        '                            If KHSQL.Connectors.Count = 0 Then
        '                                'If mventry("timestamp").IsPresent Then ''Readded 05.11.2021
        '                                'Dim timeattrib As DateTime = Convert.ToDateTime(mventry("timestamp").Value)
        '                                'Dim timenow As DateTime = DateTime.Now
        '                                'Dim dur As TimeSpan = timenow - timeattrib
        '                                'If dur.Minutes < 200 Then
        '                                'do nothing 
        '                                'Else '' Readded End
        '                                Dim rdn As String = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                DN = KHSQL.EscapeDNComponent(rdn)
        '                                Try
        '                                    KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
        '                                    KHSQLcsentry.DN = DN

        '                                    KHSQLcsentry("AbsenceID").Value = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                    KHSQLcsentry("fromBadgeNo").Value = mventry("SQLaccountName").Value
        '                                    KHSQLcsentry("objectType").Value = "person"
        '                                    If mventry("name").IsPresent Then
        '                                        KHSQLcsentry("name").Value = mventry("name").Value
        '                                    End If
        '                                    If mventry("saldoOvertid").IsPresent Then
        '                                        KHSQLcsentry("saldoOvertid").Value = mventry("saldoOvertid").Value
        '                                    End If
        '                                    If mventry("managerName").IsPresent Then
        '                                        KHSQLcsentry("manager").Value = mventry("managerName").Value
        '                                    End If
        '                                    If mventry("extCode").IsPresent Then
        '                                        KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue
        '                                    End If
        '                                    If mventry("templateextid").IsPresent Then
        '                                        KHSQLcsentry("templateextid").Value = mventry("templateextid").Value
        '                                    End If
        '                                    If mventry("templatename").IsPresent Then
        '                                        KHSQLcsentry("templatename").Value = mventry("templatename").Value
        '                                    End If
        '                                    If mventry("employmentPercent").IsPresent Then
        '                                        KHSQLcsentry("employmentPercent").Value = mventry("employmentPercent").Value
        '                                    End If
        '                                    If mventry("unitname").IsPresent Then
        '                                        KHSQLcsentry("unitname").Value = mventry("unitname").Value
        '                                    End If
        '                                    If mventry("gender").IsPresent Then
        '                                        KHSQLcsentry("gender").Value = mventry("gender").Value
        '                                    End If
        '                                    If mventry("employeeStartDate").IsPresent Then
        '                                        KHSQLcsentry("timeStamp").Value = mventry("employeeStartDate").Value
        '                                    End If
        '                                    If mventry("saldoTidsbank").IsPresent Then
        '                                        KHSQLcsentry("saldoTimebank").Value = mventry("saldoTidsbank").Value
        '                                    End If
        '                                    If mventry("saldoJubileumsuke").IsPresent Then
        '                                        KHSQLcsentry("saldoJubileumsuke").Value = mventry("saldoJubileumsuke").Value
        '                                    End If
        '                                    If mventry("saldoFerie").IsPresent Then
        '                                        KHSQLcsentry("saldoFerie").Value = mventry("saldoFerie").Value
        '                                    End If

        '                                    KHSQLcsentry.CommitNewConnector()

        '                                Catch ex As ObjectAlreadyExistsException
        '                                Catch ex As Exception
        '                                    Throw
        '                                    'IMAExtensible2CallExport 'bla
        '                                End Try
        '                            End If
        '                        End If

        '                    Else
        '                        'do nothing
        '                    End If
        '                Else
        '                    Dim yearResult As Integer = currentYear - endYear
        '                    If yearResult = 1 Then
        '                        If endMonth.Equals(12) And currentMonth.Equals(1) Then
        '                            If mventry("SQLaccountName").Value.Length > 2 Then
        '                                If KHSQL.Connectors.Count = 0 Then
        '                                    'If mventry("timestamp").IsPresent Then ''Readded 05.11.2021
        '                                    'Dim timeattrib As DateTime = Convert.ToDateTime(mventry("timestamp").Value)
        '                                    'Dim timenow As DateTime = DateTime.Now
        '                                    'Dim dur As TimeSpan = timenow - timeattrib
        '                                    'If dur.Minutes < 200 Then
        '                                    'do nothing 
        '                                    'Else '' Readded End
        '                                    Dim rdn As String = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                    DN = KHSQL.EscapeDNComponent(rdn)
        '                                    Try
        '                                        KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
        '                                        KHSQLcsentry.DN = DN

        '                                        KHSQLcsentry("AbsenceID").Value = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                        KHSQLcsentry("fromBadgeNo").Value = mventry("SQLaccountName").Value
        '                                        KHSQLcsentry("objectType").Value = "person"
        '                                        If mventry("name").IsPresent Then
        '                                            KHSQLcsentry("name").Value = mventry("name").Value
        '                                        End If
        '                                        If mventry("saldoOvertid").IsPresent Then
        '                                            KHSQLcsentry("saldoOvertid").Value = mventry("saldoOvertid").Value
        '                                        End If
        '                                        If mventry("extCode").IsPresent Then
        '                                            KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue
        '                                        End If
        '                                        If mventry("templateextid").IsPresent Then
        '                                            KHSQLcsentry("templateextid").Value = mventry("templateextid").Value
        '                                        End If
        '                                        If mventry("managerName").IsPresent Then
        '                                            KHSQLcsentry("manager").Value = mventry("managerName").Value
        '                                        End If
        '                                        If mventry("templatename").IsPresent Then
        '                                            KHSQLcsentry("templatename").Value = mventry("templatename").Value
        '                                        End If
        '                                        If mventry("employmentPercent").IsPresent Then
        '                                            KHSQLcsentry("employmentPercent").Value = mventry("employmentPercent").Value
        '                                        End If
        '                                        If mventry("unitname").IsPresent Then
        '                                            KHSQLcsentry("unitname").Value = mventry("unitname").Value
        '                                        End If
        '                                        If mventry("gender").IsPresent Then
        '                                            KHSQLcsentry("gender").Value = mventry("gender").Value
        '                                        End If
        '                                        If mventry("employeeStartDate").IsPresent Then
        '                                            KHSQLcsentry("timeStamp").Value = mventry("employeeStartDate").Value
        '                                        End If
        '                                        If mventry("saldoTidsbank").IsPresent Then
        '                                            KHSQLcsentry("saldoTimebank").Value = mventry("saldoTidsbank").Value
        '                                        End If
        '                                        If mventry("saldoJubileumsuke").IsPresent Then
        '                                            KHSQLcsentry("saldoJubileumsuke").Value = mventry("saldoJubileumsuke").Value
        '                                        End If
        '                                        If mventry("saldoFerie").IsPresent Then
        '                                            KHSQLcsentry("saldoFerie").Value = mventry("saldoFerie").Value
        '                                        End If

        '                                        KHSQLcsentry.CommitNewConnector()

        '                                    Catch ex As ObjectAlreadyExistsException
        '                                    Catch ex As Exception
        '                                        Throw
        '                                        'IMAExtensible2CallExport 'bla
        '                                    End Try
        '                                Else
        '                                    'do nothing
        '                                End If
        '                            End If
        '                        End If
        '                    End If
        '                End If
        '            Else
        '                If mventry("SQLaccountName").IsPresent Then
        '                    If mventry("SQLaccountName").Value.Length > 2 Then
        '                        If KHSQL.Connectors.Count = 0 Then
        '                            'If mventry("timestamp").IsPresent Then ''Readded 05.11.2021
        '                            'Dim timeattrib As DateTime = Convert.ToDateTime(mventry("timestamp").Value)
        '                            'Dim timenow As DateTime = DateTime.Now
        '                            'Dim dur As TimeSpan = timenow - timeattrib
        '                            'If dur.Minutes < 200 Then
        '                            'do nothing 
        '                            'Else '' Readded End
        '                            Dim rdn As String = mventry("SQLaccountName").Value & " & " & timeStamp
        '                            DN = KHSQL.EscapeDNComponent(rdn)
        '                            Try
        '                                KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
        '                                KHSQLcsentry.DN = DN

        '                                KHSQLcsentry("AbsenceID").Value = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                KHSQLcsentry("fromBadgeNo").Value = mventry("SQLaccountName").Value
        '                                KHSQLcsentry("objectType").Value = "person"
        '                                If mventry("name").IsPresent Then
        '                                    KHSQLcsentry("name").Value = mventry("name").Value
        '                                End If
        '                                If mventry("saldoOvertid").IsPresent Then
        '                                    KHSQLcsentry("saldoOvertid").Value = mventry("saldoOvertid").Value
        '                                End If
        '                                If mventry("extCode").IsPresent Then
        '                                    KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue
        '                                End If
        '                                If mventry("templateextid").IsPresent Then
        '                                    KHSQLcsentry("templateextid").Value = mventry("templateextid").Value
        '                                End If
        '                                If mventry("templatename").IsPresent Then
        '                                    KHSQLcsentry("templatename").Value = mventry("templatename").Value
        '                                End If
        '                                If mventry("managerName").IsPresent Then
        '                                    KHSQLcsentry("manager").Value = mventry("managerName").Value
        '                                End If
        '                                If mventry("employmentPercent").IsPresent Then
        '                                    KHSQLcsentry("employmentPercent").Value = mventry("employmentPercent").Value
        '                                End If
        '                                If mventry("unitname").IsPresent Then
        '                                    KHSQLcsentry("unitname").Value = mventry("unitname").Value
        '                                End If
        '                                If mventry("gender").IsPresent Then
        '                                    KHSQLcsentry("gender").Value = mventry("gender").Value
        '                                End If
        '                                If mventry("employeeStartDate").IsPresent Then
        '                                    KHSQLcsentry("timeStamp").Value = mventry("employeeStartDate").Value
        '                                End If
        '                                If mventry("saldoTidsbank").IsPresent Then
        '                                    KHSQLcsentry("saldoTimebank").Value = mventry("saldoTidsbank").Value
        '                                End If
        '                                If mventry("saldoJubileumsuke").IsPresent Then
        '                                    KHSQLcsentry("saldoJubileumsuke").Value = mventry("saldoJubileumsuke").Value
        '                                End If
        '                                If mventry("saldoFerie").IsPresent Then
        '                                    KHSQLcsentry("saldoFerie").Value = mventry("saldoFerie").Value
        '                                End If

        '                                KHSQLcsentry.CommitNewConnector()

        '                            Catch ex As ObjectAlreadyExistsException
        '                            Catch ex As Exception
        '                                Throw
        '                                'IMAExtensible2CallExport 'bla
        '                            End Try
        '                        Else
        '                            If KHSQL.Connectors.Count >= 1 Then
        '                                'If mventry("timestamp").IsPresent Then ''Readded 05.11.2021
        '                                'Dim timeattrib As DateTime = Convert.ToDateTime(mventry("timestamp").Value)
        '                                'Dim timenow As DateTime = DateTime.Now
        '                                'Dim dur As TimeSpan = timenow - timeattrib
        '                                'If dur.Minutes < 200 Then
        '                                'do nothing 
        '                                'Else '' Readded End
        '                                Dim rdn As String = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                DN = KHSQL.EscapeDNComponent(rdn)
        '                                Try
        '                                    KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
        '                                    KHSQLcsentry.DN = DN

        '                                    KHSQLcsentry("AbsenceID").Value = mventry("SQLaccountName").Value & " & " & timeStamp
        '                                    KHSQLcsentry("fromBadgeNo").Value = mventry("SQLaccountName").Value
        '                                    KHSQLcsentry("objectType").Value = "person"
        '                                    If mventry("name").IsPresent Then
        '                                        KHSQLcsentry("name").Value = mventry("name").Value
        '                                    End If
        '                                    If mventry("saldoOvertid").IsPresent Then
        '                                        KHSQLcsentry("saldoOvertid").Value = mventry("saldoOvertid").Value
        '                                    End If
        '                                    If mventry("extCode").IsPresent Then
        '                                        KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue
        '                                    End If
        '                                    If mventry("templateextid").IsPresent Then
        '                                        KHSQLcsentry("templateextid").Value = mventry("templateextid").Value
        '                                    End If
        '                                    If mventry("templatename").IsPresent Then
        '                                        KHSQLcsentry("templatename").Value = mventry("templatename").Value
        '                                    End If
        '                                    If mventry("managerName").IsPresent Then
        '                                        KHSQLcsentry("manager").Value = mventry("managerName").Value
        '                                    End If
        '                                    If mventry("employmentPercent").IsPresent Then
        '                                        KHSQLcsentry("employmentPercent").Value = mventry("employmentPercent").Value
        '                                    End If
        '                                    If mventry("unitname").IsPresent Then
        '                                        KHSQLcsentry("unitname").Value = mventry("unitname").Value
        '                                    End If
        '                                    If mventry("gender").IsPresent Then
        '                                        KHSQLcsentry("gender").Value = mventry("gender").Value
        '                                    End If
        '                                    If mventry("employeeStartDate").IsPresent Then
        '                                        KHSQLcsentry("timeStamp").Value = mventry("employeeStartDate").Value
        '                                    End If
        '                                    If mventry("saldoTidsbank").IsPresent Then
        '                                        KHSQLcsentry("saldoTimebank").Value = mventry("saldoTidsbank").Value
        '                                    End If
        '                                    If mventry("saldoJubileumsuke").IsPresent Then
        '                                        KHSQLcsentry("saldoJubileumsuke").Value = mventry("saldoJubileumsuke").Value
        '                                    End If
        '                                    If mventry("saldoFerie").IsPresent Then
        '                                        KHSQLcsentry("saldoFerie").Value = mventry("saldoFerie").Value
        '                                    End If

        '                                    KHSQLcsentry.CommitNewConnector()

        '                                Catch ex As ObjectAlreadyExistsException
        '                                Catch ex As Exception
        '                                    Throw
        '                                    'IMAExtensible2CallExport 'bla
        '                                End Try

        '                            End If
        '                            'End If
        '                            'End If
        '                        End If
        '                    End If
        '                End If
        '            End If

        '            'If mventry("SQLaccountName").IsPresent Then
        '            '        If mventry("SQLaccountName").Value.Length > 2 Then
        '            '            'If KHSQL.Connectors.Count = 0 Then
        '            '            Dim rdn As String = mventry("SQLaccountName").Value & " & " & timeStamp
        '            '            DN = KHSQL.EscapeDNComponent(rdn)

        '            '            Try
        '            '                KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
        '            '                KHSQLcsentry.DN = DN

        '            '                KHSQLcsentry("AbsenceID").Value = mventry("SQLaccountName").Value & " & " & timeStamp
        '            '                KHSQLcsentry("fromBadgeNo").Value = mventry("SQLaccountName").Value
        '            '                KHSQLcsentry("objectType").Value = "person"
        '            '                If mventry("extCode").IsPresent Then
        '            '                    KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue
        '            '                End If
        '            '                If mventry("templateextid").IsPresent Then
        '            '                    KHSQLcsentry("templateextid").Value = mventry("templateextid").Value
        '            '                End If
        '            '                If mventry("templatename").IsPresent Then
        '            '                    KHSQLcsentry("templatename").Value = mventry("templatename").Value
        '            '                End If
        '            '                If mventry("employmentPercent").IsPresent Then
        '            '                    KHSQLcsentry("employmentPercent").Value = mventry("employmentPercent").Value
        '            '                End If
        '            '                If mventry("unitname").IsPresent Then
        '            '                    KHSQLcsentry("unitname").Value = mventry("unitname").Value
        '            '                End If
        '            '                If mventry("gender").IsPresent Then
        '            '                    KHSQLcsentry("gender").Value = mventry("gender").Value
        '            '                End If
        '            '                If mventry("employeeStartDate").IsPresent Then
        '            '                    KHSQLcsentry("timeStamp").Value = mventry("employeeStartDate").Value
        '            '                End If
        '            '                If mventry("name").IsPresent Then
        '            '                    KHSQLcsentry("name").Value = mventry("name").Value
        '            '                End If
        '            '                If mventry("saldoTidsbank").IsPresent Then
        '            '                    KHSQLcsentry("saldoTimebank").Value = mventry("saldoTidsbank").Value
        '            '                End If
        '            '                If mventry("saldoJubileumsuke").IsPresent Then
        '            '                    KHSQLcsentry("saldoJubileumsuke").Value = mventry("saldoJubileumsuke").Value
        '            '                End If
        '            '                If mventry("saldoFerie").IsPresent Then
        '            '                    KHSQLcsentry("saldoFerie").Value = mventry("saldoFerie").Value
        '            '                End If


        '            '                'KHSQLcsentry("sickLevel").Value = mventry("sickLevel").Value
        '            '                'KHSQLcsentry("approvedByManager").Value = mventry("approvedByManager").Value
        '            '                KHSQLcsentry.CommitNewConnector()

        '            '            Catch ex As ObjectAlreadyExistsException
        '            '            Catch ex As Exception
        '            '                Throw
        '            '                'IMAExtensible2CallExport 'bla
        '            '            End Try
        '            '        End If
        '        End If
        '    End If
        'End If
        ''End If
        ''End If
        ''End If
        ''End If

        ''End If  -- 05.11.2021


        ''Else --- Removed 100921
        ''        'If KHSQL.Connectors.Count = 0 Then
        ''        Dim rdn As String = mventry("SQLaccountName").Value & " & " & timeStamp
        ''        DN = KHSQL.EscapeDNComponent(rdn)

        ''        Try
        ''            KHSQLcsentry = KHSQL.Connectors.StartNewConnector("absence")
        ''            KHSQLcsentry.DN = DN
        ''            KHSQLcsentry("AbsenceID").Value = mventry("SQLaccountName").Value & " & " & timeStamp
        ''            KHSQLcsentry("fromBadgeNo").Value = mventry("SQLaccountName").Value
        ''            KHSQLcsentry("objectType").Value = "person"
        ''            If mventry("extCode").IsPresent Then
        ''                KHSQLcsentry("extCode").IntegerValue = mventry("extCode").IntegerValue
        ''            End If
        ''            If mventry("templateextid").IsPresent Then
        ''                KHSQLcsentry("templateextid").Value = mventry("templateextid").Value
        ''            End If

        ''            If mventry("templatename").IsPresent Then
        ''                KHSQLcsentry("templatename").Value = mventry("templatename").Value
        ''            End If
        ''            If mventry("employmentPercent").IsPresent Then
        ''                KHSQLcsentry("employmentPercent").Value = mventry("employmentPercent").Value
        ''            End If
        ''            If mventry("unitname").IsPresent Then
        ''                KHSQLcsentry("unitname").Value = mventry("unitname").Value
        ''            End If
        ''            If mventry("gender").IsPresent Then
        ''                KHSQLcsentry("gender").Value = mventry("gender").Value
        ''            End If
        ''            If mventry("employeeStartDate").IsPresent Then
        ''                KHSQLcsentry("timeStamp").Value = mventry("employeeStartDate").Value
        ''            End If

        ''            'KHSQLcsentry("sickLevel").Value = mventry("sickLevel").Value
        ''            'KHSQLcsentry("approvedByManager").Value = mventry("approvedByManager").Value
        ''            KHSQLcsentry.CommitNewConnector()

        ''        Catch ex As ObjectAlreadyExistsException
        ''        Catch ex As Exception
        ''            Throw
        ''            'IMAExtensible2CallExport 'bla
        ''        End Try
        ''    End If
        ''End If
        '' End If
        ''End If
        ''--Fixremoval 05122024 End



    End Sub

    Public Function ShouldDeleteFromMV(ByVal csentry As CSEntry, ByVal mventry As MVEntry) As Boolean Implements IMVSynchronization.ShouldDeleteFromMV
        Select Case mventry("orphaned").BooleanValue
            Case False
                ShouldDeleteFromMV = False
            Case True
                ShouldDeleteFromMV = True
        End Select


    End Function
End Class
