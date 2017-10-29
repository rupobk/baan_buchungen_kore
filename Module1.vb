Imports System.IO
Imports System.Data.SqlClient
Imports System.Diagnostics

Module Module1
    Dim conBaan As New SqlConnection("Server=baan4;Database=baan4db;User Id=sa; Password=Baan123;")
    Public jahr As String
    Public args(10) As String
    Public stundensatz As String
    Public bewegungssatz As String

    Sub Main()
        LogSchreiben("Programm baan_buchungen_kore gestartet.")
        Console.WriteLine("Programm baan_buchungen_kore gestartet!")

        ProgrammParameterLesen()

        Select Case args(1)
            Case "FAHRZEUGKOSTEN"
                BuchenFahrzeugkosten()
            Case "WERKZEUGKOSTEN"
                BuchenWerkzeugkosten()
            Case "PRODUKTIONSKOSTEN"
                BuchenProduktionskosten()
            Case "LOGISTIKKOSTEN"
                BuchenLogistikkosten()
                'Case else ...
        End Select


        Console.WriteLine("Programm baan_buchungen_kore ohne Fehler beendet!")
        LogSchreiben("Programm baan_buchungen_kore ohne Fehler beendet!")
        Threading.Thread.Sleep(30000)
    End Sub

    Sub ProgrammParameterLesen()
        'args(0) = Programmname
        'args(1) = Prozedurname

        'Wenn Prozedurname=fahrzeugkosten, dann:
        'args(2) = Geschäftsjahr

        'Wenn Prozedurname=werkzeugkosten, dann:
        'args(2) = Geschäftsjahr
        'args(3) = Stundensatz

        'Wenn Prozedurname=produktionskosten, dann:
        'args(2) = Geschäftsjahr
        'args(3) = Stundensatz

        'Wenn Prozedurname=logistikkosten, dann:
        'args(2) = Geschäftsjahr
        'args(3) = Kosten pro Bewegung

        Try
            'Prozedurname
            args(1) = Environment.GetCommandLineArgs(1).ToUpper

            Select Case args(1)
                Case "FAHRZEUGKOSTEN"
                    args(2) = Environment.GetCommandLineArgs(2)   'Geschäftsjahr
                    jahr = args(2)
                Case "WERKZEUGKOSTEN"
                    args(2) = Environment.GetCommandLineArgs(2)  'Geschäftsjahr
                    args(3) = Environment.GetCommandLineArgs(3)  'Stundensatz
                    jahr = args(2)
                    stundensatz = args(3)
                Case "PRODUKTIONSKOSTEN"
                    args(2) = Environment.GetCommandLineArgs(2)  'Geschäftsjahr
                    args(3) = Environment.GetCommandLineArgs(3)  'Stundensatz
                    jahr = args(2)
                    stundensatz = args(3)
                Case "LOGISTIKKOSTEN"
                    args(2) = Environment.GetCommandLineArgs(2)  'Geschäftsjahr
                    args(3) = Environment.GetCommandLineArgs(3)  'Kostensatz Bewegung
                    jahr = args(2)
                    bewegungssatz = args(3)

            End Select

        Catch ex As Exception
            Console.WriteLine("")
            Console.WriteLine("")
            Console.WriteLine("Programm baan_buchungen_kore wird mit folgenden Argumenten aufgerufen:")
            Select Case args(1)
                Case "FAHRZEUGKOSTEN"
                    Console.WriteLine("- Prozedurname fahrzeugkosten")
                    Console.WriteLine("- Geschäftsjahr")
                Case "WERKZEUGKOSTEN"
                    Console.WriteLine("- Prozedurname werkzeugkosten")
                    Console.WriteLine("- Geschäftsjahr")
                    Console.WriteLine("- Stundensatz (Achtung auf '.' anstatt Komma bei Kommastellen)")
                Case "PRODUKTIONSKOSTEN"
                    Console.WriteLine("- Prozedurname produktionskosten")
                    Console.WriteLine("- Geschäftsjahr")
                    Console.WriteLine("- Stundensatz (Achtung auf '.' anstatt Komma bei Kommastellen)")
                Case "LOGISTIKKOSTEN"
                    Console.WriteLine("- Prozedurname logistikkosten")
                    Console.WriteLine("- Geschäftsjahr")
                    Console.WriteLine("- Bewegungssatz (Achtung auf '.' anstatt Komma bei Kommastellen)")
                Case Else
                    Console.WriteLine("- Prozedur: welche Metadaten sollen generiert werden (fahrzeugkosten, werkzeugkosten, produktionskosten, logistikkosten)")
                    Console.WriteLine("- weitere Parameter je nach Prozedurname")
                    Console.WriteLine("Zweck des Programmes: Generierung von Kore-Buchungen für die Deckungsbeitragsrechnung.")
            End Select
            LogSchreiben("Sub ProgrammParameterLesen: Programmaufrufparameter falsch!")
            LogSchreiben("Programm baan_buchungen_kore mit Fehler beendet.")
            Threading.Thread.Sleep(30000)
            Environment.Exit(0)
        End Try
    End Sub

    Sub BuchenFahrzeugkosten()
        Console.WriteLine("Prozedur BuchenFahrzeugkosten gestartet ...")
        LogSchreiben("Prozedur BuchenFahrzeugkosten gestartet ...")
        Try
            Using conBaan
                'Lösche alle bestehenden Bewegungen des Jahres weg
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno between 100000 and 199999 and t_year=" + jahr, conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert
                '- t_task wird aus den Stammdaten gezogen
                '- where t_rats between 1 and 3 nimmt nur die Bewegungen der Arbeiter, Fahrer u. Techniker
                '- Der Stundentarif kommt aus den Stammdaten der Personalnr. mit Kodex über 100000
                command.CommandText = "INSERT INTO ttpppc231100 " +
                                "Select t_year, t_peri, a.t_emno, ROW_NUMBER() OVER (ORDER BY a.t_emno) As t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                "'     ***' AS t_cact, '   '+t_tefx AS t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, round(t_wgrt, 2), round(t_wgrt*t_quan, 2) AS t_amoc, " +
                                "0 As t_rats, 0 As t_amos, '001' AS t_cwgt, 'costi autovettura' AS t_desc, '' AS t_cdoc, t_year AS t_cyea, t_peri AS t_cper, " +
                                "t_peri As t_fper, t_year As t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, '1' AS t_potf, t_cstl, t_ccco, '1' AS t_tetc, '2' AS t_sttl, " +
                                "'0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, '0' AS t_serc, '0' AS t_wgcs, '0' AS t_serh, '0' AS t_Refcntd, '0' AS t_refcntu " +
                                "FROM " +
                                "( " +            'hier kommt das SQL, mit dem zuerst die Stunden pro Mitarbeiter u. Periode summiert werden
                                "Select t_year, MONTH(t_rgdt) As t_peri, 100000+t_emno As t_emno, t_cprj, t_cspa, LTRIM(STR(t_year))+ " +
                                "right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28' AS t_rgdt, t_cstl, t_ccco, SUM(t_quan) AS t_quan " +
                                "From ttpppc231100 " +
                                "Where t_year = " + jahr + " And (ltrim(t_task) Between '10100' And '10530' or ltrim(t_task) Between '12100' AND '13830') " +
                                "Group By t_year, month(t_rgdt), t_emno, t_cprj, t_cspa, LTRIM(STR(t_year)) + right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28', t_cstl, t_ccco " +
                                ") a " +
                                "LEFT JOIN ttccom001100 b ON b.t_emno = a.t_emno " +
                                "WHERE t_rats BETWEEN 1 And 3"
                command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            LogSchreiben("Fehler In Sub BuchenFahrzeugkosten!")
            LogSchreiben(ex.Message)
            Console.WriteLine(ex.Message)
            LogSchreiben("Programm BuchenFahrzeugkosten mit Fehler beendet.")
            Threading.Thread.Sleep(30000)
            Environment.Exit(0)
        End Try
        Console.WriteLine("Fahrzeugkosten sind gebucht!")
        LogSchreiben("Prozedur BuchenFahrzeugkosten beendet.")

    End Sub

    Sub BuchenProduktionskosten()
        Console.WriteLine("Prozedur BuchenProduktionskosten gestartet ...")
        LogSchreiben("Prozedur BuchenProduktionskosten gestartet ...")
        Try
            Using conBaan
                'Lösche alle bestehenden Bewegungen des Jahres weg
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno=900250 And t_year=" + jahr, conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert

                command.CommandText = "INSERT INTO ttpppc231100 " +
                                      "Select t_year, t_peri, '900250' as t_emno, ROW_NUMBER() OVER (ORDER BY t_cprj) AS t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                      "'     ***' AS t_cact, '   11136' as t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, " + stundensatz + " as t_wgrt, " +
                                      "round(t_quan*" + stundensatz + ",2) As t_amoc, 0 As t_rats, 0 As t_amos, '001' AS t_cwgt, 'ribaltamento costi produzione' AS t_desc, " +
                                      "'' AS t_cdoc, t_year AS t_cyea, t_peri AS t_cper, t_peri AS t_fper, t_year AS t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, " +
                                      "'1' AS t_potf, t_cstl, '    1060' AS t_ccco, '1' AS t_tetc, '2' AS t_sttl, '0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, " +
                                      "'0' AS t_serc, '0' AS t_wgcs, '0' AS t_serh, '0' AS t_Refcntd, '0' AS t_refcntu " +
                                      "FROM (" +            'hier kommt das SQL, mit dem zuerst die Stunden pro Mitarbeiter u. Periode summiert werden
                                      "Select t_year, MONTH(t_rgdt) As t_peri, t_cprj, t_cspa, LTRIM(STR(t_year))+right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28' AS t_rgdt, " +
                                      "t_cstl, SUM(t_quan) AS t_quan " +
                                      "From ttpppc231100 " +
                                      "Where t_year = " + jahr + " And LTrim(t_task) like '104%' " +
                                      "GROUP BY t_year, Month(t_rgdt), t_cprj, t_cspa, LTrim(Str(t_year)) + Right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28', t_cstl " +
                                      ") summen"
                Command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            LogSchreiben("Fehler In Sub BuchenProduktionskosten!")
            LogSchreiben(ex.Message)
            Console.WriteLine(ex.Message)
            LogSchreiben("Programm BuchenProduktionskosten mit Fehler beendet.")
            Threading.Thread.Sleep(30000)
            Environment.Exit(0)
        End Try
        Console.WriteLine("Produktionskosten sind gebucht!")
        LogSchreiben("BuchenProduktionskosten beendet.")
    End Sub

    Sub BuchenWerkzeugkosten()
        Console.WriteLine("Prozedur BuchenWerkzeugkosten gestartet ...")
        LogSchreiben("Prozedur BuchenWerkzeugkosten gestartet ...")
        Try
            Using conBaan
                'Lösche alle bestehenden Bewegungen des Jahres weg
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno=900230 and t_year=" + jahr, conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert

                command.CommandText = "INSERT INTO ttpppc231100 " +
                                 "SELECT t_year, t_peri, '900230' as t_emno, ROW_NUMBER() OVER (ORDER BY t_cprj) AS t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                 "'     ***' AS t_cact, '   11133' as t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, " + stundensatz + " as t_wgrt, " + stundensatz +
                                 " * t_quan As t_amoc, 0 As t_rats, 0 As t_amos, '001' AS t_cwgt, 'ribaltamento costi attrezzi' AS t_desc, '' AS t_cdoc, " +
                                 "t_year AS t_cyea, t_peri AS t_cper, t_peri AS t_fper, t_year AS t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, '1' AS t_potf, t_cstl, " +
                                 "'    1060' AS t_ccco, '1' AS t_tetc, '2' AS t_sttl, '0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, '0' AS t_serc, '0' AS t_wgcs, " +
                                 "'0' AS t_serh, '0' AS t_Refcntd, '0' AS t_refcntu FROM (" +  'hier kommt das SQL, mit dem zuerst die Stunden pro Periode summiert werden
                                 "Select t_year, Month(t_rgdt) As t_peri, t_cprj, t_cspa, LTrim(Str(t_year)) + Right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28' AS t_rgdt, " +
                                 "t_cstl, SUM(t_quan) AS t_quan FROM ( " +    'hier kommt das SQL mit den Detailstunden eigene u. Dritte Arbeiterstunden
                                 "SELECT t_year, t_rgdt, t_cprj, t_cspa, t_cstl, t_quan From ttpppc231100 Where t_year = " + jahr + " And LTrim(t_task) Between '10100' AND '10160' " +
                                 "UNION all " +
                                 "Select t_year, t_rgdt, t_cprj, t_cspa, t_cstl, t_quan FROM ttpppc271100 WHERE t_year = " + jahr + " And LTrim(t_csub) BETWEEN 'S100' AND 'S150' " +
                                 ") detail " +
                                 "GROUP BY t_year, Month(t_rgdt), t_cprj, t_cspa, LTrim(Str(t_year)) + Right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28', t_cstl " +
                                 ") summen"
                command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            LogSchreiben("Fehler In Sub BuchenWerkzeugkosten!")
            LogSchreiben(ex.Message)
            Console.WriteLine(ex.Message)
            LogSchreiben("Programm BuchenWerkzeugkosten mit Fehler beendet.")
            Threading.Thread.Sleep(30000)
            Environment.Exit(0)
        End Try
        Console.WriteLine("Werkzeugkosten sind gebucht!")
        LogSchreiben("Prozedur BuchenWerkzeugkosten beendet.")
    End Sub

    Sub BuchenLogistikkosten()
        Console.WriteLine("Prozedur BuchenLogistikkosten gestartet ...")
        LogSchreiben("Prozedur BuchenLogistikkosten gestartet ...")
        Try
            Using conBaan
                'Lösche alle bestehenden Bewegungen des Jahres weg
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno=900270 and t_year=" + jahr, conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert

                command.CommandText = "INSERT INTO ttpppc231100 " +
                                      "SELECT t_year, t_peri, '900270' as t_emno, ROW_NUMBER() OVER (ORDER BY t_peri, t_cprj) AS t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                      "'     ***' AS t_cact, '   11137' as t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, " + bewegungssatz + " as t_wgrt, " + bewegungssatz +
                                      "*t_quan AS t_amoc, 0 AS t_rats, 0 AS t_amos, '001' AS t_cwgt, 'ribaltamento costi logistiche' AS t_desc, '' AS t_cdoc, t_year AS t_cyea, " +
                                      "t_peri AS t_cper, t_peri AS t_fper, t_year AS t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, '1' AS t_potf, t_cstl, '    1060' AS t_ccco, " +
                                      "'1' AS t_tetc, '2' AS t_sttl, '0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, '0' AS t_serc, '0' AS t_wgcs, '0' AS t_serh, " +
                                      "'0' AS t_Refcntd, '0' AS t_refcntu FROM ( " +
                                      "Select jahr As t_year, MONTH(datum) As t_peri, baustelle As t_cprj, '     ***' AS t_cspa, " +
                                      "LTRIM(STR(jahr))+right('0'+LTRIM(STR(MONTH(datum))), 2)+'28' AS t_rgdt, '' AS t_cstl, count(*) AS t_quan " +
                                      "From SRVPREVERO.prev_staging_prod.dbo.t_belegdetail_material Where jahr = " + jahr +
                                      "Group By jahr, Month(datum), baustelle, LTrim(Str(jahr)) + Right('0'+LTRIM(STR(MONTH(datum))), 2)+'28' " +
                                      ") summen"
                command.ExecuteNonQuery()

                'bei Ecomaster-Baustellen immer als Element 90000000 angeben
                command.CommandText = "update ttpppc231100 set t_cspa='90000000' where t_cprj like 'E%' and t_emno=900270"
                command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            LogSchreiben("Fehler In Sub BuchenLogistikkosten!")
            LogSchreiben(ex.Message)
            Console.WriteLine(ex.Message)
            LogSchreiben("Programm BuchenLogistikkosten mit Fehler beendet.")
            Threading.Thread.Sleep(30000)
            Environment.Exit(0)
        End Try
        Console.WriteLine("Logistikkosten sind gebucht!")
        LogSchreiben("Prozedur BuchenLogistikkosten beendet.")
    End Sub

    Sub LogSchreiben(ByVal sEvent As String)
        Dim sSource As String
        Dim sLog As String

        Try

            sSource = "baan_buchungen_kore"
            sLog = "Application"

            If Not EventLog.SourceExists(sSource) Then
                EventLog.CreateEventSource(sSource, sLog)
            End If

            EventLog.WriteEntry(sSource, sEvent)
            EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Warning, 234)

        Catch ex As Exception
        End Try

    End Sub

End Module
