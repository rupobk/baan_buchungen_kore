﻿Imports System.IO
Imports System.Data.SqlClient
Imports System.Diagnostics

Module Module1
    Dim conBaan As New SqlConnection("Server=baan4;Database=baan4db;User Id=sa; Password=Baan123;")
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
        Threading.Thread.Sleep(10000)
    End Sub

    Sub ProgrammParameterLesen()
        'args(0) = Programmname

        Try
            'Prozedurname
            args(1) = Environment.GetCommandLineArgs(1).ToUpper

        Catch ex As Exception
            Console.WriteLine("")
            Console.WriteLine("")
            Console.WriteLine("Programm baan_buchungen_kore wird mit folgenden Argumenten aufgerufen:")
            Console.WriteLine("fahrzeugkosten oder werkzeugkosten oder produktionskosten oder logistikkosten")
            LogSchreiben("Sub ProgrammParameterLesen: Programmaufrufparameter falsch!")
            LogSchreiben("Programm baan_buchungen_kore mit Fehler beendet.")
            Threading.Thread.Sleep(10000)
            Environment.Exit(0)
        End Try
    End Sub

    Sub BuchenFahrzeugkosten()
        Console.WriteLine("Prozedur BuchenFahrzeugkosten gestartet ...")
        LogSchreiben("Prozedur BuchenFahrzeugkosten gestartet ...")
        Try
            Using conBaan
                'Lösche alle bestehenden Bewegungen des Jahres weg
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_desc like '%(mov.auto%' and t_year=YEAR(DATEADD(MONTH, -1, GETDATE()))", conBaan)
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
                'Achtung: Fahrzeugkosten dürfen nicht auf den Mitarbeiter gebucht werden, weil sonst sieht er sie bei seinen täglichen Bewegungen u. erschreckt!
                command.CommandText = "INSERT INTO ttpppc231100 " +
                                "Select t_year, t_peri, a.t_emno as t_emno, ROW_NUMBER() OVER (ORDER BY a.t_emno)+10000 As t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                "'     ***' AS t_cact, t_tefx AS t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, round(t_wgrt, 2) as t_ratc, round(t_wgrt*t_quan, 2) AS t_amoc, " +
                                "0 As t_rats, 0 As t_amos, '001' AS t_cwgt, 'costi autovettura (mov.autom.)' AS t_desc, '' AS t_cdoc, t_year AS t_cyea, t_peri AS t_cper, " +
                                "t_peri As t_fper, t_year As t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, '1' AS t_potf, t_cstl, t_ccco, '1' AS t_tetc, '2' AS t_sttl, " +
                                "'0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, '0' AS t_serc, '0' AS t_wgcs, '0' AS t_serh, '0' AS t_Refcntd, '0' AS t_refcntu " +
                                "FROM (" +            'hier kommt das SQL, mit dem zuerst die Stunden pro Mitarbeiter u. Periode summiert werden
                                "select t_year, month(t_rgdt) as t_peri, 100000+t_emno as t_emno, t_cprj, t_cspa, '2020-'+right('0'+ltrim(str(month(t_rgdt))), 2)+'-28' as t_rgdt, " +
                                "t_cstl, t_ccco, sum(t_quan) as t_quan From ttpppc231100 " +
                                "Where t_year = Year(DateAdd(Month, -1, GETDATE())) And (LTrim(t_task) Between '10100' And '10530' or ltrim(t_task) Between '12100' AND '13830') " +
                                "And Not (LTrim(t_task) in ('11135', '13450')) Group By t_year, month(t_rgdt), t_emno, t_cprj, t_cspa, month(t_rgdt), t_cstl, t_ccco) a " +
                                "LEFT JOIN ttccom001100 b ON b.t_emno = a.t_emno WHERE a.t_emno IN (SELECT t_emno FROM ttccom001100 WHERE t_emno between 100000 and 199999)"
                command.ExecuteNonQuery()

                'bei Energie/Müll-Baustellen immer als Element 90000000 angeben
                command.CommandText = "UPDATE ttpppc231100 SET t_cspa='90000000' WHERE t_emno>100000 AND LTRIM(t_cspa)='***' AND t_cprj IN (SELECT t_cprj FROM ttpptc130100 WHERE t_cspa='90000000')"
                command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            LogSchreiben("Fehler In Sub BuchenFahrzeugkosten!")
            LogSchreiben(ex.Message)
            Console.WriteLine(ex.Message)
            LogSchreiben("Programm BuchenFahrzeugkosten mit Fehler beendet.")
            Threading.Thread.Sleep(10000)
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
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno=900250 and t_year=YEAR(DATEADD(MONTH, -1, GETDATE()))", conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert
                '- der Stundentarif wird aus den Stammdaten vom Mitarbeiter '900250' gezogen
                '- Das Jahr wird immer bis Ende Januar des Folgejahres mit dem Vorjahr gebucht u. anschliessend mit dem Folgejahr

                command.CommandText = "INSERT INTO ttpppc231100 " +
                                      "Select t_year, t_peri, '900250' as t_emno, ROW_NUMBER() OVER (ORDER BY t_cprj) AS t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                      "'     ***' AS t_cact, '   11136' as t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, b.t_wgrt, " +
                                      "round(t_quan*b.t_wgrt,2) As t_amoc, 0 As t_rats, 0 As t_amos, '001' AS t_cwgt, 'ribaltamento costi produzione' AS t_desc, " +
                                      "'' AS t_cdoc, t_year AS t_cyea, t_peri AS t_cper, t_peri AS t_fper, t_year AS t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, " +
                                      "'1' AS t_potf, t_cstl, '    1060' AS t_ccco, '1' AS t_tetc, '2' AS t_sttl, '0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, " +
                                      "'0' AS t_serc, '0' AS t_wgcs, '0' AS t_serh, '0' AS t_Refcntd, '0' AS t_refcntu " +
                                      "FROM (" +            'hier kommt das SQL, mit dem zuerst die Stunden pro Mitarbeiter u. Periode summiert werden
                                      "select t_year, month(t_rgdt) as t_peri, t_cprj, t_cspa, '2020-'+right('0'+ltrim(str(month(t_rgdt))), 2)+'-28' " +
                                      "as t_rgdt, t_cstl, sum(t_quan) as t_quan From ttpppc231100 " +
                                      "Where t_year = Year(DateAdd(Month, -1, GETDATE())) And (LTrim(t_task)) like '104%' " +
                                      "Group By t_year, month(t_rgdt), t_cprj, t_cspa, month(t_rgdt), t_cstl " +
                                      ") summen LEFT JOIN ttccom001100 b ON b.t_emno=900250"
                '                                      "Select t_year, Case when year(getdate())>t_year then 12 else Month(getdate()) end as t_peri, t_cprj, t_cspa, getdate() AS t_rgdt, t_cstl, SUM(t_quan) AS t_quan From ttpppc231100 " +
                '                                      "Where t_year = YEAR(DATEADD(MONTH, -1, GETDATE())) And LTrim(t_task) like '104%' " +
                '                                     "GROUP BY t_year, Month(t_rgdt), t_cprj, t_cspa, LTrim(Str(t_year)) + Right('0'+LTRIM(STR(MONTH(t_rgdt))), 2)+'28', t_cstl " +

                command.ExecuteNonQuery()

                'Auch Produktionskosten neu rechnen mit Plantarif
                command.CommandText = "UPDATE ttihra100100 SET t_wgca=round(t_hrea*b.t_wgrt, 2) " +
                                      "FROM ttihra100100 LEFT JOIN ttccom001100 b ON b.t_emno=900250 WHERE t_year=YEAR(DATEADD(MONTH, -1, GETDATE())) And t_tano<160"
                command.ExecuteNonQuery()

                'bei Energie/Müll-Baustellen immer als Element 90000000 angeben
                command.CommandText = "UPDATE ttpppc231100 SET t_cspa='90000000' WHERE t_emno>900000 AND LTRIM(t_cspa)='***' AND t_cprj IN (SELECT t_cprj FROM ttpptc130100 WHERE t_cspa='90000000')"
                command.ExecuteNonQuery()
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
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno=900230 and t_year=YEAR(DATEADD(MONTH, -1, GETDATE()))", conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert
                '- der Stundentarif wird aus den Stammdaten vom Mitarbeiter '900230' gezogen
                '- Das Jahr wird immer bis Ende Januar des Folgejahres mit dem Vorjahr gebucht u. anschliessend mit dem Folgejahr

                command.CommandText = "INSERT INTO ttpppc231100 " +
                                 "SELECT t_year, t_peri, '900230' as t_emno, ROW_NUMBER() OVER (ORDER BY t_cprj) AS t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                 "'     ***' AS t_cact, '   11133' as t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, b.t_wgrt, round(t_quan*b.t_wgrt, 2) " +
                                 "As t_amoc, 0 As t_rats, 0 As t_amos, '001' AS t_cwgt, 'ribaltamento costi attrezzi' AS t_desc, '' AS t_cdoc, " +
                                 "t_year AS t_cyea, t_peri AS t_cper, t_peri AS t_fper, t_year AS t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, '1' AS t_potf, t_cstl, " +
                                 "'    1060' AS t_ccco, '1' AS t_tetc, '2' AS t_sttl, '0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, '0' AS t_serc, '0' AS t_wgcs, " +
                                 "'0' AS t_serh, '0' AS t_Refcntd, '0' AS t_refcntu FROM (" +  'hier kommt das SQL, mit dem zuerst die Stunden pro Periode summiert werden
                                 "Select t_year, month(t_rgdt) as t_peri, t_cprj, t_cspa, '2020-'+right('0'+ltrim(str(month(t_rgdt))), 2)+'-28' AS t_rgdt, t_cstl, SUM(t_quan) AS t_quan FROM (" +
                                 "SELECT t_year, t_rgdt, t_cprj, t_cspa, t_cstl, t_quan From ttpppc231100 Where t_year = YEAR(DATEADD(MONTH, -1, GETDATE())) And LTrim(t_task) Between '10100' AND '10160' " +
                                 "UNION all " +
                                 "Select t_year, t_rgdt, t_cprj, t_cspa, t_cstl, t_quan FROM ttpppc271100 WHERE t_year = YEAR(DATEADD(MONTH, -1, GETDATE())) And LTrim(t_csub) BETWEEN 'S100' AND 'S150' " +
                                 ") detail GROUP BY t_year, month(t_rgdt), t_cprj, t_cspa, month(t_rgdt), t_cstl " +
                                 ") summen LEFT JOIN ttccom001100 b ON b.t_emno=900230"
                command.ExecuteNonQuery()

                'bei Energie/Müll-Baustellen immer als Element 90000000 angeben
                command.CommandText = "UPDATE ttpppc231100 SET t_cspa='90000000' WHERE t_emno>900000 AND LTRIM(t_cspa)='***' AND t_cprj IN (SELECT t_cprj FROM ttpptc130100 WHERE t_cspa='90000000')"
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
                Dim command As SqlCommand = New SqlCommand("DELETE from ttpppc231100 where t_emno=900270 and t_year=YEAR(DATEADD(MONTH, -1, GETDATE()))", conBaan)
                conBaan.Open()
                command.Connection = conBaan
                command.ExecuteNonQuery()
                'Buche jetzt neu das komplette Jahr
                'Wichtig:
                '- als t_rgdt nehme ich den 28. des jeweiligen Monats, auf diese Weise muss ich nicht umständlich berechnen, welcher der letzte Tag im Monat ist
                '- mit der Funktion ROW_NUMBER() wird die fortlaufende Bewegungsnummer generiert
                '- der Stundentarif wird aus den Stammdaten vom Mitarbeiter '900270' gezogen
                '- Das Jahr wird immer bis Ende Januar des Folgejahres mit dem Vorjahr gebucht u. anschliessend mit dem Folgejahr

                command.CommandText = "INSERT INTO ttpppc231100 " +
                                      "SELECT t_year, t_peri, '900270' as t_emno, ROW_NUMBER() OVER (ORDER BY t_peri, t_cprj) AS t_sern, t_cprj, t_cspa, '  1' AS t_cpla, " +
                                      "'     ***' AS t_cact, '   11137' as t_task, GETDATE() AS t_ltdt, t_rgdt, t_quan, b.t_wgrt, round(t_quan*b.t_wgrt, 2) " +
                                      "AS t_amoc, 0 AS t_rats, 0 AS t_amos, '001' AS t_cwgt, 'ribaltamento costi logistiche' AS t_desc, '' AS t_cdoc, t_year AS t_cyea, " +
                                      "t_peri AS t_cper, t_peri AS t_fper, t_year AS t_fyea, '100' AS t_ncmp, '2' AS t_cfpo, '1' AS t_potf, t_cstl, '    1060' AS t_ccco, " +
                                      "'1' AS t_tetc, '2' AS t_sttl, '0' AS t_txta, 'damii' AS t_loco, '0' AS t_hemp, '0' AS t_serc, '0' AS t_wgcs, '0' AS t_serh, " +
                                      "'0' AS t_Refcntd, '0' AS t_refcntu FROM ( " +
                                      "Select jahr As t_year, month(datum) as t_peri, baustelle As t_cprj, '     ***' AS t_cspa, '2020-'+right('0'+ltrim(str(month(datum))), 2)+'-28' AS t_rgdt, '' AS t_cstl, " +
                                      "count(*) AS t_quan " +
                                      "From [SRVATZDCBZ040\PREVERO,1434].prev_staging_prod.dbo.t_belegdetail_material Where jahr = YEAR(DATEADD(MONTH, -1, GETDATE())) " +
                                      "and ze1=1 Group By jahr, month(datum), baustelle, month(datum) " +
                                      ") summen LEFT JOIN ttccom001100 b ON b.t_emno=900270"
                command.ExecuteNonQuery()

                'bei Energie/Müll-Baustellen immer als Element 90000000 angeben
                command.CommandText = "UPDATE ttpppc231100 SET t_cspa='90000000' WHERE t_emno>900000 AND LTRIM(t_cspa)='***' AND t_cprj IN (SELECT t_cprj FROM ttpptc130100 WHERE t_cspa='90000000')"
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
