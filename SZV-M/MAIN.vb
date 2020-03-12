Imports System.Text.UTF8Encoding
Imports System.Security.Cryptography
'' TODO - VALIDATE AND DETECT MISSING FROM DO/PO
Public Class MAIN
#Region "Globs"
    Dim MODE_PO As Boolean
    Dim StatusName() As String = {"База готова к импорту файлов ДО", "Импорт ДО завершен, база готова к импорту ПОСЛЕ"}
    Dim Calc_Month As Short
    Dim DB_RA As Short
#End Region



    Dim x As Process = Process.GetCurrentProcess()
    Dim CurrState As String = ""

    Private Sub MAIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LBL_STATE.Text = StatusName(GetDBStatus)
        TS_MEMUSE.Text = ""
        If IO.Directory.Exists(Application.StartupPath & "\DO") = False Then
            IO.Directory.CreateDirectory(Application.StartupPath & "\DO")
        End If
        If IO.Directory.Exists(Application.StartupPath & "\PO") = False Then
            IO.Directory.CreateDirectory(Application.StartupPath & "\PO")
        End If
        'CleanUP(0)
        '' MsgBox("START " & My.Settings.MAN_md5)
        Optimize_Base()
        CB_MONTH.SelectedIndex = Date.Now.Month - 1
    End Sub

    Private Sub BTN_START_Click(sender As Object, e As EventArgs) Handles BTN_START.Click
        If GetDBStatus() = 1 Then
            If MsgBox("Базы ДО уже импортированы. Очистить и импортировать снова?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                CleanUP(CleanType.ONLY_DO_PO)
            Else
                Exit Sub
            End If
        End If
        Calc_Month = CB_MONTH.SelectedIndex + 1
        BTN_START.Enabled = False
        BTN_START_PO.Enabled = False
        BTN_SETTINGS.Enabled = False
        CB_MONTH.Enabled = False
        BW_IndexBases.RunWorkerAsync()
    End Sub
    Private Sub Optimize_Base()
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()

        cmd = sqlite_con.CreateCommand

        cmd.CommandText = "PRAGMA synchronous = OFF"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "PRAGMA journal_mode = OFF"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "PRAGMA page_size = 10000"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "VACUUM"
        cmd.ExecuteNonQuery()
    End Sub
    Enum CleanType
        ALL
        ONLY_DO
        ONLY_PO
        ONLY_DO_PO
        ONLY_MAN
        ONLY_PE
        ONLY_WPR
    End Enum

    Public Sub CleanUP(Typ As CleanType)

        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()

        cmd = sqlite_con.CreateCommand
        Select Case Typ
            Case CleanType.ALL
                '' MsgBox("ALL")
                cmd.CommandText = "DELETE FROM MAN;DELETE FROM PE; DELETE FROM PAY_DO; DELETE FROM PAY_PO; DELETE FROM POPEN_DO; DELETE FROM POPEN_PO; DELETE FROM RECIP_DO; DELETE FROM RECIP_PO;DELETE FROM IND4; DELETE FROM WPR"
                SetDBParam("PE_MD5", "0")
                SetDBParam("MAN_MD5", "0")
                '  My.Settings.WPR_md5 = ""
                cmd.ExecuteNonQuery()
                SetDBStatus(0)
            Case CleanType.ONLY_DO
                cmd.CommandText = "DELETE FROM MAN;DELETE FROM PE; DELETE FROM PAY_DO; DELETE FROM POPEN_DO; DELETE FROM RECIP_DO;DELETE FROM IND4"
                cmd.ExecuteNonQuery()
            Case CleanType.ONLY_PO
                cmd.CommandText = "DELETE FROM PAY_PO; DELETE FROM POPEN_PO; DELETE FROM RECIP_PO;DELETE FROM IND4"
                cmd.ExecuteNonQuery()
            Case CleanType.ONLY_DO_PO
                cmd.CommandText = "DELETE FROM PAY_DO; DELETE FROM PAY_PO; DELETE FROM POPEN_DO; DELETE FROM POPEN_PO; DELETE FROM RECIP_DO; DELETE FROM RECIP_PO;DELETE FROM IND4"
                cmd.ExecuteNonQuery()
                SetDBStatus(0)
            Case CleanType.ONLY_MAN
                SetDBParam("MAN_MD5", "0")
                cmd.CommandText = "DELETE FROM MAN"
                cmd.ExecuteNonQuery()
            Case CleanType.ONLY_PE
                SetDBParam("PE_MD5", "0")
                cmd.CommandText = "DELETE FROM PE"
                cmd.ExecuteNonQuery()
            Case CleanType.ONLY_WPR
                'My.Settings.WPR_md5 = ""
                cmd.CommandText = "DELETE FROM WPR"
                cmd.ExecuteNonQuery()
        End Select
        cmd.CommandText = "VACUUM"
        cmd.ExecuteNonQuery()
        My.Settings.Save()
        'LBL_STATE.Text = StatusName(GetDBStatus)

    End Sub
    Private Sub Mem_Timer_Tick(sender As Object, e As EventArgs) Handles Mem_Timer.Tick
        x = Process.GetCurrentProcess()
        Dim inf As String
        inf = "Использование памяти: " & Decimal.Round(x.WorkingSet64 / 1024 / 1024, 2, MidpointRounding.AwayFromZero) & " MB / 1536 MB"
        TS_MEMUSE.Text = inf
    End Sub

    Private Sub BW_IndexBases_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_IndexBases.DoWork
        'Dim man_strings(), pay_strings(), pe_strings(), popen_strings(), recip_strings(), recip_po_strings(), pay_po_strings(), popen_po_strings() As String

        Dim total_strings As Integer = 0
        ''connect to
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        Dim transact As SQLite.SQLiteTransaction
        cmd = sqlite_con.CreateCommand
        ' cmd.CommandText = "INSERT INTO MAN"
        Dim counter As Integer = 0

        If MODE_PO = False Then


            Try
                Dim MAN_MD5 As String = CreateMD5Sum(Application.StartupPath & "\DO\man.csv")

                If MAN_MD5 = GetDBParam("MAN_MD5") Then
                    '    MsgBox("SKIP " & MAN_MD5 & " " & GetDBParam("MAN_MD5"))
                    ''INSERT CODE IF NEEDED
                Else



                    CleanUP(CleanType.ONLY_MAN)
                    SetDBParam("MAN_MD5", MAN_MD5)
                    Dim SR As New IO.StreamReader(Application.StartupPath & "\DO\man.csv", System.Text.Encoding.GetEncoding(1251))
                    total_strings = GetNumberOfLines(Application.StartupPath & "\DO\man.csv")
                    CurrState = "Индексация MAN "
                    transact = sqlite_con.BeginTransaction()
                    While SR.Peek >= 0
                        Dim cols() As String = SR.ReadLine.Split(";")

                        cmd.CommandText = "INSERT INTO MAN (ADR_INDEX,FA,ID,IM,NPERS,OT,RA,RDAT) VALUES ('" & cols(0) & "','" & cols(1).Replace("'", " ") & "','" & cols(2) & "','" & cols(3).Replace("'", " ") & "','" & cols(4) & "','" & cols(5).Replace("'", " ") & "','" & cols(6) & "','" & cols(7) & "')"
                        cmd.ExecuteNonQuery()
                        counter += 1
                        If counter Mod 10000 = 0 Or counter = total_strings Or counter = 0 Then

                            BW_IndexBases.ReportProgress(counter, total_strings)
                        End If
                    End While
                    transact.Commit()
                End If
            Catch ex As Exception
                MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))


                Exit Sub
            End Try



            'CLEANUP MAN
            ' Erase man_strings


            Try

                ' input_strings = IO.File.ReadAllLines(Application.StartupPath & "\DO\pay.csv", System.Text.Encoding.GetEncoding(1251))
                Dim SR As New IO.StreamReader(Application.StartupPath & "\DO\pay.csv", System.Text.Encoding.GetEncoding(1251))
                total_strings = GetNumberOfLines(Application.StartupPath & "\DO\pay.csv")
                counter = 0
                CurrState = "Индексация PAY DO "
                transact = sqlite_con.BeginTransaction()
                While SR.Peek >= 0
                    Dim cols() As String = SR.ReadLine.Split(";")
                    If cols.Length = 18 Then
                        cols = cols.Take(cols.Count - 1).ToArray
                    End If
                    cmd.CommandText = "INSERT INTO PAY_DO (ACTION_TYPE,AMOUNT,FILE_ID,ID,INFO_TYPE,IST,MAN_ID,MONTH,PARENT_ID,RA,RAZDEL,RE,RECIPIENT_ID,SOURCE_ID,SUBVED_ID,VERSION,YEAR) VALUES('" & String.Join("','", cols) & "')"
                    cmd.ExecuteNonQuery()
                    counter += 1
                    If counter Mod 1000000 = 0 Then
                        transact.Commit()
                        transact = sqlite_con.BeginTransaction()
                    End If
                    If counter Mod 100000 = 0 Or counter = total_strings Then

                        BW_IndexBases.ReportProgress(counter, total_strings)
                    End If
                End While
                transact.Commit()
            Catch ex As Exception


                MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))


                Exit Sub
            End Try


            'Try
            '    '  input_strings = IO.File.ReadAllLines(Application.StartupPath & "\DO\recip.csv", System.Text.Encoding.GetEncoding(1251))
            '    Dim SR As New IO.StreamReader(Application.StartupPath & "\DO\recip.csv", System.Text.Encoding.GetEncoding(1251))
            '    total_strings = GetNumberOfLines(Application.StartupPath & "\DO\recip.csv")
            '    '''''''''''''''''''''''''''
            '    Dim tick As DateTime
            '    counter = 0
            '    CurrState = "Индексация RECIP DO "
            '    transact = sqlite_con.BeginTransaction()
            '    While SR.Peek >= 0
            '        tick = DateTime.Now
            '        Dim cols() As String = SR.ReadLine.Split(";")
            '        If cols.Length = 17 Then
            '            cols = cols.Take(cols.Count - 1).ToArray
            '        End If
            '        If counter Mod 1000000 = 0 Then
            '            transact.Commit()
            '            transact = sqlite_con.BeginTransaction()
            '        End If
            '        cmd.CommandText = "INSERT INTO RECIP_DO (ACCOUNT,CLOSE_CODE,DATE_FROM,DATE_TO,ID,MAN_ID,OPERATED,OPERATION,PAY_AREA_NUM,PAY_DAY,RA,RAZDEL,RECIPIENT_ID,RN,SV,WPR_ID) VALUES('" & String.Join("','", cols) & "')"
            '        cmd.ExecuteNonQuery()
            '        counter += 1
            '        If counter Mod 100000 = 0 Or counter = total_strings Or counter = 0 Then

            '            BW_IndexBases.ReportProgress(counter, total_strings)
            '        End If
            '    End While
            '    transact.Commit()
            '    '''''''''''''''''''''''''''
            'Catch ex As Exception
            '    MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))
            '    Exit Sub
            'End Try






            Try
                ' input_strings = IO.File.ReadAllLines(Application.StartupPath & "\DO\popen.csv", System.Text.Encoding.GetEncoding(1251))

                Dim SR As New IO.StreamReader(Application.StartupPath & "\DO\popen.csv", System.Text.Encoding.GetEncoding(1251))
                total_strings = GetNumberOfLines(Application.StartupPath & "\DO\popen.csv")
                counter = 0
                CurrState = "Индексация POPEN DO "
                transact = sqlite_con.BeginTransaction()
                While SR.Peek >= 0
                    Dim cols() As String = SR.ReadLine.Split(";")
                    If cols.Length = 17 Then
                        cols = cols.Take(cols.Count - 1).ToArray
                    End If
                    If cols.Length = 20 Then
                        cols = cols.Take(cols.Count - 1).ToArray
                    End If
                    cmd.CommandText = "INSERT INTO POPEN_DO (PO_ID,PO_RA,DPW,ID,NP,NVP,PW,SPOSOB,SROKPO,SROKS,STATUSR,TEQSROKPO,TEQSROKS,U_GO1,U_GO2,U_TR,V_GO1,V_GO2,V_TR) VALUES('" & String.Join("','", cols) & "')"
                    cmd.ExecuteNonQuery()
                    counter += 1
                    If counter Mod 100000 = 0 Or counter = total_strings Or counter = 0 Then

                        BW_IndexBases.ReportProgress(counter, total_strings)
                    End If
                End While

                transact.Commit()
            Catch ex As Exception

                MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))

                Exit Sub
            End Try


            Try
                Dim PE_MD5 As String = CreateMD5Sum(Application.StartupPath & "\DO\pe.csv")
                If PE_MD5 = GetDBParam("PE_MD5") Then
                    'MsgBox("SKIP")
                    ''INSERT CODE IF NEEDED
                Else

                    CleanUP(CleanType.ONLY_PE)
                    SetDBParam("PE_MD5", PE_MD5)
                    '  input_strings = IO.File.ReadAllLines(Application.StartupPath & "\DO\pe.csv", System.Text.Encoding.GetEncoding(1251))
                    Dim SR As New IO.StreamReader(Application.StartupPath & "\DO\pe.csv", System.Text.Encoding.GetEncoding(1251))
                    total_strings = GetNumberOfLines(Application.StartupPath & "\DO\pe.csv")
                    counter = 0
                    CurrState = "Индексация PE "
                    transact = sqlite_con.BeginTransaction()
                    While SR.Peek >= 0
                        Dim cols() As String = SR.ReadLine.Split(";")
                        If cols.Length = 8 Then
                            cols = cols.Take(cols.Count - 1).ToArray
                        End If
                        cmd.CommandText = "INSERT INTO PE (MAN_ID,RA,DAT,DATNP_TR,DIVP,OTKAZ,PGP) VALUES('" & String.Join("','", cols) & "')"
                        cmd.ExecuteNonQuery()
                        counter += 1
                        If counter Mod 100000 = 0 Or counter = total_strings Or counter = 0 Then

                            BW_IndexBases.ReportProgress(counter, total_strings)
                        End If
                    End While
                    transact.Commit()
                End If
            Catch ex As Exception
                MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))
                Exit Sub
            End Try


            'Try
            '    If IO.File.Exists(Application.StartupPath & "\DO\wpr.csv") = False Then
            '        Exit Try
            '    End If
            '    Dim WPR_MD5 As String = CreateMD5Sum(Application.StartupPath & "\DO\wpr.csv")
            '    If WPR_MD5 = My.Settings.WPR_md5 Then
            '        'MsgBox("SKIP")
            '        ''INSERT CODE IF NEEDED
            '    Else
            '        My.Settings.WPR_md5 = WPR_MD5
            '        CleanUP(CleanType.ONLY_WPR)

            '        Dim SR As New IO.StreamReader(Application.StartupPath & "\DO\wpr.csv", System.Text.Encoding.GetEncoding(1251))
            '        total_strings = GetNumberOfLines(Application.StartupPath & "\DO\wpr.csv")
            '        counter = 0
            '        CurrState = "Индексация WPR "
            '        transact = sqlite_con.BeginTransaction()
            '        While SR.Peek >= 0
            '            Dim cols() As String = SR.ReadLine.Split(";")
            '            If cols.Length = 5 Then
            '                cols = cols.Take(cols.Count - 1).ToArray
            '            ElseIf cols.Length > 5 Then
            '                Continue While
            '            End If
            '            cols(1) = cols(1).Replace("'", "")
            '            cmd.CommandText = "INSERT INTO WPR (ID,BANK,DOP_ID,RA) VALUES('" & String.Join("','", cols) & "')"
            '            cmd.ExecuteNonQuery()
            '            counter += 1
            '            If counter Mod 1000 = 0 Or counter = total_strings Or counter = 0 Then

            '                BW_IndexBases.ReportProgress(counter, total_strings)
            '            End If
            '        End While
            '        transact.Commit()
            '    End If
            'Catch ex As Exception
            '    MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))
            '    Exit Sub
            'End Try





            SetDBStatus(1)

        End If
        If MODE_PO = True Then




            '    Try
            '    '   input_strings = IO.File.ReadAllLines(Application.StartupPath & "\PO\recip.csv", System.Text.Encoding.GetEncoding(1251))
            '    Dim SR As New IO.StreamReader(Application.StartupPath & "\PO\recip.csv", System.Text.Encoding.GetEncoding(1251))
            '    total_strings = GetNumberOfLines(Application.StartupPath & "\PO\recip.csv")
            '    counter = 0
            '    CurrState = "Индексация RECIP PO "
            '    transact = sqlite_con.BeginTransaction()
            '    While SR.Peek >= 0

            '        Dim cols() As String = SR.ReadLine.Split(";")
            '        If cols.Length = 17 Then
            '            cols = cols.Take(cols.Count - 1).ToArray
            '        End If

            '        cmd.CommandText = "INSERT INTO RECIP_PO (ACCOUNT,CLOSE_CODE,DATE_FROM,DATE_TO,ID,MAN_ID,OPERATED,OPERATION,PAY_AREA_NUM,PAY_DAY,RA,RAZDEL,RECIPIENT_ID,RN,SV,WPR_ID) VALUES('" & String.Join("','", cols) & "')"
            '        cmd.ExecuteNonQuery()
            '        counter += 1
            '        If counter Mod 1000000 = 0 Then
            '            transact.Commit()
            '            transact = sqlite_con.BeginTransaction()
            '        End If
            '        If counter Mod 100000 = 0 Or counter = total_strings Or counter = 0 Then

            '            BW_IndexBases.ReportProgress(counter, total_strings)
            '        End If
            '    End While
            '    transact.Commit()

            'Catch ex As Exception
            '    MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))
            '    Exit Sub
            'End Try





            Try
                '  input_strings = IO.File.ReadAllLines(Application.StartupPath & "\PO\pay.csv", System.Text.Encoding.GetEncoding(1251))
                Dim SR As New IO.StreamReader(Application.StartupPath & "\PO\pay.csv", System.Text.Encoding.GetEncoding(1251))
                total_strings = GetNumberOfLines(Application.StartupPath & "\PO\pay.csv")

                counter = 0
                CurrState = "Индексация PAY PO "
                transact = sqlite_con.BeginTransaction()
                While SR.Peek >= 0
                    Dim cols() As String = SR.ReadLine.Split(";")
                    If cols.Length = 18 Then
                        cols = cols.Take(cols.Count - 1).ToArray
                    End If
                    cmd.CommandText = "INSERT INTO PAY_PO (ACTION_TYPE,AMOUNT,FILE_ID,ID,INFO_TYPE,IST,MAN_ID,MONTH,PARENT_ID,RA,RAZDEL,RE,RECIPIENT_ID,SOURCE_ID,SUBVED_ID,VERSION,YEAR) VALUES('" & String.Join("','", cols) & "')"
                    cmd.ExecuteNonQuery()
                    counter += 1
                    If counter Mod 1000000 = 0 Then
                        transact.Commit()
                        transact = sqlite_con.BeginTransaction()
                    End If
                    If counter Mod 100000 = 0 Or counter = total_strings Or counter = 0 Then

                        BW_IndexBases.ReportProgress(counter, total_strings)
                    End If
                End While
                transact.Commit()
                cmd.CommandText = "SELECT RA FROM PAY_DO LIMIT 1"
                DB_RA = cmd.ExecuteScalar
            Catch ex As Exception
                MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))
                Exit Sub
            End Try




            Try
                'input_strings = IO.File.ReadAllLines(Application.StartupPath & "\PO\popen.csv", System.Text.Encoding.GetEncoding(1251))
                Dim SR As New IO.StreamReader(Application.StartupPath & "\PO\popen.csv", System.Text.Encoding.GetEncoding(1251))
                total_strings = GetNumberOfLines(Application.StartupPath & "\PO\popen.csv")
                CurrState = "Индексация POPEN PO "
                counter = 0
                transact = sqlite_con.BeginTransaction()
                While SR.Peek >= 0
                    Dim cols() As String = SR.ReadLine.Split(";")
                    If cols.Length = 20 Then
                        cols = cols.Take(cols.Count - 1).ToArray
                    End If
                    cmd.CommandText = "INSERT INTO POPEN_PO (PO_ID,PO_RA,DPW,ID,NP,NVP,PW,SPOSOB,SROKPO,SROKS,STATUSR,TEQSROKPO,TEQSROKS,U_GO1,U_GO2,U_TR,V_GO1,V_GO2,V_TR) VALUES('" & String.Join("','", cols) & "')"
                    cmd.ExecuteNonQuery()
                    counter += 1
                    If counter Mod 100000 = 0 Or counter = total_strings Or counter = 0 Then

                        BW_IndexBases.ReportProgress(counter, total_strings)
                    End If
                End While
                transact.Commit()
            Catch ex As Exception
                MsgBox(String.Join(Environment.NewLine, "Обнаружена ошибка!", ex.Message, cmd.CommandText))
                Exit Sub
            End Try


        End If
        ' Erase popen_po_strings
        cmd.CommandText = " PRAGMA optimize"
        cmd.ExecuteNonQuery()
        sqlite_con.Close()
        transact.Dispose()
        sqlite_con.Dispose()
    End Sub

    Private Sub BW_IndexBases_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BW_IndexBases.ProgressChanged
        If e.ProgressPercentage Mod 10000 = 0 Or e.ProgressPercentage = e.UserState Then
            LBL_STATE.Text = CurrState & e.ProgressPercentage & "/" & e.UserState
            PB_PROGRESS.Maximum = e.UserState
            PB_PROGRESS.Value = e.ProgressPercentage
        End If

    End Sub
    Structure ERR_REPORT
        Public Property MAN_ID As String
        Public Property ERRTEXT As String
        Public Function ToHTMLTableRow()
            Dim stmp As String
            Dim inf As MAN_INFO = MAIN.GetManInfo(MAN_ID)
            Dim RDATE As Date


            If inf.NAZN = "" Then
                ERRTEXT &= Environment.NewLine & "НЕ УДАЛОСЬ ОПРЕДЕЛИТЬ ДАТУ НАЗНАЧЕНИЯ!" & Environment.NewLine

                stmp = "<tr style=""background-color: white""><td><pre>" & String.Join(" ", inf.FA, inf.IM, inf.OT) & Environment.NewLine & "(<b>" & inf.NPERS & "</b>) Д. Назн. - <b>НЕ ОПРЕДЕЛЕНА</b>" & Environment.NewLine & String.Format("</b>Дата рождения - <b>{0}</b> ({1} лет)", inf.RDAT.ToString().Replace(" 0:00:00", ""), MAIN.GetCurrentAge(inf.RDAT)) & "</pre></td><td><pre>" & ERRTEXT & "</pre></td></tr>"

            Else
                Dim DateNazn As Date = inf.NAZN
                stmp = "<tr style=""background-color: white""><td><pre>" & String.Join(" ", inf.FA, inf.IM, inf.OT) & Environment.NewLine & "(<b>" & inf.NPERS & "</b>) Д. Назн. - <b>" & DateNazn.ToString().Replace(" 0:00:00", "") & Environment.NewLine & String.Format("</b>Дата рождения - <b>{0}</b> ({1} лет)", inf.RDAT.ToString().Replace(" 0:00:00", ""), MAIN.GetCurrentAge(inf.RDAT)) & "</pre></td><td><pre>" & ERRTEXT & "</pre></td></tr>"

            End If

            Return stmp
        End Function
        Public Function ToCSVTableRow()
            Dim stmp As String = ""
            Dim infostring As String = ""
            Dim inf As MAN_INFO = MAIN.GetManInfo(MAN_ID)
            Dim winf As WORK_STATUS = MAIN.GetWorkState(MAN_ID)
            Dim NazDate As String = ""
            Dim count As Integer = 0
            If inf.NAZN = "" Then
                NazDate = "Н/Д"
            Else
                Dim DateNazn As Date = inf.NAZN
                NazDate = DateNazn.ToString().Replace(" 0:00:00", "")

            End If
            stmp = ""
            For Each entry In ERRTEXT.Split(Environment.NewLine)
                entry = entry.Replace(vbCr, "").Replace(vbLf, "")
                If count = 0 Then
                    infostring = String.Join(";", String.Format("{0} {1} {2}", inf.FA, inf.IM, inf.OT), inf.NPERS, NazDate, inf.RDAT, MAIN.GetCurrentAge(inf.RDAT), winf.SPOSOB)
                Else
                    infostring = ";;;;;"
                End If
                count += 1
                stmp &= String.Join(";", infostring, entry) & Environment.NewLine

            Next



            Return stmp
        End Function
    End Structure
    Dim Reports As New List(Of ERR_REPORT)
    Structure SUMM_LIST
        Public Property MAN_ID As String
        Public Property AMOUNT As Decimal
        Public Property PRIZNAK As Short
        Public Property ACTION_TYPE
    End Structure
    Structure MAN_INFO
        Public Property FA As String
        Public Property IM As String
        Public Property OT As String
        Public Property RA As String
        Public Property NPERS As String
        Public Property RDAT As String
        Public Property NAZN As String
        Public Property ADR_INDEX As String
    End Structure
    Structure INDEX4_STRING
        Public Property MAN_ID As String
        Public Property RA As String
        Public Property ADR_INDEX As String
        Public Property FA As String
        Public Property IM As String
        Public Property OT As String
        Public Property NPERS As String
        Public Property RDAT As String
        Public Property NAZN As String
        Public Property FW_PENS_DO As Decimal
        Public Property FW_PENS_PO As Decimal
        Public Property SC_PENS_DO As Decimal
        Public Property SC_PENS_PO As Decimal
        Public Property FW_DOPL_DO As Decimal
        Public Property FW_DOPL_PO As Decimal
        Public Property SC_DOPL_DO As Decimal
        Public Property SC_DOPL_PO As Decimal
        Public Property WORK_DO As String
        Public Property WORK_PO As String
        Public Property SPOSOB As String
        Public Function CSV_PREPARE()
            Dim resp, resp2 As String
            resp = String.Join(";", RA, ADR_INDEX, FA, IM, OT, NPERS, RDAT, NAZN, FW_PENS_DO, FW_PENS_PO, FW_PENS_PO - FW_PENS_DO, FW_DOPL_DO, FW_DOPL_PO, FW_DOPL_PO - FW_DOPL_DO, WORK_DO, WORK_PO, SPOSOB, "ФВ")
            resp2 = String.Join(";", RA, ADR_INDEX, FA, IM, OT, NPERS, RDAT, NAZN, SC_PENS_DO, SC_PENS_PO, SC_PENS_PO - SC_PENS_DO, SC_DOPL_DO, SC_DOPL_PO, SC_DOPL_PO - SC_DOPL_DO, WORK_DO, WORK_PO, SPOSOB, "СЧ")
            Return String.Join(Environment.NewLine, resp, resp2)
        End Function
    End Structure
    Structure WORK_STATUS
        Public Property WORK_DO As String
        Public Property WORK_PO As String
        Public Property SPOSOB As String
    End Structure
    Dim Summ_do As New List(Of SUMM_LIST)
    Dim Summ_po As New List(Of SUMM_LIST)

    Private Function GetNumberOfLines(path As String) As Integer
        Dim count As Integer
        Dim SR As New IO.StreamReader(path)
        While SR.ReadLine <> Nothing
            count += 1
        End While
        Return count
    End Function

    Private Function GetCurrentAge(ByVal dob As Date) As Integer ''thx StackOverflow
        Dim age As Integer
        age = Today.Year - dob.Year
        If (dob > Today.AddYears(-age)) Then age -= 1
        Return age
    End Function

    Private Function GetManInfo(man_id As String) As MAN_INFO
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "SELECT * FROM MAN WHERE ID='" & man_id & "'"
        Dim dbreader As SQLite.SQLiteDataReader = cmd.ExecuteReader
        Dim Result As New MAN_INFO
        While dbreader.Read
            With Result
                .ADR_INDEX = dbreader(0)
                .FA = dbreader(1)
                .IM = dbreader(3)
                .NPERS = dbreader(4)
                .OT = dbreader(5)
                .RA = dbreader(6)
                .RDAT = dbreader(7)
            End With

        End While
        dbreader.Close()
        cmd.CommandText = "SELECT CASE WHEN MAX(DATNP_TR)>MAX(DIVP) THEN MAX(DATNP_TR) ELSE MAX(DIVP) END NAZ FROM PE WHERE MAN_ID='" & man_id & "'"
        dbreader = cmd.ExecuteReader
        While dbreader.Read
            If dbreader.IsDBNull(0) = True Then
                With Result
                    .NAZN = ""

                End With
            Else
                With Result
                    .NAZN = dbreader(0)
                End With
            End If

        End While
        Return Result
    End Function
    Private Function GetWorkState(man_id As String)
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "SELECT POPEN_DO.STATUSR, POPEN_PO.STATUSR, SPOSOB.NAME FROM popen_do LEFT JOIN POPEN_PO ON POPEN_DO.ID = POPEN_PO.ID LEFT JOIN SPOSOB ON POPEN_PO.SPOSOB = SPOSOB.CODE WHERE POPEN_DO.ID='" & man_id & "'"
        Dim dbreader As SQLite.SQLiteDataReader = cmd.ExecuteReader
        Dim Result As New WORK_STATUS
        While dbreader.Read
            With Result
                If dbreader.IsDBNull(0) = True Then
                    .WORK_DO = ""
                Else
                    .WORK_DO = dbreader(0)
                End If
                If dbreader.IsDBNull(1) = True Then
                    .WORK_PO = ""
                Else
                    .WORK_PO = dbreader(1)
                End If
                If dbreader.IsDBNull(2) = True Then
                    .SPOSOB = ""
                Else
                    .SPOSOB = dbreader(2)
                End If

            End With

        End While
        Return Result
    End Function



    Private Sub BW_OPERATION_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_OPERATION.DoWork
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        Dim dbreader As SQLite.SQLiteDataReader
        Dim count As Integer = 0
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET DO
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "PRAGMA OPTIMIZE"
        cmd.ExecuteNonQuery()
        CurrState = "Подготовка базы к анализу:"
        BW_OPERATION.ReportProgress(count)
        ' MsgBox(count)

        cmd.CommandText = "INSERT INTO IND4 (MAN_ID) SELECT ID FROM MAN ORDER BY ID;"
        cmd.ExecuteNonQuery()
        ''10 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(10)
        cmd.CommandText = "UPDATE IND4 SET FW_DO = (SELECT PAY_DO.AMOUNT FROM PAY_DO LEFT OUTER JOIN PAYTYPE ON PAY_DO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 1 AND PAY_DO.ACTION_TYPE = '4' AND PAY_DO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''20 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(20)
        cmd.CommandText = "UPDATE IND4 SET SC_DO = (SELECT PAY_DO.AMOUNT FROM PAY_DO LEFT OUTER JOIN PAYTYPE ON PAY_DO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 2 AND PAY_DO.ACTION_TYPE = '4' AND PAY_DO.MAN_ID = IND4.MAN_ID); "
        cmd.ExecuteNonQuery()
        ''30 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(30)
        cmd.CommandText = "UPDATE IND4 SET SC_DOPL_DO = (SELECT PAY_DO.AMOUNT FROM PAY_DO LEFT OUTER JOIN PAYTYPE ON PAY_DO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 2 AND PAY_DO.ACTION_TYPE = '5' AND PAY_DO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''40 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(40)
        cmd.CommandText = "UPDATE IND4 SET FW_DOPL_DO = (SELECT PAY_DO.AMOUNT FROM PAY_DO LEFT OUTER JOIN PAYTYPE ON PAY_DO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 1 AND PAY_DO.ACTION_TYPE = '5' AND PAY_DO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''50 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(50)
        cmd.CommandText = "UPDATE IND4 SET FW_PO = (SELECT PAY_PO.AMOUNT FROM PAY_PO LEFT OUTER JOIN PAYTYPE ON PAY_PO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 1 AND PAY_PO.ACTION_TYPE = '4' AND PAY_PO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''60 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(60)
        cmd.CommandText = "UPDATE IND4 SET SC_PO = (SELECT PAY_PO.AMOUNT FROM PAY_PO LEFT OUTER JOIN PAYTYPE ON PAY_PO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 2 AND PAY_PO.ACTION_TYPE = '4' AND PAY_PO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''70 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(70)
        cmd.CommandText = "UPDATE IND4 SET SC_DOPL_PO = (SELECT PAY_PO.AMOUNT FROM PAY_PO LEFT OUTER JOIN PAYTYPE ON PAY_PO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 2 AND PAY_PO.ACTION_TYPE = '5' AND PAY_PO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''80 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(80)
        cmd.CommandText = "UPDATE IND4 SET FW_DOPL_PO = (SELECT PAY_PO.AMOUNT FROM PAY_PO LEFT OUTER JOIN PAYTYPE ON PAY_PO.IST = PAYTYPE.IST WHERE PAYTYPE.PRIZNAK = 1 AND PAY_PO.ACTION_TYPE = '5' AND PAY_PO.MAN_ID = IND4.MAN_ID);"
        cmd.ExecuteNonQuery()
        ''90 процентов
        CurrState = "Подготовка базы к анализу. Процент готовности:"
        BW_OPERATION.ReportProgress(90)

        cmd.CommandText = "SELECT * from ind4 where abs(FW_DOPL_DO - FW_DOPL_PO) >0.1 OR ABS(FW_DO-FW_PO)>0.1 OR ABS(SC_DO-SC_PO)>0.1 or ABS(SC_DOPL_DO-SC_DOPL_PO)>0.1;"
        '  Dim MEN As New List(Of String)

        dbreader = cmd.ExecuteReader
        CurrState = "Обнажуено изменений:"
        Dim stmp As String = "РАЙОН;ИНДЕКС;ФАМИЛИЯ;ИМЯ;ОТЧЕСТВО;СНИЛС;ДАТА РОЖДЕНИЯ;ДАТА НАЗ.;ПЕНСИЯ ДО;ПЕНСИЯ ПОСЛЕ;ИЗМ. ПЕНСИИ;ДОПЛАТА ДО;ДОПЛАТА ПОСЛЕ;ИЗМ. ДОПЛАТЫ;РАБОТА ДО;РАБОТА ПОСЛЕ;СПОСОБ ПОЛУЧЕНИЯ;ВИД ВЫПЛ." & Environment.NewLine
        ' Dim INDX4 As New List(Of INDEX4_STRING)
        While dbreader.Read


            Dim ind4 As New INDEX4_STRING
            ind4.MAN_ID = dbreader(0)
            ind4.FW_PENS_DO = dbreader(1)
            ind4.FW_PENS_PO = dbreader(2)
            ind4.SC_PENS_DO = dbreader(3)
            ind4.SC_PENS_PO = dbreader(4)
            ind4.FW_DOPL_DO = dbreader(5)
            ind4.FW_DOPL_PO = dbreader(6)
            ind4.SC_DOPL_DO = dbreader(7)
            ind4.SC_DOPL_PO = dbreader(8)
            Dim man_inf As MAN_INFO = GetManInfo(dbreader(0))
            ind4.ADR_INDEX = man_inf.ADR_INDEX
            ind4.RA = man_inf.RA
            ind4.FA = man_inf.FA
            ind4.IM = man_inf.IM
            ind4.OT = man_inf.OT
            ind4.NPERS = man_inf.NPERS

            ind4.NAZN =  man_inf.NAZN.ToString().Replace(" 0:00:00", "")
            ind4.RDAT = man_inf.RDAT
            Dim wrk_inf As WORK_STATUS = GetWorkState(dbreader(0))
            ind4.WORK_DO = wrk_inf.WORK_DO
            ind4.WORK_PO = wrk_inf.WORK_PO
            ind4.SPOSOB = wrk_inf.SPOSOB
            Validate_Summ(ind4)
            stmp &= ind4.CSV_PREPARE & Environment.NewLine
            ''update state
            count += 1
            BW_OPERATION.ReportProgress(count)

        End While
        dbreader.Close()
        ''validate workstat
        cmd.CommandText = "SELECT * FROM IND4 WHERE ((((ABS(FW_DO -FW_PO) < 0.1 AND FW_DO <> 0) and FW_DOPL_DO = FW_DOPL_PO) OR ((ABS(SC_DO - SC_PO)<0.1 AND SC_DO <> 0)AND SC_DOPL_DO=SC_DOPL_PO))) AND IND4.MAN_ID IN (SELECT POPEN_PO.ID FROM popen_do LEFT JOIN POPEN_PO ON POPEN_DO.ID = POPEN_PO.ID WHERE POPEN_DO.STATUSR ='Д' AND POPEN_PO.STATUSR='Н');"

        dbreader = cmd.ExecuteReader
        CurrState = "Завершение проверки: "
        While dbreader.Read


            Dim ind4 As New INDEX4_STRING
            ind4.MAN_ID = dbreader(0)
            ind4.FW_PENS_DO = dbreader(1)
            ind4.FW_PENS_PO = dbreader(2)
            ind4.SC_PENS_DO = dbreader(3)
            ind4.SC_PENS_PO = dbreader(4)
            ind4.FW_DOPL_DO = dbreader(5)
            ind4.FW_DOPL_PO = dbreader(6)
            ind4.SC_DOPL_DO = dbreader(7)
            ind4.SC_DOPL_PO = dbreader(8)
            Dim man_inf As MAN_INFO = GetManInfo(dbreader(0))
            ind4.ADR_INDEX = man_inf.ADR_INDEX
            ind4.RA = man_inf.RA
            ind4.FA = man_inf.FA
            ind4.IM = man_inf.IM
            ind4.OT = man_inf.OT
            ind4.NPERS = man_inf.NPERS

            ind4.NAZN = man_inf.NAZN.ToString().Replace(" 0:00:00", "")
            ind4.RDAT = man_inf.RDAT
            Dim wrk_inf As WORK_STATUS = GetWorkState(dbreader(0))
            ind4.WORK_DO = wrk_inf.WORK_DO
            ind4.WORK_PO = wrk_inf.WORK_PO
            ind4.SPOSOB = wrk_inf.SPOSOB
            Validate_Summ(ind4)
            stmp &= ind4.CSV_PREPARE & Environment.NewLine
            ''update state
            count += 1
            BW_OPERATION.ReportProgress(count)

        End While

        dbreader.Close()
        ''validate workstat 2
        cmd.CommandText = "SELECT * FROM IND4 WHERE ((((ABS(FW_DO -FW_PO) < 0.1 AND FW_DO <> 0) and FW_DOPL_DO <> FW_DOPL_PO) OR ((ABS(SC_DO - SC_PO)<0.1 AND SC_DO <> 0)AND SC_DOPL_DO<>SC_DOPL_PO))) AND IND4.MAN_ID IN (SELECT POPEN_PO.ID FROM popen_do LEFT JOIN POPEN_PO ON POPEN_DO.ID = POPEN_PO.ID WHERE POPEN_DO.STATUSR ='Н' AND POPEN_PO.STATUSR='Д')"

        dbreader = cmd.ExecuteReader
        CurrState = "Завершение проверки 2: "
        While dbreader.Read


            Dim ind4 As New INDEX4_STRING
            ind4.MAN_ID = dbreader(0)
            ind4.FW_PENS_DO = dbreader(1)
            ind4.FW_PENS_PO = dbreader(2)
            ind4.SC_PENS_DO = dbreader(3)
            ind4.SC_PENS_PO = dbreader(4)
            ind4.FW_DOPL_DO = dbreader(5)
            ind4.FW_DOPL_PO = dbreader(6)
            ind4.SC_DOPL_DO = dbreader(7)
            ind4.SC_DOPL_PO = dbreader(8)
            Dim man_inf As MAN_INFO = GetManInfo(dbreader(0))
            ind4.ADR_INDEX = man_inf.ADR_INDEX
            ind4.RA = man_inf.RA
            ind4.FA = man_inf.FA
            ind4.IM = man_inf.IM
            ind4.OT = man_inf.OT
            ind4.NPERS = man_inf.NPERS
            Dim DateNazn As Date = man_inf.NAZN
            ind4.NAZN = DateNazn.ToString().Replace(" 0:00:00", "")
            ind4.RDAT = man_inf.RDAT
            Dim wrk_inf As WORK_STATUS = GetWorkState(dbreader(0))
            ind4.WORK_DO = wrk_inf.WORK_DO
            ind4.WORK_PO = wrk_inf.WORK_PO
            ind4.SPOSOB = wrk_inf.SPOSOB
            Validate_Summ(ind4)
            stmp &= ind4.CSV_PREPARE & Environment.NewLine
            ''update state
            count += 1
            BW_OPERATION.ReportProgress(count)

        End While



        ValidatePensInfo()

        BW_OPERATION.ReportProgress(-1)
        Dim FileToSave As String = Application.StartupPath & "\INDEX_4(" & DB_RA & ")-" & String.Join(".", Date.Today.Day.ToString(), Date.Today.Month.ToString(), Date.Today.Year.ToString) & " " & String.Join("-", DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString()) & ".CSV"
        IO.File.WriteAllText(FileToSave, stmp)
        ' If MsgBox("Открыть список пенсионеров, у которых изменилась пенсия?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then Process.Start(FileToSave)
        Try
            CurrState = "Формирование отчета "
            BW_OPERATION.ReportProgress(count)
            GenerateCSVReport()
            ' GenerateHTMLReport()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        dbreader.Close()

        sqlite_con.Close()
        sqlite_con.Dispose()
        CleanUP(CleanType.ONLY_DO_PO)

    End Sub
    Private Function GetSumms(man_id As String, type As Byte) As List(Of SUMM_LIST)
        Dim prefix As String = "SUMM_PO"
        Dim kmnd As String
        Select Case type
            Case 1 'DO
                prefix = "SUMM_DO"
                kmnd = "SELECT PAY_DO.MAN_ID, PAY_DO.AMOUNT, paytype.PRIZNAK, PAY_DO.ACTION_TYPE, PAY_DO.ID FROM PAY_DO LEFT OUTER JOIN PAYTYPE ON PAY_DO.IST = PAYTYPE.IST WHERE ((PAYTYPE.PRIZNAK = 1 OR PAYTYPE.PRIZNAK = 2) AND PAY_DO.ACTION_TYPE = '4' OR ((PAYTYPE.PRIZNAK = 1 OR PAYTYPE.PRIZNAK = 2) AND PAY_DO.ACTION_TYPE = '5')) AND PAY_DO.MAN_ID = '" & man_id & "'"
            Case 2 'po
                prefix = "SUMM_PO"
                kmnd = "SELECT PAY_PO.MAN_ID, PAY_PO.AMOUNT, paytype.PRIZNAK, PAY_PO.ACTION_TYPE, PAY_PO.ID FROM PAY_PO LEFT OUTER JOIN PAYTYPE ON PAY_PO.IST = PAYTYPE.IST WHERE ((PAYTYPE.PRIZNAK = 1 OR PAYTYPE.PRIZNAK = 2) AND PAY_PO.ACTION_TYPE = '4' OR ((PAYTYPE.PRIZNAK = 1 OR PAYTYPE.PRIZNAK = 2) AND PAY_PO.ACTION_TYPE = '5')) AND PAY_PO.MAN_ID = '" & man_id & "'"

        End Select
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        Dim Summs As New List(Of SUMM_LIST)
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = kmnd
        'cmd.CommandText = "SELECT * FROM " & prefix & " WHERE MAN_ID = '" & man_id & "'"
        Dim dbreader As SQLite.SQLiteDataReader = cmd.ExecuteReader
        While dbreader.Read()
            Dim summ As New SUMM_LIST
            With summ
                .MAN_ID = dbreader(0)
                .AMOUNT = dbreader(1)
                .PRIZNAK = dbreader(2)
                .ACTION_TYPE = dbreader(3)
            End With
            Summs.Add(summ)


        End While
        Return Summs
    End Function
    Private Function CreateMD5Sum(ByVal filename As String) As String ''THANKS TO VBFORUMS
        Using md5 As MD5 = MD5.Create()

            Using stream = IO.File.OpenRead(filename)
                Return BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", String.Empty)
            End Using
        End Using
    End Function
    Private Sub Validate_Summ(ind4 As INDEX4_STRING)
        Dim Err_record As New ERR_REPORT
        Err_record.MAN_ID = ind4.MAN_ID
        Err_record.ERRTEXT = ""
        If ind4.NAZN = "" Then
            Err_record.ERRTEXT &= "НЕТ ДАТЫ НАЗНАЧЕНИЯ " & Environment.NewLine
            Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_DO, ind4.FW_DOPL_DO, ind4.FW_PENS_DO + ind4.FW_DOPL_DO) & Environment.NewLine
            Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_PO, ind4.FW_DOPL_PO, ind4.FW_PENS_PO + ind4.FW_DOPL_PO) & Environment.NewLine
            Err_record.ERRTEXT &= String.Format("Размер CЧ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_DO, ind4.SC_DOPL_DO, ind4.SC_PENS_DO + ind4.SC_DOPL_DO) & Environment.NewLine
            Err_record.ERRTEXT &= String.Format("Размер СЧ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_PO, ind4.SC_DOPL_PO, ind4.SC_PENS_PO + ind4.SC_DOPL_PO) & Environment.NewLine
            Err_record.ERRTEXT &= "Работал ДО: " & ind4.WORK_DO & " ПОСЛЕ: " & ind4.WORK_PO & ""
            Reports.Add(Err_record)
            Exit Sub '' НЕТ ДАТЫ - НЕТ ОПРЕДЕЛЕНИЯ
        End If
        Dim NAZN_Date As Date = ind4.NAZN
        Dim Multipl_S As Decimal = 1.0
        Dim Multipl_F As Decimal = 1.0
        Dim IsCorrect As Boolean = False

        ''ASSIGN VALUES FOR INDICIES
        Dim sindex0, sindex162, sindex172, sindex174, sindex181, sindex191, sindex201, findex162, findex172, findex174, findex181, findex191, findex201 As Decimal
        sindex0 = 71.41
        sindex162 = 74.27
        sindex172 = 78.28
        sindex174 = 78.58
        sindex181 = 81.49
        sindex191 = 87.24
        sindex201 = 93.0
        findex162 = (41639392.0 / 32100231)
        findex172 = (36701269.0 / 29425117)
        findex174 = (22829043.0 / 19291475)
        findex181 = (22829043.0 / 19291475)
        findex191 = (9858699.0 / 8639251)
        findex201 = 1.066
        Dim SIndicies As New List(Of Decimal)
        Dim FIndicies As New List(Of Decimal)
        Dim FW_NACH_DO As Decimal = ind4.FW_DOPL_DO + ind4.FW_PENS_DO
        Dim FW_NACH_PO As Decimal = ind4.FW_DOPL_PO + ind4.FW_PENS_PO
        Dim SC_NACH_DO As Decimal = ind4.SC_DOPL_DO + ind4.SC_PENS_DO
        Dim SC_NACH_PO As Decimal = ind4.SC_DOPL_PO + ind4.SC_PENS_PO

        If NAZN_Date < "01.02.2016" Then
            SIndicies.Add(sindex0)

        End If
        If NAZN_Date <= "31.01.2017" Then
            SIndicies.Add(sindex162)
            FIndicies.Add(findex162)

        End If
        If NAZN_Date <= "31.03.2017" Then
            SIndicies.Add(sindex172)
            FIndicies.Add(findex172)

        End If
        If NAZN_Date <= "31.12.2017" Then
            SIndicies.Add(sindex174)
            FIndicies.Add(findex174)

        End If
        If NAZN_Date <= "31.12.2018" Then
            SIndicies.Add(sindex181)
            FIndicies.Add(findex181)

        End If
        If NAZN_Date <= "31.12.2019" Then
            SIndicies.Add(sindex191)
            FIndicies.Add(findex191)

        End If
        If NAZN_Date <= "31.12.2020" Then
            ' SIndicies.Add(sindex201)
            ' FIndicies.Add(findex201)
        End If
        If ind4.WORK_DO = "Д" And ind4.WORK_PO = "Н" Then ''ОЖИДАЕТСЯ УВЕЛИЧЕНИЕ СУММЫ
            ''Анализ размера доплаты
            Dim P_DELTA_FW As Decimal = Math.Abs(ind4.FW_PENS_DO - ind4.FW_PENS_PO)
            Dim P_DELTA_SC As Decimal = Math.Abs(ind4.SC_PENS_DO - ind4.SC_PENS_PO)
            Dim D_DELTA_FW As Decimal = Math.Abs(ind4.FW_DOPL_DO - ind4.FW_DOPL_PO)
            Dim D_DELTA_SC As Decimal = Math.Abs(ind4.SC_DOPL_DO - ind4.SC_DOPL_PO)
            Dim EXPECT_DOPL_FW As Decimal
            Dim EXPECT_DOPL_SC As Decimal

            Select Case Calc_Month
                Case 1
                    LBL_DBG.Text = 1
                    EXPECT_DOPL_FW = Decimal.Round(P_DELTA_FW / findex201 * 2 + P_DELTA_FW, 2, MidpointRounding.AwayFromZero)
                    EXPECT_DOPL_SC = Decimal.Round(P_DELTA_SC / sindex201 * sindex191 * 2 + P_DELTA_SC, 2, MidpointRounding.AwayFromZero)
                Case 2
                    LBL_DBG.Text = 2
                    EXPECT_DOPL_FW = Decimal.Round(P_DELTA_FW / findex201 + P_DELTA_FW * 2, 2, MidpointRounding.AwayFromZero)
                    EXPECT_DOPL_SC = Decimal.Round(P_DELTA_SC / sindex201 * sindex191 + P_DELTA_SC * 2, 2, MidpointRounding.AwayFromZero)
                Case >= 3
                    LBL_DBG.Text = "3 and more"
                    EXPECT_DOPL_FW = Decimal.Round(P_DELTA_FW * 3, 2, MidpointRounding.AwayFromZero)
                    EXPECT_DOPL_SC = Decimal.Round(P_DELTA_SC * 3, 2, MidpointRounding.AwayFromZero)

            End Select

            If (EXPECT_DOPL_FW < D_DELTA_FW And Math.Abs(EXPECT_DOPL_FW - D_DELTA_FW) > 20) Or (EXPECT_DOPL_SC < D_DELTA_SC And Math.Abs(EXPECT_DOPL_SC - D_DELTA_SC) > 20) Then
                Err_record.ERRTEXT &= "Проверьте доплату" & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Доплата ФВ: {0} Ожидаемая Сумма: {1}", D_DELTA_FW, EXPECT_DOPL_FW) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Доплата СЧ: {0} Ожидаемая Сумма: {1}", D_DELTA_SC, EXPECT_DOPL_SC) & Environment.NewLine


            End If

            If ((ind4.FW_PENS_DO >= ind4.FW_PENS_PO Or ind4.FW_DOPL_DO = ind4.FW_DOPL_PO) And (FW_NACH_DO <> 0 Or FW_NACH_PO <> 0)) Or ((ind4.SC_PENS_DO >= ind4.SC_PENS_PO Or ind4.SC_DOPL_DO = ind4.SC_DOPL_PO) And (SC_NACH_DO <> 0 Or SC_NACH_PO <> 0)) Then
                Err_record.ERRTEXT &= "Размер ФВ или СЧ пенсии уменьшился или не изменился, либо не изменилась доплата, когда ожидалось его увеличение" & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_DO, ind4.FW_DOPL_DO, FW_NACH_DO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_PO, ind4.FW_DOPL_PO, FW_NACH_PO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер CЧ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_DO, ind4.SC_DOPL_DO, SC_NACH_DO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер СЧ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_PO, ind4.SC_DOPL_PO, SC_NACH_PO) & Environment.NewLine
            Else
                For Each findex As Decimal In FIndicies

                    If FW_NACH_PO - FW_NACH_DO * findex > (FW_NACH_PO - FW_NACH_DO) Then

                    Else
                        IsCorrect = True
                    End If
                Next
                If IsCorrect = False Then
                    Err_record.ERRTEXT &= "Размер ФВ пенсии возможно некорректен." & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_DO, ind4.FW_DOPL_DO, FW_NACH_DO) & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_PO, ind4.FW_DOPL_PO, FW_NACH_PO) & Environment.NewLine
                Else
                    IsCorrect = False
                End If
                For Each sindex In SIndicies
                    If SC_NACH_PO - SC_NACH_DO / sindex * sindex201 > (SC_NACH_PO - SC_NACH_DO) Then

                    Else
                        IsCorrect = True
                    End If

                Next
                If IsCorrect = False Then
                    Err_record.ERRTEXT &= "Размер СЧ пенсии возможно некорректен." & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер CЧ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_DO, ind4.SC_DOPL_DO, SC_NACH_DO) & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер СЧ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_PO, ind4.SC_DOPL_PO, SC_NACH_PO) & Environment.NewLine
                Else
                    IsCorrect = False
                End If
            End If

        ElseIf ind4.WORK_DO = "Н" And ind4.WORK_PO = "Д" Then '' ожидается Уменьшение суммы
            If (ind4.FW_PENS_DO < ind4.FW_PENS_PO Or ind4.FW_DOPL_DO <> ind4.FW_DOPL_PO) Or (ind4.SC_PENS_DO < ind4.SC_PENS_PO Or ind4.SC_DOPL_DO <> ind4.SC_DOPL_PO) Then
                Err_record.ERRTEXT &= "Размер ФВ или СЧ пенсии увеличился или не изменился, либо изменилась выплата когда ожидалось его уменьшение" & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_DO, ind4.FW_DOPL_DO, FW_NACH_DO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_PO, ind4.FW_DOPL_PO, FW_NACH_PO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер CЧ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_DO, ind4.SC_DOPL_DO, SC_NACH_DO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер СЧ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_PO, ind4.SC_DOPL_PO, SC_NACH_PO) & Environment.NewLine
            Else

                For Each findex As Decimal In FIndicies

                    If FW_NACH_PO - FW_NACH_DO / findex > (FW_NACH_DO - FW_NACH_PO) Then


                    Else
                        IsCorrect = True

                    End If
                Next
                If IsCorrect = False Then

                    Err_record.ERRTEXT &= "Размер ФВ пенсии возможно некорректен." & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_DO, ind4.FW_DOPL_DO, FW_NACH_DO) & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_PO, ind4.FW_DOPL_PO, FW_NACH_PO) & Environment.NewLine
                Else
                    IsCorrect = False
                End If

                For Each sindex In SIndicies
                    If SC_NACH_PO - SC_NACH_DO / sindex201 * sindex > (SC_NACH_DO - SC_NACH_PO) Then

                    Else
                        IsCorrect = True
                    End If

                Next
                If IsCorrect = False Then

                    Err_record.ERRTEXT &= "Размер СЧ пенсии возможно некорректен." & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер CЧ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_DO, ind4.SC_DOPL_DO, SC_NACH_DO) & Environment.NewLine
                    Err_record.ERRTEXT &= String.Format("Размер СЧ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_PO, ind4.SC_DOPL_PO, SC_NACH_PO) & Environment.NewLine

                End If
                IsCorrect = False
            End If

        ElseIf ind4.WORK_DO = ind4.WORK_PO Then '' НИХУЯ НЕ ОЖИДАЕТСЯ 'НО МОЖЕТ БЫТЬ МЕНЬШЕ (ред. 17,02)
            If (FW_NACH_DO < FW_NACH_PO) Or (SC_NACH_DO < SC_NACH_PO) Then
                Err_record.ERRTEXT &= "Размер ФВ или СЧ пенсии увеличился , когда его изменения не ожидалось" & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_DO, ind4.FW_DOPL_DO, FW_NACH_DO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер ФВ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.FW_PENS_PO, ind4.FW_DOPL_PO, FW_NACH_PO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер CЧ пенсии ДО: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_DO, ind4.SC_DOPL_DO, SC_NACH_DO) & Environment.NewLine
                Err_record.ERRTEXT &= String.Format("Размер СЧ пенсии ПОСЛЕ: {0} + Доплата: {1} = ИТОГО: {2}", ind4.SC_PENS_PO, ind4.SC_DOPL_PO, SC_NACH_PO) & Environment.NewLine

            End If

        End If
        ''АНАЛИЗ ДОПЛАТЫ



        If Err_record.ERRTEXT <> "" Then
            Err_record.ERRTEXT &= "Работал ДО: " & ind4.WORK_DO & " ПОСЛЕ: " & ind4.WORK_PO & ""
            Reports.Add(Err_record)

        End If
    End Sub
    Private Sub GenerateHTMLReport()
        Dim Count As Integer
        Dim stmp As String = "<!DOCTYPE HTML><html><head><meta charset=""utf-8""><title>Результаты проверки</title></head><body><table border=""1""><caption style=""font-weight: bold; font-size: 15pt; "">Результаты проверки</caption><tr><th>Гражданин</th><th>Замечание</th></tr>"
        For Each entry In Reports
            stmp &= entry.ToHTMLTableRow
            Count += 1
        Next
        stmp &= "<tr><td colspan=2>ИТОГО: <b>" & Count & "</b> ЧЕЛОВЕК С ЗАМЕЧАНИЯМИ</td></tr></table></body></html>"
        Dim FileToSave As String = Application.StartupPath & "\SPISOK-RESULT(" & DB_RA & ")-" & String.Join(".", Date.Today.Day.ToString(), Date.Today.Month.ToString(), Date.Today.Year.ToString) & " " & String.Join("-", DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString()) & ".html"
        IO.File.WriteAllText(FileToSave, stmp)

        stmp = String.Format("Итого: {0} Записей", Count) & Environment.NewLine & "Открыть отчет о сравнении сумм?"
        If MsgBox(stmp, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Process.Start(FileToSave)
        End If

    End Sub

    Private Sub GenerateCSVReport()
        Dim stmp As String = "ФИО;СНИЛС;Д. Назн.;Д. Рожд.;Возраст;Способ доставки;Замечание" & Environment.NewLine
        Dim count As Integer
        For Each entry In Reports
            stmp &= entry.ToCSVTableRow
            count += 1
        Next

        Dim FileToSave As String = Application.StartupPath & "\SPISOK-RESULT(" & DB_RA & ")-" & String.Join(".", Date.Today.Day.ToString(), Date.Today.Month.ToString(), Date.Today.Year.ToString) & " " & String.Join("-", DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString()) & ".CSV"
        IO.File.WriteAllText(FileToSave, stmp, GetEncoding(1251))

        stmp = String.Format("Итого: {0} Записей", count) & Environment.NewLine & "Открыть CSV отчет о сравнении сумм?"
        If MsgBox(stmp, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Process.Start(FileToSave)
        End If


    End Sub

    Private Sub BW_OPERATION_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BW_OPERATION.ProgressChanged
        If e.ProgressPercentage = -1 Then
            PB_PROGRESS.Style = ProgressBarStyle.Blocks
            CurrState = "Операция завершена"
            LBL_STATE.Text = CurrState
            Exit Sub
        End If
        PB_PROGRESS.Style = ProgressBarStyle.Marquee
        LBL_STATE.Text = CurrState & " " & e.ProgressPercentage
    End Sub

    Private Sub BW_IndexBases_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_IndexBases.RunWorkerCompleted
        LBL_STATE.Text = StatusName(GetDBStatus)
        If MODE_PO = True Then

            BW_OPERATION.RunWorkerAsync()

            MODE_PO = False
        Else
            BTN_START.Enabled = True
            BTN_START_PO.Enabled = True
            BTN_SETTINGS.Enabled = True
            CB_MONTH.Enabled = True
        End If
        My.Settings.Save()
    End Sub

    Private Sub BTN_START_PO_Click(sender As Object, e As EventArgs) Handles BTN_START_PO.Click
        Calc_Month = CB_MONTH.SelectedIndex + 1
        MODE_PO = True
        BTN_START.Enabled = False
        BTN_START_PO.Enabled = False
        BTN_SETTINGS.Enabled = False
        CB_MONTH.Enabled = False
        BW_IndexBases.RunWorkerAsync()

    End Sub
    Private Sub ValidatePensInfo()

        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "SELECT * FROM POPEN_DO LEFT OUTER JOIN POPEN_PO ON POPEN_DO.ID=POPEN_PO.ID where (POPEN_DO.NP<>POPEN_PO.NP AND (POPEN_PO.NP = 'ПРИ' OR POPEN_PO.NP = 'ПРЕ' OR POPEN_PO.NP = 'СНЯ')) OR POPEN_DO.SROKPO <> POPEN_PO.SROKPO"
        Dim dbreader As SQLite.SQLiteDataReader = cmd.ExecuteReader
        Dim stmp As String = ""
        While dbreader.Read
            Try
                stmp = Environment.NewLine & "Изменились значимые данные: " & Environment.NewLine & String.Format("ДО: Операция <{0}> Срок по: {1}", dbreader(4), dbreader(8).ToString().Replace(" 0:00:00", "")) & Environment.NewLine & String.Format("ПО: Операция <{0}> Срок по: {1}", dbreader(23), dbreader(27).ToString().Replace(" 0:00:00", ""))

            Catch ex As Exception
                stmp = Environment.NewLine & "Изменились значимые данные: " & Environment.NewLine & String.Format("ДО: Операция <{0}> Срок по: -", dbreader(4)) & Environment.NewLine & String.Format("ПО: Операция <{0}> Срок по: -", dbreader(23))

            End Try

            If Reports.Exists(Function(x) x.MAN_ID = dbreader(0)) = True Then
                Dim report As ERR_REPORT = Reports.Find(Function(x) x.MAN_ID = dbreader(0))
                Reports.RemoveAt(Reports.FindIndex(Function(x) x.MAN_ID = dbreader(0)))

                report.ERRTEXT &= stmp
                Reports.Add(report)
            Else
                Dim report As New ERR_REPORT With {.MAN_ID = dbreader(0), .ERRTEXT = stmp}
                Reports.Add(report)
            End If
        End While

    End Sub

    Private Sub BW_OPERATION_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BW_OPERATION.RunWorkerCompleted
        BTN_START.Enabled = True
        BTN_START_PO.Enabled = True
        BTN_SETTINGS.Enabled = True
        CB_MONTH.Enabled = True
        LBL_STATE.Text = "Завершено. Файл с отчетом находится в папке с программой."
        PB_PROGRESS.Style = ProgressBarStyle.Blocks
    End Sub
    Public Function GetDBStatus() As Integer
        Dim response As Integer
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "SELECT PARAM_VALUE FROM DB_SETTINGS WHERE PARAM_NAME = 'DB_STATE'"
        response = cmd.ExecuteScalar
        Return response
    End Function
    Public Sub SetDBStatus(status As String)

        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "UPDATE DB_SETTINGS SET PARAM_VALUE ='" & status & "' WHERE PARAM_NAME = 'DB_STATE'"
        cmd.ExecuteNonQuery()
    End Sub
    Public Function GetDBParam(ParamName As String) As String
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "Select PARAM_VALUE from DB_SETTINGS where PARAM_NAME ='" & ParamName & "'"
        Return cmd.ExecuteScalar()


    End Function
    Public Sub SetDBParam(ParamName As String, ParamValue As String)
        '   MsgBox(String.Format("PARAM UPDTATE {0} : {1}", ParamName, ParamValue))
        Dim sqlite_con As New SQLite.SQLiteConnection()
        Dim cmd As SQLite.SQLiteCommand
        sqlite_con.ConnectionString = "Data Source=" & Application.StartupPath & "\database.sqlite;"
        sqlite_con.Open()
        ''GET info
        cmd = sqlite_con.CreateCommand
        cmd.CommandText = "UPDATE DB_SETTINGS SET PARAM_VALUE ='" & ParamValue & "' WHERE PARAM_NAME = '" & ParamName & "'"
        cmd.ExecuteNonQuery()

    End Sub
    Private Sub BTN_SETTINGS_Click(sender As Object, e As EventArgs) Handles BTN_SETTINGS.Click
        SZVM_CONFIG.Show()
    End Sub
End Class
