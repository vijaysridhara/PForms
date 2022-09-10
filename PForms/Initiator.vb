'***********************************************************************
'Copyright 2005-2022 Vijay Sridhara

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at

'http://www.apache.org/licenses/LICENSE-2.0

'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'***********************************************************************
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Data.SQLite

Public Class Initiator

    Private content() As String
    Event MessageSent(ByVal msg As String)
    Event ClearMessagearea()
    Dim conn As System.Data.SQLite.SQLiteConnection
    Dim fpath As String
    Dim finalFormPath As String

    Private MetaForm As Meta
    Private FDEF As FormDef

    Private formMetaWithuser As String = "\[\s*META:\s*([A-Za-z0-9\s]+)\s*:\s*([A-Za-z\s\@\.]+)\s*\]"
    Private formMeta As String = "\[\s*META:\s*([A-Za-z0-9\s]+)\s*\]"
    Private formDef As String = "\[FORMDEF\s*:\s*([A-Za-z0-9\s]+)\s*:\s*([A-Za-z_]+)\s*\]"
    Private formFieldsSourced As String = "\[\s*FIELD\s*:\s*([A-Za-z]+)\s*:\s*(ROCOMBO|COMBO)\s*:\s*(.+)\s*\]"
    Private formFieldsListed As String = "\[\s*FIELD\s*:\s*([A-Za-z]+)\s*:\s*LIST\s*:\s*([A-Za-z0-9\._,@#\$\%\&\*\+\-\?]+)\s*\]"
    Private formUserLabel As String = "\[\s*FIELD\s*:\s*USER\s*:\s*LABEL\s*\]"
    Private formSerialLabel As String = "\[\s*FIELD\s*:\s*([A-Za-z]+)\s*:\s*SERIAL\s*\]"
    Private formLabels As String = "\[\s*FIELD\s*:\s*([A-Za-z]+)\s*:\s*LABEL\s*:\s*(.*)\s*\]"
    Private formFields As String = "\[\s*FIELD\s*:\s*([A-Za-z]+)\s*:\s*(TEXT|MULTILINE|LARGETEXT|CHECK|RADIO|DATE|FILE|PICTURE|NUMERIC|MONTHYEAR|RICHTEXT)\s*\]"
    Private formReport As String = "\[\s*REPORT\s*:\s*([A-Za-z\s]+)\s*:\s*([A-Za-z_]+)\s*:\s*(.+)\s*\]"
    Public ReadOnly Property FormPath As String
        Get
            Return finalFormPath
        End Get
    End Property
    Public Function ParseContent(ByVal fpath As String) As Boolean
        Dim errormessage As String = ""
        Try
            Me.fpath = fpath
            Dim sr As New IO.StreamReader(fpath)
            Dim def As String = sr.ReadToEnd()
            Me.content = def.Split(vbLf)
            sr.Close()
            sr.Dispose()
            Dim parseEntries As New Dictionary(Of String, String)
            MetaForm = New Meta
            Dim linenumber As Integer
            RaiseEvent ClearMessagearea()
            Dim metafound As Boolean
            For Each ln As String In content
                ln = ln.Trim
                linenumber += 1
                Debug.Print(ln)
                If ln.Length = 0 Then Continue For
                If ln.Substring(0, 1) = "#" Then Continue For
                Dim lnSTart As Integer = 0
                Dim lnEnd As Integer = 0
                If metafound = True Then GoTo NEXTLINE
                Dim rgx As Regex = New Regex(formMetaWithuser, RegexOptions.IgnoreCase)
                Dim mtch As Match = rgx.Match(ln)
                If mtch.Success Then
                    MetaForm.Name = mtch.Groups(1).Value
                    MetaForm.Username = mtch.Groups(2).Value
                    MetaForm.Version = FORM_VER
                    metafound = True
                    Continue For
                Else
                    rgx = New Regex(formMeta, RegexOptions.IgnoreCase)
                    mtch = rgx.Match(ln)
                    If mtch.Success Then
                        MetaForm.Name = mtch.Groups(1).Value
                        MetaForm.Username = "#UNDEFINED#"
                        MetaForm.Version = FORM_VER
                        metafound = True
                        Continue For
                    Else
                        Continue For
                    End If
                End If
NEXTLINE:       Dim eline As Integer
                If metafound And ln.Substring(0, 8) = "[FORMDEF" Then
                    lnSTart = linenumber - 1
                    eline = CheckDefinition(content, lnSTart)
                    If eline = -1 Then
                        RaiseEvent MessageSent("Compilation stopped. Please correct errors")
                        Return False
                    Else
                        Exit For
                    End If
                Else
                    Continue For
                End If
            Next
            If metafound = False Then
                RaiseEvent MessageSent("No Meta defined for this form.So exiting")
                Return False
            End If
            RaiseEvent MessageSent("Successfully compiled. No Syntax errors found")

            Dim fname As String = IO.Path.GetDirectoryName(fpath) & "\" & IO.Path.GetFileNameWithoutExtension(fpath) & ".pf1"
            Dim sw As New IO.StreamWriter(fname, False)
            sw.Close()
            sw.Dispose()
            finalFormPath = fname
            Return SetupDB(fname)

        Catch ex As Exception
            RaiseEvent MessageSent(ex.Message)
        End Try

EXITTOERROR:
        RaiseEvent MessageSent(errormessage)
        Return False
    End Function


    Private Function SetupDB(ByVal fname As String) As Boolean
        Dim successfullydone As Boolean
        Try
            conn = New SQLite.SQLiteConnection("Data Source=" + fname + ";Version=3;")
            conn.Open()
            Dim cmd As New SQLite.SQLiteCommand

            With cmd
                .Connection = conn
                .CommandText = "CREATE TABLE META(NAME VARCHAR(100),VER CHAR(10),USER VARCHAR(100),PRIMARY KEY(NAME))"
                .ExecuteNonQuery()
                .CommandText = "CREATE TABLE FORM_DEF ( NAME VARCHAR(100), DATASTORE VARCHAR(50), PRIMARY KEY(NAME) )"
                .ExecuteNonQuery()
                .CommandText = "CREATE TABLE FORM_FIELDS ( NAME VARCHAR(100), FORMNAME VARCHAR(100), TYPE VARCHAR(10), SOURCE VARCHAR(100), PRIMARY KEY (NAME,FORMNAME))"
                .ExecuteNonQuery()
                .CommandText = "CREATE TABLE FORM_REPORTS ( NAME VARCHAR(200), FIELDS VARCHAR(1000),INBUILT CHAR(1),IMPORTTO VARCHAR(50),PRIMARY KEY (NAME))		"
                .ExecuteNonQuery()
                .Parameters.Clear()
                .CommandText = "INSERT INTO META(NAME,VER,USER) VALUES(@NAME,@VER,@USER)"
                .Parameters.Add("NAME", DbType.String)
                .Parameters.Add("VER", DbType.String)
                .Parameters.Add("USER", DbType.String)
                .Parameters("NAME").Value = MetaForm.Name
                .Parameters("VER").Value = MetaForm.Version
                .Parameters("USER").Value = MetaForm.Username
                .ExecuteNonQuery()
                For Each FDEF As FormDef In MetaForm.Forms.Values
                    .CommandText = "INSERT INTO FORM_DEF(NAME,DATASTORE) VALUES(@NAME,@DATASTORE)"
                    .Parameters.Add("NAME", DbType.String)
                    .Parameters("NAME").Value = FDEF.Title
                    .Parameters.Add("DATASTORE", DbType.String)
                    .Parameters("DATASTORE").Value = FDEF.DataStore
                    .ExecuteNonQuery()
                    .Parameters.Clear()
                    .CommandText = "INSERT INTO FORM_FIELDS(NAME,FORMNAME,TYPE,SOURCE) VALUES(@NAME,@FORMNAME,@TYPE,@SOURCE)"
                    .Parameters.Add("NAME", DbType.String)
                    .Parameters.Add("FORMNAME", DbType.String)
                    .Parameters.Add("TYPE", DbType.String)
                    .Parameters.Add("SOURCE", DbType.String)
                    Dim usfd As New Field
                    With usfd
                        .Name = "USER"
                        .FieldType = PDFieldType.Label
                    End With

                    Dim tsfd As New Field
                    With tsfd
                        .Name = "LASTUPT"
                        .FieldType = PDFieldType.Timestamp
                    End With
                    FDEF.Fields.Add(usfd.Name, usfd)
                    FDEF.Fields.Add(tsfd.Name, tsfd)
                    For Each FD As Field In FDEF.Fields.Values
                        .Parameters("SOURCE").Value = ""
                        .Parameters("NAME").Value = FD.Name
                        .Parameters("FORMNAME").Value = FDEF.Title
                        Select Case FD.FieldType
                            Case PDFieldType.Check
                                .Parameters("TYPE").Value = "CHECK"
                            Case PDFieldType.Combo
                                .Parameters("TYPE").Value = "COMBO"
                                .Parameters("SOURCE").Value = FD.DataSource
                            Case PDFieldType.ROCombo
                                .Parameters("TYPE").Value = "ROCOMBO"
                                .Parameters("SOURCE").Value = FD.DataSource
                            Case PDFieldType.List
                                .Parameters("TYPE").Value = "LIST"
                                .Parameters("SOURCE").Value = FD.DataSource
                            Case PDFieldType.Date
                                .Parameters("TYPE").Value = "DATE"
                            Case PDFieldType.File
                                .Parameters("TYPE").Value = "FILE"
                            Case PDFieldType.Multiline
                                .Parameters("TYPE").Value = "MULTILINE"
                            Case PDFieldType.Largetext
                                .Parameters("TYPE").Value = "LARGETEXT"
                            Case PDFieldType.Picture
                                .Parameters("TYPE").Value = "PICTURE"
                            Case PDFieldType.Radio
                                .Parameters("TYPE").Value = "RADIO"
                            Case PDFieldType.Text
                                .Parameters("TYPE").Value = "TEXT"
                            Case PDFieldType.Timestamp
                                .Parameters("TYPE").Value = "TIMESTAMP"
                                '.Parameters("SOURCE").Value = FD.DataSource
                            Case PDFieldType.Label
                                .Parameters("TYPE").Value = "LABEL"
                                .Parameters("SOURCE").Value = FD.DataSource
                            Case PDFieldType.Numeric
                                .Parameters("TYPE").Value = "NUMERIC"
                                .Parameters("SOURCE").Value = "0"
                            Case PDFieldType.Serial
                                .Parameters("TYPE").Value = "SERIAL"
                            Case PDFieldType.MonthYear
                                .Parameters("TYPE").Value = "MONTHYEAR"
                            Case PDFieldType.RichText
                                .Parameters("TYPE").Value = "RICHTEXT"
                        End Select
                        .ExecuteNonQuery()
                    Next
                    If FDEF.AllFieldsLabels Then Continue For
                    .Parameters.Clear()
                    .CommandText = "CREATE TABLE " & FDEF.DataStore & " ("
                    For Each fd As Field In FDEF.Fields.Values
                        .CommandText += fd.Name
                        Select Case fd.FieldType
                            Case PDFieldType.Check, PDFieldType.Radio
                                .CommandText += " CHAR(1),"
                            Case PDFieldType.Combo, PDFieldType.Text, PDFieldType.Label, PDFieldType.ROCombo
                                .CommandText += " VARCHAR(200),"
                            Case PDFieldType.List
                                .CommandText += " VARCHAR(1000),"
                            Case PDFieldType.MonthYear
                                .CommandText += " CHAR(7),"
                            Case PDFieldType.Date
                                .CommandText += " DATE,"
                            Case PDFieldType.File, PDFieldType.Picture, PDFieldType.RichText
                                .CommandText += " BLOB,"
                            Case PDFieldType.Multiline, PDFieldType.Largetext
                                .CommandText += " TEXT,"

                            Case PDFieldType.Picture
                                .CommandText += " BLOB,"
                            Case PDFieldType.Timestamp
                                If fd.Name = "LASTUPT" Then
                                    .CommandText += " TIMESTAMP"
                                Else
                                    .CommandText += " TIMESTAMP,"
                                End If
                            Case PDFieldType.Numeric
                                .CommandText += " NUMERIC,"
                            Case PDFieldType.Serial
                                .CommandText += " INTEGER PRIMARY KEY AUTOINCREMENT,"

                        End Select
                    Next
                    .CommandText += ")"
                    .ExecuteNonQuery()
                    .Parameters.Clear()
                Next
                .CommandText = "INSERT INTO FORM_REPORTS(NAME,FIELDS,INBUILT,IMPORTTO) VALUES(@NAME,@FIELDS,@INBUILT,@IMPORTTO)"
                .Parameters.Add("NAME", DbType.String)
                .Parameters.Add("FIELDS", DbType.String)
                .Parameters.Add("INBUILT", DbType.String)
                .Parameters.Add("IMPORTTO", DbType.String)
                For Each RP As Report In MetaForm.GridReport.Values
                    .Parameters("NAME").Value = RP.Title
                    .Parameters("FIELDS").Value = RP.SQL
                    .Parameters("INBUILT").Value = IIf(RP.Inbuilt, "Y", "N")
                    .Parameters("IMPORTTO").Value = RP.ImportDataStore
                    .ExecuteNonQuery()
                Next
                .Parameters.Clear()
            End With
            successfullydone = True
            RaiseEvent MessageSent("Completed compiling the form definition")

        Catch ex As Exception
            RaiseEvent MessageSent(ex.Message)
            Return False
        Finally
            conn.Close()
            conn.Dispose()
        End Try
        Return True
    End Function
    Private Function CheckDefinition(ByVal lines() As String, ByVal startix As Integer) As Integer
        Try
            Dim errormessage As String = ""
            Dim hasformfields, hasformdef, hasreport As Boolean
            Dim regex As System.Text.RegularExpressions.Regex, mtch As System.Text.RegularExpressions.Match
            Dim nextFormFound As Boolean
            Dim linenum As Integer = 0
            Dim FDEF As FormDef
            Dim hasallfieldslabels As Boolean = True
            For linenum = startix To lines.Length - 1
                Dim ln As String = content(linenum)
                Debug.Print(ln)
                ln = ln.Trim
                If ln.Length = 0 Then Continue For
                If ln.Substring(0, 1) = "#" Then Continue For
                If ln.Substring(0, 8) = "[FORMDEF" And linenum > startix Then 'this check is required, because it finds on first line any way
                    nextFormFound = True
                    Exit For
                End If

                If hasformdef = False Then
                    regex = New Regex(formDef, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then
                        FDEF = New FormDef
                        FDEF.Title = mtch.Groups(1).Value
                        hasformdef = True
                        FDEF.DataStore = mtch.Groups(2).Value
                    End If
                Else
                    regex = New Regex(formFields, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then
                        Dim fld As New Field
                        With fld
                            .Name = mtch.Groups(1).Value
                            If mtch.Groups(2).Value <> "LABEL" Then
                                hasallfieldslabels = False
                            End If
                            Select Case mtch.Groups(2).Value
                                Case "TEXT"
                                    .FieldType = PDFieldType.Text
                                Case "MULTILINE"
                                    .FieldType = PDFieldType.Multiline
                                Case "CHECK"
                                    .FieldType = PDFieldType.Check
                                Case "RADIO"
                                    .FieldType = PDFieldType.Radio
                                Case "DATE"
                                    .FieldType = PDFieldType.Date
                                Case "FILE"
                                    .FieldType = PDFieldType.File
                                Case "PICTURE"
                                    .FieldType = PDFieldType.Picture
                                Case "NUMERIC"
                                    .FieldType = PDFieldType.Numeric
                                Case "LARGETEXT"
                                    .FieldType = PDFieldType.Largetext
                                Case "MONTHYEAR"
                                    .FieldType = PDFieldType.MonthYear
                                Case "RICHTEXT"
                                    .FieldType = PDFieldType.RichText
                            End Select
                            FDEF.Fields.Add(fld.Name, fld)
                            hasformfields = True
                            Continue For
                        End With
                    End If
                    regex = New Regex(formFieldsSourced, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then
                        Dim fld As New Field
                        With fld
                            If mtch.Groups(2).Value = "ROCOMBO" Then
                                .FieldType = PDFieldType.ROCombo
                            Else
                                .FieldType = PDFieldType.Combo
                            End If

                            .Name = mtch.Groups(1).Value
                            .DataSource = mtch.Groups(3).Value
                        End With
                        FDEF.Fields.Add(fld.Name, fld)
                        hasformfields = True
                        Continue For
                    End If
                    regex = New Regex(formSerialLabel, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then
                        Dim fld As New Field
                        With fld
                            .Name = mtch.Groups(1).Value
                            .FieldType = PDFieldType.Serial
                        End With
                        FDEF.Fields.Add(fld.Name, fld)
                        hasformfields = True
                        Continue For
                    End If
                    regex = New Regex(formFieldsListed, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then
                        Dim fld As New Field
                        With fld
                            .FieldType = PDFieldType.List
                            .Name = mtch.Groups(1).Value
                            .DataSource = mtch.Groups(2).Value
                        End With
                        FDEF.Fields.Add(fld.Name, fld)
                        hasformfields = True
                        Continue For
                    End If
                    regex = New Regex(formLabels, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then
                        Dim fld As New Field
                        With fld
                            .FieldType = PDFieldType.Label
                            .Name = mtch.Groups(1).Value
                            .DataSource = mtch.Groups(2).Value
                        End With
                        FDEF.Fields.Add(fld.Name, fld)
                        hasformfields = True
                        Continue For
                    End If
                    regex = New Regex(formReport, RegexOptions.IgnoreCase)
                    mtch = regex.Match(ln)
                    If mtch.Success Then

                        Dim rp As New Report
                        rp.Title = mtch.Groups(1).Value
                        Dim sql As String = mtch.Groups(3).Value
                        If sql.Contains(",USER,LASTUPT") = False Then
                            RaiseEvent MessageSent("The report should have USER,LASTUPT as the last columns, if the report is to be importable: " & rp.Title)
                            rp.Inbuilt = False
                        Else
                            rp.Inbuilt = True
                        End If
                        Dim hdrfnd As Boolean = False
                        rp.SQL = sql

                        rp.ImportDataStore = mtch.Groups(2).Value
                        hasreport = True
                        MetaForm.GridReport.Add(rp.Title, rp)
                        Continue For
                    End If
                End If

            Next
            If hasformdef = False Then
                errormessage = "There is no form definition found"
                GoTo EXITTOERROR
            ElseIf hasformdef And hasformfields = False Then
                errormessage = "The form definition has no fields defined"
                GoTo EXITTOERROR
            Else

            End If
            FDEF.AllFieldsLabels = hasallfieldslabels
            MetaForm.Forms.Add(FDEF.Title, FDEF)
            RaiseEvent MessageSent("Completed for form : " & FDEF.Title)
            If nextFormFound Then Return CheckDefinition(lines, linenum)
            Return linenum
ExitToError:
            RaiseEvent MessageSent(errormessage)
            Return -1
        Catch ex As Exception
            RaiseEvent MessageSent(ex.Message)
            Return -1
        End Try
    End Function

    Public Function ReadForm(ByVal ffile As String) As Meta
        RaiseEvent MessageSent("Reading the form to load....")
        Dim dbfile As String, connString = "Data Source=#FILE#;Version=3;"
        Dim cmd As New SQLiteCommand, rdr As SQLiteDataReader
        Try
            dbfile = ffile
            connString = connString.Replace("#FILE#", dbfile)
            RaiseEvent MessageSent("Opening up the form...")
            conn = New SQLiteConnection(connString)
            conn.Open()
            cmd.Connection = conn
            With cmd
                .CommandText = "SELECT NAME,VER,USER FROM META"
                rdr = .ExecuteReader
                If rdr.Read Then
                    MetaForm = New Meta
                    MetaForm.Name = rdr(0)
                    MetaForm.Version = rdr(1)
                    MetaForm.Username = rdr(2)


                Else
                    RaiseEvent MessageSent("Meta definition missing... so exiting")
                    Return Nothing
                End If
                rdr.Close()

                .CommandText = "SELECT NAME,DATASTORE FROM FORM_DEF"
                rdr = .ExecuteReader
                While rdr.Read

                    Dim fdef = New FormDef
                    fdef.Title = rdr(0)
                    fdef.DataStore = rdr(1)
                    MetaForm.Forms.Add(fdef.Title, fdef)
                End While
                rdr.Close()
                For Each FDEF As FormDef In MetaForm.Forms.Values
                    .CommandText = "SELECT NAME,TYPE,SOURCE FROM  FORM_FIELDS WHERE FORMNAME='" & FDEF.Title & "'"
                    rdr = .ExecuteReader
                    Dim fd As Field
                    Dim hasallfieldslabels As Boolean = True
                    While rdr.Read
                        fd = New Field
                        With fd
                            .Name = rdr(0)
                            If rdr(1) <> "LABEL" And Not (rdr(0) = "LASTUPT" And rdr(1) = "TIMESTAMP") Then
                                hasallfieldslabels = False
                            End If
                            Select Case rdr(1)
                                Case "TEXT"
                                    fd.FieldType = PDFieldType.Text
                                Case "COMBO"
                                    fd.FieldType = PDFieldType.Combo
                                    fd.DataSource = rdr(2)
                                Case "ROCOMBO"
                                    fd.FieldType = PDFieldType.ROCombo
                                    fd.DataSource = rdr(2)
                                Case "DATE"
                                    fd.FieldType = PDFieldType.Date
                                Case "PICTURE"
                                    fd.FieldType = PDFieldType.Picture
                                Case "FILE"
                                    fd.FieldType = PDFieldType.File
                                Case "CHECK"
                                    fd.FieldType = PDFieldType.Check
                                Case "RADIO"
                                    fd.FieldType = PDFieldType.Radio
                                Case "MULTILINE"
                                    fd.FieldType = PDFieldType.Multiline
                                Case "LARGETEXT"
                                    fd.FieldType = PDFieldType.Largetext
                                Case "TIMESTAMP"
                                    fd.FieldType = PDFieldType.Timestamp
                                Case "LABEL"
                                    fd.FieldType = PDFieldType.Label
                                    fd.DataSource = IIf(rdr.IsDBNull(2), "", rdr(2))
                                Case "NUMERIC"
                                    fd.FieldType = PDFieldType.Numeric
                                Case "LIST"
                                    fd.FieldType = PDFieldType.List
                                    fd.DataSource = rdr(2)
                                Case "SERIAL"
                                    fd.FieldType = PDFieldType.Serial
                                Case "MONTHYEAR"
                                    fd.FieldType = PDFieldType.MonthYear
                                Case "RICHTEXT"
                                    fd.FieldType = PDFieldType.RichText
                                Case Else
                                    fd.FieldType = PDFieldType.Text

                            End Select
                        End With
                        FDEF.Fields.Add(fd.Name, fd)
                    End While
                    FDEF.AllFieldsLabels = hasallfieldslabels
                    rdr.Close()
                Next
            End With
            cmd.CommandText = "SELECT NAME,FIELDS,INBUILT,IMPORTTO FROM FORM_REPORTS"
            rdr = cmd.ExecuteReader
            While rdr.Read
                Dim RP As New Report
                RP.Title = rdr(0)
                RP.SQL = rdr(1)

                If rdr(2) = "Y" Then
                    RP.Inbuilt = True
                Else
                    RP.Inbuilt = False
                End If
                RP.ImportDataStore = rdr(3)
                MetaForm.GridReport.Add(RP.Title, RP)
            End While
            rdr.Close()
            conn.Close()
            conn.Dispose()

            RaiseEvent MessageSent("Your form is read with contents")
            Return MetaForm
        Catch ex As Exception
            RaiseEvent MessageSent(ex.Message)
            Return Nothing
        End Try

    End Function
End Class
