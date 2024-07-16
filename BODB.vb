Imports System.Data.SQLite
Imports System.Text.RegularExpressions
Imports System.Linq

Public Module BODB
    Public Function DBSQLite(ByVal PString) As (status As String, cargo As String)
        'PString: fileName '|' table '|' field '|' value 

        'Check for vertical bar char
        Dim verticalBarPos As Integer = InStr(PString, "|", CompareMethod.Text)
        If verticalBarPos = 0 Then
            Return ("BOBD-E-NO_VERTICAL_BAR_CHARS", PString)
        End If

        'Split on vertical bar chars
        Dim parts = Split(PString, "|", -1, CompareMethod.Text).ToList()

        Dim partsCount = parts.Count
        'check all fields present (4 in total)
        If parts.Count = 0 Then
            Return ("BODB-E-NOT_ENOUGH_FIELDS_PRESENT", PString)
        End If

        'get and check filename
        Dim idx As Integer = 0
        Dim fileName As String = parts(idx)
        If fileName = vbNullString Then
            Return ("BOBD-E-EMPTY_FILENAME", "")
        End If
        If Not IO.File.Exists(fileName) Then
            Return ("BOBD-E-FILENAME_NOT_FOUND", fileName)
        End If

        'connect to sqlite file
        Dim conn As New SQLiteConnection()

        'fileName might be a connection string
        Dim connectionString As String = ReworkConnectionString(fileName)
        'conn.ConnectionString = connectionString
        'conn.Open()

        'get and check tablename
        idx = 1
        Dim tableName As String = parts(idx)
        If tableName = vbNullString Then
            Return ("BODB-E-EMPTY_TABLENAME", PString)
        End If

        Dim tableList = GetListOfTables(connectionString)

        If tableList.IndexOf(tableName) = -1 Then
            Return ("BODB-E-TABLENAME_NOT_FOUND", PString)
        End If

        'get and check fieldname
        Dim fieldName As String = parts(2)

        'get and check comparison value
        Dim comparisonValue As String = parts(3)

    End Function

    Private Function GetListOfTables(connectionString As String) As List(Of String)
        Dim tables As New List(Of String)
        Dim reader = SQLiteCommand.Execute("PRAGMA table_list;", SQLiteExecuteType.Reader, connectionString)

        Do
            Dim values = reader.GetValues()
            tables.Add(values("name"))
        Loop Until Not reader.Read()

        Return tables
    End Function

    Private Function ReworkConnectionString(fileName As String) As String
        Dim source As String = fileName
        If Right(source, 1) <> ";" Then
            source &= ";"
        End If

        Dim answer = UnPattern(source, "Data Source=(.*?);")
        Dim dataSource As String = answer.result
        source = answer.newSource

        answer = UnPattern(source, "Version=(\d);")
        Dim version As String = answer.result
        source = answer.newSource

        answer = UnPattern(source, "New=(True|False);")
        Dim newDB As String = answer.result
        source = answer.newSource

        answer = UnPattern(source, "UseUTF16Encoding=True;")
        Dim useUTF16Encoding As String = answer.result
        source = answer.newSource

        answer = UnPattern(source, "Password=(.*?);")
        source = answer.newSource
        Dim password As String = answer.result

        answer = UnPattern(source, "Legacy Format=(True|False);")
        source = answer.newSource
        Dim legacyFormat As String = answer.result

        answer = UnPattern(source, "Pooling=(True|False);")
        source = answer.newSource
        Dim pooling As String = answer.result

        answer = UnPattern(source, "Max Pool Size=(\d+?);")
        source = answer.newSource
        Dim maxPoolSize As String = answer.result

        answer = UnPattern(source, "BinaryGUID=(False|True);")
        source = answer.newSource
        Dim binaryGUID As String = answer.result

        answer = UnPattern(source, "Cache Size=(\d+?);")
        source = answer.newSource
        Dim cacheSize As String = answer.result

        answer = UnPattern(source, "Page Size=(\d+?);")
        source = answer.newSource
        Dim pageSize As String = answer.result

        answer = UnPattern(source, "Enlist=(N|Y);")
        source = answer.newSource
        Dim enlist As String = answer.result

        answer = UnPattern(source, "FailIfMissing=(True|False);")
        source = answer.newSource
        Dim failIfMissing As String = answer.result

        answer = UnPattern(source, "Max Page Count=(\d+?);")
        source = answer.newSource
        Dim maxPageCount As String = answer.result

        answer = UnPattern(source, "Journal Mode=(Off|On);")
        source = answer.newSource
        Dim journalMode As String = answer.result

        answer = UnPattern(source, "Synchronous=(Full|Normal);")
        source = answer.newSource
        Dim synchronous As String = answer.result


        If dataSource = vbNullString Then
            Return $"Data Source={fileName};"
        Else
            Return Join({
                        IIf(dataSource <> vbNullString, $"Data Source={dataSource};", vbNullString),
                        IIf(version <> vbNullString, $"Version={version};", vbNullString),
                        IIf(binaryGUID <> vbNullString, $"BinaryGUID={binaryGUID};", vbNullString),
                        IIf(cacheSize <> vbNullString, $"Cache Size={cacheSize};", vbNullString),
                        IIf(enlist <> vbNullString, $"Enlist={enlist};", vbNullString),
                        IIf(failIfMissing <> vbNullString, $"FailIfMissing={failIfMissing};", vbNullString),
                        IIf(journalMode <> vbNullString, $"Journal Mode={journalMode};", vbNullString),
                        IIf(legacyFormat <> vbNullString, $"Legacy Format={legacyFormat};", vbNullString),
                        IIf(maxPageCount <> vbNullString, $"Max Page Count={maxPageCount};", vbNullString),
                        IIf(maxPoolSize <> vbNullString, $"Max Pool Size={maxPoolSize};", vbNullString),
                        IIf(newDB <> vbNullString, $"New={newDB};", vbNullString),
                        IIf(pageSize <> vbNullString, $"Page Size={pageSize};", vbNullString),
                        IIf(password <> vbNullString, $"Password={password};", vbNullString),
                        IIf(pooling <> vbNullString, $"Pooling={pooling};", vbNullString),
                        IIf(synchronous <> vbNullString, $"Synchronous={synchronous};", vbNullString),
                        IIf(useUTF16Encoding <> vbNullString, $"UseUTF16Encoding={useUTF16Encoding};", vbNullString)
                    }, vbNullString)
        End If
    End Function

    Private Function UnPattern(source As String, pattern As String) As (newSource As String, result As String)
        Dim result As String = vbNullString
        Dim newSource As String = vbNullString
        If source <> vbNullString Then
            Dim regex As New Regex(pattern, RegexOptions.IgnoreCase)
            If regex.IsMatch(source) Then
                Dim matches = regex.Matches(source)
                result = matches.Item(0).Groups(1).Value.Trim()
                newSource = source.Replace(matches.Item(0).Groups(0).Value, vbNullString)
            End If
        End If

        Return (newSource, result)
    End Function

End Module
