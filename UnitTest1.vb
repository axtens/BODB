Imports BODB

Namespace BODB_Test
    Public Class Tests
        <NUnit.Framework.SetUp>
        Public Sub Setup()

        End Sub

        <NUnit.Framework.Test>
        Public Sub NoBarChar()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo("BOBD-E-NO_VERTICAL_BAR_CHARS"))
        End Sub
        <NUnit.Framework.Test>
        Public Sub FieldCountTooLow()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo("BOBD-E-NO_VERTICAL_BAR_CHARS"))
        End Sub
        <NUnit.Framework.Test>
        Public Sub EmptyFileName()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("|tablename|fieldname|value")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo("BOBD-E-EMPTY_FILENAME"))
        End Sub
        <NUnit.Framework.Test>
        Public Sub NotFoundFileName()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("c:\temp\notthere.db|tablename|fieldname|value")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo("BOBD-E-FILE_NOT_FOUND"))
        End Sub
        <NUnit.Framework.Test>
        Public Sub BadTableName()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("C:\temp\tether.db||fieldname|value")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo(""))
        End Sub
        <NUnit.Framework.Test>
        Public Sub BadFieldName()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("C:\temp\msys_678959127.db|tablename||value")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo("BODB-E-TABLENAME_NOT_FOUND"))
        End Sub
        <NUnit.Framework.Test>
        Public Sub GetAllTetherRecords()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("C:\temp\tether.db|Tethers")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo(""))
        End Sub
        <NUnit.Framework.Test>
        Public Sub GetAllGaFieldInTetherRecords()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("C:\temp\tether.db|Tethers|ga")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo(""))
        End Sub
        <NUnit.Framework.Test>
        Public Sub GetAllGaFieldWhere2InTetherRecords()
            Dim status As String = Nothing, cargo As String = Nothing
            (status, cargo) = DBSQLite("C:\temp\tether.db|Tethers|ga|2")
            NUnit.Framework.Assert.That(status, NUnit.Framework.Is.EqualTo(""))
        End Sub

    End Class
End Namespace
