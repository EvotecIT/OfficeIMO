using System.Data;
using Xunit;

namespace OfficeIMO.Odbc.Tests {
    public sealed class CompatibilityContractTests {
        [Fact]
        public unsafe void DiagnosticHeaderWidthsMatchTheirFieldIdentifiers() {
            using NativeHandles handles = new();
            delegate* unmanaged<short, nint, short, short, nint, short, short*, short> get = &Exports.SQLGetDiagFieldW;
            long* output = stackalloc long[2];
            output[0] = 123;
            output[1] = 456;
            Assert.Equal(-1, get(3, handles.Statement, 0, 1, (nint)output, 0, null));
            Assert.Equal(123, output[0]);
            Assert.Equal(456, output[1]);
            Assert.Equal(0, get(3, handles.Statement, 0, 3, (nint)output, 0, null));
            Assert.Equal(-1, output[0]);
            Assert.Equal(456, output[1]);
        }

        [Theory]
        [InlineData(false, "DELETE FROM Customers")]
        [InlineData(true, "DELETE FROM Customers")]
        [InlineData(false, "SELECT ?1")]
        [InlineData(true, "SELECT ?1")]
        [InlineData(false, "markers")]
        [InlineData(true, "markers")]
        public unsafe void FailedReplacementLeavesTheStatementUnprepared(bool direct, string rejected) {
            using NativeHandles handles = new();
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            statement.Sql = "SELECT @p1";
            statement.ParameterCount = 1;
            delegate* unmanaged<nint, nint, int, short> prepare = direct ? &Exports.SQLExecDirectW : &Exports.SQLPrepareW;
            delegate* unmanaged<nint, short> execute = &Exports.SQLExecute;
            if (rejected == "markers") rejected = "SELECT " + string.Join(",", Enumerable.Repeat("?", 257));
            fixed (char* sql = rejected) Assert.Equal(-1, prepare(handles.Statement, (nint)sql, -3));
            Assert.Null(statement.Sql);
            Assert.Equal(0, statement.ParameterCount);
            Assert.Equal(-1, execute(handles.Statement));
            Assert.Equal("HY010", Assert.Single(statement.Diagnostics).State);
        }

        [Theory]
        [InlineData(-12)]
        [InlineData(12)]
        public unsafe void FetchNextIgnoresTheOffset(int offset) {
            using NativeHandles handles = new();
            DataTable table = new();
            table.Columns.Add("Value", typeof(string));
            table.Rows.Add("row");
            ((StatementHandle)DriverHandles.All[handles.Statement]).Reader = table.CreateDataReader();
            delegate* unmanaged<nint, short, nint, short> fetch = &Exports.SQLFetchScroll;
            Assert.Equal(0, fetch(handles.Statement, 1, offset));
            Assert.Equal(100, fetch(handles.Statement, 1, offset));
            Assert.Equal(-1, fetch(handles.Statement, 2, 0));
            Assert.Equal("HY106", Assert.Single(DriverHandles.All[handles.Statement].Diagnostics).State);
        }

        [Theory]
        [InlineData(1, 0)]
        [InlineData(1, 1)]
        [InlineData(2, 0)]
        [InlineData(2, 1)]
        public unsafe void TransactionlessCompletionIsASuccessfulNoOp(short kind, short completion) {
            using NativeHandles handles = new();
            delegate* unmanaged<short, nint, short, short> end = &Exports.SQLEndTran;
            Assert.Equal(0, end(kind, kind == 1 ? handles.Environment : handles.Connection, completion));
            Assert.Equal(-1, end(2, handles.Connection, 5));
            Assert.Equal("HY012", Assert.Single(DriverHandles.All[handles.Connection].Diagnostics).State);
            Assert.Equal(-2, end(3, handles.Statement, 0));
        }

        [Fact]
        public unsafe void CapabilitiesMatchSqliteIdentifierAndSubqueryBehavior() {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            delegate* unmanaged<nint, ushort, nint, short, short*, short> get = &Exports.SQLGetInfoW;
            ushort identifierCase = 0;
            uint subqueries = 0;
            Assert.Equal(0, get(handles.Connection, 93, (nint)(&identifierCase), 2, null));
            Assert.Equal(4, identifierCase);
            Assert.Equal(0, get(handles.Connection, 95, (nint)(&subqueries), 4, null));
            Assert.Equal(23u, subqueries);
            WorkbookSnapshot snapshot = ((ConnectionHandle)DriverHandles.All[handles.Connection]).Snapshot!;
            using System.Data.Common.DbDataReader reader = snapshot.Query("SELECT \"name\" FROM \"customers\" WHERE \"amount\" IN (SELECT Amount FROM Customers)", new Dictionary<string, object?>(), CancellationToken.None);
            Assert.True(reader.Read());
        }

        [Fact]
        public unsafe void CatalogTruncationHasADiagnosticAndDisconnectDropsChildren() {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            nint[] descriptors = statement.Descriptors.Values.ToArray();
            delegate* unmanaged<nint, int, nint, int, int*, short> get = &Exports.SQLGetConnectAttrW;
            char* text = stackalloc char[2];
            int length = 0;
            Assert.Equal(1, get(handles.Connection, 109, (nint)text, 4, &length));
            Assert.Equal(Path.GetFileName(fixture.FilePath).Length * 2, length);
            Assert.Equal("01004", Assert.Single(DriverHandles.All[handles.Connection].Diagnostics).State);
            delegate* unmanaged<nint, short> disconnect = &Exports.SQLDisconnect;
            Assert.Equal(0, disconnect(handles.Connection));
            Assert.True(statement.Freed);
            Assert.False(DriverHandles.All.ContainsKey(handles.Statement));
            Assert.All(descriptors, id => Assert.False(DriverHandles.All.ContainsKey(id)));
            Assert.Equal(-2, disconnect(handles.Statement));
        }

        [Fact]
        public unsafe void TypeAndColumnMetadataDistinguishBinaryAndDecimalPrecision() {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            delegate* unmanaged<nint, short, short> types = &Exports.SQLGetTypeInfoW;
            Assert.Equal(0, types(handles.Statement, 8));
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            Assert.True(statement.Reader!.Read());
            Assert.Equal(53, statement.Reader.GetInt32(2));
            Assert.Equal(2, statement.Reader.GetInt32(17));
            statement.CloseCursor();
            DataTable table = new();
            table.Columns.Add("Text", typeof(string));
            table.Columns.Add("Number", typeof(double));
            statement.Reader = table.CreateDataReader();
            delegate* unmanaged<nint, ushort, ushort, nint, short, short*, nint*, short> column = &Exports.SQLColAttributeW;
            nint output = 0;
            Assert.Equal(0, column(handles.Statement, 1, 4, 0, 0, null, &output));
            Assert.Equal((nint)0, output);
            Assert.Equal(0, column(handles.Statement, 2, 4, 0, 0, null, &output));
            Assert.Equal((nint)15, output);
            Assert.Equal(0, column(handles.Statement, 2, 1005, 0, 0, null, &output));
            Assert.Equal((nint)53, output);
        }

        [Theory]
        [InlineData(typeof(double), 24)]
        [InlineData(typeof(long), 20)]
        [InlineData(typeof(int), 11)]
        [InlineData(typeof(short), 6)]
        [InlineData(typeof(byte[]), -4)]
        [InlineData(typeof(string), -4)]
        public unsafe void DisplayMetadataIncludesSignsExponentsAndBinaryEncoding(Type type, int expected) {
            using NativeHandles handles = new();
            DataTable table = new();
            table.Columns.Add("Value", type);
            ((StatementHandle)DriverHandles.All[handles.Statement]).Reader = table.CreateDataReader();
            delegate* unmanaged<nint, ushort, ushort, nint, short, short*, nint*, short> column = &Exports.SQLColAttributeW;
            nint output = 0;
            Assert.Equal(0, column(handles.Statement, 1, 6, 0, 0, null, &output));
            Assert.Equal((nint)expected, output);
            if (type == typeof(byte[]) || type == typeof(string)) {
                Assert.Equal(0, column(handles.Statement, 1, 1013, 0, 0, null, &output));
                Assert.Equal((nint)(-4), output);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public unsafe void DescriptorNamesSupportNullBufferLengthQueries(bool wide) {
            using NativeHandles handles = new();
            DataTable table = new();
            table.Columns.Add("Name", typeof(string));
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            statement.Reader = table.CreateDataReader();
            delegate* unmanaged<nint, short, short, nint, int, int*, short> get = wide ? &Exports.SQLGetDescFieldW : &Exports.SQLGetDescField;
            int length = -1;
            Assert.Equal(0, get(statement.Descriptors[10012], 1, 1011, 0, 0, &length));
            Assert.Equal(wide ? 8 : 4, length);
            Assert.Equal(-1, get(statement.Descriptors[10012], 1, 1005, 0, 0, &length));
            Assert.Equal("HY009", Assert.Single(DriverHandles.All[statement.Descriptors[10012]].Diagnostics).State);
        }

        [Fact]
        public unsafe void FileAndStatementCapabilitiesMatchTheSupportedWorkflow() {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            delegate* unmanaged<nint, ushort, nint, short, short*, short> info = &Exports.SQLGetInfoW;
            ushort usage = 0;
            Assert.Equal(0, info(handles.Connection, 84, (nint)(&usage), 2, null));
            Assert.Equal(2, usage);
            delegate* unmanaged<nint, short, short> types = &Exports.SQLGetTypeInfoW;
            Assert.Equal(0, types(handles.Statement, 0));
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            List<short> values = [];
            while (statement.Reader!.Read()) values.Add(statement.Reader.GetInt16(1));
            Assert.Equal(new short[] { -9, -7, -5, 4, 5, 8, 93 }, values);
            delegate* unmanaged<nint, short> close = &Exports.SQLCloseCursor;
            Assert.Equal(0, close(handles.Statement));
            Assert.Equal(-1, close(handles.Statement));
            Assert.Equal("24000", Assert.Single(statement.Diagnostics).State);
            delegate* unmanaged<nint, ushort, short> free = &Exports.SQLFreeStmt;
            Assert.Equal(0, free(handles.Statement, 0));
            delegate* unmanaged<nint, int, nint, int, short> set = &Exports.SQLSetStmtAttrW;
            delegate* unmanaged<nint, int, nint, int, int*, short> get = &Exports.SQLGetStmtAttrW;
            nint noscan = 0;
            Assert.Equal(0, get(handles.Statement, 2, (nint)(&noscan), 0, null));
            Assert.Equal((nint)1, noscan);
            Assert.Equal(-1, set(handles.Statement, 2, 0, 0));
            Assert.Equal("HYC00", Assert.Single(statement.Diagnostics).State);
            Assert.Equal(0, set(handles.Statement, 2, 1, 0));
            Assert.Equal(0, set(handles.Statement, 4, 0, 0));
            Assert.Equal(-1, set(handles.Statement, 4, 1, 0));
            Assert.Equal("HYC00", Assert.Single(statement.Diagnostics).State);
        }

        [Theory]
        [InlineData(0, false)]
        [InlineData(1, false)]
        [InlineData(3, false)]
        [InlineData(0, true)]
        [InlineData(1, true)]
        [InlineData(3, true)]
        public unsafe void CompleteConnectionStringsNeedNoDialog(ushort completion, bool wide) {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new();
            byte[] input = NativeText.Bytes("DBQ={" + fixture.FilePath + "};", wide);
            delegate* unmanaged<nint, nint, nint, short, nint, short, short*, ushort, short> connect = wide ? &Exports.SQLDriverConnectW : &Exports.SQLDriverConnect;
            fixed (byte* pointer = input)
                Assert.Equal(0, connect(handles.Connection, 0, (nint)pointer, checked((short)(input.Length / (wide ? 2 : 1))), 0, 0, null, completion));
        }

        [Fact]
        public unsafe void UnsupportedParameterConversionsAreRejectedAndDefaultIntegerIsTyped() {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            delegate* unmanaged<nint, ushort, short, short, short, nuint, short, nint, nint, nint, short> bind = &Exports.SQLBindParameter;
            char* text = stackalloc char[] { '0', '0', '1', '\0' };
            Assert.Equal(-1, bind(handles.Statement, 1, 1, -8, 4, 10, 0, (nint)text, 8, 0));
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            Assert.Equal("07006", Assert.Single(statement.Diagnostics).State);
            Assert.Empty(statement.Parameters);
            int number = 1;
            Assert.Equal(0, bind(handles.Statement, 1, 1, 99, 4, 10, 0, (nint)(&number), 4, 0));
            Assert.Equal(4, statement.Parameters[1].SqlType);
            Assert.Equal(4, statement.Parameters[1].Type);
            string sql = "SELECT ? = 1";
            delegate* unmanaged<nint, nint, int, short> execute = &Exports.SQLExecDirectW;
            fixed (char* pointer = sql) Assert.Equal(0, execute(handles.Statement, (nint)pointer, -3));
            Assert.True(statement.Reader!.Read());
            Assert.Equal(1, statement.Reader.GetInt64(0));
        }

        [Fact]
        public unsafe void ParameterDeclaredLengthsAndTimestampPrecisionAreEnforced() {
            char* text = stackalloc char[] { 'a', 'b', 'c', '\0' };
            ParameterValueException error = Assert.Throws<ParameterValueException>(() => NativeValue.Read(new(-8, (nint)text, 8, 0, -9, 2, 0)));
            Assert.Equal("22001", error.State);
            byte* timestamp = stackalloc byte[16];
            ushort* parts = (ushort*)timestamp;
            parts[0] = 2026; parts[1] = 10; parts[2] = 12; parts[3] = 0; parts[4] = 0; parts[5] = 0;
            *(uint*)(timestamp + 12) = 123_400_000;
            ParameterBinding binding = new(93, (nint)timestamp, 16, 0, 93, 27, 4);
            Assert.IsType<DateTime>(NativeValue.Read(binding));
            error = Assert.Throws<ParameterValueException>(() => NativeValue.Read(binding with { Scale = 3 }));
            Assert.Equal("22008", error.State);
            *(uint*)(timestamp + 12) = 123_400_001;
            error = Assert.Throws<ParameterValueException>(() => NativeValue.Read(binding));
            Assert.Equal("22008", error.State);
        }

        [Fact]
        public unsafe void PreparedMetadataDoesNotEvaluateRowsOrReadParameterBuffers() {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            delegate* unmanaged<nint, nint, int, short> prepare = &Exports.SQLPrepareW;
            delegate* unmanaged<nint, short*, short> columns = &Exports.SQLNumResultCols;
            delegate* unmanaged<nint, ushort, short, short, short, nuint, short, nint, nint, nint, short> bind = &Exports.SQLBindParameter;
            // A metadata call must not dereference this validly bound but unavailable value buffer.
            Assert.Equal(0, bind(handles.Statement, 1, 1, 4, 4, 10, 0, 1, 4, 0));
            string sql = "SELECT Name AS same, Amount AS same, abs(-9223372036854775808) AS unevaluated FROM Customers LIMIT ?";
            fixed (char* pointer = sql) Assert.Equal(0, prepare(handles.Statement, (nint)pointer, -3));
            short count = 0;
            Assert.Equal(0, columns(handles.Statement, &count));
            Assert.Equal(3, count);
            Assert.Null(statement.Reader);
            Assert.Equal("same", statement.ColumnName(0));
            Assert.Equal("same", statement.ColumnName(1));
            Assert.Equal(-9, statement.ColumnType(0));
            Assert.Equal(8, statement.ColumnType(1));
            Assert.Equal(0, statement.ColumnType(2));
            delegate* unmanaged<nint, ushort, nint, short, short*, short*, nuint*, short*, short*, short> describe = &Exports.SQLDescribeColW;
            short type = 0;
            char* name = stackalloc char[32];
            Assert.Equal(0, describe(handles.Statement, 1, (nint)name, 32, null, &type, null, null, null));
            Assert.Equal("same", new string(name));
            Assert.Equal(-9, type);
            delegate* unmanaged<nint, ushort, ushort, nint, short, short*, nint*, short> attribute = &Exports.SQLColAttributeW;
            nint numericType = -1;
            Assert.Equal(0, attribute(handles.Statement, 3, 2, 0, 0, null, &numericType));
            Assert.Equal((nint)0, numericType);
            delegate* unmanaged<nint, short, short, nint, int, int*, short> descriptor = &Exports.SQLGetDescFieldW;
            Assert.Equal(0, descriptor(statement.Descriptors[10012], 0, 1001, (nint)(&count), 2, null));
            Assert.Equal(3, count);
            delegate* unmanaged<nint, short> fetch = &Exports.SQLFetch;
            Assert.Equal(-1, fetch(handles.Statement));
            Assert.Equal("24000", Assert.Single(statement.Diagnostics).State);
            sql = "SELECT Name FROM Customers LIMIT ?";
            fixed (char* pointer = sql) Assert.Equal(0, prepare(handles.Statement, (nint)pointer, -3));
            Assert.Equal(0, columns(handles.Statement, &count));
            Assert.Equal(1, count);
            int limit = 1;
            Assert.Equal(0, bind(handles.Statement, 1, 1, 4, 4, 10, 0, (nint)(&limit), 4, 0));
            delegate* unmanaged<nint, short> execute = &Exports.SQLExecute;
            Assert.Equal(0, execute(handles.Statement));
            Assert.Equal(0, fetch(handles.Statement));
            Assert.Equal("Alicja Żółć", statement.Reader!.GetString(0));
            Assert.Equal(100, fetch(handles.Statement));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public unsafe void CatalogDiscoveryMatchesAndReturnsTheWorkbookFilename(bool wide) {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            string catalog = Path.GetFileName(fixture.FilePath);
            byte[] encoded = NativeText.Bytes(catalog + "\0", wide);
            delegate* unmanaged<nint, nint, short, nint, short, nint, short, nint, short, short> tables = wide ? &Exports.SQLTablesW : &Exports.SQLTables;
            delegate* unmanaged<nint, nint, short, nint, short, nint, short, nint, short, short> columns = wide ? &Exports.SQLColumnsW : &Exports.SQLColumns;
            fixed (byte* pointer = encoded) {
                Assert.Equal(0, tables(handles.Statement, (nint)pointer, -3, 0, 0, 0, 0, 0, 0));
                Assert.True(statement.Reader!.Read());
                Assert.Equal(catalog, statement.Reader.GetString(0));
                Assert.Equal("Customers", statement.Reader.GetString(2));
                Assert.Equal(0, columns(handles.Statement, (nint)pointer, -3, 0, 0, 0, 0, 0, 0));
                Assert.True(statement.Reader!.Read());
                Assert.Equal(catalog, statement.Reader.GetString(0));
                Assert.Equal("Name", statement.Reader.GetString(3));
            }
            byte[] all = NativeText.Bytes("%\0", wide), empty = NativeText.Bytes("\0", wide);
            fixed (byte* pattern = all) fixed (byte* blank = empty) {
                Assert.Equal(0, tables(handles.Statement, (nint)pattern, -3, (nint)blank, 0, (nint)blank, 0, 0, 0));
                Assert.True(statement.Reader!.Read());
                Assert.Equal(catalog, statement.Reader.GetString(0));
                Assert.False(statement.Reader.Read());
            }
        }

        [Theory]
        [InlineData(91, 2026, 13, 1, 0, 0, 0)]
        [InlineData(91, 2026, 2, 30, 0, 0, 0)]
        [InlineData(91, -1, 1, 1, 0, 0, 0)]
        [InlineData(93, 2026, 1, 1, 24, 0, 0)]
        [InlineData(93, 2026, 1, 1, 0, 60, 0)]
        [InlineData(93, 2026, 1, 1, 0, 0, 60)]
        public unsafe void InvalidInputDatetimeFieldsReturnTheDatetimeDiagnostic(short type, short year, short month, short day, short hour, short minute, short second) {
            using WorkbookFixture fixture = new();
            using NativeHandles handles = new(fixture.FilePath);
            short* parts = stackalloc short[] { year, month, day, hour, minute, second, 0, 0 };
            delegate* unmanaged<nint, ushort, short, short, short, nuint, short, nint, nint, nint, short> bind = &Exports.SQLBindParameter;
            Assert.Equal(0, bind(handles.Statement, 1, 1, type, type, 27, 7, (nint)parts, 16, 0));
            string sql = "SELECT ?";
            delegate* unmanaged<nint, nint, int, short> direct = &Exports.SQLExecDirectW;
            fixed (char* pointer = sql) Assert.Equal(-1, direct(handles.Statement, (nint)pointer, -3));
            Assert.Equal("22008", Assert.Single(DriverHandles.All[handles.Statement].Diagnostics).State);
        }

        private sealed class NativeHandles : IDisposable {
            internal nint Environment { get; }
            internal nint Connection { get; }
            internal nint Statement { get; }
            internal NativeHandles(string? path = null) {
                Assert.Equal(0, DriverHandles.Allocate(1, 0, out nint environment));
                Environment = environment;
                Assert.Equal(0, DriverHandles.Allocate(2, environment, out nint connection));
                Connection = connection;
                if (path is not null) ((ConnectionHandle)DriverHandles.All[connection]).Snapshot = new WorkbookSnapshot(path, CancellationToken.None);
                Assert.Equal(0, DriverHandles.Allocate(3, connection, out nint statement));
                Statement = statement;
            }
            public unsafe void Dispose() {
                delegate* unmanaged<short, nint, short> free = &Exports.SQLFreeHandle;
                if (DriverHandles.All.ContainsKey(Statement)) Assert.Equal(0, free(3, Statement));
                Assert.Equal(0, free(2, Connection));
                Assert.Equal(0, free(1, Environment));
            }
        }
    }
}
