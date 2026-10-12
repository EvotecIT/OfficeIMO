using System.Data;
using Xunit;

namespace OfficeIMO.Odbc.Tests {
    public sealed class NativeContractTests {
        [Fact]
        public unsafe void IntegerDiagnosticsAndDescriptorFieldsPreserveAdjacentMemory() {
            using HandleFamily handles = new();
            DriverHandles.All[handles.Statement].Fail("HY000", "Test diagnostic");
            int* output = stackalloc int[2];
            output[1] = 0x12345678;
            delegate* unmanaged<short, nint, short, short, nint, short, short*, short> diagnostic = &Exports.SQLGetDiagFieldW;
            Assert.Equal(0, diagnostic(3, handles.Statement, 1, -1247, (nint)output, 0, null));
            Assert.Equal(-1, output[0]);
            Assert.Equal(0x12345678, output[1]);
            StatementHandle statement = (StatementHandle)DriverHandles.All[handles.Statement];
            delegate* unmanaged<nint, short, short, nint, int, int*, short> descriptor = &Exports.SQLGetDescFieldW;
            int length = 0;
            Assert.Equal(0, descriptor(statement.Descriptors[10010], 0, 25, (nint)output, 0, &length));
            Assert.Equal(0, output[0]);
            Assert.Equal(sizeof(int), length);
            Assert.Equal(0x12345678, output[1]);
        }

        [Theory]
        [InlineData(0, 30)]
        [InlineData(4, 1)]
        [InlineData(6, 0)]
        [InlineData(7, 1)]
        [InlineData(11, 1)]
        public unsafe void StatementIntegersInitializeTheWholeNativeWord(int attribute, int expected) {
            using HandleFamily handles = new();
            nint output = -1;
            delegate* unmanaged<nint, int, nint, int, int*, short> get = &Exports.SQLGetStmtAttrW;
            Assert.Equal(0, get(handles.Statement, attribute, (nint)(&output), 0, null));
            Assert.Equal((nint)expected, output);
        }

        [Fact]
        public unsafe void EnvironmentRemembersTheSelectedOdbcVersion() {
            using HandleFamily handles = new();
            delegate* unmanaged<nint, int, nint, int, short> set = &Exports.SQLSetEnvAttr;
            delegate* unmanaged<nint, int, nint, int, int*, short> get = &Exports.SQLGetEnvAttr;
            int output = 0;
            Assert.Equal(0, set(handles.Environment, 200, 380, 0));
            Assert.Equal(0, get(handles.Environment, 200, (nint)(&output), 0, null));
            Assert.Equal(380, output);
        }

        [Fact]
        public unsafe void UnsupportedConformanceInfoFailsInsteadOfReturningAnUndefinedLevel() {
            using HandleFamily handles = new();
            delegate* unmanaged<nint, ushort, nint, short, short*, short> get = &Exports.SQLGetInfoW;
            uint output = 0;
            Assert.Equal(-1, get(handles.Connection, 152, (nint)(&output), sizeof(uint), null));
            Assert.Equal("HY096", Assert.Single(DriverHandles.All[handles.Connection].Diagnostics).State);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public unsafe void CharacterParametersWithoutIndicatorsStopAtTheTerminator(bool wide) {
            short type = wide ? (short)-8 : (short)1;
            byte[] text = NativeText.Bytes("Poland\0padding", wide);
            fixed (byte* pointer = text) {
                Assert.Equal("Poland", NativeValue.Read(new(type, (nint)pointer, 0, 0, -9, 32, 0)));
                Assert.Equal("Poland", NativeValue.Read(new(type, (nint)pointer, text.Length, 0, -9, 32, 0)));
            }
        }

        [Fact]
        public unsafe void EmptyGetDataBufferReportsTruncationWithoutConsumingTheValue() {
            DataTable table = new();
            table.Columns.Add("Value", typeof(string));
            table.Rows.Add("abc");
            StatementHandle statement = new(0) { Reader = table.CreateDataReader(), OnRow = true };
            using (statement.Reader) {
                Assert.True(statement.Reader.Read());
                char output = 'x';
                nint length = 0;
                Assert.Equal(1, NativeValue.Write(statement, 1, -8, (nint)(&output), 0, (nint)(&length), true));
                Assert.Equal("01004", Assert.Single(statement.Diagnostics).State);
                Assert.Equal(6, (long)length);
                Assert.Equal('x', output);
                char* full = stackalloc char[4];
                Assert.Equal(0, NativeValue.Write(statement, 1, -8, (nint)full, 8, (nint)(&length), true));
                Assert.Equal("abc", new string(full));
            }
        }

        [Fact]
        public async Task ABlockedStatementDoesNotBlockAnIndependentConnection() {
            using HandleFamily first = new();
            using HandleFamily second = new();
            using ManualResetEventSlim entered = new();
            using ManualResetEventSlim release = new();
            Task<short> blocked = Task.Run(() => DriverHandles.With<StatementHandle>(first.Statement, _ => {
                entered.Set();
                if (!release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException();
                return 0;
            }));
            Task<short>? independent = null;
            try {
                Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
                independent = Task.Run(() => DriverHandles.With<StatementHandle>(second.Statement, _ => 0));
                Assert.Equal(0, await independent.WaitAsync(TimeSpan.FromSeconds(5)));
            } finally {
                release.Set();
                Assert.Equal(0, await blocked.WaitAsync(TimeSpan.FromSeconds(5)));
                if (independent is not null) await independent.WaitAsync(TimeSpan.FromSeconds(5));
            }
        }

        private sealed class HandleFamily : IDisposable {
            internal nint Environment { get; }
            internal nint Connection { get; }
            internal nint Statement { get; }
            internal HandleFamily() {
                Assert.Equal(0, DriverHandles.Allocate(1, 0, out nint environment));
                Environment = environment;
                Assert.Equal(0, DriverHandles.Allocate(2, environment, out nint connection));
                Connection = connection;
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
