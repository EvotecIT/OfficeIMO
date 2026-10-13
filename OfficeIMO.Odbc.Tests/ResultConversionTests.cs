using System.Data;
using Xunit;

namespace OfficeIMO.Odbc.Tests {
    public sealed class ResultConversionTests {
        [Theory]
        [InlineData(7, "1,2", "22018")]
        [InlineData(8, "1,2", "22018")]
        [InlineData(7, "NaN", "22018")]
        [InlineData(8, "Infinity", "22018")]
        [InlineData(7, "1e999", "22003")]
        [InlineData(8, "1e999", "22003")]
        [InlineData(7, "1e-999", "22003")]
        [InlineData(8, "1e-999", "22003")]
        public unsafe void FloatingConversionsRejectInvalidOrOutOfRangeLiterals(short target, string text, string state) {
            StatementHandle statement = Row(text);
            using (statement.Reader) {
                long output = 0x12345678;
                Assert.Equal(-1, NativeValue.Write(statement, 1, target, (nint)(&output), 8, 0, true));
                Assert.Equal(state, Assert.Single(statement.Diagnostics).State);
                Assert.Equal(0x12345678, output);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public unsafe void BinaryTextChunksPreserveWholeHexPairs(bool wide) {
            StatementHandle statement = Row(new byte[] { 0, 0xAB, 0xFF });
            using (statement.Reader) {
                byte* output = stackalloc byte[8];
                nint length = 0;
                foreach (string expected in new[] { "00", "AB", "FF" }) {
                    short status = NativeValue.Write(statement, 1, wide ? (short)-8 : (short)1,
                        (nint)output, wide ? 8 : 4, (nint)(&length), true);
                    Assert.Equal(expected == "FF" ? 0 : 1, status);
                    Assert.Equal(expected, NativeText.Read((nint)output, -3, wide));
                }
                Assert.Equal(100, NativeValue.Write(statement, 1, wide ? (short)-8 : (short)1, (nint)output, 8, 0, true));
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public unsafe void NumericTextBuffersMustPreserveWholeDigits(bool wide) {
            byte* output = stackalloc byte[8];
            StatementHandle statement = Row(12345L);
            using (statement.Reader) {
                output[0] = 0xCC;
                Assert.Equal(-1, NativeValue.Write(statement, 1, wide ? (short)-8 : (short)1, (nint)output, wide ? 6 : 3, 0, true));
                Assert.Equal("22003", Assert.Single(statement.Diagnostics).State);
                Assert.Equal(0xCC, output[0]);
            }
            statement = Row(123.456);
            using (statement.Reader) {
                Assert.Equal(1, NativeValue.Write(statement, 1, wide ? (short)-8 : (short)1, (nint)output, wide ? 8 : 4, 0, true));
                Assert.Equal("01004", Assert.Single(statement.Diagnostics).State);
                Assert.Equal("123", NativeText.Read((nint)output, -3, wide));
            }
            statement = Row(true);
            using (statement.Reader) {
                Assert.Equal(-1, NativeValue.Write(statement, 1, wide ? (short)-8 : (short)1, (nint)output, wide ? 2 : 1, 0, true));
                Assert.Equal("22003", Assert.Single(statement.Diagnostics).State);
            }
        }

        [Fact]
        public unsafe void FloatingParameterValuesMustBeFinite() {
            double value = double.PositiveInfinity;
            ParameterBinding binding = new(8, (nint)(&value), 8, 0, 8, 15, 0);
            ParameterValueException error = Assert.Throws<ParameterValueException>(() => NativeValue.Read(binding));
            Assert.Equal("22003", error.State);
        }

        [Fact]
        public unsafe void NonBinaryResultsDoNotInventABinaryRepresentation() {
            foreach (object value in new object[] { 123L, "abc" }) {
                StatementHandle statement = Row(value);
                using (statement.Reader) {
                    long output = 0x12345678;
                    Assert.Equal(-1, NativeValue.Write(statement, 1, -2, (nint)(&output), 8, 0, true));
                    Assert.Equal("07006", Assert.Single(statement.Diagnostics).State);
                    Assert.Equal(0x12345678, output);
                }
            }
        }

        private static StatementHandle Row(object value) {
            DataTable table = new();
            table.Columns.Add("Value", value.GetType());
            table.Rows.Add(new object[] { value });
            StatementHandle statement = new(0) { Reader = table.CreateDataReader(), OnRow = true };
            Assert.True(statement.Reader.Read());
            return statement;
        }
    }
}
