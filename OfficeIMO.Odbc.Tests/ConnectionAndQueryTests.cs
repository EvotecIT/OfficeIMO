using Xunit;

namespace OfficeIMO.Odbc.Tests {

    public sealed class ConnectionAndQueryTests {
        [Fact]
        public void ConnectionAttributesSupportOdbcBracesAndEscapedClosingBraces() {
            string path = OperatingSystem.IsWindows() ? "C:\\data\\a;b}c.xlsx" : "/data/a;b}c.xlsx";
            Assert.Equal(path, ConnectionSettings.WorkbookPath("DRIVER={OfficeIMO};DBQ={" + path.Replace("}", "}}") + "};ReadOnly=1;"));
        }

        [Theory]
        [InlineData("DBQ=relative.xlsx")]
        [InlineData("DBQ={unterminated")]
        [InlineData("DBQ=a;DBQ=b")]
        [InlineData("UID=admin;DBQ=a")]
        [InlineData("DBQ=//server/share/book.xlsx")]
        [InlineData("DBQ=\\/server/share/book.xlsx")]
        [InlineData("DBQ=/\\server/share/book.xlsx")]
        [InlineData("DBQ=\\\\?\\UNC\\server\\share\\book.xlsx")]
        public void InvalidConnectionsFailExplicitly(string connection) =>
            Assert.ThrowsAny<ArgumentException>(() => ConnectionSettings.WorkbookPath(connection));

        [Fact]
        public void SqlMarkersIgnoreQuotedTextIdentifiersAndComments() {
            string sql = "/* ? */ SELECT '?' AS [a?], \"b?\" FROM Customers WHERE Country=? AND Name=?; -- ?";
            Assert.Equal("/* ? */ SELECT '?' AS [a?], \"b?\" FROM Customers WHERE Country=@p1 AND Name=@p2; -- ?", QueryEnvelope.Prepare(sql, out int count));
            Assert.Equal(2, count);
        }

        [Theory]
        [InlineData("SELECT 'a;b'; /* valid */")]
        [InlineData("-- leading comment\n SELECT \"semi;colon\" FROM Customers")]
        public void StatementEnvelopeAllowsSqlLiteralsAndComments(string sql) => QueryEnvelope.Validate(sql);
    }
}
