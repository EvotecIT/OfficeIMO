using System.Globalization;
using OfficeIMO.GoogleWorkspace;
using Xunit;

namespace OfficeIMO.GoogleWorkspace.Tests {
    public class GoogleWorkspaceCheckpointFormatTests {
        [Fact]
        public void FormatsHighPrecisionNumbersIdenticallyAcrossCulturesAndRuntimes() {
            CultureInfo original = CultureInfo.CurrentCulture;
            try {
                FormattableString fingerprint = $"{1.2345678901234567d}|{1.2345678f}|{1234.50m}";
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                string french = GoogleWorkspaceCheckpointFormat.Format(fingerprint);
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
                Assert.Equal("1.2345678901234567|1.23456776|1234.5", french);
                Assert.Equal(french, GoogleWorkspaceCheckpointFormat.Format(fingerprint));
            } finally {
                CultureInfo.CurrentCulture = original;
            }
        }
    }
}
