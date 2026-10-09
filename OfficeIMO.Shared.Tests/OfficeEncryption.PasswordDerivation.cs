using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class AgilePasswordDerivationTests {
        [Theory]
        [InlineData("agile-password-sha1-aes256.xlsx", "Zażółć-🙂-密碼")]
        [InlineData("agile-password-sha256-aes192.xlsx", "Другой-🔑-päss")]
        public void DecryptPackage_MatchesIndependentUnicodePasswordVectors(string filename, string password) {
            byte[] encrypted = ReadFixture(filename);
            Assert.Equal(ReadFixture("agile-aes256-sha512.plain.xlsx"),
                OfficeEncryption.DecryptPackage(encrypted, password));
            Assert.Throws<CryptographicException>(() =>
                OfficeEncryption.DecryptPackage(encrypted, password + "x"));
        }

        [Fact]
        public async Task DecryptPackage_ConcurrentOperationsKeepPasswordStateSeparate() {
            byte[] expected = ReadFixture("agile-aes256-sha512.plain.xlsx");
            byte[] first = ReadFixture("agile-password-sha1-aes256.xlsx");
            byte[] second = ReadFixture("agile-password-sha256-aes192.xlsx");
            Task[] operations = Enumerable.Range(0, 8).Select(index => Task.Run(() => {
                bool useFirst = (index & 1) == 0;
                byte[] encrypted = useFirst ? first : second;
                string password = useFirst ? "Zażółć-🙂-密碼" : "Другой-🔑-päss";
                Assert.Equal(expected, OfficeEncryption.DecryptPackage(encrypted, password));
                Assert.Throws<CryptographicException>(() =>
                    OfficeEncryption.DecryptPackage(encrypted,
                        useFirst ? "Другой-🔑-päss" : "Zażółć-🙂-密碼"));
            })).ToArray();
            await Task.WhenAll(operations);
        }

        private static byte[] ReadFixture(string name) => File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory, "Documents", "ExcelEncryptionCorpus", name));
    }
}
