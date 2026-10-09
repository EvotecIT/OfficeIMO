using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void PooledUtf8TextWriter_PreservesEncoderStateBeforeAscii(bool longInput, bool completePair) {
            string prefix = longInput ? new string('a', 4096) : "prefix ";
            string continuation = completePair ? "\ude80 ASCII" : "ASCII";
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), 16, leaveOpen: true)) {
                writer.Write(prefix + "\ud83d");
                writer.Flush();
                writer.Write(continuation);
                writer.Flush();
                writer.Write(" Ελληνικά ");
                writer.Write(new string('z', 50_000));
            }
            byte[] expected = Encoding.UTF8.GetBytes(prefix + "\ud83d" + continuation + " Ελληνικά " + new string('z', 50_000));
            Assert.Equal(expected, output.ToArray());
        }
    }
}
