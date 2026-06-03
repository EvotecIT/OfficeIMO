using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    private static string Normalize(string value) {
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static int CountOccurrences(string value, string text) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(text, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += text.Length;
        }

        return count;
    }
}
