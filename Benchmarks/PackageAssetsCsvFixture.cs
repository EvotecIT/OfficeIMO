using System.Text;

namespace OfficeIMO.Benchmarks;

/// <summary>Column decoration used by the NCsvPerf PackageAssets shapes that Sep publishes.</summary>
public enum PackageAssetsCsvShape {
    /// <summary>The original unquoted rows.</summary>
    Plain,
    /// <summary>Every column except 0, 1 and 4 is wrapped in quotes.</summary>
    Quotes,
    /// <summary>The quoted columns also carry one space outside and inside each quote.</summary>
    SpacesQuotes
}

/// <summary>
/// Hash-pinned NCsvPerf PackageAssets rows expanded to the 50,000-line inputs used by Sep's
/// comparison benchmarks. Benchmark projects link this source file; it is not a runtime package.
/// </summary>
internal static class PackageAssetsCsvFixture {
    internal const string SourceCommit = "c7bff62e601eb1e2ce97c95a5b54ef94eff4b6ab";
    internal const string FileName = "PackageAssets.csv";
    internal const string Hash = "5344E99AB70D3D68EDCF41F3F787E4EF330EEDAE5A84CDB65144DBA17485503D";
    internal const int Lines = 50_000;
    internal const int Columns = 25;

    private static readonly Uri Source = new(
        "https://raw.githubusercontent.com/joelverhagen/NCsvPerf/" + SourceCommit + "/NCsvPerf/TestData/" + FileName);

    /// <summary>Downloads or verifies the source rows and returns the UTF-8 input without a BOM.</summary>
    internal static byte[] CreateInput(PackageAssetsCsvShape shape) {
        string[] source = ReadSourceLines();
        var text = new StringBuilder(Lines * 300);
        for (int line = 0; line < Lines; line++) {
            string row = SourceRow(source, line);
            string[] values = row.Split(',');
            for (int column = 0; column < values.Length; column++) {
                if (column > 0) text.Append(',');
                if (shape == PackageAssetsCsvShape.Plain || column is 0 or 1 or 4) {
                    text.Append(values[column]);
                } else if (shape == PackageAssetsCsvShape.Quotes) {
                    text.Append('"').Append(values[column]).Append('"');
                } else {
                    text.Append(" \" ").Append(values[column]).Append(" \" ");
                }
            }
            text.Append("\r\n");
        }
        return new UTF8Encoding(false).GetBytes(text.ToString());
    }

    /// <summary>
    /// Returns the decoded rows every headerless reader must produce. Spaces-and-quotes input is read with
    /// whitespace outside quotes trimmed and quotes removed, so its quoted values keep the inner spaces.
    /// </summary>
    internal static IEnumerable<string[]> ExpectedRows(PackageAssetsCsvShape shape) {
        string[] source = ReadSourceLines();
        for (int line = 0; line < Lines; line++) {
            string[] values = SourceRow(source, line).Split(',');
            if (values.Length != Columns) throw new InvalidDataException($"PackageAssets row has {values.Length} columns.");
            if (shape == PackageAssetsCsvShape.SpacesQuotes) {
                for (int column = 0; column < values.Length; column++) {
                    if (column is not (0 or 1 or 4)) values[column] = " " + values[column] + " ";
                }
            }
            yield return values;
        }
    }

    // Sep's line cache keeps line 0 and cycles the remaining source rows. The file has no header,
    // so every reader treats all lines as data.
    private static string SourceRow(string[] source, int line) =>
        source[line == 0 ? 0 : (line - 1) % (source.Length - 1) + 1];

    private static string[] ReadSourceLines() {
        MarkPflug65KFixture.EnsureAuthentic(FileName, Hash, Source);
        return File.ReadAllLines(Path.Combine(MarkPflug65KFixture.Root, FileName));
    }
}
