using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Minimal, zero-dependency text extractor for simple PDFs produced by OfficeIMO.Pdf
/// and common external PDFs with basic text operators and common content-stream filters.
/// Not a general-purpose PDF parser; designed as a pragmatic starting point.
/// </summary>
public static partial class PdfTextExtractor {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
    private static readonly char[] SpaceSplitChars = new[] { ' ' };
    private static readonly char[] CsvQuoteChars = new[] { ',', '"', '\r', '\n' };
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.*?)/Type\s*/Page\b(?:.*?)/Contents\s+(?:(?<single>\d+)\s+0\s+R|\[(?<array>[^\]]*)\])(?:.*?)/?>>", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex RefRegex = new Regex(@"(\d+)\s+0\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex HexTjRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*Tj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex QuoteLiteralRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*'", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex QuoteHexRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*'", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex DoubleQuoteLiteralRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+\((?<txt>(?:\\.|[^\\\)])*)\)\s*""", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex DoubleQuoteHexRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+<(?<txt>[0-9A-Fa-f\s]+)>\s*""", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+0\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex InfoRefRegex = new Regex(@"/Info\s+(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex PageObjRegex = new Regex(@"<<(?:.|\n|\r)*?/Type\s*/Page\b(?:.|\n|\r)*?/Contents\s+(?:(?<single>\d+)\s+0\s+R|\[(?<array>[^\]]*)\])(?:.|\n|\r)*?>>", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex RefRegex = new Regex(@"(\d+)\s+0\s+R", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex TjRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*Tj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex HexTjRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*Tj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex QuoteLiteralRegex = new Regex(@"\((?<txt>(?:\\.|[^\\\)])*)\)\s*'", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex QuoteHexRegex = new Regex(@"<(?<txt>[0-9A-Fa-f\s]+)>\s*'", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex DoubleQuoteLiteralRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+\((?<txt>(?:\\.|[^\\\)])*)\)\s*""", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex DoubleQuoteHexRegex = new Regex(@"(?<ws>[+-]?\d*\.?\d+)\s+(?<cs>[+-]?\d*\.?\d+)\s+<(?<txt>[0-9A-Fa-f\s]+)>\s*""", RegexOptions.Compiled, RegexTimeout);
#endif
}
