using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using OfficeIMO;
using OfficeIMO.Excel;

if (args is ["--startup", var workload]) {
    LegacyXlsColdStart.Measure(workload);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(LegacyXlsProjectionBenchmarks).Assembly).Run(args);

[MemoryDiagnoser]
public class LegacyXlsProjectionBenchmarks {
    private const string EditedText = "Edited in place";
    private byte[] _binary = Array.Empty<byte>();
    private byte[] _openXml = Array.Empty<byte>();
    private readonly ExcelSaveOptions _xls = new() { Format = ExcelFileFormat.Xls };

    [Params("calculation", "workbook options")]
    public string Contents { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT")
            ?? throw new InvalidOperationException("Point OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT to the master capture.");
        string testName = Contents == "calculation"
            ? "LegacyXls_NativeSave_WritesCalculationSettings"
            : "LegacyXls_NativeSave_WritesWorkbookOptions";
        string folder = Path.Combine(root, "test-files", testName);
        _binary = File.ReadAllBytes(Path.Combine(folder, "001.xls"));
        _openXml = File.ReadAllBytes(Path.Combine(folder, "001.projected.xlsx"));

        string expected = Contents == "calculation" ? "Calculation settings" : "Workbook options";
        CheckText(_binary, expected);
        byte[] output = ConvertToXls();
        CheckText(output, expected);
        byte[] edited = EditXls();
        CheckText(edited, EditedText);
        Console.WriteLine($"Input XLS {Convert.ToHexString(SHA256.HashData(_binary))}");
        Console.WriteLine($"Input XLSX {Convert.ToHexString(SHA256.HashData(_openXml))}");
        Console.WriteLine($"Output XLS {Convert.ToHexString(SHA256.HashData(output))}");
        Console.WriteLine($"Edited XLS {Convert.ToHexString(SHA256.HashData(edited))}");
    }

    [Benchmark]
    public int ImportXls() {
        using var stream = new MemoryStream(_binary, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream);
        return document.Sheets.Count;
    }

    [Benchmark]
    public byte[] ConvertToXls() {
        // A fresh workbook per operation keeps the save on the XLSX-to-XLS path, before XLS adoption.
        using var stream = new MemoryStream(_openXml, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream);
        return document.ToBytes(_xls);
    }

    [Benchmark]
    public byte[] EditXls() {
        using var stream = new MemoryStream(_binary, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream);
        document.Sheets[0].Cell(1, 1).SetValue(EditedText);
        return document.ToBytes(_xls);
    }

    private static void CheckText(byte[] binary, string expected) {
        using var stream = new MemoryStream(binary, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream);
        if (document.SourceFormat != ExcelFileFormat.Xls || document.Sheets.Count != 1
            || !document.Sheets[0].TryGetCellText(1, 1, out string? text) || text != expected) {
            throw new InvalidDataException($"The benchmark workbook must contain '{expected}' in A1.");
        }
    }
}

[MemoryDiagnoser]
public class LegacyXlsStreamingBenchmarks {
    // Rotate distinct themes to measure style conversion on a cache miss.
    private const string CustomBodyFont = "Courier New";
    private static readonly OfficeTheme CustomTheme = OfficeTheme.Modern
        .WithFonts(OfficeTheme.Modern.Fonts.With(body: OfficeTheme.Modern.Fonts.Body.With(family: CustomBodyFont)))
        .WithPalette(OfficeTheme.Modern.Palette.With(text: OfficeColor.FromRgb(0x12, 0x45, 0x78)));
    private static readonly OfficeTheme[] Themes = Enumerable.Range(0, 16)
        .Select(index => CustomTheme.WithName("XLS streaming " + index)).ToArray();
    private static readonly int[] Data = Enumerable.Range(1, 16).ToArray();
    private static readonly string[] Columns = { "Label", "Value" };
    private int _nextTheme;

    [GlobalSetup]
    public void Setup() {
        byte[] output = WriteWithNewTheme();
        WriteWithCachedTheme();
        using var stream = new MemoryStream(output, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream);
        if (document.SourceFormat != ExcelFileFormat.Xls || document.Sheets.Count != 1
            || !document.Sheets[0].TryGetCellText(2, 1, out string? label) || label != "Sample") {
            throw new InvalidDataException("The streaming workload did not write its rows.");
        }
        if (document.Sheets[0].Cell(2, 1).GetStyle().FontName != CustomBodyFont) {
            throw new InvalidDataException("The streaming workbook lost its custom body font.");
        }
        Console.WriteLine($"Streaming XLS {Convert.ToHexString(SHA256.HashData(output))}");
    }

    [Benchmark]
    public byte[] WriteWithNewTheme() {
        OfficeTheme theme = Themes[_nextTheme];
        _nextTheme = (_nextTheme + 1) % Themes.Length;
        return Write(theme);
    }

    [Benchmark]
    public byte[] WriteWithCachedTheme() => Write(OfficeTheme.Modern);

    private static byte[] Write(OfficeTheme theme) {
        using var output = new MemoryStream();
        using (ExcelWorkbookWriter writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, output, theme: theme, leaveOpen: true)) {
            writer.WriteTable("Data", Data, Columns, static (row, value) => {
                row.Write("Sample");
                row.Write(value);
            });
        }
        return output.ToArray();
    }
}

internal static class LegacyXlsColdStart {
    internal static void Measure(string workload) {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT")
            ?? throw new InvalidOperationException("Point OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT to the master capture.");
        byte[] input = workload switch {
            "import" => File.ReadAllBytes(Path.Combine(root, "test-files", "LegacyXls_NativeSave_WritesCalculationSettings", "001.xls")),
            "plain-import" => LegacyXlsPlainImportBenchmarks.LoadFixture(root),
            _ => Array.Empty<byte>()
        };
        long allocated = GC.GetAllocatedBytesForCurrentThread();
        long started = Stopwatch.GetTimestamp();
        int result = workload switch {
            "import" => Import(input),
            "plain-import" => LegacyXlsPlainImportBenchmarks.ReadPlainCells(input),
            "stream" => Stream(),
            _ => throw new ArgumentOutOfRangeException(nameof(workload))
        };
        TimeSpan elapsed = Stopwatch.GetElapsedTime(started);
        allocated = GC.GetAllocatedBytesForCurrentThread() - allocated;
        if (result <= 0) throw new InvalidDataException("The first call produced no workbook output.");
        string loadedSdk = string.Join(",", AppDomain.CurrentDomain.GetAssemblies()
            .Select(assembly => assembly.GetName().Name)
            .Where(name => name?.StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal) == true)
            .OrderBy(name => name, StringComparer.Ordinal));
        Console.WriteLine(FormattableString.Invariant($"startup,{workload},{elapsed.TotalMilliseconds:F3},{allocated},{loadedSdk}"));
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static int Import(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream);
        return document.Sheets.Count;
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static int Stream() => new LegacyXlsStreamingBenchmarks().WriteWithNewTheme().Length;
}
