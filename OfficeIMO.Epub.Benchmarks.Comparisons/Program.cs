using System.Text.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Epub.Benchmarks.Comparisons;

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    IReadOnlyList<EpubComparisonReport> reports = EpubComparisonValidation.ValidateAll();
    string? outputPath = ReadOption(args, "--json");
    if (!string.IsNullOrWhiteSpace(outputPath)) {
        string fullPath = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllText(fullPath, JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }

    foreach (EpubComparisonReport report in reports) {
        Console.WriteLine(
            $"{report.Scale,-6} input {report.InputBytes,10:N0} bytes | " +
            $"chapters {report.OfficeIMO.ChapterCount,4:N0} | " +
            $"XHTML {report.OfficeIMO.ContentCharacters,10:N0} chars | " +
            $"text {report.OfficeIMO.TextCharacters,10:N0} chars | hashes match");
    }

    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);

static string? ReadOption(string[] values, string name) {
    for (int index = 0; index < values.Length - 1; index++) {
        if (string.Equals(values[index], name, StringComparison.OrdinalIgnoreCase)) return values[index + 1];
    }

    return null;
}
