using System.Text.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Rtf.Benchmarks.Comparisons;

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    IReadOnlyList<RtfHtmlComparisonReport> reports = RtfHtmlComparisonValidation.ValidateAll();
    string? outputPath = ReadOption(args, "--json");
    if (!string.IsNullOrWhiteSpace(outputPath)) {
        string fullPath = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllText(fullPath, JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }

    foreach (RtfHtmlComparisonReport report in reports) {
        Console.WriteLine(
            $"{report.Scale,-8} input {report.InputBytes,10:N0} bytes | " +
            $"OfficeIMO {report.OfficeIMO.OutputBytes,10:N0} bytes | " +
            $"RtfPipe {report.RtfPipe.OutputBytes,10:N0} bytes | " +
            $"records {report.OfficeIMO.RecordCount,5:N0}");
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
