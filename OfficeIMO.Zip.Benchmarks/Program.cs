using System.Text.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Zip.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    IReadOnlyList<ZipComparisonReport> reports = ZipComparisonValidation.ValidateAll();
    string? outputPath = ReadOption(args, "--json");
    if (!string.IsNullOrWhiteSpace(outputPath)) {
        string fullPath = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllText(fullPath, JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }

    foreach (ZipComparisonReport report in reports) {
        Console.WriteLine(
            $"{report.Scale,-6} {report.EntryCount,5:N0} entries | " +
            $"{report.InputBytes,10:N0} ZIP bytes | {report.TotalUncompressedBytes,10:N0} uncompressed bytes | " +
            report.StructuralFingerprint);
    }
    return 0;
}

if (args.Length > 0 && string.Equals(args[0], "--probe", StringComparison.OrdinalIgnoreCase)) {
    return ZipEvidenceRunner.RunProbe(args.Skip(1).ToArray());
}

if (args.Any(argument => string.Equals(argument, "--evidence", StringComparison.OrdinalIgnoreCase))) {
    return ZipEvidenceRunner.RunEvidence(args);
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;

static string? ReadOption(string[] values, string name) {
    for (int index = 0; index < values.Length - 1; index++) {
        if (string.Equals(values[index], name, StringComparison.OrdinalIgnoreCase)) return values[index + 1];
    }
    return null;
}
