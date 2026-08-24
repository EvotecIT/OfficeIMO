using System.Text.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Email.Benchmarks.Comparisons;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = EmailMimeEvidenceRunner.RunProbe(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = EmailMimeEvidenceRunner.RunEvidence(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    IReadOnlyList<EmailMimeComparisonReport> reports = EmailMimeComparisonValidation.ValidateAll();
    string? outputPath = ReadOption(args, "--json");
    if (!string.IsNullOrWhiteSpace(outputPath)) {
        string fullPath = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllText(fullPath, JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }

    foreach (EmailMimeComparisonReport report in reports) {
        Console.WriteLine(
            $"{report.Scale,-8} input {report.InputBytes,10:N0} bytes | " +
            $"OfficeIMO {report.OfficeIMOOutputBytes,10:N0} bytes | " +
            $"MimeKit {report.MimeKitOutputBytes,10:N0} bytes | " +
            $"attachments {report.AttachmentCount,2:N0}/{report.DecodedAttachmentBytes,10:N0} bytes");
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
