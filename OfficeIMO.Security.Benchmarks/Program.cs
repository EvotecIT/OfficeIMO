using BenchmarkDotNet.Running;
using OfficeIMO.Security.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = SecurityCmsEvidenceRunner.RunProbe(args[1..]);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = SecurityCmsEvidenceRunner.Run(args[1..]);
    return;
}

if (args.Length == 1 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    foreach (string scale in SecurityCmsBenchmarkCorpus.Scales) {
        using SecurityCmsBenchmarkFixture fixture = SecurityCmsBenchmarkCorpus.Create(scale);
        SecurityCmsValidationSnapshot snapshot = SecurityCmsBenchmarkValidation.Validate(fixture);
        Console.WriteLine(
            $"{scale,-6} content {snapshot.ContentBytes,10:N0} bytes | " +
            $"OfficeIMO CMS {snapshot.OfficeSignatureBytes,6:N0} bytes | " +
            $"platform CMS {snapshot.PlatformSignatureBytes,6:N0} bytes | " +
            $"ratio {(double)snapshot.OfficeSignatureBytes / snapshot.PlatformSignatureBytes:F3}x");
    }
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
