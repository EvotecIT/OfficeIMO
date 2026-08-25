using BenchmarkDotNet.Running;
using OfficeIMO.GoogleWorkspace.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
    return await GoogleWorkspaceTransportEvidence.RunAsync(args.Skip(1).ToArray()).ConfigureAwait(false);
}
if (args.Length > 0 && string.Equals(args[0], "probe", StringComparison.OrdinalIgnoreCase)) {
    return await GoogleWorkspaceTransportEvidence.RunProbeAsync(args.Skip(1).ToArray()).ConfigureAwait(false);
}
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
