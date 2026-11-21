using BenchmarkDotNet.Running;
using OfficeIMO.Toml.Benchmarks;
using OfficeIMO.Toml.Benchmarks.LargeCorpus;

if (args.Length > 0) {
    string[] rest = args.Skip(1).ToArray();
    switch (args[0].ToLowerInvariant()) {
        case "--validate-workloads":
            Environment.ExitCode = TomlWorkloadCheck.Run();
            return;
        case "--fetch-corpus":
            Environment.ExitCode = TomlLargeCorpus.Fetch(rest);
            return;
        case "--refresh-corpus":
            Environment.ExitCode = TomlLargeCorpus.Refresh(rest);
            return;
        case "--large-corpus":
            Environment.ExitCode = TomlLargeCorpusRun.Run(rest);
            return;
        case "--mutations":
            Environment.ExitCode = TomlMutationRun.Run(rest);
            return;
    }
}

BenchmarkSwitcher.FromAssembly(typeof(TomlParseBenchmarks).Assembly).Run(args);
