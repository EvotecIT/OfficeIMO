using BenchmarkDotNet.Running;

namespace OfficeIMO.Data.Conversions.Benchmarks;

public static class Program {
    public static void Main(string[] args) {
        if (args.Length > 0 && args[0] == "--verify") {
            Environment.ExitCode = IndependentReaders.Run();
            return;
        }

        if (args.Length > 0 && args[0] == "--profile") {
            Profile.Run();
            return;
        }

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}
