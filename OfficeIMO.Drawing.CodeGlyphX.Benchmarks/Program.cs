using System;
using BenchmarkDotNet.Running;

if (args.Length == 1 && args[0].Equals("--complexity", StringComparison.OrdinalIgnoreCase)) {
    SvgComplexityReporter.Write(Console.Out);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
