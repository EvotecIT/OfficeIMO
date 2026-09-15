using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using OfficeIMO.Invoicing.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--verify-budgets", StringComparison.OrdinalIgnoreCase)) {
    return await InvoicePerformanceBudgetRunner.RunAsync(args.Skip(1).ToArray()).ConfigureAwait(false);
}

ManualConfig config = ManualConfig.Create(DefaultConfig.Instance)
    .AddDiagnoser(MemoryDiagnoser.Default)
    .AddExporter(JsonExporter.Full)
    .WithSummaryStyle(SummaryStyle.Default.WithRatioStyle(RatioStyle.Percentage))
    .AddColumn(StatisticColumn.OperationsPerSecond);

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
return 0;
