using OfficeIMO.Pdf.Benchmarks.Comparisons;

if (args.Length > 0 && string.Equals(args[0], "html-static-budget-worker", StringComparison.OrdinalIgnoreCase)) {
    return await HtmlStaticBudgetWorker.RunAsync(args);
}

string[] runnerArgs = new[] { "html-static-budget" }.Concat(args).ToArray();
return await HtmlStaticBudgetRunner.RunAsync(runnerArgs);
