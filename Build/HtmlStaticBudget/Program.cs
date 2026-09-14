using OfficeIMO.Pdf.Benchmarks.Comparisons;

if (args.Length > 0 && string.Equals(args[0], "html-static-budget-worker", StringComparison.OrdinalIgnoreCase)) {
    return await HtmlStaticBudgetWorker.RunAsync(args);
}

return await HtmlStaticBudgetRunner.RunAsync(args);
