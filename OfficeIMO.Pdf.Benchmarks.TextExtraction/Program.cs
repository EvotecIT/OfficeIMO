using OfficeIMO.Pdf.Benchmarks.TextExtraction;

if (args.Length == 0) {
    Console.Error.WriteLine("""
        Usage:
          run --suite name=directory [--suite ...] --output directory [--mode warm|cold] [--engines a,b]
              [--timeout 60] [--shard i/n] [--write-text]
          parity --output directory [--pairs left:right,...]
          summary --output directory [--pattern results-*.csv]
          worker --engine name      (internal: one engine child process)
        Engines: officeimo, officeimo-read, pdfoxide, pdfium, pdfpig
        """);
    return 2;
}

return args[0] switch {
    "run" => CorpusRunner.Run(args[1..]),
    "parity" => TextParity.Run(args[1..]),
    "summary" => ResultSummary.Run(args[1..]),
    "worker" => ExtractionWorker.Run(Options.Get(args, "--engine") ?? throw new ArgumentException("--engine is required.")),
    _ => throw new ArgumentException($"Unknown command '{args[0]}'.")
};
