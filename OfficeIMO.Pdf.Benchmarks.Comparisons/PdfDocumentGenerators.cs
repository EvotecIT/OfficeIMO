namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfDocumentGenerators {
    internal static byte[] Generate(PdfBenchmarkProducer producer, PdfBenchmarkScenario scenario) => producer switch {
        PdfBenchmarkProducer.OfficeIMO => OfficeImoPdfGenerator.Generate(scenario),
        PdfBenchmarkProducer.QuestPDF => QuestPdfGenerator.Generate(scenario),
        PdfBenchmarkProducer.PeachPDF => PeachPdfGenerator.Generate(PdfHtmlScenarioBuilder.Create(scenario)),
        PdfBenchmarkProducer.MigraDoc => MigraDocPdfGenerator.Generate(scenario),
        PdfBenchmarkProducer.IText => ITextPdfGenerator.Generate(scenario),
        _ => throw new ArgumentOutOfRangeException(nameof(producer))
    };
}
