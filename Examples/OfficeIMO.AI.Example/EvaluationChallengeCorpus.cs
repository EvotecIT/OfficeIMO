using System.Text;
using OfficeIMO.AI;
using OfficeIMO.Pdf;

internal static class EvaluationChallengeCorpus {
    public static IReadOnlyList<EvaluationCase> Create() {
        var fields = new[] { new OfficeAiFieldDefinition("reference"), new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal),
            new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, "dd.MM.yyyy") };
        var request = new OfficeAiRequest { Operation = OfficeAiOperation.ExtractFields, Culture = "pl-PL", Fields = fields,
            Instruction = "Extract the credit reference, refund amount, and date. Do not infer missing values." };
        var cases = new List<EvaluationCase> {
            Text("refund-pl-development", "Credit reference: KR-204. Refund: -1 234,50 PLN. Date: 29.02.2028.", request,
                new(Fields: new[] { Present("reference", "KR-204"), Present("amount", "-1234.50"), Present("date", "2028-02-29") }), "development"),
            Text("refund-pl-heldout", "Korekta KOR/731/B. Data wystawienia 07.11.2029. Kwota zwrotu: -8 076,09 zł.", request,
                new(Fields: new[] { Present("reference", "KOR/731/B"), Present("amount", "-8076.09"), Present("date", "2029-11-07") })),
            Text("refund-en-heldout", "Credit memo CM-9Z. Refund amount USD -2,031.08. Issued 2032-02-29.", request with {
                Culture = "en-US", Fields = new[] { fields[0], fields[1], new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, "yyyy-MM-dd") }
            }, new(Fields: new[] { Present("reference", "CM-9Z"), Present("amount", "-2031.08"), Present("date", "2032-02-29") })),
            Text("missing-date-heldout", "Credit reference CM-03. Refund 19.25 USD. The issue date is not recorded.", request with { Culture = "en-US" },
                new(Fields: new[] { Present("reference", "CM-03"), Present("amount", "19.25"), new EvaluationFieldGold("date", OfficeAiFieldStatus.Missing) })),
            Text("ambiguous-date-heldout", "Credit reference CM-04. Refund 81.90 USD. Date 03/04/2030; the source explicitly leaves day/month order unspecified.",
                request with { Culture = "en-US", Fields = new[] { fields[0], fields[1], new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, "dd/MM/yyyy") } },
                new(Fields: new[] { Present("reference", "CM-04"), Present("amount", "81.90"), new EvaluationFieldGold("date", OfficeAiFieldStatus.Ambiguous) })),
            Text("conflicting-pages-heldout", "Warehouse A reports quantity 17. Warehouse B reports quantity 23 for the same item and date. Neither report is authoritative.",
                new() { Operation = OfficeAiOperation.ExtractFields, Instruction = "Extract the quantity; preserve conflicting observations.",
                    Fields = new[] { new OfficeAiFieldDefinition("quantity", OfficeAiFieldType.Integer) } },
                new(Fields: new[] { new EvaluationFieldGold("quantity", OfficeAiFieldStatus.Conflicting) }))
        };
        string[][] table = { new[] { "Code", "Count", "Amount" }, new[] { "AA-7", "0", "-3.25" },
            new[] { "BB-2", "19", "1024.80" }, new[] { "CC-8", "2", "0.00" } };
        var fontOptions = new PdfOptions { DefaultFontSize = 14 };
        if (!fontOptions.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) throw new InvalidOperationException("Evaluation requires an embeddable font.");
        byte[] pdf = PdfDocument.Create(builder => builder.Content(content => content.H1("Reconciliation").Table(table)),
            fontOptions).ToBytes();
        var tableGold = new EvaluationGold(Table: new(table[0], table.Skip(1).Select(row => (IReadOnlyList<string>)row).ToArray()));
        var parse = new OfficeAiRequest { Operation = OfficeAiOperation.Parse, Instruction = "Transcribe the reconciliation heading and every table cell exactly." };
        cases.Add(new("table-heldout", ".pdf", pdf, parse, false, "3 headers and 9 position-sensitive cells", tableGold) { Split = "heldout" });
        cases.Add(new("table-vision-heldout", ".pdf", pdf, parse, true, "3 headers and 9 position-sensitive cells", tableGold) { Split = "heldout" });
        // Distinct facts separated by long irrelevant text exercise batching and final synthesis.
        string padding = string.Concat(Enumerable.Repeat("This section contains routine handling guidance without additional quantities or deadlines. ", 400));
        string longText = "Northern depot: 17 crates. " + padding + "\n\nSouthern depot: 23 crates. " + padding
            + "\n\nCombined dispatch date: 2031-11-19. " + padding;
        cases.Add(Text("long-summary-heldout", longText, new() { Operation = OfficeAiOperation.Summarize,
            Instruction = "Summarize the two depot quantities and the combined dispatch date. Exclude generic handling guidance.",
            Limits = new() { MaxRequestCharacters = 18000, MaxRequests = 40, Timeout = TimeSpan.FromMinutes(8) }
        }, new(FactMarkers: new[] { "17", "23", "2031-11-19" }, RequireSynthesis: true)));
        return cases;
    }

    private static EvaluationFieldGold Present(string name, string value) => new(name, OfficeAiFieldStatus.Present, value);
    private static EvaluationCase Text(string id, string text, OfficeAiRequest request, EvaluationGold gold, string split = "heldout") =>
        new(id, ".txt", Encoding.UTF8.GetBytes(text), request, false, "Independent typed gold; semantic claims require review", gold) { Split = split };
}
