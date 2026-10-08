using System.Security.Cryptography;
using OfficeIMO.AI;

// Reuse the unchanged, licensed KoSIT producer fixtures already maintained by Invoicing.
// XML is supplied as literal text: this lane measures field interpretation, not scanned-invoice OCR.
internal static class EvaluationIndependentInvoiceCorpus {
    public static IReadOnlyList<EvaluationCase> Create() => new[] {
        Create("01.01a", "74fb09c609d5fba15a8c543060998d3b92858f56a81fb5b0ed244d6794e498d1",
            "123456XX", "2016-04-04", "336.9", "314.86"),
        Create("01.05a", "09cff10c3ad2a6934e4825b1f5e1ab595410443ff2f150d3cda634d2f0eedc6c",
            "PRG1502112", "2015-04-24", "10555.3", "8870")
    };

    private static EvaluationCase Create(string id, string hash, string reference, string issued, string payable, string beforeTax) {
        using Stream stream = typeof(EvaluationIndependentInvoiceCorpus).Assembly.GetManifestResourceStream("EvaluationInvoices." + id + ".xml")
            ?? throw new InvalidOperationException("Independent invoice fixture is missing.");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        byte[] bytes = buffer.ToArray();
        if (!string.Equals(Convert.ToHexString(SHA256.HashData(bytes)), hash, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Independent invoice fixture differs from its labelled source.");
        return new("kosit-invoice-" + id, ".txt", bytes, new OfficeAiRequest {
            Operation = OfficeAiOperation.ExtractFields,
            Instruction = "Extract the invoice identifier, issue date, document currency, payable amount and tax-exclusive total. Use invoice-level values, not line-item amounts, order references or supplier identifiers.",
            Fields = [new("invoiceId"), new("issueDate", OfficeAiFieldType.Date, "yyyy-MM-dd"), new("currency"),
                new("payable", OfficeAiFieldType.Decimal), new("taxExclusive", OfficeAiFieldType.Decimal)]
        }, false, "Independent KoSIT XRechnung invoice XML; five declared field labels; no scan or OCR accuracy claim.",
            new(Fields: [Present("invoiceId", reference), Present("issueDate", issued), Present("currency", "EUR"),
                Present("payable", payable), Present("taxExclusive", beforeTax)])) { Split = "independent-producer" };
    }

    private static EvaluationFieldGold Present(string name, string value) => new(name, OfficeAiFieldStatus.Present, value);
}
