using System.IO;
using System.Runtime.CompilerServices;
using OfficeIMO.Core.Internal;
using OfficeIMO.Invoicing;

namespace PSWriteOffice.Services.Invoicing;

/// <summary>Retains source-loss evidence while exposing the editable invoice as ordinary pipeline data.</summary>
internal static class InvoiceCommandUtilities {
    private static readonly ConditionalWeakTable<Invoice, InvoiceReadResult> Sources = new();

    internal static InvoiceReadResult Read(string path) {
        using var input = File.OpenRead(path);
        var result = InvoiceParser.Read(OfficeStreamReader.ReadAllBytes(input, InvoiceProfileDeclaration.MaximumXmlBytes));
        Sources.Add(result.Invoice, result);
        return result;
    }

    internal static byte[] Write(Invoice invoice, InvoiceXmlOptions target, bool allowUnmappedLoss) =>
        Sources.TryGetValue(invoice, out var source)
            ? source.Write(target, allowUnmappedLoss)
            : InvoiceSerializer.Write(invoice, target);

    internal static InvoiceModelValidationResult Validate(Invoice invoice) =>
        Sources.TryGetValue(invoice, out var source)
            ? InvoiceModelValidator.ValidateSource(source)
            : InvoiceModelValidator.Validate(invoice);
}
