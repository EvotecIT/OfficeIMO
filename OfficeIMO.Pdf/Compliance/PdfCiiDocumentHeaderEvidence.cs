namespace OfficeIMO.Pdf;

internal sealed class PdfCiiDocumentHeaderEvidence {
    internal PdfCiiDocumentHeaderEvidence(string? id, string? typeCode, string? issueDateTime, bool hasSupplyChainTradeTransaction) {
        Id = string.IsNullOrWhiteSpace(id) ? null : id!.Trim();
        TypeCode = string.IsNullOrWhiteSpace(typeCode) ? null : typeCode!.Trim();
        IssueDateTime = string.IsNullOrWhiteSpace(issueDateTime) ? null : issueDateTime!.Trim();
        HasSupplyChainTradeTransaction = hasSupplyChainTradeTransaction;
    }

    internal string? Id { get; }

    internal string? TypeCode { get; }

    internal string? IssueDateTime { get; }

    internal bool HasSupplyChainTradeTransaction { get; }
}
