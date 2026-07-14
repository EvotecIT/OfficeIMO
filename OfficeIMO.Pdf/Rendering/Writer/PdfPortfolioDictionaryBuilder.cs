namespace OfficeIMO.Pdf;

internal static class PdfPortfolioDictionaryBuilder {
    internal static string Build(PdfPortfolioOptions options, IReadOnlyList<PdfEmbeddedFile> embeddedFiles) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(embeddedFiles, nameof(embeddedFiles));
        if (embeddedFiles.Count == 0) {
            throw new InvalidOperationException("A PDF portfolio requires at least one embedded file.");
        }

        if (options.InitialDocumentFileName != null &&
            !embeddedFiles.Any(file => string.Equals(file.FileName, options.InitialDocumentFileName, StringComparison.Ordinal))) {
            throw new InvalidOperationException("The PDF portfolio initial document must match a generated embedded file name.");
        }

        IReadOnlyList<PdfPortfolioField> fields = options.Fields;
        if (options.SortBy.HasValue && !fields.Any(field => field.Kind == options.SortBy.Value)) {
            throw new InvalidOperationException("The PDF portfolio sort field must also be present in the portfolio schema.");
        }

        var sb = new StringBuilder("<< /Type /Collection /View /");
        sb.Append(GetViewName(options.View));
        if (options.InitialDocumentFileName != null) {
            sb.Append(" /D ").Append(PdfSyntaxEscaper.TextString(options.InitialDocumentFileName));
        }

        if (fields.Count > 0) {
            sb.Append(" /Schema <<");
            foreach (PdfPortfolioField field in fields.OrderBy(field => field.Order)) {
                sb.Append(" /").Append(PdfSyntaxEscaper.Name(field.Key))
                    .Append(" << /Type /CollectionField /N ").Append(PdfSyntaxEscaper.TextString(field.DisplayName))
                    .Append(" /Subtype /").Append(GetSubtypeName(field.Kind))
                    .Append(" /O ").Append(field.Order.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" /V ").Append(field.Visible ? "true" : "false")
                    .Append(" /E ").Append(field.Editable ? "true" : "false")
                    .Append(" >>");
            }
            sb.Append(" >>");
        }

        if (options.SortBy.HasValue) {
            string key = fields.Single(field => field.Kind == options.SortBy.Value).Key;
            sb.Append(" /Sort << /S /").Append(PdfSyntaxEscaper.Name(key))
                .Append(" /A ").Append(options.SortAscending ? "true" : "false")
                .Append(" >>");
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    private static string GetViewName(PdfPortfolioView view) => view switch {
        PdfPortfolioView.Details => "D",
        PdfPortfolioView.Tile => "T",
        PdfPortfolioView.Hidden => "H",
        _ => throw new ArgumentOutOfRangeException(nameof(view), view, "Unsupported PDF portfolio view.")
    };

    internal static string GetSubtypeName(PdfPortfolioFieldKind kind) => kind switch {
        PdfPortfolioFieldKind.FileName => "F",
        PdfPortfolioFieldKind.Description => "Desc",
        PdfPortfolioFieldKind.CreationDate => "CreationDate",
        PdfPortfolioFieldKind.ModificationDate => "ModDate",
        PdfPortfolioFieldKind.Size => "Size",
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported PDF portfolio field kind.")
    };
}
