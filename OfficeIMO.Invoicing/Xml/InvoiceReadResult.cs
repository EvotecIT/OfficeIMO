namespace OfficeIMO.Invoicing;

/// <summary>Parsed editable invoice together with source identity and all unmapped business data.</summary>
public sealed class InvoiceReadResult {
    private readonly byte[] _source;
    internal InvoiceReadResult(Invoice invoice, InvoiceProfileDeclaration declaration, byte[] source, IReadOnlyList<InvoiceDiagnostic> unmapped) {
        Invoice = invoice; Declaration = declaration; _source = source; UnmappedData = unmapped;
    }
    /// <summary>Editable semantic model. Declared source amounts remain available for comparison.</summary>
    public Invoice Invoice { get; }
    /// <summary>Source syntax and guideline declaration.</summary>
    public InvoiceProfileDeclaration Declaration { get; }
    /// <summary>Elements, attributes or semantic variations that the model could not preserve.</summary>
    public IReadOnlyList<InvoiceDiagnostic> UnmappedData { get; }
    /// <summary>True when the supported model can represent all observed source business data.</summary>
    public bool HasCompleteMapping => UnmappedData.Count == 0;
    /// <summary>Returns the original bytes unchanged, regardless of subsequent model edits.</summary>
    public byte[] GetOriginalBytes() => (byte[])_source.Clone();
    /// <summary>Writes the edited model. By default, any unmapped source data blocks rewriting.</summary>
    public byte[] Write(InvoiceXmlOptions? options = null, bool allowUnmappedDataLoss = false) {
        if (!HasCompleteMapping && !allowUnmappedDataLoss)
            throw new InvalidDataException("Rewriting would discard unmapped source data. Inspect UnmappedData and explicitly accept that loss before writing.");
        if (options == null) {
            if (!Declaration.Profile.HasValue) throw new InvalidDataException("An explicit supported output guideline is required for this source.");
            options = new InvoiceXmlOptions(Declaration.Syntax, Declaration.Profile.Value);
        }
        return InvoiceSerializer.Write(Invoice, options);
    }
}

/// <summary>Bounded syntax conversion result. Failure returns diagnostics and no output bytes.</summary>
public sealed class InvoiceConversionResult {
    private readonly byte[]? _xml;
    internal InvoiceConversionResult(byte[]? xml, IReadOnlyList<InvoiceDiagnostic> diagnostics) { _xml = xml; Diagnostics = diagnostics; }
    /// <summary>True when conversion produced XML without discarding observed business data.</summary>
    public bool Succeeded => _xml != null;
    /// <summary>Converted XML as a defensive copy, or null when conversion is blocked.</summary>
    public byte[]? Xml => _xml == null ? null : (byte[])_xml.Clone();
    /// <summary>Unmapped fields, unsupported target fields, model errors and profile changes.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
}

/// <summary>Native CII/UBL conversion over the same model used for authoring and presentation.</summary>
public static class InvoiceConverter {
    /// <summary>Converts only when the source mapping, model and target contract are complete; bounded diagnostics retain the highest omitted severity.</summary>
    public static InvoiceConversionResult Convert(byte[] xml, InvoiceXmlOptions target) {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (target == null) throw new ArgumentNullException(nameof(target));
        InvoiceReadResult source;
        try { source = InvoiceParser.Read(xml); }
        catch (Exception exception) when (exception is InvalidDataException || exception is System.Xml.XmlException) {
            return new InvoiceConversionResult(null, new[] { new InvoiceDiagnostic("INV-CONVERSION-INPUT", exception.Message, "Source") });
        }
        var diagnostics = new InvoiceDiagnosticBuffer();
        diagnostics.AddRange(source.UnmappedData);
        diagnostics.AddRange(InvoiceSerializer.InspectTarget(source.Invoice, target));
        if (source.Declaration.Profile != target.Profile)
            diagnostics.Add(new InvoiceDiagnostic("INV-PROFILE-CHANGE", "The output guideline differs from the source. Run the target's pinned validation rules before use.", "Guideline", InvoiceDiagnosticSeverity.Information));
        byte[]? output = null;
        if (!diagnostics.HasErrors) {
            try { output = InvoiceSerializer.Write(source.Invoice, target); }
            catch (InvalidDataException exception) {
                diagnostics.Add(new InvoiceDiagnostic("INV-CONVERSION-OUTPUT", exception.Message, "Target"));
            }
        }
        return new InvoiceConversionResult(output, diagnostics.ToList().AsReadOnly());
    }
}
