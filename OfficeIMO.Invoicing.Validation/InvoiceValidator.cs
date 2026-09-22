using System.Security.Cryptography;
using System.Xml;

namespace OfficeIMO.Invoicing.Validation;

/// <summary>Validates exact invoice bytes using pinned schema and business-rule artifacts.</summary>
public sealed class InvoiceValidator {
    private readonly InvoiceRuleBundle _bundle;
    private readonly SaxonInvoiceRulesRunner? _runner;
    /// <summary>Creates a validator. Without a runner, business-rule status remains NotRun and IsValid remains false.</summary>
    public InvoiceValidator(InvoiceRuleBundle bundle, SaxonInvoiceRulesRunner? runner = null) { _bundle = bundle ?? throw new ArgumentNullException(nameof(bundle)); _runner = runner; }
    /// <summary>Runs local schema and configured Schematron validation, preserving cancellation and distinguishing engine failure from invalid input.</summary>
    public async Task<InvoiceValidationReport> ValidateAsync(byte[] xml, InvoiceSpecificationRelease release, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(xml);
        if (xml.Length == 0 || xml.Length > InvoiceProfileDeclaration.MaximumXmlBytes) throw new InvalidDataException("Invoice XML must contain between 1 byte and 16 MiB.");
        if (!Enum.IsDefined(release)) throw new ArgumentOutOfRangeException(nameof(release));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] snapshot = (byte[])xml.Clone();
        string hash = Convert.ToHexString(SHA256.HashData(snapshot));
        var diagnostics = new InvoiceDiagnosticBuffer();
        InvoiceValidationStatus schema = InvoiceValidationStatus.NotRun, rules = InvoiceValidationStatus.NotRun;
        string? runnerIdentity = null;
        InvoiceProfileDeclaration declaration;
        try {
            declaration = InvoiceProfileDeclaration.Read(snapshot);
        } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
            schema = InvoiceValidationStatus.Invalid;
            diagnostics.Add(new InvoiceDiagnostic("INV-XML", exception.Message, "Invoice")); return Report();
        }
        if (!declaration.Profile.HasValue) {
            diagnostics.Add(new InvoiceDiagnostic("INV-RULESET-PROFILE", "Invoice syntax/guideline does not match the explicitly selected validation release.", "Guideline"));
            return Report();
        }
        try { _ = new InvoiceXmlOptions(release, declaration.Syntax, declaration.Profile.Value); }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException) {
            diagnostics.Add(new InvoiceDiagnostic("INV-RULESET-PROFILE", exception.Message, "Guideline"));
            return Report();
        }
        bool credit;
        using (var input = new MemoryStream(snapshot, false))
        using (XmlReader reader = XmlReader.Create(input, InvoiceRuleBundle.XmlSettings())) {
            reader.MoveToContent(); credit = reader.LocalName == "CreditNote";
        }
        try {
            List<InvoiceDiagnostic> schemaDiagnostics = InvoiceSchemaValidation.Validate(snapshot, _bundle, release, declaration.Profile!.Value, declaration.Syntax, credit, cancellationToken);
            diagnostics.AddRange(schemaDiagnostics);
            schema = schemaDiagnostics.Any(static diagnostic => diagnostic.Code == InvoiceDiagnosticBuffer.IncompleteCode)
                ? InvoiceValidationStatus.Failed
                : diagnostics.HasErrors ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed;
        } catch (Exception exception) when (exception is XmlException or System.Xml.Schema.XmlSchemaException or InvalidDataException) {
            schema = InvoiceValidationStatus.Failed; diagnostics.Add(new InvoiceDiagnostic("INV-SCHEMA-ENGINE", exception.Message, "Schema"));
        }
        if (schema != InvoiceValidationStatus.Passed) return Report();
        if (_runner == null || release == InvoiceSpecificationRelease.PeppolBis_3_0_21 && !_bundle.HasPeppolRules ||
            release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2 && !_bundle.HasFacturXRules) {
            string message = _runner == null ? "Configure the Saxon runner to execute business rules." :
                release == InvoiceSpecificationRelease.PeppolBis_3_0_21 ? "Load the pinned Peppol Schematron source." : "Load the pinned Factur-X validation archive.";
            diagnostics.Add(new InvoiceDiagnostic("INV-RULES-NOT-RUN", message, "BusinessRules", InvoiceDiagnosticSeverity.Warning));
            return Report();
        }
        try {
            var overrides = _bundle.SeverityOverrides(release, declaration.Syntax, credit);
            bool incompleteRules = false;
            foreach (InvoiceRuleSource rule in _bundle.Rules(release, declaration.Profile.Value, declaration.Syntax)) {
                IReadOnlyList<InvoiceDiagnostic> result = await _runner.RunAsync(snapshot, rule.Bytes, rule.Compile, overrides, cancellationToken,
                    rule.SupportingFiles, rule.SourcePath,
                    () => runnerIdentity = _runner.Identity).ConfigureAwait(false);
                diagnostics.AddRange(result);
                if (result.Any(static diagnostic => diagnostic.Code == InvoiceDiagnosticBuffer.IncompleteCode)) {
                    incompleteRules = true;
                    break;
                }
            }
            rules = incompleteRules ? InvoiceValidationStatus.Failed
                : diagnostics.HasErrors ? InvoiceValidationStatus.Invalid : InvoiceValidationStatus.Passed;
        } catch (Exception exception) when (exception is not OperationCanceledException) {
            rules = InvoiceValidationStatus.Failed; diagnostics.Add(new InvoiceDiagnostic("INV-RULES-ENGINE", exception.Message, "BusinessRules"));
        }
        return Report();
        InvoiceValidationReport Report() => new InvoiceValidationReport(hash, snapshot.Length, release, schema, rules, diagnostics.ToList().AsReadOnly(), runnerIdentity);
    }
}
