namespace OfficeIMO.Pdf;

/// <summary>Creates and transactionally edits AcroForm fields in existing PDFs.</summary>
internal static partial class PdfAcroFormEditor {
    /// <summary>Applies field-tree, widget, calculation-order, tab-order, and selective-flatten edits as one validated full rewrite.</summary>
    public static PdfAcroFormEditResult Edit(byte[] pdf, Action<PdfAcroFormEditSession> edit, PdfReadOptions? readOptions = null, PdfFormFillerOptions? appearanceOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(edit, nameof(edit));
        PdfReadDocument source = PdfReadDocument.Open(pdf, readOptions);
        if (source.AcroFormXfa is not null) throw new NotSupportedException("Transactional AcroForm editing does not modify XFA packets. Remove or convert XFA before editing the AcroForm field tree.");

        var session = new PdfAcroFormEditSession(source.ReadOptions.Limits);
        edit(session);
        if (session.Commands.Count == 0) throw new ArgumentException("At least one AcroForm edit command is required.", nameof(edit));
        string[] fieldNames = session.Commands.SelectMany(GetCommandFieldNames).Distinct(StringComparer.Ordinal).ToArray();
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAcroForm, readOptions, fieldNames);

        var refillValues = new Dictionary<string, string>(StringComparer.Ordinal);
        var flattenNames = new List<string>();
        var operations = new List<string>(session.Commands.Count);
        int[] pageObjectNumbers = source.Pages.Select(static page => page.ObjectNumber).ToArray();
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => {
            ApplyCommands(objects, security, pageObjectNumbers, session.Commands, refillValues, flattenNames, operations, appearanceOptions);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });

        PdfReadOptions savedReadOptions = PdfReadOptions.WithMinimumInputBytes(source.ReadOptions, output.LongLength);
        _ = PdfReadDocument.Open(output, savedReadOptions).FormWidgetJavaScriptCount;
        if (refillValues.Count > 0) {
            IReadOnlyDictionary<string, PdfFormField> rewrittenFields = PdfInspector.Inspect(output, savedReadOptions).FormFieldsByName;
            foreach (string fieldName in refillValues.Keys.ToArray()) {
                if (rewrittenFields.TryGetValue(fieldName, out PdfFormField? field) && field.IsPushButton) {
                    refillValues.Remove(fieldName);
                }
            }
        }
        if (refillValues.Count > 0) {
            output = PdfFormFiller.FillFieldsWithinPlannedRewrite(output, ToFieldValues(refillValues), appearanceOptions, savedReadOptions);
            savedReadOptions = PdfReadOptions.WithMinimumInputBytes(source.ReadOptions, output.LongLength);
        }
        if (flattenNames.Count > 0) {
            output = PdfFormFiller.FlattenFieldsWithinPlannedRewrite(output, flattenNames, readOptions: savedReadOptions);
            savedReadOptions = PdfReadOptions.WithMinimumInputBytes(source.ReadOptions, output.LongLength);
        }

        PdfDocumentInfo saved = PdfInspector.Inspect(output, savedReadOptions);
        IReadOnlyList<string> calculationOrder = ReadCalculationOrder(output, savedReadOptions);
        ValidateReadback(saved, calculationOrder, session.Commands);
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = source.ReadOptions,
            RewrittenReadOptions = savedReadOptions,
            PreserveForms = false,
            PreserveAnnotations = false,
            PreserveRevisionStructure = false,
            PreserveSecurityState = !session.Commands.Any(static command => command.Options?.Kind == PdfFormFieldCreationKind.Signature)
        };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(pdf, output, preservationOptions);
        return new PdfAcroFormEditResult(output, plan, preservation, saved.FormFields, calculationOrder, operations.AsReadOnly(), savedReadOptions);
    }

    private static IEnumerable<string> GetCommandFieldNames(PdfAcroFormEditSession.EditCommand command) {
        if (!string.IsNullOrWhiteSpace(command.Name)) yield return command.Name!;
        if (command.Kind == PdfAcroFormEditSession.EditKind.Rename && !string.IsNullOrWhiteSpace(command.Value)) yield return command.Value!;
        if (command.Options is not null && !string.IsNullOrWhiteSpace(command.Options.Name)) yield return command.Options.Name;
        if (command.Names is not null) for (int i = 0; i < command.Names.Length; i++) if (!string.IsNullOrWhiteSpace(command.Names[i])) yield return command.Names[i];
    }

    private static Dictionary<string, PdfFormFieldValue> ToFieldValues(Dictionary<string, string> values) {
        var result = new Dictionary<string, PdfFormFieldValue>(values.Count, StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> entry in values) result[entry.Key] = PdfFormFieldValue.From(entry.Value);
        return result;
    }
}
