using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Assesses several mutation families against one shared preflight snapshot.</summary>
    /// <remarks>
    /// This is a portfolio view over the existing mutation planner, not a second capability table.
    /// It explains which workflows can be offered for one input before any mutation is attempted.
    /// </remarks>
    public PdfMutationPortfolioReport AssessMutations(
        IEnumerable<PdfMutationOperation>? operations = null,
        IEnumerable<string>? fieldNames = null,
        PdfLoadOptions? options = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) =>
        AssessMutations(CancellationToken.None, operations, fieldNames, options, executionPreference);

    /// <summary>Assesses mutation families with cooperative cancellation and one shared preflight snapshot.</summary>
    /// <remarks>
    /// Observes cancellation while collecting requested operations and field names, during source
    /// parsing and preflight, and between individual mutation plans. No document content is changed.
    /// </remarks>
    public PdfMutationPortfolioReport AssessMutations(
        CancellationToken cancellationToken,
        IEnumerable<PdfMutationOperation>? operations = null,
        IEnumerable<string>? fieldNames = null,
        PdfLoadOptions? options = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        cancellationToken.ThrowIfCancellationRequested();
        var requestedSet = new HashSet<PdfMutationOperation>();
        if (operations is not null) {
            foreach (PdfMutationOperation operation in operations) {
                cancellationToken.ThrowIfCancellationRequested();
                requestedSet.Add(operation);
            }
        } else {
#pragma warning disable CA2263 // Generic Enum.GetValues is unavailable on netstandard2.0 and net472.
            foreach (PdfMutationOperation operation in global::OfficeIMO.Internal.EnumCompat.GetValues<PdfMutationOperation>())
                requestedSet.Add(operation);
#pragma warning restore CA2263
        }
        cancellationToken.ThrowIfCancellationRequested();
        PdfMutationOperation[] requested = requestedSet.OrderBy(static operation => operation).ToArray();
        if (requested.Length == 0) throw new ArgumentException("At least one mutation operation is required.", nameof(operations));

        List<string>? requestedFieldNames = null;
        if (fieldNames is not null) {
            requestedFieldNames = new List<string>();
            foreach (string fieldName in fieldNames) {
                cancellationToken.ThrowIfCancellationRequested();
                requestedFieldNames.Add(fieldName);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = GetReadSnapshot(options, cancellationToken);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(
            snapshot.Bytes, snapshot.Options, () => snapshot.Document, cancellationToken);
        var plans = new PdfMutationPlan[requested.Length];
        for (int index = 0; index < requested.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            plans[index] = PdfMutationPlanner.Plan(
                preflight, snapshot.Bytes, requested[index], requestedFieldNames,
                executionPreference, snapshot.Options);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfMutationPortfolioReport(preflight, Array.AsReadOnly(plans));
    }
}
