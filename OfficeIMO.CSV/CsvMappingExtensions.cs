#nullable enable

using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using OfficeIMO.Data;

namespace OfficeIMO.CSV;

/// <summary>Typed projections over a materialized CSV document.</summary>
public static class CsvMappingExtensions {
    /// <summary>
    /// Projects rows by matching CSV headers to writable public properties.
    /// Matching is case-insensitive and also ignores spaces and punctuation.
    /// </summary>
    public static IEnumerable<T> RowsAs<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this CsvDocument document) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        return EnumerateAutomatic<T>(document);
    }

    /// <summary>Projects rows using explicit, AOT-friendly column assignments.</summary>
    public static IEnumerable<T> RowsAs<T>(
        this CsvDocument document,
        Action<RowMapper<T>> configure) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        return EnumerateExplicit(document, configure);
    }

    /// <summary>
    /// Projects rows with a caller-supplied factory.
    /// This overload supports constructor-bound and other models without a public parameterless constructor.
    /// </summary>
    /// <param name="document">Document whose rows are projected.</param>
    /// <param name="factory">Creates one model instance from the current row.</param>
    public static IEnumerable<T> RowsAs<T>(
        this CsvDocument document,
        Func<IDataRecord, T> factory) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        return EnumerateFactory(document, factory);
    }

    private static IEnumerable<T> EnumerateAutomatic<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        CsvDocument document) where T : new() {
        AutomaticRowMappingPlan<T> plan = AutomaticRowMappingPlan<T>.Create(document.Header);
        foreach (CsvRow row in document.AsEnumerable()) {
            yield return plan.MapRow(
                index => row[index],
                document.Culture,
                document.DateTimeFormats);
        }
    }

    private static IEnumerable<T> EnumerateExplicit<T>(
        CsvDocument document,
        Action<RowMapper<T>> configure) where T : new() {
        ExplicitRowMappingPlan<T> plan = ExplicitRowMappingPlan<T>.Create(document.Header, configure);
        if (plan.IsEmpty) yield break;
        foreach (CsvRow row in document.AsEnumerable()) {
            yield return plan.MapRow(
                index => row[index],
                document.Culture,
                document.DateTimeFormats);
        }
    }

    private static IEnumerable<T> EnumerateFactory<T>(
        CsvDocument document,
        Func<IDataRecord, T> factory) {
        using DbDataReader reader = document.CreateDataReader();
        foreach (T row in reader.RowsAs(factory)) {
            yield return row;
        }
    }
}
