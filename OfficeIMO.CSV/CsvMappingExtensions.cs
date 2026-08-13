#nullable enable

using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
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

    /// <summary>
    /// Projects rows in bounded parallel batches by matching CSV headers to writable public properties.
    /// Results retain source order.
    /// </summary>
    public static IEnumerable<T> RowsAsParallel<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this CsvDocument document,
        ParallelRowMappingOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        return EnumerateParallelAutomatic<T>(document, options, cancellationToken);
    }

    /// <summary>
    /// Projects rows in bounded parallel batches using explicit, AOT-friendly column assignments.
    /// Results retain source order.
    /// </summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this CsvDocument document,
        Action<RowMapper<T>> configure,
        ParallelRowMappingOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        return EnumerateParallelExplicit(document, configure, options, cancellationToken);
    }

    /// <summary>
    /// Projects rows in bounded parallel batches with a caller-supplied factory.
    /// Results retain source order.
    /// </summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this CsvDocument document,
        Func<IDataRecord, T> factory,
        ParallelRowMappingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        return EnumerateParallelFactory(document, factory, options, cancellationToken);
    }

    private static IEnumerable<T> EnumerateAutomatic<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        CsvDocument document) where T : new() {
        AutomaticRowMappingPlan<T> plan = AutomaticRowMappingPlan<T>.Create(document.Header);
        foreach (CsvRow row in document.AsEnumerable()) {
            yield return plan.MapRow(
                index => row[index],
                document.Culture,
                document.DateTimeFormats,
                errorValuePolicy: document.MappingErrorValuePolicy);
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
                document.DateTimeFormats,
                errorValuePolicy: document.MappingErrorValuePolicy);
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

    private static IEnumerable<T> EnumerateParallelAutomatic<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        CsvDocument document,
        ParallelRowMappingOptions? options,
        CancellationToken cancellationToken) where T : new() {
        using DbDataReader reader = document.CreateDataReader();
        foreach (T row in reader.RowsAsParallel<T>(options, cancellationToken)) yield return row;
    }

    private static IEnumerable<T> EnumerateParallelExplicit<T>(
        CsvDocument document,
        Action<RowMapper<T>> configure,
        ParallelRowMappingOptions? options,
        CancellationToken cancellationToken) where T : new() {
        using DbDataReader reader = document.CreateDataReader();
        foreach (T row in reader.RowsAsParallel(configure, options, cancellationToken)) yield return row;
    }

    private static IEnumerable<T> EnumerateParallelFactory<T>(
        CsvDocument document,
        Func<IDataRecord, T> factory,
        ParallelRowMappingOptions? options,
        CancellationToken cancellationToken) {
        using DbDataReader reader = document.CreateDataReader();
        foreach (T row in reader.RowsAsParallel(factory, options, cancellationToken)) yield return row;
    }
}
