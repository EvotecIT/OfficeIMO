using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using OfficeIMO.Data;

namespace OfficeIMO.CSV;

/// <summary>Typed row projections over a loaded Csv document.</summary>
public static class CsvMappingExtensions {
    /// <summary>Projects rows through the shared data-reader mapping contract.</summary>
    public static IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
        this CsvDocument document,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new CsvReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAs<T>()) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows through the shared data-reader mapping contract.</summary>
    public static IEnumerable<T> RowsAs<T>(
        this CsvDocument document,
        Action<RowMapper<T>> configure,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new CsvReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAs(configure)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows through the shared data-reader mapping contract.</summary>
    public static IEnumerable<T> RowsAs<T>(
        this CsvDocument document,
        Func<IDataRecord, T> factory,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new CsvReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAs(factory)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows in bounded parallel batches, preserving source order.</summary>
    public static IEnumerable<T> RowsAsParallel<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
        this CsvDocument document,
        ParallelRowMappingOptions? parallelOptions = null,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new CsvReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAsParallel<T>(parallelOptions, linked.Token)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows in bounded parallel batches, preserving source order.</summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this CsvDocument document,
        Action<RowMapper<T>> configure,
        ParallelRowMappingOptions? parallelOptions = null,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new CsvReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAsParallel(configure, parallelOptions, linked.Token)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows in bounded parallel batches, preserving source order.</summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this CsvDocument document,
        Func<IDataRecord, T> factory,
        ParallelRowMappingOptions? parallelOptions = null,
        CsvReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new CsvReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAsParallel(factory, parallelOptions, linked.Token)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

}
