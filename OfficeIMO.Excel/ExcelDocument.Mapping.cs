using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using OfficeIMO.Data;

// The 3.6 projections mirror ExcelSheet and differ by a required mapping argument.
#pragma warning disable RS0026

namespace OfficeIMO.Excel;

/// <summary>Typed row projections over a loaded Excel document.</summary>
public static class ExcelDocumentMappingExtensions {
    /// <summary>Projects rows through the shared data-reader mapping contract.</summary>
    public static IEnumerable<T> RowsAs<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
        this ExcelDocument document,
        ExcelReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new ExcelReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAs<T>()) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows through the shared data-reader mapping contract.</summary>
    public static IEnumerable<T> RowsAs<T>(
        this ExcelDocument document,
        Action<RowMapper<T>> configure,
        ExcelReaderOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new ExcelReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAs(configure)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows through the shared data-reader mapping contract.</summary>
    public static IEnumerable<T> RowsAs<T>(
        this ExcelDocument document,
        Func<IDataRecord, T> factory,
        ExcelReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new ExcelReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAs(factory)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows in bounded parallel batches, preserving source order.</summary>
    public static IEnumerable<T> RowsAsParallel<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicConstructors)] T>(
        this ExcelDocument document,
        ParallelRowMappingOptions? parallelOptions = null,
        ExcelReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new ExcelReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAsParallel<T>(parallelOptions, linked.Token)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows in bounded parallel batches, preserving source order.</summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this ExcelDocument document,
        Action<RowMapper<T>> configure,
        ParallelRowMappingOptions? parallelOptions = null,
        ExcelReaderOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new ExcelReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAsParallel(configure, parallelOptions, linked.Token)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    /// <summary>Projects rows in bounded parallel batches, preserving source order.</summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this ExcelDocument document,
        Func<IDataRecord, T> factory,
        ParallelRowMappingOptions? parallelOptions = null,
        ExcelReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken, options?.CancellationToken ?? default);
        linked.Token.ThrowIfCancellationRequested();
        var effectiveOptions = options?.Clone() ?? new ExcelReaderOptions();
        effectiveOptions.CancellationToken = linked.Token;
        using var reader = document.CreateDataReader(effectiveOptions);
        foreach (T row in reader.RowsAsParallel(factory, parallelOptions, linked.Token)) {
            linked.Token.ThrowIfCancellationRequested();
            yield return row;
        }
    }

}

#pragma warning restore RS0026
