#nullable enable

#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Data;

/// <summary>Asynchronous typed row projections for forward-only data readers.</summary>
public static class DataReaderAsyncMappingExtensions {
    /// <summary>
    /// Asynchronously projects remaining rows by matching column names to writable public properties.
    /// The caller retains ownership of the reader.
    /// </summary>
    public static IAsyncEnumerable<T> RowsAsAsync<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this DbDataReader reader,
        CancellationToken cancellationToken = default) where T : new() {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        return EnumerateAutomatic<T>(reader, cancellationToken);
    }

    /// <summary>
    /// Asynchronously projects remaining rows using explicit, AOT-safe column assignments.
    /// The caller retains ownership of the reader.
    /// </summary>
    public static IAsyncEnumerable<T> RowsAsAsync<T>(
        this DbDataReader reader,
        Action<RowMapper<T>> configure,
        CancellationToken cancellationToken = default) where T : new() {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        return EnumerateExplicit(reader, configure, cancellationToken);
    }

    /// <summary>
    /// Asynchronously projects remaining rows with a caller-supplied factory.
    /// The caller retains ownership of the reader.
    /// </summary>
    public static IAsyncEnumerable<T> RowsAsAsync<T>(
        this DbDataReader reader,
        Func<IDataRecord, T> factory,
        CancellationToken cancellationToken = default) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        return EnumerateFactory(reader, factory, cancellationToken);
    }

    private static async IAsyncEnumerable<T> EnumerateAutomatic<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        DbDataReader reader,
        [EnumeratorCancellation] CancellationToken cancellationToken) where T : new() {
        if (reader.FieldCount == 0) yield break;

        DataReaderMappingExtensions.GetConversionOptions(
            reader,
            out CultureInfo culture,
            out IReadOnlyList<string>? dateTimeFormats,
            out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            out bool requireAllColumnsMapped,
            out DataMappingErrorValuePolicy errorValuePolicy);
        AutomaticRowMappingPlan<T> plan = AutomaticRowMappingPlan<T>.Create(
            DataReaderMappingExtensions.GetHeaders(reader),
            requireAllColumnsMapped);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) {
            T row = plan.MapReaderRow(
                reader,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
            cancellationToken.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    private static async IAsyncEnumerable<T> EnumerateExplicit<T>(
        DbDataReader reader,
        Action<RowMapper<T>> configure,
        [EnumeratorCancellation] CancellationToken cancellationToken) where T : new() {
        if (reader.FieldCount == 0) yield break;

        ExplicitRowMappingPlan<T> plan = ExplicitRowMappingPlan<T>.Create(
            DataReaderMappingExtensions.GetHeaders(reader),
            configure);
        if (plan.IsEmpty) yield break;
        DataReaderMappingExtensions.GetConversionOptions(
            reader,
            out CultureInfo culture,
            out IReadOnlyList<string>? dateTimeFormats,
            out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            out _,
            out DataMappingErrorValuePolicy errorValuePolicy);
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) {
            T row = plan.MapReaderRow(
                reader,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
            cancellationToken.ThrowIfCancellationRequested();
            yield return row;
        }
    }

    private static async IAsyncEnumerable<T> EnumerateFactory<T>(
        DbDataReader reader,
        Func<IDataRecord, T> factory,
        [EnumeratorCancellation] CancellationToken cancellationToken) {
        if (reader.FieldCount == 0) yield break;
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false)) {
            T row = factory(reader);
            cancellationToken.ThrowIfCancellationRequested();
            yield return row;
        }
    }
}
#endif
