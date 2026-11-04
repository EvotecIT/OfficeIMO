using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Data;

namespace OfficeIMO.Core.Internal;

internal static partial class ObjectDataHelpers {
    /// <summary>Selects columns before enforcing native bounds or reading their values.</summary>
    internal static string[] SelectTableColumns(IEnumerable<string> names, OfficeTableOptions? options) {
        if (options == null) return names.ToArray();
        return OrderTableColumns(names.Where(name => (options.Include.Count == 0 || options.Include.Contains(name))
            && !options.Exclude.Contains(name)), options);
    }

    private static string[] OrderTableColumns(IEnumerable<string> names, OfficeTableOptions options) {
        var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (string name in options.Order)
            if (!order.ContainsKey(name)) order.Add(name, order.Count);
        return names.OrderBy(name => order.TryGetValue(name, out int index) ? index : int.MaxValue).ToArray();
    }

    internal static string GetTableHeader(string name, OfficeTableOptions? options) =>
        options != null && options.HeaderRenames.TryGetValue(name, out string? header) ? header
        : options?.HeaderTransform?.Invoke(name) ?? name;

    internal static object? FormatTableValue(string name, object? value, OfficeTableOptions? options) =>
        options != null && options.Formatters.TryGetValue(name, out var formatter) ? formatter(value) : value;

    /// <summary>Applies the common table controls to the existing bounded flattening engine.</summary>
    internal static void ConfigureTableProjection(ObjectFlattenerOptions target, OfficeTableOptions? options) {
        if (options == null) return;
        if (options.Include.Count != 0) target.IncludeProperties = options.Include.ToArray();
        target.ExcludeProperties = target.ExcludeProperties.Concat(options.Exclude).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (options.Order.Count != 0) {
            target.PinnedFirst = options.Order.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (target.Columns != null && target.Columns.Length > 0)
                target.Columns = OrderTableColumns(target.Columns, options);
        }
        foreach (var entry in options.Formatters) target.Formatters[entry.Key] = value => entry.Value(value);
        target.TableOptions = options;
    }

    internal static void ValidateTableAlignments(OfficeTableOptions? options) {
        if (options == null) return;
        foreach (OfficeAlignment? alignment in options.Alignments)
            if (alignment.HasValue && !Enum.IsDefined(typeof(OfficeAlignment), alignment.Value))
                throw new ArgumentOutOfRangeException(nameof(options), "Unknown table column alignment.");
    }
}
