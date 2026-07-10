using System;
using System.Collections.Generic;

namespace OfficeIMO.Data;

/// <summary>
/// Configures conversion from common tabular inputs to a <see cref="System.Data.DataTable"/>.
/// </summary>
public sealed class TabularDataOptions {
    /// <summary>Optional table name for newly created tables.</summary>
    public string? TableName { get; set; }

    /// <summary>Copies an existing DataTable input instead of returning the original table.</summary>
    public bool CopyExistingDataTable { get; set; } = true;

    /// <summary>Expands a single enumerable input into rows.</summary>
    public bool ExpandSingleEnumerableInput { get; set; } = true;

    /// <summary>Preserves explicit null input items as scalar rows.</summary>
    public bool PreserveNullRows { get; set; }

    /// <summary>Column name used for scalar rows.</summary>
    public string ScalarColumnName { get; set; } = "Value";

    /// <summary>Controls whether object columns come from the first row or all rows.</summary>
    public TabularColumnDiscoveryMode ColumnDiscoveryMode { get; set; } = TabularColumnDiscoveryMode.AllRows;

    /// <summary>Unwraps host-specific wrapper objects before tabular conversion.</summary>
    public Func<object?, object?>? UnwrapValue { get; set; }

    /// <summary>Projects host-specific objects into a column/value dictionary, using established columns when available.</summary>
    public Func<object?, IReadOnlyList<string>?, IReadOnlyDictionary<string, object?>?>? ProjectObject { get; set; }

    /// <summary>Converts cell values before they are stored in the DataTable.</summary>
    public Func<object?, object?>? NormalizeValue { get; set; }
}
