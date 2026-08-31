using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Data;

/// <summary>Requests compiler-generated <see cref="RowMapper{T}"/> configuration for a model.</summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, AllowMultiple = false, Inherited = false)]
public sealed class GenerateRowMapperAttribute : Attribute { }

/// <summary>Declares a primary tabular column name and optional aliases for a model property.</summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class DataColumnAttribute : Attribute, IDataColumnAliasProvider {
    /// <summary>Initializes a column mapping.</summary>
    public DataColumnAttribute(string name, params string[] aliases) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Column name cannot be null or whitespace.", nameof(name));
        }

        Name = name;
        Aliases = aliases?.Where(static alias => !string.IsNullOrWhiteSpace(alias)).ToArray()
            ?? Array.Empty<string>();
    }

    /// <summary>Gets the primary column name.</summary>
    public string Name { get; }

    /// <summary>Gets additional accepted column names.</summary>
    public IReadOnlyList<string> Aliases { get; }

    IReadOnlyList<string> IDataColumnAliasProvider.ColumnAliases =>
        new[] { Name }.Concat(Aliases).ToArray();
}
