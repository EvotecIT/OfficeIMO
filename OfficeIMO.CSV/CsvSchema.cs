#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Describes the expected structure of a CSV document.
/// </summary>
public sealed class CsvSchema
{
    internal CsvSchema(IReadOnlyList<CsvSchemaColumn> columns)
    {
        Columns = columns;
        ColumnIndexLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columns.Count; i++)
        {
            ColumnIndexLookup[columns[i].Name] = i;
        }
    }

    /// <summary>
    /// Gets the ordered set of columns expected by the schema.
    /// </summary>
    public IReadOnlyList<CsvSchemaColumn> Columns { get; }

    internal Dictionary<string, int> ColumnIndexLookup { get; }
}

/// <summary>
/// Describes a single column inside a CSV schema.
/// </summary>
public sealed class CsvSchemaColumn
{
    internal CsvSchemaColumn(string name)
    {
        Name = name;
    }

    /// <summary>
    /// Gets the column name.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Gets the expected data type for the column, when specified.
    /// </summary>
    public Type? DataType { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the column must be present and non-null.
    /// </summary>
    public bool IsRequired { get; internal set; }

    /// <summary>
    /// Gets the default value used when the field is missing or null.
    /// </summary>
    public object? DefaultValue { get; internal set; }

    /// <summary>
    /// Gets the validation rules attached to this column.
    /// </summary>
    public IReadOnlyList<CsvColumnRule> Validators => _validators;

    internal List<CsvColumnRule> _validators { get; } = new();
}

/// <summary>
/// Represents an individual validation rule for a column.
/// </summary>
public sealed class CsvColumnRule
{
    internal CsvColumnRule(Func<object?, bool> predicate, string message)
    {
        Predicate = predicate;
        Message = message;
    }

    /// <summary>
    /// Gets the validation error message emitted when the rule fails.
    /// </summary>
    public string Message { get; }

    internal Func<object?, bool> Predicate { get; }
}

/// <summary>
/// Fluent builder used to create <see cref="CsvSchema"/> instances.
/// </summary>
public sealed class CsvSchemaBuilder
{
    private readonly List<CsvSchemaColumn> _columns = new();

    /// <summary>
    /// Defines a column with the given name.
    /// </summary>
    public CsvColumnBuilder Column(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("Column name cannot be null or empty.", nameof(name));
        }

        var column = new CsvSchemaColumn(name);
        _columns.Add(column);
        return new CsvColumnBuilder(column, this);
    }

    internal CsvSchema Build()
    {
        var immutableColumns = _columns.Select(CloneColumn).ToList();
        return new CsvSchema(immutableColumns);
    }

    private static CsvSchemaColumn CloneColumn(CsvSchemaColumn column)
    {
        var clone = new CsvSchemaColumn(column.Name)
        {
            DataType = column.DataType,
            DefaultValue = column.DefaultValue,
            IsRequired = column.IsRequired
        };

        foreach (var validator in column.Validators)
        {
            clone._validators.Add(new CsvColumnRule(validator.Predicate, validator.Message));
        }

        return clone;
    }
}

/// <summary>
/// Fluent configuration for a single schema column.
/// </summary>
public sealed class CsvColumnBuilder
{
    private readonly CsvSchemaColumn _column;
    private readonly CsvSchemaBuilder _owner;

    internal CsvColumnBuilder(CsvSchemaColumn column, CsvSchemaBuilder owner)
    {
        _column = column;
        _owner = owner;
    }

    /// <summary>
    /// Marks the column as required.
    /// </summary>
    public CsvColumnBuilder Required()
    {
        _column.IsRequired = true;
        return this;
    }

    /// <summary>
    /// Marks the column as optional.
    /// </summary>
    public CsvColumnBuilder Optional()
    {
        _column.IsRequired = false;
        return this;
    }

    /// <summary>
    /// Sets the expected data type to <see cref="int"/>.
    /// </summary>
    public CsvColumnBuilder AsInt32() => AsType(typeof(int));

    /// <summary>
    /// Sets the expected data type to <see cref="string"/>.
    /// </summary>
    public CsvColumnBuilder AsString() => AsType(typeof(string));

    /// <summary>
    /// Sets the expected data type to <see cref="DateTime"/>.
    /// </summary>
    public CsvColumnBuilder AsDateTime() => AsType(typeof(DateTime));

    /// <summary>
    /// Sets the expected data type to <see cref="bool"/>.
    /// </summary>
    public CsvColumnBuilder AsBoolean() => AsType(typeof(bool));

    /// <summary>
    /// Sets a custom expected data type.
    /// </summary>
    public CsvColumnBuilder AsType(Type type)
    {
        _column.DataType = type ?? throw new ArgumentNullException(nameof(type));
        return this;
    }

    /// <summary>
    /// Specifies a default value when the field is missing or null.
    /// </summary>
    public CsvColumnBuilder WithDefault(object? value)
    {
        _column.DefaultValue = value;
        return this;
    }

    /// <summary>
    /// Adds a custom validation rule.
    /// </summary>
    public CsvColumnBuilder Validate(Func<object?, bool> predicate, string message)
    {
        if (predicate is null)
        {
            throw new ArgumentNullException(nameof(predicate));
        }

        _column._validators.Add(new CsvColumnRule(predicate, message));
        return this;
    }

    /// <summary>
    /// Returns the parent schema builder so configuration can continue.
    /// </summary>
    public CsvSchemaBuilder Done() => _owner;

    /// <summary>
    /// Begins configuration of another column on the same schema.
    /// </summary>
    public CsvColumnBuilder Column(string name) => _owner.Column(name);
}
