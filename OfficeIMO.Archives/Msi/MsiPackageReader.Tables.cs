using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    private void LoadSchemas() {
        var bootstrap = new[] { new MsiColumn("Table", 's', 1), new MsiColumn("Number", 'i', 2), new MsiColumn("Name", 's', 3), new MsiColumn("Type", 'i', 4) };
        object?[][] definitions = ReadRows("_Columns", bootstrap);
        var schemas = new Dictionary<string, List<MsiColumn>>(StringComparer.OrdinalIgnoreCase);
        foreach (object?[] row in definitions) {
            _token.ThrowIfCancellationRequested();
            string table = RequiredString(row[0], "_Columns.Table");
            string name = RequiredString(row[2], "_Columns.Name");
            int position = RequiredInteger(row[1], "_Columns.Number");
            int type = RequiredInteger(row[3], "_Columns.Type");
            if (position <= 0 || position > 32) throw new InvalidDataException("MSI column position is invalid.");
            if (!schemas.TryGetValue(table, out List<MsiColumn>? columns)) schemas.Add(table, columns = new List<MsiColumn>());
            if (columns.Any(column => column.Position == position || string.Equals(column.Name, name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException("MSI table has duplicate column definitions.");
            char kind = (type & 0x800) != 0 ? 's' : (type & 0xff) == 4 ? 'I' : (type & 0xff) == 2 ? 'i' : '?';
            if ((type & ~0x1000) == 0x900) kind = 'b';
            columns.Add(new MsiColumn(name, kind, position, kind == 's' ? type & 0xff : 0));
        }
        _schemas = schemas.ToDictionary(pair => pair.Key, pair => pair.Value.OrderBy(column => column.Position).ToArray(), StringComparer.OrdinalIgnoreCase);
    }
    private object?[][] Table(string name, params string[] requiredColumns) {
        if (!_schemas.TryGetValue(name, out MsiColumn[]? schema))
            throw new InvalidDataException("Required MSI table schema is absent or unsupported: " + name);
        for (int column = 0; column < schema.Length; column++)
            if (schema[column].Position != column + 1) throw new InvalidDataException("MSI table column positions are incomplete: " + name);
        var projection = new int[requiredColumns.Length];
        for (int column = 0; column < requiredColumns.Length; column++) {
            int index = Array.FindIndex(schema, field => string.Equals(field.Name, requiredColumns[column], StringComparison.Ordinal));
            if (index < 0 || schema[index].Kind == '?' || schema[index].Kind == 'b')
                throw new InvalidDataException("Required MSI table column is absent or unsupported: " + name + "." + requiredColumns[column]);
            char expected = requiredColumns[column] == "FileSize" ? 'I'
                : requiredColumns[column] == "Attributes" || requiredColumns[column] == "DiskId"
                    || requiredColumns[column] == "Sequence" || requiredColumns[column] == "LastSequence" ? 'i' : 's';
            // Large packages may widen both sequence fields from i2 to i4.
            bool wideSequence = (name == "File" && requiredColumns[column] == "Sequence"
                || name == "Media" && requiredColumns[column] == "LastSequence") && schema[index].Kind == 'I';
            if (schema[index].Kind != expected && !wideSequence)
                throw new InvalidDataException("MSI column has an incompatible type: " + name + "." + requiredColumns[column]);
            projection[column] = index;
        }
        return ReadRows(name, schema).Select(row => projection.Select(index => row[index]).ToArray()).ToArray();
    }
    private object?[][] ReadRows(string name, MsiColumn[] schema) {
        byte[] bytes = ReadMetadata(name);
        int Width(MsiColumn column) => column.Kind == 's' ? _stringWidth : column.Kind == 'i' || column.Kind == 'b' ? 2 : column.Kind == 'I' ? 4 : throw new InvalidDataException("Unsupported MSI column encoding.");
        int rowSize = schema.Sum(Width);
        if (rowSize == 0 || bytes.Length % rowSize != 0) throw new InvalidDataException("MSI table contains an incomplete row: " + name);
        int count = bytes.Length / rowSize;
        if (count > _limits.Rows) throw new InvalidDataException("MSI table exceeds the configured row limit: " + name);
        var rows = Enumerable.Range(0, count).Select(_ => new object?[schema.Length]).ToArray();
        int offset = 0;
        for (int column = 0; column < schema.Length; column++) for (int row = 0; row < count; row++) {
            _token.ThrowIfCancellationRequested();
            MsiColumn field = schema[column];
            int width = Width(field);
            uint value = width == 4 ? UInt32(bytes, offset) : UInt16(bytes, offset);
            if (width == 3) value |= (uint)bytes[offset + 2] << 16;
            offset += width;
            if (value == 0 || field.Kind == 'b') continue;
            if (field.Kind == 's') {
                if (value >= _strings.Length) throw new InvalidDataException("MSI table references an absent string: " + name);
                // Standard columns may be widened by their authored schema; zero is unbounded.
                string text = _strings[value] ?? throw new InvalidDataException("MSI table references an absent string: " + name);
                if (field.StringLength != 0 && text.Length > field.StringLength)
                    throw new InvalidDataException("MSI string exceeds its authored column length: " + name + "." + field.Name);
                rows[row][column] = text;
            } else rows[row][column] = field.Kind == 'I' ? (int)(value ^ 0x80000000) : (int)(short)(value ^ 0x8000);
        }
        return rows;
    }
    private static string RequiredString(object? value, string field) => value is string text && text.Length != 0 ? text : throw new InvalidDataException("MSI requires a nonempty string in " + field + ".");
    /// <summary>Validates Identifier-typed standard fields independently of target filesystem paths.</summary>
    private static string RequiredIdentifier(object? value, string field) {
        string text = RequiredString(value, field);
        bool Letter(char character) => character >= 'A' && character <= 'Z' || character >= 'a' && character <= 'z';
        if (!(Letter(text[0]) || text[0] == '_') || text.Any(character =>
            !Letter(character) && !(character >= '0' && character <= '9') && character != '_' && character != '.'))
            throw new InvalidDataException("MSI contains an invalid identifier in " + field + ".");
        return text;
    }
    private static int RequiredInteger(object? value, string field) => value is int number ? number : throw new InvalidDataException("MSI requires an integer in " + field + ".");
    private static Dictionary<string, object?[]> Index(object?[][] rows, string table, StringComparer? comparer = null) {
        var result = new Dictionary<string, object?[]>(comparer ?? StringComparer.Ordinal);
        foreach (object?[] row in rows) {
            string key = RequiredIdentifier(row[0], table);
            if (result.ContainsKey(key)) throw new InvalidDataException("MSI table has duplicate identifiers: " + table);
            result.Add(key, row);
        }
        return result;
    }
}
