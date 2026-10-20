using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.XmlTables;

/// <summary>Explicit, bounded table schema mapping; it does not infer a workbook from arbitrary XML structure.</summary>
internal static class ExcelXmlTableCodec {
    private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";

    internal static DataTable Read(Stream input, ExcelXmlTableOptions options) {
        XNamespace ns = options.NamespaceUri;
        XElement root = Utilities.ExcelXmlInput.Read(input, options.MaxInputBytes, options.MaxXmlDepth, options.MaxXmlElements, options.CancellationToken);
        return ReadRoot(root, options);
    }

    internal static DataTable Read(string input, ExcelXmlTableOptions options) => ReadRoot(Utilities.ExcelXmlInput.ReadText(input, options.MaxInputBytes, options.MaxXmlDepth, options.MaxXmlElements, options.CancellationToken), options);

    private static DataTable ReadRoot(XElement root, ExcelXmlTableOptions options) {
        XNamespace ns = options.NamespaceUri;
        if (root.Name != ns + options.RootElementName) throw new InvalidDataException("XML table root does not match the configured schema.");
        if (root.Elements().Any(item => item.Name != ns + options.RowElementName)) throw new InvalidDataException("Unexpected XML table row element.");
        var table = new DataTable();
        foreach (var mapping in options.Columns) table.Columns.Add(mapping.Name, DataType(mapping.ValueType));
        foreach (XElement row in root.Elements()) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (table.Rows.Count >= options.MaxRows) throw new InvalidDataException("XML table exceeds MaxRows.");
            object[] values;
            if (options.Columns.Count == 0) {
                XElement[] cells = row.Elements().ToArray();
                if (cells.Length > options.MaxColumns) throw new InvalidDataException("XML table exceeds MaxColumns.");
                if (cells.Any(item => item.Name != ns + options.CellElementName)) throw new InvalidDataException("Unexpected positional XML table cell.");
                if ((long)(table.Rows.Count + 1) * Math.Max(table.Columns.Count, cells.Length) > options.MaxCells) throw new InvalidDataException("XML table exceeds MaxCells.");
                while (table.Columns.Count < cells.Length) table.Columns.Add("Column" + (table.Columns.Count + 1), typeof(string));
                values = cells.Select(cell => ReadElement(cell, ExcelXmlTableValueType.Text)).ToArray();
            } else {
                if ((long)(table.Rows.Count + 1) * table.Columns.Count > options.MaxCells) throw new InvalidDataException("XML table exceeds MaxCells.");
                values = new object[options.Columns.Count];
                for (int index = 0; index < values.Length; index++) {
                    var mapping = options.Columns[index];
                    if (mapping.IsAttribute) values[index] = row.Attribute(mapping.XmlName) is XAttribute attribute ? Parse(attribute.Value, mapping.ValueType) : DBNull.Value;
                    else {
                        XElement[] matches = row.Elements(ns + mapping.XmlName).Take(2).ToArray();
                        if (matches.Length > 1) throw new InvalidDataException("Multiple XML elements match one scalar table column.");
                        values[index] = matches.Length == 0 ? DBNull.Value : ReadElement(matches[0], mapping.ValueType);
                    }
                }
            }
            table.Rows.Add(values);
        }
        return table;
    }

    internal static byte[] Write(DataTable table, ExcelXmlTableOptions options) {
        if (table.Rows.Count > options.MaxRows || table.Columns.Count > options.MaxColumns || (long)table.Rows.Count * table.Columns.Count > options.MaxCells) throw new InvalidDataException("Worksheet range exceeds the XML table resource limits.");
        if (options.Columns.Count != 0 && options.Columns.Count != table.Columns.Count) throw new InvalidDataException("XML column mapping count must match the exported worksheet range.");
        using var output = new OfficeBoundedMemoryStream(options.MaxOutputBytes);
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), NewLineHandling = NewLineHandling.Entitize, CloseOutput = false })) {
            writer.WriteStartDocument(); writer.WriteStartElement(options.RootElementName, options.NamespaceUri);
            foreach (DataRow row in table.Rows) {
                options.CancellationToken.ThrowIfCancellationRequested(); writer.WriteStartElement(options.RowElementName, options.NamespaceUri);
                // XML attributes must precede child elements regardless of the worksheet column order.
                for (int index = 0; index < options.Columns.Count; index++) if (options.Columns[index].IsAttribute && row[index] != DBNull.Value)
                    writer.WriteAttributeString(options.Columns[index].XmlName, Format(row[index], options.Columns[index].ValueType));
                for (int index = 0; index < table.Columns.Count; index++) {
                    ExcelXmlTableColumn? mapping = options.Columns.Count > 0 ? options.Columns[index] : null;
                    if (mapping?.IsAttribute == true) continue;
                    bool missing = row[index] == DBNull.Value;
                    if (missing && mapping != null) continue;
                    writer.WriteStartElement(mapping?.XmlName ?? options.CellElementName, options.NamespaceUri);
                    if (missing) writer.WriteAttributeString("xsi", "nil", Xsi.NamespaceName, "true");
                    else writer.WriteString(Format(row[index], mapping?.ValueType ?? ExcelXmlTableValueType.Text));
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); writer.WriteEndDocument();
        }
        return output.ToArray();
    }

    private static object ReadElement(XElement element, ExcelXmlTableValueType type) {
        if (element.HasElements) throw new InvalidDataException("An XML table scalar cannot contain nested elements.");
        if (element.Attribute(Xsi + "nil") is XAttribute nil && XmlConvert.ToBoolean(nil.Value)) {
            if (element.Value.Length != 0) throw new InvalidDataException("A nil XML table cell cannot contain text.");
            return DBNull.Value;
        }
        return Parse(element.Value, type);
    }

    private static object Parse(string text, ExcelXmlTableValueType type) {
        if (text.Length > 32767) throw new InvalidDataException("XML table cell exceeds Excel's text limit.");
        try {
            switch (type) {
                case ExcelXmlTableValueType.Text: return text;
                case ExcelXmlTableValueType.Number:
                    double number = XmlConvert.ToDouble(text);
                    return !double.IsNaN(number) && !double.IsInfinity(number) ? number : throw new InvalidDataException("XML table number must be finite.");
                case ExcelXmlTableValueType.Boolean: return XmlConvert.ToBoolean(text);
                case ExcelXmlTableValueType.DateTime: return XmlConvert.ToDateTime(text, XmlDateTimeSerializationMode.Unspecified);
                default: throw new InvalidDataException("Unsupported XML table scalar type.");
            }
        } catch (FormatException exception) { throw new InvalidDataException("XML table value does not match its explicit column type.", exception); }
        catch (OverflowException exception) { throw new InvalidDataException("XML table value exceeds its explicit column type.", exception); }
    }

    private static string Format(object value, ExcelXmlTableValueType type) {
        if (type == ExcelXmlTableValueType.Text) {
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            if (text.Length > 32767) throw new InvalidDataException("XML table cell exceeds Excel's text limit.");
            return text;
        }
        object scalar = DataType(type).IsInstanceOfType(value) ? value : Parse(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, type);
        return type switch {
            ExcelXmlTableValueType.Number => !double.IsNaN((double)scalar) && !double.IsInfinity((double)scalar) ? XmlConvert.ToString((double)scalar) : throw new InvalidDataException("XML table number must be finite."),
            ExcelXmlTableValueType.Boolean => XmlConvert.ToString((bool)scalar),
            ExcelXmlTableValueType.DateTime => XmlConvert.ToString((DateTime)scalar, XmlDateTimeSerializationMode.Unspecified),
            _ => throw new InvalidDataException("Unsupported XML table scalar type.") };
    }

    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields | DynamicallyAccessedMemberTypes.PublicProperties)]
    private static Type DataType(ExcelXmlTableValueType type) => type switch {
        ExcelXmlTableValueType.Text => typeof(string), ExcelXmlTableValueType.Number => typeof(double),
        ExcelXmlTableValueType.Boolean => typeof(bool), ExcelXmlTableValueType.DateTime => typeof(DateTime),
        _ => throw new InvalidDataException("Unsupported XML table scalar type.") };
}
