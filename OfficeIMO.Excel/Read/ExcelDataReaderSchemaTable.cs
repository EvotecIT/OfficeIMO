using System;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Builds the canonical ADO.NET schema table shared by Excel data-reader implementations.
    /// </summary>
    internal static class ExcelDataReaderSchemaTable {
        private const string IsReadOnlyColumn = "IsReadOnly";
        private const string IsRowVersionColumn = "IsRowVersion";
        private const string IsAutoIncrementColumn = "IsAutoIncrement";
        private const string BaseCatalogNameColumn = "BaseCatalogName";

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2111",
            Justification = "The schema table stores Type values as data and does not reflect over Type members.")]
        internal static DataTable Create(
            int fieldCount,
            Func<int, string> getName,
            Func<int, Type> getFieldType) {
            var schema = new DataTable("SchemaTable");
            schema.Columns.Add(SchemaTableColumn.ColumnName, typeof(string));
            schema.Columns.Add(SchemaTableColumn.ColumnOrdinal, typeof(int));
            schema.Columns.Add(SchemaTableColumn.ColumnSize, typeof(int));
            schema.Columns.Add(SchemaTableColumn.NumericPrecision, typeof(short));
            schema.Columns.Add(SchemaTableColumn.NumericScale, typeof(short));
            schema.Columns.Add(SchemaTableColumn.DataType, typeof(Type));
            schema.Columns.Add(SchemaTableColumn.ProviderType, typeof(int));
            schema.Columns.Add(SchemaTableColumn.IsLong, typeof(bool));
            schema.Columns.Add(SchemaTableColumn.AllowDBNull, typeof(bool));
            schema.Columns.Add(IsReadOnlyColumn, typeof(bool));
            schema.Columns.Add(IsRowVersionColumn, typeof(bool));
            schema.Columns.Add(SchemaTableColumn.IsUnique, typeof(bool));
            schema.Columns.Add(SchemaTableColumn.IsKey, typeof(bool));
            schema.Columns.Add(IsAutoIncrementColumn, typeof(bool));
            schema.Columns.Add(SchemaTableColumn.BaseSchemaName, typeof(string));
            schema.Columns.Add(BaseCatalogNameColumn, typeof(string));
            schema.Columns.Add(SchemaTableColumn.BaseTableName, typeof(string));
            schema.Columns.Add(SchemaTableColumn.BaseColumnName, typeof(string));

            for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
                string name = getName(ordinal);
                DataRow row = schema.NewRow();
                row[SchemaTableColumn.ColumnName] = name;
                row[SchemaTableColumn.ColumnOrdinal] = ordinal;
                row[SchemaTableColumn.ColumnSize] = -1;
                row[SchemaTableColumn.NumericPrecision] = DBNull.Value;
                row[SchemaTableColumn.NumericScale] = DBNull.Value;
                row[SchemaTableColumn.DataType] = getFieldType(ordinal);
                row[SchemaTableColumn.ProviderType] = 0;
                row[SchemaTableColumn.IsLong] = false;
                row[SchemaTableColumn.AllowDBNull] = true;
                row[IsReadOnlyColumn] = true;
                row[IsRowVersionColumn] = false;
                row[SchemaTableColumn.IsUnique] = false;
                row[SchemaTableColumn.IsKey] = false;
                row[IsAutoIncrementColumn] = false;
                row[SchemaTableColumn.BaseSchemaName] = DBNull.Value;
                row[BaseCatalogNameColumn] = DBNull.Value;
                row[SchemaTableColumn.BaseTableName] = DBNull.Value;
                row[SchemaTableColumn.BaseColumnName] = name;
                schema.Rows.Add(row);
            }

            return schema;
        }
    }
}
