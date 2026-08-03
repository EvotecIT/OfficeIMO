using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;
using System.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void OpenDataReader_XlsbFastPathPreservesTheCanonicalSchemaContract() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"));

        DataReaderSchemaContractAssertions.AssertCanonicalSchema(reader);
    }
}

internal static class DataReaderSchemaContractAssertions {
    internal static void AssertCanonicalSchema(DbDataReader reader) {
        string[] expectedColumns = {
            SchemaTableColumn.ColumnName,
            SchemaTableColumn.ColumnOrdinal,
            SchemaTableColumn.ColumnSize,
            SchemaTableColumn.NumericPrecision,
            SchemaTableColumn.NumericScale,
            SchemaTableColumn.DataType,
            SchemaTableColumn.ProviderType,
            SchemaTableColumn.IsLong,
            SchemaTableColumn.AllowDBNull,
            "IsReadOnly",
            "IsRowVersion",
            SchemaTableColumn.IsUnique,
            SchemaTableColumn.IsKey,
            "IsAutoIncrement",
            SchemaTableColumn.BaseSchemaName,
            "BaseCatalogName",
            SchemaTableColumn.BaseTableName,
            SchemaTableColumn.BaseColumnName
        };

        DataTable schema = reader.GetSchemaTable();
        Assert.Equal(expectedColumns, schema.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        Assert.Equal(reader.FieldCount, schema.Rows.Count);

        for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) {
            string name = reader.GetName(ordinal);
            DataRow row = schema.Rows[ordinal];
            Assert.Equal(name, row[SchemaTableColumn.ColumnName]);
            Assert.Equal(ordinal, row[SchemaTableColumn.ColumnOrdinal]);
            Assert.Equal(-1, row[SchemaTableColumn.ColumnSize]);
            Assert.Equal(reader.GetFieldType(ordinal), row[SchemaTableColumn.DataType]);
            Assert.Equal(true, row[SchemaTableColumn.AllowDBNull]);
            Assert.Equal(true, row["IsReadOnly"]);
            Assert.Equal(false, row["IsRowVersion"]);
            Assert.Equal(false, row[SchemaTableColumn.IsUnique]);
            Assert.Equal(false, row[SchemaTableColumn.IsKey]);
            Assert.Equal(false, row["IsAutoIncrement"]);
            Assert.Equal(name, row[SchemaTableColumn.BaseColumnName]);
        }

#if NET8_0_OR_GREATER
        ReadOnlyCollection<DbColumn> columns = reader.GetColumnSchema();
        Assert.Equal(reader.FieldCount, columns.Count);
        for (int ordinal = 0; ordinal < columns.Count; ordinal++) {
            DbColumn column = columns[ordinal];
            Assert.Equal(reader.GetName(ordinal), column.ColumnName);
            Assert.Equal(ordinal, column.ColumnOrdinal);
            Assert.Equal(-1, column.ColumnSize);
            Assert.True(column.AllowDBNull);
            Assert.True(column.IsReadOnly);
            Assert.False(column.IsUnique);
            Assert.False(column.IsKey);
            Assert.False(column.IsAutoIncrement);
            Assert.Equal(reader.GetName(ordinal), column.BaseColumnName);
        }
#endif
    }
}
