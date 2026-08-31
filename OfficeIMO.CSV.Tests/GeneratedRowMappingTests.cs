using System.Linq;
#if NET8_0_OR_GREATER
using System.Collections.Generic;
using System.Threading.Tasks;
#endif
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class GeneratedRowMappingTests {
    [Fact]
    public void GeneratedMapperUsesPrimaryNamesAliasesAndTypedValuesWithoutReflection() {
        CsvDocument document = CsvDocument.Parse("Id,Display Name,Total,Source\n42,Alpha,12.50,imported\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });

        GeneratedInvoiceRow row = reader
            .RowsAs<GeneratedInvoiceRow>(GeneratedInvoiceRowRowMapping.Configure)
            .Single();

        Assert.Equal(42, row.InvoiceId);
        Assert.Equal("Alpha", row.Name);
        Assert.Equal(12.50m, row.Total);
        Assert.Equal("imported", row.Source);
    }

    [Fact]
    public void GeneratedMapperPrefersPrimaryColumnWhenAliasIsAlsoPresent() {
        CsvDocument document = CsvDocument.Parse(
            "Invoice Id,Id,Name,Total,Origin\n42,99,Alpha,12.50,imported\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });

        GeneratedInvoiceRow row = reader
            .RowsAs<GeneratedInvoiceRow>(GeneratedInvoiceRowRowMapping.Configure)
            .Single();

        Assert.Equal(42, row.InvoiceId);
    }

    [Fact]
    public void GeneratedMappersWithPreviouslyCollidingHintNamesBothCompile() {
        var first = new RowMapper<GeneratorCollision.A.B.C_D>();
        GeneratorCollision.A.B.C_DRowMapping.Configure(first);
        var second = new RowMapper<GeneratorCollision.A.B_C_D>();
        GeneratorCollision.A.B_C_DRowMapping.Configure(second);

        Assert.NotNull(first);
        Assert.NotNull(second);
    }

    [Theory]
    [InlineData("Inherited Id")]
    [InlineData("Legacy Id")]
    public void GeneratedMapperHonorsInheritedDataColumnAttributeOnOverride(string columnName) {
        CsvDocument document = CsvDocument.Parse(columnName + "\n42\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });

        GeneratedOverrideRow row = reader
            .RowsAs<GeneratedOverrideRow>(GeneratedOverrideRowRowMapping.Configure)
            .Single();

        Assert.Equal(42, row.Id);
    }

    [Fact]
    public void GeneratedMapperDoesNotFallBackToBasePropertiesHiddenByIneligibleDerivedMembers() {
        CsvDocument document = CsvDocument.Parse("Value,Included\n99,5\n");

        using (var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true })) {
            GeneratedReadOnlyHiddenRow row = reader
                .RowsAs<GeneratedReadOnlyHiddenRow>(GeneratedReadOnlyHiddenRowRowMapping.Configure)
                .Single();
            Assert.Equal(7, ((GeneratedHiddenBase)row).Value);
            Assert.Equal(5, row.Included);
        }

        using (var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true })) {
            GeneratedInitOnlyHiddenRow row = reader
                .RowsAs<GeneratedInitOnlyHiddenRow>(GeneratedInitOnlyHiddenRowRowMapping.Configure)
                .Single();
            Assert.Equal(7, ((GeneratedHiddenBase)row).Value);
            Assert.Equal(5, row.Included);
        }

        using (var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true })) {
            GeneratedStaticHiddenRow row = reader
                .RowsAs<GeneratedStaticHiddenRow>(GeneratedStaticHiddenRowRowMapping.Configure)
                .Single();
            Assert.Equal(7, ((GeneratedHiddenBase)row).Value);
            Assert.Equal(5, row.Included);
        }
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task GeneratedMapperProjectsAsyncReaderRowsWithoutRuntimeDiscovery() {
        CsvDocument document = CsvDocument.Parse("Id,Name,Total,Origin\n1,Alpha,12.50,a\n2,Beta,20.00,b\n");
        using var reader = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true });
        var rows = new List<GeneratedInvoiceRow>();

        await foreach (GeneratedInvoiceRow row in reader.RowsAsAsync<GeneratedInvoiceRow>(
                           GeneratedInvoiceRowRowMapping.Configure)) {
            rows.Add(row);
        }

        Assert.Equal(2, rows.Count);
        Assert.Equal("Beta", rows[1].Name);
        Assert.Equal(20m, rows[1].Total);
        Assert.Equal("b", rows[1].Source);
    }
#endif
}

[GenerateRowMapper]
public class GeneratedRowBase {
    [DataColumn("Origin", "Source")]
    public string Source { get; set; } = "base";
}

[GenerateRowMapper]
public sealed class GeneratedInvoiceRow : GeneratedRowBase {
    [DataColumn("Invoice Id", "Id")]
    public int InvoiceId { get; set; }

    [DataColumn("Name", "Display Name")]
    public string Name { get; set; } = string.Empty;

    public decimal Total { get; set; }
}

public class GeneratedAttributedBase {
    [DataColumn("Inherited Id", "Legacy Id")]
    public virtual int Id { get; set; }
}

[GenerateRowMapper]
public sealed class GeneratedOverrideRow : GeneratedAttributedBase {
    public override int Id { get; set; }
}

public class GeneratedHiddenBase {
    public int Value { get; set; } = 7;

    public int Included { get; set; }
}

[GenerateRowMapper]
public sealed class GeneratedReadOnlyHiddenRow : GeneratedHiddenBase {
    public new int Value => 42;
}

[GenerateRowMapper]
public sealed class GeneratedInitOnlyHiddenRow : GeneratedHiddenBase {
    public new int Value { get; init; } = 42;
}

[GenerateRowMapper]
public sealed class GeneratedStaticHiddenRow : GeneratedHiddenBase {
    public new static int Value { get; set; } = 42;
}
