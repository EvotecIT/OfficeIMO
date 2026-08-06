using System.Data.Common;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void DataReaderApi_UsesOpenForSourcesAndCreateForOpenDocuments() {
        Assert.Contains(typeof(ExcelWorkbookDataReader), typeof(ExcelDocument).Assembly.GetExportedTypes());

        MethodInfo[] methods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(string)
            && method.ReturnType == typeof(ExcelWorkbookDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Stream)
            && method.ReturnType == typeof(ExcelWorkbookDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "OpenDataReader"
            && method.IsStatic
            && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(byte[])
            && method.ReturnType == typeof(ExcelWorkbookDataReader));
        Assert.Contains(methods, static method =>
            method.Name == "CreateDataReader"
            && !method.IsStatic
            && method.ReturnType == typeof(ExcelWorkbookDataReader));
    }

    [Fact]
    public void OpenDataReader_ExposesWorksheetsAsOrderedResults() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.DataReader.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "One");
                var second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, "Two");
                document.Save();
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(new[] { "First", "Second" }, reader.SheetNames);
            Assert.Equal("First", reader.CurrentSheetName);
            Assert.Equal(0, reader.CurrentSheetIndex);
            Assert.Equal(0, reader.CurrentResultIndex);
            Assert.True(reader.Read());
            Assert.Equal("One", reader.GetString(0));
            Assert.True(reader.NextResult());
            Assert.Equal(1, reader.CurrentSheetIndex);
            Assert.Equal(1, reader.CurrentResultIndex);
            Assert.Equal("Second", reader.CurrentSheetName);
            Assert.True(reader.Read());
            Assert.Equal("Two", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_SelectsWorksheetIndexAndA1Range() {
        using var document = ExcelDocument.Create(new MemoryStream());
        document.AddWorksheet("Ignore").CellValue(1, 1, "Ignored");
        ExcelSheet selected = document.AddWorksheet("Data");
        selected.CellValue(1, 1, "Skip");
        selected.CellValue(1, 2, "Name");
        selected.CellValue(2, 1, 1);
        selected.CellValue(2, 2, "Ada");

        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            document.ToBytes(),
            new ExcelReadOptions { SheetIndex = 1, A1Range = "B1:B2" });

        Assert.Equal("Data", reader.CurrentSheetName);
        Assert.Equal(1, reader.CurrentSheetIndex);
        Assert.Equal(0, reader.CurrentResultIndex);
        Assert.Equal("Name", reader.GetName(0));
        Assert.True(reader.Read());
        Assert.Equal("Ada", reader.GetString(0));
        Assert.False(reader.NextResult());
    }

    [Fact]
    public void OpenDataReader_RejectsSheetNameAndIndexTogether() {
        using var document = ExcelDocument.Create(new MemoryStream());
        document.AddWorksheet("Data").CellValue(1, 1, "Value");

        Assert.Throws<ArgumentException>(() => ExcelDocument.OpenDataReader(
            document.ToBytes(),
            new ExcelReadOptions { SheetName = "Data", SheetIndex = 0 }));
    }

    [Fact]
    public void OpenDataReader_XlsbSupportsA1RangeThroughCanonicalEntryPoint() {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelReadOptions { A1Range = "B1:B3" });

        Assert.Equal("Amount", reader.GetName(0));
        Assert.True(reader.Read());
        Assert.Equal(42, reader.GetInt32(0));
        Assert.True(reader.Read());
        Assert.Equal(50, reader.GetInt32(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void OpenDataReader_GetFieldValueSupportsNullablePrimitives() {
        using var document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Value");
        sheet.CellValue(2, 1, 42);
        byte[] workbook = document.ToBytes();

        using DbDataReader reader = ExcelDocument.OpenDataReader(workbook);

        Assert.True(reader.Read());
        Assert.Equal(42, reader.GetFieldValue<int?>(0));
    }

    [Fact]
    public void OpenDataReader_RejectsCaseInsensitiveDuplicateOpenXmlWorksheetNames() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DuplicateOpenXmlWorksheetName.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Data").CellValue(1, 1, "First");
                document.AddWorksheet("Other").CellValue(1, 1, "Second");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Sheet[] sheets = package.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().ToArray();
                sheets[1].Name = "DATA";
                package.WorkbookPart.Workbook.Save();
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("duplicate worksheet name", exception.Message, StringComparison.OrdinalIgnoreCase);
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_SupportsLegacyXlsThroughTheSameEntryPoint() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.DataReader.{Guid.NewGuid():N}.xls");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Legacy");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 7);
                sheet.CellValue(2, 2, "Ada");
                document.Save();
            }

            using (DbDataReader reader = ExcelDocument.OpenDataReader(path)) {
                Assert.Equal("Id", reader.GetName(0));
                Assert.Equal("Name", reader.GetName(1));
                Assert.True(reader.Read());
                Assert.Equal(7, reader.GetInt32(0));
                Assert.Equal("Ada", reader.GetString(1));
                Assert.False(reader.Read());
            }

            byte[] bytes = File.ReadAllBytes(path);
            using var stream = new MemoryStream(bytes, writable: false);
            using DbDataReader streamReader = ExcelDocument.OpenDataReader(stream);
            Assert.True(streamReader.Read());
            Assert.Equal(7, streamReader.GetInt32(0));
            Assert.Equal("Ada", streamReader.GetString(1));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_LegacyImportUsesTheConfiguredInputLimit() {
        var options = new ExcelReadOptions {
            MaxInputBytes = 128L * 1024L * 1024L
        };

        OfficeIMO.Excel.LegacyXls.LegacyXlsImportOptions importOptions =
            ExcelWorkbookDataReader.CreateLegacyImportOptions(options);

        Assert.Equal(128 * 1024 * 1024, importOptions.MaxInputBytes);
    }

    [Fact]
    public void OpenDataReader_ReadsSeekableWorkbookStreamFromCurrentPositionAndRestoresIt() {
        byte[] workbook = File.ReadAllBytes(GetDataReaderXlsbFixture("basic-values-formula.xlsb"));
        byte[] prefix = Encoding.UTF8.GetBytes("already-consumed-envelope");
        using var stream = new MemoryStream(prefix.Length + workbook.Length);
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(workbook, 0, workbook.Length);
        stream.Position = prefix.Length;

        using DbDataReader reader = ExcelDocument.OpenDataReader(stream);

        Assert.Equal(prefix.Length, stream.Position);
        Assert.True(reader.Read());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal(42, reader.GetInt32(1));
    }

    [Fact]
    public void OpenDataReader_RejectsUnknownPathExtensionsInsteadOfGuessing() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            ExcelDocument.OpenDataReader("workbook.unknown"));

        Assert.Contains(".xlsx", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".xlsb", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".xls", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OpenDataReader_SelectsOneWorksheetByName() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.DataReaderSheet.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Ignore").CellValue(1, 1, "Ignored");
                var selected = document.AddWorksheet("Data");
                selected.CellValue(1, 1, "Value");
                selected.CellValue(2, 1, "Selected");
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { SheetName = "data" });
            Assert.True(reader.Read());
            Assert.Equal("Selected", reader.GetString(0));
            Assert.False(reader.NextResult());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_StreamsXlsbValuesAndCachedFormulas() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(GetDataReaderXlsbFixture("basic-values-formula.xlsb"));

        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Amount", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal("Alpha", reader.GetString(0));
        Assert.Equal(42, reader.GetInt32(1));
        Assert.True(reader.Read());
        Assert.Equal(50, reader.GetInt32(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsFormulaTokensWhenCachedResultsAreDisabled() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            ExcelDocument.OpenDataReader(
                GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
                new ExcelReadOptions { UseCachedFormulaResult = false }));
        Assert.Contains("formula-token", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OpenDataReader_XlsxRejectsUnexpandedSharedFormulaFollowersInFormulaTextMode() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.SharedFormulaDataReader.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellValue(1, 1, "Value");
                sheet.CellFormula(2, 1, "B2+1");
                sheet.CellFormula(3, 1, "B3+1");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell master = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                Cell follower = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A3");
                master.CellFormula = new CellFormula("B2+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 7U,
                    Reference = "A2:A3"
                };
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 7U
                };
                worksheet.Save();
            }

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                ExcelDocument.OpenDataReader(
                    path,
                    new ExcelReadOptions { UseCachedFormulaResult = false }));

            Assert.Contains("shared-formula follower", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_DefersFormulaValidationUntilTheWorksheetIsOpened() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.DeferredFormulaValidation.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, "Ready");
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellFormula(2, 1, "B2+1");
                second.CellFormula(3, 1, "B3+1");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Sheet secondSheet = workbookPart.Workbook.Sheets!
                    .Elements<Sheet>()
                    .Single(sheet => sheet.Name?.Value == "Second");
                Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(secondSheet.Id!))
                    .Worksheet;
                Cell master = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                Cell follower = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A3");
                master.CellFormula = new CellFormula("B2+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 9U,
                    Reference = "A2:A3"
                };
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 9U
                };
                worksheet.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { UseCachedFormulaResult = false });
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));

            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => reader.NextResult());
            Assert.Contains("shared-formula follower", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
            Assert.True(reader.IsClosed);
            using var exclusive = new FileStream(
                path,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_IgnoresExtensionDescendantsWhenCheckingFormulaCache() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.FormulaExtension.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellValue(1, 1, "Value");
                sheet.CellFormula(2, 1, "B2+1");
                sheet.CellFormula(3, 1, "B3+1");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell master = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                Cell follower = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A3");
                master.CellFormula = new CellFormula("B2+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 11U,
                    Reference = "A2:A3"
                };
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 11U
                };
                follower.CellValue = null;
                var extension = new OpenXmlUnknownElement(
                    "x",
                    "extension",
                    "urn:officeimo:formula-test");
                extension.AppendChild(new OpenXmlUnknownElement(
                    "x",
                    "v",
                    "urn:officeimo:formula-test"));
                follower.AppendChild(extension);
                worksheet.Save();
            }

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                ExcelDocument.OpenDataReader(path));

            Assert.Contains("shared-formula follower", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_ValidatesSharedFormulaFollowersInStrictWorksheetNamespace() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.StrictSharedFormula.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellValue(1, 1, "Value");
                sheet.CellFormula(2, 1, "B2+1");
                sheet.CellFormula(3, 1, "B3+1");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell master = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                Cell follower = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A3");
                master.CellFormula = new CellFormula("B2+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 12U,
                    Reference = "A2:A3"
                };
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 12U
                };
                follower.CellValue = null;
                worksheet.Save();
            }
            ReplaceFirstWorksheetNamespace(
                path,
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "http://purl.oclc.org/ooxml/spreadsheetml/main");

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                ExcelDocument.OpenDataReader(path));

            Assert.Contains("shared-formula follower", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxRejectsUnexpandedSharedFormulaFollowerWithoutCachedValue() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.SharedFormulaMissingCache.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellValue(1, 1, "Value");
                sheet.CellFormula(2, 1, "B2+1");
                sheet.CellFormula(3, 1, "B3+1");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell master = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                Cell follower = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A3");
                master.CellFormula = new CellFormula("B2+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 8U,
                    Reference = "A2:A3"
                };
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 8U
                };
                follower.CellValue = null;
                worksheet.Save();
            }

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                ExcelDocument.OpenDataReader(path));

            Assert.Contains("shared-formula follower", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxAcceptsSharedFormulaFollowerWithCachedValue() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.SharedFormulaCached.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Shared");
                sheet.CellValue(1, 1, "Value");
                sheet.CellFormula(2, 1, "B2+1");
                sheet.CellFormula(3, 1, "B3+1");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Cell master = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A2");
                Cell follower = worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "A3");
                master.CellFormula = new CellFormula("B2+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 10U,
                    Reference = "A2:A3"
                };
                master.CellValue = new CellValue(2);
                follower.CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 10U
                };
                follower.CellValue = new CellValue(3);
                worksheet.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal(2, reader.GetInt32(0));
            Assert.True(reader.Read());
            Assert.Equal(3, reader.GetInt32(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_MissingWorksheetReleasesTheFile() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.MissingSheet.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            Assert.Throws<KeyNotFoundException>(() =>
                ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Missing" }));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("A1:Z1000")]
    [InlineData("A1:A2")]
    public void OpenDataReader_XlsxDiscoversActualBoundsWhenDeclaredDimensionIsStale(
        string declaredDimension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Dimension.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Name");
                sheet.CellValue(2, 1, 1);
                sheet.CellValue(2, 2, "Ada");
                sheet.CellValue(3, 1, 2);
                sheet.CellValue(3, 2, "Grace");
                document.Save();
            }

            ReplaceWorksheetDimension(path, declaredDimension);

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Id", reader.GetName(0));
            Assert.Equal("Name", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.Equal("Ada", reader.GetString(1));
            Assert.True(reader.Read());
            Assert.Equal(2, reader.GetInt32(0));
            Assert.Equal("Grace", reader.GetString(1));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbDiscoversActualColumnsWhenDeclaredDimensionIsStale() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Dimension.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceXlsbWorksheetLastColumn(path, 0);

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Name", reader.GetName(0));
            Assert.Equal("Amount", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetString(0));
            Assert.Equal(42, reader.GetInt32(1));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_HeaderlessXlsbDiscoversColumnsBeyondDeclaredDimension() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.HeaderlessDimension.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceXlsbWorksheetLastColumn(path, 0);

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { HasHeaderRow = false });

            Assert.Equal(2, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal("Name", reader.GetString(0));
            Assert.Equal("Amount", reader.GetString(1));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxUsesConfiguredCultureAndParsesGuidText() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Culture.{Guid.NewGuid():N}.xlsx");
        Guid identifier = Guid.NewGuid();
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(1, 2, "Identifier");
                sheet.CellValue(2, 1, "1,5");
                sheet.CellValue(2, 2, identifier.ToString("D"));
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { Culture = CultureInfo.GetCultureInfo("de-DE") });
            Assert.True(reader.Read());
            Assert.Equal(1.5m, reader.GetDecimal(0));
            Assert.Equal(identifier, reader.GetGuid(1));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbMissingWorksheetReleasesTheFile() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.MissingSheet.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            Assert.Throws<KeyNotFoundException>(() =>
                ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = "Missing" }));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbHasRowsRemainsStableAfterLastRow() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"));

        Assert.True(reader.HasRows);
        while (reader.Read()) {
        }

        Assert.True(reader.HasRows);
    }

    [Fact]
    public void OpenDataReader_XlsxHasRowsRemainsStableAfterLastRow() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.HasRows.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "One");
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.HasRows);
            while (reader.Read()) {
            }

            Assert.True(reader.HasRows);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbEnforcesSharedStringLimits() {
        string path = GetDataReaderXlsbFixture("basic-values-formula.xlsb");

        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MaxSharedStringItems = 1 }));
        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MaxSharedStringItemCharacters = 3 }));
        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(path, new ExcelReadOptions { MaxSharedStringCharacters = 5 }));
    }

    [Fact]
    public void OpenDataReader_ProjectedXlsbEnforcesSharedStringCharacterLimitsDuringImport() {
        string path = GetDataReaderXlsbFixture("basic-values-formula.xlsb");

        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    A1Range = "A1:B5",
                    MaxSharedStringItemCharacters = 3
                }));
        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    A1Range = "A1:B5",
                    MaxSharedStringCharacters = 5
                }));
    }

    [Fact]
    public void OpenDataReader_PreCancelledLegacyXlsStopsBeforeLoading() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            ExcelDocument.OpenDataReader(
                "not-opened.xls",
                new ExcelReadOptions { CancellationToken = cancellation.Token }));
    }

    [Fact]
    public void CreateDataReader_PreCancelledOpenDocumentDoesNotMaterializeDeferredRows() {
        using var document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.InsertObjects(
            new[] { new { Name = "North", Score = 10 } },
            ("Name", row => row.Name),
            ("Score", row => row.Score));
        Assert.True(document.HasDeferredDirectDataSetImport);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            document.CreateDataReader(
                new ExcelReadOptions { CancellationToken = cancellation.Token }));
        Assert.True(document.HasDeferredDirectDataSetImport);
    }

    [Fact]
    public void CreateDataReader_ObservesCancellationDuringDeferredDataSetMaterialization() {
        var dataSet = new System.Data.DataSet("Export");
        var table = new System.Data.DataTable("Data");
        for (int columnIndex = 0; columnIndex < 8; columnIndex++) {
            table.Columns.Add("Column" + columnIndex.ToString(CultureInfo.InvariantCulture), typeof(int));
        }

        var values = new object[8];
        for (int rowIndex = 0; rowIndex < 50_000; rowIndex++) {
            for (int columnIndex = 0; columnIndex < values.Length; columnIndex++) {
                values[columnIndex] = rowIndex + columnIndex;
            }

            table.Rows.Add(values);
        }

        dataSet.Tables.Add(table);
        using var document = ExcelDocument.Create(new MemoryStream());
        document.InsertDataSet(dataSet, createTables: false);
        Assert.True(document.HasDeferredDirectDataSetImport);

        using var cancellation = new CancellationTokenSource();
        using var cancelThreadReady = new ManualResetEventSlim();
        int materializationObserved = 0;
        var cancelThread = new Thread(() => {
            cancelThreadReady.Set();
            if (SpinWait.SpinUntil(
                    () => document.IsMaterializingDeferredDataSetImport,
                    TimeSpan.FromSeconds(10))) {
                Interlocked.Exchange(ref materializationObserved, 1);
                cancellation.Cancel();
            }
        });
        cancelThread.Start();
        cancelThreadReady.Wait();
        try {
            Assert.Throws<OperationCanceledException>(() =>
                document.CreateDataReader(
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));
        } finally {
            Assert.True(cancelThread.Join(TimeSpan.FromSeconds(10)));
        }
        Assert.Equal(1, Volatile.Read(ref materializationObserved));

        using DbDataReader reader = document.CreateDataReader();
        int rowCount = 0;
        while (reader.Read()) {
            Assert.Equal(rowCount, reader.GetInt32(0));
            Assert.Equal(rowCount + 7, reader.GetInt32(7));
            rowCount++;
        }

        Assert.Equal(50_000, rowCount);
        Assert.False(document.HasDeferredDirectDataSetImport);
    }

    [Fact]
    public void PendingDirectCellValueMaterializationCanResumeAfterCancellation() {
        using var document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        for (int row = 1; row <= 50_000; row++) {
            sheet.CellValue(row, 1, row);
        }
        Assert.True(document.HasPendingDirectCellValues);

        using var cancellation = new CancellationTokenSource();
        using var cancelThreadReady = new ManualResetEventSlim();
        int materializationObserved = 0;
        var cancelThread = new Thread(() => {
            cancelThreadReady.Set();
            if (SpinWait.SpinUntil(
                    () => sheet.IsMaterializingPendingDirectCellValues,
                    TimeSpan.FromSeconds(10))) {
                Interlocked.Exchange(ref materializationObserved, 1);
                cancellation.Cancel();
            }
        });
        cancelThread.Start();
        cancelThreadReady.Wait();
        try {
            Assert.Throws<OperationCanceledException>(() =>
                sheet.MaterializePendingDirectCellValues(cancellation.Token));
        } finally {
            Assert.True(cancelThread.Join(TimeSpan.FromSeconds(10)));
        }
        Assert.Equal(1, Volatile.Read(ref materializationObserved));
        Assert.True(document.HasPendingDirectCellValues);

        sheet.MaterializePendingDirectCellValues();

        Assert.False(document.HasPendingDirectCellValues);
        Assert.True(sheet.TryGetCellText(1, 1, out string first));
        Assert.Equal("1", first);
        Assert.True(sheet.TryGetCellText(50_000, 1, out string last));
        Assert.Equal("50000", last);
    }

    [Fact]
    public void OpenDataReader_PreCancelledLargeSeekableStreamStopsBeforeSizingItsSnapshot() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var stream = new LargeLengthSeekableStream((long)int.MaxValue + 1);

        Assert.Throws<OperationCanceledException>(() =>
            ExcelDocument.OpenDataReader(
                stream,
                new ExcelReadOptions {
                    CancellationToken = cancellation.Token,
                    MaxInputBytes = long.MaxValue
                }));
        Assert.False(stream.ReadAttempted);
    }

    [Fact]
    public void OpenDataReader_XlsbObservesCancellationDuringTraversal() {
        using var cancellation = new CancellationTokenSource();
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelReadOptions { CancellationToken = cancellation.Token });

        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }

    [Fact]
    public void Load_XlsbImportObservesConfiguredCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => ExcelDocument.Load(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelLoadOptions {
                AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly,
                XlsbImportOptions = new Xlsb.XlsbImportOptions {
                    CancellationToken = cancellation.Token
                }
            }));
    }

    [Fact]
    public void OpenDataReader_LegacyXlsObservesCancellationWhileBufferingInput() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Cancel.{Guid.NewGuid():N}.xls");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 2000; row++) {
                    sheet.CellValue(row, 1, "Value " + row.ToString(CultureInfo.InvariantCulture));
                }
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            using var stream = new CancelingReadStream(File.ReadAllBytes(path), cancellation, 1024);
            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocument.OpenDataReader(
                    stream,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxInfersSchemaIndependentlyOfDataTableInference() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Schema.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 42);
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    InferSchema = true,
                    InferDataTableColumnTypes = false
                });

            Assert.Equal(typeof(double), reader.GetFieldType(0));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_EmptyHeaderlessXlsxHasNoRows() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Empty.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Empty");
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions { HasHeaderRow = false });

            Assert.Equal(0, reader.FieldCount);
            Assert.False(reader.HasRows);
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbHonorsTheCellValueConverter() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            GetDataReaderXlsbFixture("basic-values-formula.xlsb"),
            new ExcelReadOptions {
                CellValueConverter = context =>
                    context.RawText == "42"
                        ? new ExcelCellValue("converted")
                        : ExcelCellValue.NotHandled
            });

        Assert.True(reader.Read());
        Assert.Equal("converted", reader.GetString(1));
        Assert.Equal(typeof(object), reader.GetFieldType(1));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_XlsbRejectsFormulaWithoutMandatoryPayloadTail(bool useConverter) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.TruncatedFormula.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            TruncateFirstXlsbNumericFormulaTail(path);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    path,
                    new ExcelReadOptions {
                        CellValueConverter = useConverter
                            ? static _ => ExcelCellValue.NotHandled
                            : null
                    }));
            Assert.Contains("formula record", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("token-byte count", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxInfersSchemaWithinSmallChunkLimit() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.SmallChunkSchema.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, 42);
                document.Save();
            }

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    InferSchema = true,
                    MaxDataReaderChunkRows = 1
                });

            Assert.Equal(typeof(double), reader.GetFieldType(0));
            Assert.True(reader.Read());
            Assert.Equal(42, reader.GetInt32(0));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsMissingCellStyleDuringDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidStyle.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            ReplaceFirstXlsbDataCellStyleIndex(path, 0x00FFFFFEU);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("missing cell format", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(47)]
    [InlineData(618)]
    public void OpenDataReader_XlsbRejectsTruncatedStylesPart(int truncateAfterRecordType) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.TruncatedStyles.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            TruncateXlsbStylesAfterRecord(path, truncateAfterRecordType);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("styles part", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("boundary", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsbRejectsMismatchedCellFormatCount() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.InvalidStyleCount.{Guid.NewGuid():N}.xlsb");
        File.Copy(GetDataReaderXlsbFixture("basic-values-formula.xlsb"), path);
        try {
            IncrementXlsbStyleCollectionDeclaredCount(path, 617);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(
                () => ExcelDocument.OpenDataReader(path));

            Assert.Contains("cell-format collection", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("declares", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("contains", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void ReadUsedRangeAsDataReaderObservesPreCancelledDiscovery() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.CancelDiscovery.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "One");
                document.Save();
            }

            using ExcelDocumentReader owner = ExcelDocumentReader.Open(
                path,
                new ExcelReadOptions());
            ExcelSheetReader sheetReader = owner.GetSheet("Data");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                sheetReader.ReadUsedRangeAsDataReader(
                    ct: cancellation.Token));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_XlsxObservesCancellationWhileBufferingInput() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Cancel.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 2000; row++) {
                    sheet.CellValue(row, 1, "Value " + row.ToString(CultureInfo.InvariantCulture));
                }
                document.Save();
            }

            using var cancellation = new CancellationTokenSource();
            using var stream = new CancelingReadStream(File.ReadAllBytes(path), cancellation, 1024);
            Assert.Throws<OperationCanceledException>(() =>
                ExcelDocument.OpenDataReader(
                    stream,
                    new ExcelReadOptions { CancellationToken = cancellation.Token }));
            Assert.True(stream.CanRead);
        } finally {
            File.Delete(path);
        }
    }

    private static string GetDataReaderXlsbFixture(string name) =>
        Path.Combine(
            AppContext.BaseDirectory,
            "Documents",
            "XlsbCorpus",
            "excel-generated",
            name);

    private static void ReplaceWorksheetDimension(string path, string declaredDimension) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? throw new InvalidDataException("The generated workbook has no first worksheet part.");
        string xml;
        using (var reader = new StreamReader(originalEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
            xml = reader.ReadToEnd();
        }

        int dimensionStart = xml.IndexOf("<dimension ref=\"", StringComparison.Ordinal);
        Assert.True(dimensionStart >= 0);
        int valueStart = dimensionStart + "<dimension ref=\"".Length;
        int valueEnd = xml.IndexOf('"', valueStart);
        Assert.True(valueEnd > valueStart);
        xml = xml.Substring(0, valueStart) + declaredDimension + xml.Substring(valueEnd);

        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/worksheets/sheet1.xml",
            CompressionLevel.Optimal);
        using var writer = new StreamWriter(
            replacement.Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(xml);
    }

    private static void ReplaceFirstWorksheetNamespace(
        string path,
        string sourceNamespace,
        string targetNamespace) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
            ?? throw new InvalidDataException("The generated workbook has no first worksheet part.");
        string xml;
        using (var reader = new StreamReader(
            originalEntry.Open(),
            Encoding.UTF8,
            detectEncodingFromByteOrderMarks: true)) {
            xml = reader.ReadToEnd();
        }

        Assert.Contains(sourceNamespace, xml, StringComparison.Ordinal);
        xml = xml.Replace(sourceNamespace, targetNamespace);

        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/worksheets/sheet1.xml",
            CompressionLevel.Optimal);
        using var writer = new StreamWriter(
            replacement.Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(xml);
    }

    private static void ReplaceXlsbWorksheetLastColumn(string path, uint lastColumn) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/worksheets/sheet1.bin")
            ?? throw new InvalidDataException("The XLSB fixture has no first worksheet part.");
        byte[] bytes;
        using (Stream input = originalEntry.Open()) {
            using var output = new MemoryStream();
            input.CopyTo(output);
            bytes = output.ToArray();
        }

        bool replaced = false;
        int position = 0;
        while (position < bytes.Length) {
            int firstTypeByte = bytes[position++];
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                type |= (bytes[position++] & 0x7F) << 7;
            }

            int size = 0;
            for (int index = 0; index < 4; index++) {
                int current = bytes[position++];
                size |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0 || index == 3) {
                    break;
                }
            }

            if (type == 148) {
                Assert.True(size >= 16);
                WriteUInt32LittleEndian(bytes, position + 12, lastColumn);
                replaced = true;
                break;
            }

            position = checked(position + size);
        }

        Assert.True(replaced);
        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/worksheets/sheet1.bin",
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, bytes.Length);
    }

    private static void TruncateFirstXlsbNumericFormulaTail(string path) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/worksheets/sheet1.bin")
            ?? throw new InvalidDataException("The XLSB fixture has no first worksheet part.");
        byte[] bytes;
        using (Stream input = originalEntry.Open()) {
            using var output = new MemoryStream();
            input.CopyTo(output);
            bytes = output.ToArray();
        }

        const int brtFmlaNum = 9;
        const int cachedPayloadBytes = sizeof(int) + sizeof(uint) + sizeof(double);
        byte[]? truncated = null;
        int position = 0;
        while (position < bytes.Length) {
            int firstTypeByte = bytes[position++];
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                type |= (bytes[position++] & 0x7F) << 7;
            }

            int sizeHeaderStart = position;
            int size = 0;
            int sizeHeaderLength = 0;
            for (int index = 0; index < 4; index++) {
                int current = bytes[position++];
                sizeHeaderLength++;
                size |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0) {
                    break;
                }
            }

            if (type == brtFmlaNum) {
                Assert.True(size > cachedPayloadBytes);
                Assert.Equal(1, sizeHeaderLength);
                int bytesToRemove = size - cachedPayloadBytes;
                truncated = new byte[bytes.Length - bytesToRemove];
                Buffer.BlockCopy(bytes, 0, truncated, 0, position + cachedPayloadBytes);
                truncated[sizeHeaderStart] = cachedPayloadBytes;
                Buffer.BlockCopy(
                    bytes,
                    position + size,
                    truncated,
                    position + cachedPayloadBytes,
                    bytes.Length - position - size);
                break;
            }

            position = checked(position + size);
        }

        Assert.NotNull(truncated);
        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/worksheets/sheet1.bin",
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(truncated, 0, truncated.Length);
    }

    private static void ReplaceFirstXlsbDataCellStyleIndex(
        string path,
        uint styleIndex,
        string entryName = "xl/worksheets/sheet1.bin") {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry(entryName)
            ?? throw new InvalidDataException($"The XLSB fixture has no '{entryName}' worksheet part.");
        byte[] bytes;
        using (Stream input = originalEntry.Open()) {
            using var output = new MemoryStream();
            input.CopyTo(output);
            bytes = output.ToArray();
        }

        bool replaced = false;
        int rowCount = 0;
        int position = 0;
        while (position < bytes.Length) {
            int firstTypeByte = bytes[position++];
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                type |= (bytes[position++] & 0x7F) << 7;
            }

            int size = 0;
            for (int index = 0; index < 4; index++) {
                int current = bytes[position++];
                size |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0 || index == 3) {
                    break;
                }
            }

            if (type == 0) {
                rowCount++;
            } else if (rowCount >= 2
                && (type is >= 1 and <= 11 || type == 62)) {
                Assert.True(size >= sizeof(int) + sizeof(uint));
                WriteUInt32LittleEndian(bytes, position + sizeof(int), styleIndex);
                replaced = true;
                break;
            }

            position = checked(position + size);
        }

        Assert.True(replaced);
        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            entryName,
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, bytes.Length);
    }

    private static void TruncateXlsbStylesAfterRecord(string path, int recordType) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/styles.bin")
            ?? throw new InvalidDataException("The XLSB fixture has no styles part.");
        byte[] bytes;
        using (Stream input = originalEntry.Open()) {
            using var output = new MemoryStream();
            input.CopyTo(output);
            bytes = output.ToArray();
        }

        int truncatedLength = -1;
        int position = 0;
        while (position < bytes.Length) {
            int firstTypeByte = bytes[position++];
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                type |= (bytes[position++] & 0x7F) << 7;
            }

            int size = 0;
            for (int index = 0; index < 4; index++) {
                int current = bytes[position++];
                size |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0 || index == 3) {
                    break;
                }
            }

            position = checked(position + size);
            if (type == recordType) {
                truncatedLength = position;
            }
        }

        Assert.InRange(truncatedLength, 1, bytes.Length - 1);
        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/styles.bin",
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, truncatedLength);
    }

    private static void IncrementXlsbStyleCollectionDeclaredCount(string path, int recordType) {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry originalEntry = archive.GetEntry("xl/styles.bin")
            ?? throw new InvalidDataException("The XLSB fixture has no styles part.");
        byte[] bytes;
        using (Stream input = originalEntry.Open()) {
            using var output = new MemoryStream();
            input.CopyTo(output);
            bytes = output.ToArray();
        }

        bool replaced = false;
        int position = 0;
        while (position < bytes.Length) {
            int firstTypeByte = bytes[position++];
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                type |= (bytes[position++] & 0x7F) << 7;
            }

            int size = 0;
            for (int index = 0; index < 4; index++) {
                int current = bytes[position++];
                size |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0 || index == 3) {
                    break;
                }
            }

            if (type == recordType) {
                Assert.Equal(sizeof(uint), size);
                uint declared = (uint)(
                    bytes[position]
                    | bytes[position + 1] << 8
                    | bytes[position + 2] << 16
                    | bytes[position + 3] << 24);
                WriteUInt32LittleEndian(bytes, position, checked(declared + 1U));
                replaced = true;
                break;
            }

            position = checked(position + size);
        }

        Assert.True(replaced);
        originalEntry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(
            "xl/styles.bin",
            CompressionLevel.Optimal);
        using Stream destination = replacement.Open();
        destination.Write(bytes, 0, bytes.Length);
    }

    private static void WriteUInt32LittleEndian(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private sealed class CancelingReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        private readonly int _maximumReadSize;

        internal CancelingReadStream(
            byte[] bytes,
            CancellationTokenSource cancellation,
            int maximumReadSize)
            : base(bytes, writable: false) {
            _cancellation = cancellation;
            _maximumReadSize = maximumReadSize;
        }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, Math.Min(count, _maximumReadSize));
            if (read > 0) {
                _cancellation.Cancel();
            }
            return read;
        }
    }

    private sealed class LargeLengthSeekableStream : Stream {
        private readonly long _length;
        private long _position;

        internal LargeLengthSeekableStream(long length) {
            _length = length;
        }

        internal bool ReadAttempted { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position {
            get => _position;
            set => _position = value;
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            ReadAttempted = true;
            throw new InvalidOperationException("The cancelled stream must not be read.");
        }

        public override long Seek(long offset, SeekOrigin origin) {
            _position = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => checked(_position + offset),
                SeekOrigin.End => checked(_length + offset),
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            return _position;
        }

        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
