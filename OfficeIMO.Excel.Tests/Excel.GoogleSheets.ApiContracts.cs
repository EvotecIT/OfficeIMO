using System.Reflection;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void GoogleSheetsPlanningApisUseBuildVocabulary() {
        string[] names = typeof(ExcelGoogleSheetsExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("BuildGoogleSheetsPlan", names);
        Assert.Contains("BuildGoogleSheetsBatch", names);
        Assert.DoesNotContain("CreateGoogleSheetsTranslationPlan", names);
        Assert.DoesNotContain("CreateGoogleSheetsBatch", names);
    }

    [Fact]
    public void GoogleSheetsReplacementDeletesConflictingTitlesBeforeAddingDesiredSheets() {
        string path = Path.Combine(_directoryWithFiles, "GoogleSheetsReplacementOrdering.xlsx");
        try {
            using var document = ExcelDocument.Create(path);
            document.AddWorksheet("Sheet1").CellValue(1, 1, "One");
            document.AddWorksheet("Data").CellValue(1, 1, "Two");
            GoogleSheetsBatch batch = document.BuildGoogleSheetsBatch();
            var existingSheets = new Dictionary<int, string> {
                [10] = "Sheet1",
                [20] = "Data",
            };
            IReadOnlyDictionary<string, int> desiredIds = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch, existingSheets.Keys);

            GoogleSheetsApiBatchUpdatePayload payload = GoogleSheetsApiPayloadBuilder.BuildReplaceSpreadsheetPayload(batch, existingSheets, desiredIds);

            GoogleSheetsApiRequestPayload keeperAdd = Assert.Single(payload.Requests, request => request.AddSheet?.Properties.Title == "__OfficeIMO_Replacement_Keeper__");
            int keeperId = keeperAdd.AddSheet!.Properties.SheetId;
            int lastExistingDelete = payload.Requests.FindLastIndex(request => request.DeleteSheet != null && existingSheets.ContainsKey(request.DeleteSheet.SheetId));
            int firstDesiredAdd = payload.Requests.FindIndex(request => request.AddSheet != null && desiredIds.ContainsKey(request.AddSheet.Properties.Title));
            Assert.True(lastExistingDelete < firstDesiredAdd);
            Assert.Equal(keeperId, payload.Requests.Last().DeleteSheet?.SheetId);
            Assert.All(desiredIds, pair => Assert.Contains(payload.Requests, request => request.AddSheet?.Properties.Title == pair.Key && request.AddSheet.Properties.SheetId == pair.Value));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GoogleSheetsCreationSizesGridFromWorkbookUsedRange() {
        string path = Path.Combine(_directoryWithFiles, "GoogleSheetsGridSizing.xlsx");
        try {
            using var document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 27, "AA");
            sheet.CellValue(1001, 1, "Row");

            GoogleSheetsBatch batch = document.BuildGoogleSheetsBatch();
            GoogleSheetsAddSheetRequest addSheet = Assert.Single(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == "Data");
            Assert.Equal(1001, addSheet.RowCount);
            Assert.Equal(27, addSheet.ColumnCount);

            GoogleSheetsApiCreateSpreadsheetPayload createPayload = GoogleSheetsApiPayloadBuilder.BuildCreateSpreadsheetPayload(batch);
            GoogleSheetsApiSheetPropertiesPayload createProperties = Assert.Single(createPayload.Sheets).Properties;
            Assert.Equal(1001, createProperties.GridProperties.RowCount);
            Assert.Equal(27, createProperties.GridProperties.ColumnCount);

            var existingSheets = new Dictionary<int, string> { [10] = "Data" };
            IReadOnlyDictionary<string, int> desiredIds = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch, existingSheets.Keys);
            GoogleSheetsApiBatchUpdatePayload replacementPayload = GoogleSheetsApiPayloadBuilder.BuildReplaceSpreadsheetPayload(batch, existingSheets, desiredIds);
            GoogleSheetsApiSheetPropertiesPayload replacementProperties = Assert.Single(
                replacementPayload.Requests,
                request => request.AddSheet?.Properties.Title == "Data").AddSheet!.Properties;
            Assert.Equal(1001, replacementProperties.GridProperties.RowCount);
            Assert.Equal(27, replacementProperties.GridProperties.ColumnCount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GoogleSheetsNativeTablesUseDoubleForOrdinaryNumbers() {
        string path = Path.Combine(_directoryWithFiles, "GoogleSheetsNumericTable.xlsx");
        try {
            using var document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Amount");
            sheet.CellValue(2, 1, 10.5d);
            sheet.CellValue(3, 1, 20.25d);
            sheet.AddTable("A1:A3", hasHeader: true, name: "Amounts", style: TableStyle.TableStyleMedium2);

            GoogleSheetsBatch batch = document.BuildGoogleSheetsBatch();
            GoogleSheetsAddTableRequest table = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>());
            Assert.Equal("DOUBLE", Assert.Single(table.Columns).ColumnType);

            GoogleSheetsApiBatchUpdatePayload payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
            GoogleSheetsApiTableColumnPropertiesPayload column = Assert.Single(
                Assert.Single(payload.Requests, request => request.AddTable != null).AddTable!.Table.ColumnProperties!);
            Assert.Equal("DOUBLE", column.ColumnType);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
