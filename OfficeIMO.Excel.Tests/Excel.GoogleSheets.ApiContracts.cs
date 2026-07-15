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
}
