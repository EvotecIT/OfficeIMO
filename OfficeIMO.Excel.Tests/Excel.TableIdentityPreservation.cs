using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EditedSavePreservesSparseTableColumnIdsAndDependentBindings(bool preflight, bool repairName) {
        using var input = new MemoryStream();
        using (var seed = ExcelDocument.Create(input)) {
            seed.AddSheet("Data");
            seed.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "SalesQuery",
                WorksheetName = "Data",
                TableName = "SalesResults",
                ColumnNames = new[] { "Region", "Date", "Amount" }
            });
            seed.Save();
        }
        input.Position = 0;
        using (var package = SpreadsheetDocument.Open(input, true)) {
            WorkbookPart workbook = package.WorkbookPart!;
            TableDefinitionPart tablePart = workbook.WorksheetParts.Single().TableDefinitionParts.Single();
            Table table = tablePart.Table!;
            table.Id = 5U;
            TableColumn[] columns = table.TableColumns!.Elements<TableColumn>().ToArray();
            uint[] ids = { 4U, 1U, 2U };
            for (int i = 0; i < ids.Length; i++) columns[i].Id = ids[i];
            columns[0].SetAttribute(new OpenXmlAttribute("p", "marker", "urn:preservation", "region"));
            table.TableColumns.SetAttribute(new OpenXmlAttribute("p", "marker", "urn:preservation", "columns"));
            if (repairName) {
                columns[0].Name = string.Empty;
                table.TableColumns.Count = 99U;
            }
            table.Save();
            QueryTable query = tablePart.QueryTableParts.Single().QueryTable!;
            QueryTableField[] fields = query.QueryTableRefresh!.QueryTableFields!.Elements<QueryTableField>().ToArray();
            for (int i = 0; i < ids.Length; i++) fields[i].TableColumnId = ids[i];
            query.Save();

            SlicerCachePart cache = workbook.AddNewPart<SlicerCachePart>();
            const string cacheXml = """
                <slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="RegionSlicer" sourceName="Region">
                  <extLst><ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}">
                    <tableSlicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" tableId="5" column="4" />
                  </ext></extLst>
                </slicerCacheDefinition>
                """;
            using (var content = new MemoryStream(Encoding.UTF8.GetBytes(cacheXml))) cache.FeedData(content);
            var extensions = workbook.Workbook.GetFirstChild<WorkbookExtensionList>();
            if (extensions == null) {
                extensions = new WorkbookExtensionList();
                workbook.Workbook.Append(extensions);
            }
            extensions.Append(new WorkbookExtension(new X14.SlicerCaches(new X14.SlicerCache { Id = workbook.GetIdOfPart(cache) })) {
                Uri = "{BBE1A952-AA13-448e-AADC-164F8A28A991}"
            });
            workbook.Workbook.Save();
        }

        input.Position = 0;
        using var document = ExcelDocument.Load(input);
        document["Data"].Cell(2, 1).SetValue("West");
        using var saved = new MemoryStream();
        document.Save(saved, new ExcelSaveOptions { SafePreflight = preflight });
        saved.Position = 0;
        using var reopened = SpreadsheetDocument.Open(saved, false);
        TableDefinitionPart persistedPart = reopened.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single();
        Table persisted = persistedPart.Table!;
        TableColumn[] persistedColumns = persisted.TableColumns!.Elements<TableColumn>().ToArray();
        Assert.Equal(5U, persisted.Id!.Value);
        Assert.Equal(new[] { 4U, 1U, 2U }, persistedColumns.Select(column => column.Id!.Value).ToArray());
        Assert.Equal("Region", persistedColumns[0].Name!.Value);
        Assert.Equal(3U, persisted.TableColumns.Count!.Value);
        Assert.Equal("region", persistedColumns[0].GetAttribute("marker", "urn:preservation").Value);
        Assert.Equal("columns", persisted.TableColumns.GetAttribute("marker", "urn:preservation").Value);
        Assert.Equal(new[] { 4U, 1U, 2U }, persistedPart.QueryTableParts.Single().QueryTable!
            .QueryTableRefresh!.QueryTableFields!.Elements<QueryTableField>().Select(field => field.TableColumnId!.Value).ToArray());
        Assert.Equal(new[] { 1U, 2U, 3U }, persistedColumns.Select(column => column.QueryTableFieldId!.Value).ToArray());
        using var cacheStream = reopened.WorkbookPart.SlicerCacheParts.Single().GetStream();
        XElement binding = XDocument.Load(cacheStream).Descendants()
            .Single(element => element.Name.LocalName == "tableSlicerCache");
        Assert.Equal("5", (string?)binding.Attribute("tableId"));
        Assert.Equal("4", (string?)binding.Attribute("column"));
    }
}
