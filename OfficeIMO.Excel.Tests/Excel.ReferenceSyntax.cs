using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("$B$4", 4, 2, true, true)]
        [InlineData("B$4", 4, 2, true, false)]
        [InlineData("$B4", 4, 2, false, true)]
        [InlineData("'Sales 2026'!B4", 4, 2, false, false)]
        public void Test_ReferenceSyntax_ParsesA1Cells(
            string text,
            int row,
            int column,
            bool rowAbsolute,
            bool columnAbsolute) {
            ExcelReference reference = ExcelReference.Parse(text);

            Assert.Equal(ExcelReferenceKind.Cell, reference.Kind);
            Assert.Equal(row, reference.Start.Row);
            Assert.Equal(column, reference.Start.Column);
            Assert.Equal(rowAbsolute, reference.Start.RowAbsolute);
            Assert.Equal(columnAbsolute, reference.Start.ColumnAbsolute);
            Assert.Equal(text, reference.ToString(ExcelReferenceStyle.A1));
        }

        [Fact]
        public void Test_ReferenceSyntax_ConvertsRelativeR1C1AgainstAnchor() {
            ExcelReference reference = ExcelReference.Parse(
                "R[-2]C[3]:R5C1",
                ExcelReferenceStyle.R1C1,
                anchorRow: 10,
                anchorColumn: 4);

            Assert.Equal("G8:$A$5", reference.ToString(ExcelReferenceStyle.A1));
            Assert.Equal(
                "R[-2]C[3]:R5C1",
                reference.ToString(ExcelReferenceStyle.R1C1, anchorRow: 10, anchorColumn: 4));
        }

        [Fact]
        public void Test_ReferenceSyntax_ProvidesRectangularAlgebra() {
            ExcelReference source = ExcelReference.Parse("'Data'!A1:D4");
            ExcelReference cutout = ExcelReference.Parse("'Data'!B2:C3");

            Assert.True(source.Contains(4, 4));
            Assert.True(source.Intersects(cutout));
            Assert.Equal("'Data'!B2:C3", source.Intersect(cutout)!.ToString());
            Assert.Equal("'Data'!A1:E5", source.BoundingUnion(ExcelReference.Parse("'Data'!E5")).ToString());
            Assert.Equal(
                new[] { "'Data'!A1:D1", "'Data'!A4:D4", "'Data'!A2:A3", "'Data'!D2:D3" },
                source.Except(cutout).Select(range => range.ToString()).ToArray());
        }

        [Fact]
        public void Test_FormulaSyntaxTree_PreservesLiteralsAndRewritesReferencesOnce() {
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(
                "=SUM('Sales 2026'!A1:B2,\"A1 and \"\"quoted\"\" B2\",Table1[Amount],[Book.xlsx]Sheet1!C3#)");

            ExcelFormulaReferenceSyntax[] references = tree.Nodes.OfType<ExcelFormulaReferenceSyntax>().ToArray();
            Assert.Equal(2, references.Length);
            Assert.Equal("'Sales 2026'!A1:B2", references[0].Text);
            Assert.Equal("[Book.xlsx]Sheet1!C3#", references[1].Text);
            ExcelFormulaStructuredReferenceSyntax structured = Assert.Single(
                tree.Nodes.OfType<ExcelFormulaStructuredReferenceSyntax>());
            Assert.Equal("Table1", structured.TableName);
            Assert.Equal("[Amount]", structured.Selector);

            string rewritten = tree.Rewrite(reference => reference.Offset(1, 1));
            Assert.Equal(
                "=SUM('Sales 2026'!B2:C3,\"A1 and \"\"quoted\"\" B2\",Table1[Amount],[Book.xlsx]Sheet1!D4#)",
                rewritten);
        }

        [Fact]
        public void Test_FormulaSyntaxTree_ModelsAndRewritesNamesAndStructuredReferences() {
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse("=TaxRate*Table1[[#Data],[Net]]+[@Net]");

            Assert.Equal(new[] { "TaxRate" }, tree.Nodes.OfType<ExcelFormulaNameSyntax>().Select(node => node.Name));
            Assert.Equal(2, tree.Nodes.OfType<ExcelFormulaStructuredReferenceSyntax>().Count());
            Assert.Equal(
                "=Rate2026*Table1[[#Data],[Net]]+[@Net]",
                tree.RewriteNames(name => name == "TaxRate" ? "Rate2026" : name));
            Assert.Equal(
                "=TaxRate*Ledger[[#Data],[Amount]]+[@Amount]",
                tree.RewriteStructuredReferences((table, selector) =>
                    (table == "Table1" ? "Ledger" : table) + selector.Replace("Net", "Amount")));
        }

        [Fact]
        public void Test_FormulaSyntaxTree_DoesNotTreatQualifiedDefinedNameAsTableName() {
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse("=Sales!TaxRate+Sales[Amount]");

            Assert.Equal(new[] { "Sales!TaxRate" }, tree.Nodes.OfType<ExcelFormulaNameSyntax>().Select(node => node.Name));
            Assert.Equal("=Sales!TaxRate+Ledger[Amount]", tree.RewriteTableNames(name => name == "Sales" ? "Ledger" : name));
        }

        [Fact]
        public void Test_FormulaSyntaxTree_ConvertsReferencesToR1C1() {
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse("=A1+$B$2+C$3+$D4");

            Assert.Equal(
                "=R[-4]C[-4]+R2C2+R3C[-2]+R[-1]C4",
                tree.ConvertReferences(ExcelReferenceStyle.R1C1, anchorRow: 5, anchorColumn: 5));
        }

        [Fact]
        public void Test_FormulaInspection_ReportsExplicitLifecycleAndArrayStates() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 2);
            sheet.CellFormula(1, 2, "A1*2");
            sheet.CellFormula(1, 3, "UNSUPPORTED(A1)");
            sheet.SetArrayFormula("D1:D2", "A1:A2*2");

            ExcelFormulaCellInfo deferred = sheet.GetFormulaCells().Single(item => item.CellReference == "B1");
            Assert.True(deferred.State.HasFlag(ExcelFormulaState.Authored));
            Assert.True(deferred.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(deferred.State.HasFlag(ExcelFormulaState.Cached));

            Assert.Equal(1, sheet.RecalculateSupportedFormulas());
            ExcelFormulaCellInfo evaluated = sheet.GetFormulaCells().Single(item => item.CellReference == "B1");
            Assert.True(evaluated.State.HasFlag(ExcelFormulaState.Cached));
            Assert.True(evaluated.State.HasFlag(ExcelFormulaState.Evaluated));
            Assert.False(evaluated.State.HasFlag(ExcelFormulaState.Deferred));

            ExcelFormulaCellInfo unsupported = sheet.GetFormulaCells().Single(item => item.CellReference == "C1");
            Assert.True(unsupported.State.HasFlag(ExcelFormulaState.Unsupported));
            Assert.True(unsupported.State.HasFlag(ExcelFormulaState.Deferred));

            ExcelFormulaCellInfo array = sheet.GetFormulaCells().Single(item => item.CellReference == "D1");
            Assert.Equal("D1:D2", array.Array!.Range);
            Assert.False(array.IsDynamicArray);
        }

        [Fact]
        public void Test_FormulaInspection_ResolvesDynamicArrayMetadataPerCell() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetArrayFormula("A1:A2", "ROW(A1:A2)");
            sheet.SetArrayFormula("B1:B2", "ROW(B1:B2)");
            Cell[] formulaCells = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Where(cell => cell.CellFormula != null).ToArray();
            formulaCells.Single(cell => cell.CellReference?.Value == "A1").SetAttribute(new OpenXmlAttribute("", "cm", "", "1"));
            formulaCells.Single(cell => cell.CellReference?.Value == "B1").SetAttribute(new OpenXmlAttribute("", "cm", "", "2"));
            CellMetadataPart metadataPart = document.WorkbookPartRoot.AddNewPart<CellMetadataPart>();
            metadataPart.Metadata = new Metadata {
                InnerXml =
                    "<x:metadataTypes xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\"><x:metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\"/><x:metadataType name=\"OTHER\" minSupportedVersion=\"120000\"/></x:metadataTypes>" +
                    "<x:futureMetadata xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"XLDAPR\" count=\"2\"><x:bk><x:extLst><x:ext uri=\"{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}\"><xda:dynamicArrayProperties xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\" fDynamic=\"1\" fCollapsed=\"0\"/></x:ext></x:extLst></x:bk><x:bk><x:extLst><x:ext uri=\"{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}\"><xda:dynamicArrayProperties xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\" fDynamic=\"1\" fCollapsed=\"1\"/></x:ext></x:extLst></x:bk></x:futureMetadata>" +
                    "<x:cellMetadata xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\"><x:bk><x:rc t=\"1\" v=\"0\"/></x:bk><x:bk><x:rc t=\"2\" v=\"1\"/></x:bk></x:cellMetadata>"
            };
            MetadataBlock dynamicCellBlock = metadataPart.Metadata.GetFirstChild<CellMetadata>()!
                .Elements<MetadataBlock>().First();
            MetadataRecord dynamicRecord = Assert.Single(dynamicCellBlock.Elements<MetadataRecord>());
            Assert.Equal(1U, dynamicRecord.TypeIndex!.Value);
            Assert.Equal("XLDAPR", metadataPart.Metadata.GetFirstChild<MetadataTypes>()!
                .Elements<MetadataType>().First().Name!.Value);
            FutureMetadataBlock dynamicFutureBlock = metadataPart.Metadata.Elements<FutureMetadata>()
                .Single().Elements<FutureMetadataBlock>().First();
            Assert.Contains(dynamicFutureBlock.Descendants(), item => item.LocalName == "dynamicArrayProperties");

            ExcelFormulaCellInfo dynamic = sheet.GetFormulaCells().Single(item => item.CellReference == "A1");
            ExcelFormulaCellInfo ordinary = sheet.GetFormulaCells().Single(item => item.CellReference == "B1");
            Assert.True(dynamic.Array!.IsDynamic, metadataPart.Metadata.OuterXml);
            Assert.False(dynamic.Array.IsCollapsed);
            Assert.False(ordinary.Array!.IsDynamic);
            Assert.False(ordinary.Array.IsCollapsed);
        }

        [Fact]
        public void Test_FormulaSearch_MatchesFunctionsTextAndParsedReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellFormula(1, 1, "SUM(B1:B4)");
            data.CellFormula(2, 1, "IF(B2>0,\"SUM(B3)\",0)");
            summary.CellFormula(1, 1, "_xlfn.XLOOKUP(A1,'Data'!B1:B4,'Data'!C1:C4)");

            Assert.Single(data.SearchFormulas(new ExcelFormulaSearchOptions { Function = "SUM" }));
            Assert.Equal(2, data.SearchFormulas(new ExcelFormulaSearchOptions { Text = "B" }).Count);
            ExcelFormulaCellInfo[] referenceMatches = document.SearchFormulas(
                new ExcelFormulaSearchOptions { Reference = "'Data'!B2" }).ToArray();
            Assert.Equal(3, referenceMatches.Length);
            Assert.Contains(referenceMatches, item => item.SheetName == "Summary" && item.CellReference == "A1");
            Assert.Single(document.SearchFormulas(new ExcelFormulaSearchOptions { Function = "XLOOKUP" }));
        }

        [Fact]
        public void Test_FormulaSearch_DistinguishesExternalWorkbookQualifiers() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "[BookA.xlsx]Data!B2");
            sheet.CellFormula(2, 1, "[BookB.xlsx]Data!B2");

            ExcelFormulaCellInfo match = Assert.Single(document.SearchFormulas(
                new ExcelFormulaSearchOptions { Reference = "[BookA.xlsx]Data!B2" }));

            Assert.Equal("A1", match.CellReference);
        }
    }
}
