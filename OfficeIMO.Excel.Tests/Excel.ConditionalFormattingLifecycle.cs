using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ConditionalFormattingLifecycle_UnifiesStandardAndOfficeExtensionRules() {
            string path = Path.Combine(_directoryWithFiles, "ConditionalFormattingLifecycle.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                for (int row = 1; row <= 3; row++) sheet.CellAt(row, 1).SetValue(row - 2);

                ExcelConditionalFormattingInfo extension = sheet.AddConditionalFormattingRule(
                    new ExcelConditionalFormattingInfo {
                        Source = ExcelConditionalFormattingSource.Office2010Extension,
                        Range = "A1:A3",
                        Type = "DataBar",
                        Priority = 2,
                        DataBarColor = "FF4472C4",
                        DataBarBorderColor = "FF203864",
                        DataBarNegativeColor = "FFC00000",
                        DataBarAxisColor = "FF000000",
                        DataBarMinimumLength = 5,
                        DataBarMaximumLength = 95,
                        DataBarBorder = true,
                        DataBarGradient = false,
                        DataBarDirection = "LeftToRight",
                        DataBarAxisPosition = "Middle",
                        DataBarThresholds = new[] {
                            new ExcelConditionalFormatThreshold { Type = "AutoMin" },
                            new ExcelConditionalFormatThreshold { Type = "AutoMax" }
                        }
                    });
                ExcelConditionalFormattingInfo standard = sheet.AddConditionalFormattingRule(
                    new ExcelConditionalFormattingInfo {
                        Range = "A1:A3",
                        Type = "Expression",
                        Priority = 1,
                        StopIfTrue = true,
                        Formulas = new[] { "A1<0" }
                    });

                IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules("A1:A3");
                Assert.Equal(2, rules.Count);
                ExcelConditionalFormattingInfo inspectedExtension = Assert.Single(rules, item =>
                    item.Source == ExcelConditionalFormattingSource.Office2010Extension);
                Assert.Equal("DataBar", inspectedExtension.Type);
                Assert.Equal("FF4472C4", inspectedExtension.DataBarColor);
                Assert.Equal("FFC00000", inspectedExtension.DataBarNegativeColor);
                Assert.Equal((uint)5, inspectedExtension.DataBarMinimumLength);
                Assert.Equal("middle", inspectedExtension.DataBarAxisPosition, ignoreCase: true);

                extension.Range = "B1:B3";
                extension.DataBarMaximumLength = 90;
                sheet.UpdateConditionalFormattingRule(extension);
                ExcelConditionalFormattingInfo clone = sheet.CloneConditionalFormattingRule(extension, "C1:C3");
                sheet.ReorderConditionalFormattingRules(new[] { extension, clone, standard });
                Assert.Equal(1, extension.Priority);
                Assert.Equal(2, clone.Priority);
                Assert.Equal(3, standard.Priority);
                Assert.Equal(new[] { 1, 2, 3 }, sheet.GetConditionalFormattingRules()
                    .Select(item => item.Priority)
                    .OrderBy(priority => priority));
                Assert.NotEqual(extension.ExtensionId, clone.ExtensionId);

                ExcelFeatureFinding feature = Assert.Single(document.InspectFeatures().FindFeatures("Conditional formatting"));
                Assert.Equal(OfficeFeatureSupportLevel.Editable, feature.SupportLevel);
                Assert.Equal(3, feature.Count);

                OfficeImageExportResult rendered = sheet.Range("B1:B3").ExportImage(OfficeImageExportFormat.Png);
                OfficeImageExportDiagnostic approximation = Assert.Single(rendered.Diagnostics, item =>
                    item.Code == ExcelImageExportDiagnosticCodes.ConditionalExtensionApproximation);
                Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, approximation.Severity);
                Assert.Equal(OfficeConversionLossKind.Approximation, approximation.LossKind);

                sheet.RemoveConditionalFormattingRule(standard);
                document.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                ExcelSheet sheet = document.Sheets[0];
                IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules();
                Assert.Equal(2, rules.Count);
                Assert.All(rules, item => Assert.Equal(ExcelConditionalFormattingSource.Office2010Extension, item.Source));
                Assert.Equal(new[] { "B1:B3", "C1:C3" }, rules.OrderBy(item => item.Priority).Select(item => item.Range));
                Assert.All(rules, item => Assert.Equal((uint)90, item.DataBarMaximumLength));
                Assert.Empty(document.ValidateOpenXml());
                sheet.ClearConditionalFormatting("B1:B3");
                Assert.Single(sheet.GetConditionalFormattingRules());
                sheet.ClearConditionalFormatting();
                Assert.Empty(sheet.GetConditionalFormattingRules());
            }
        }

        [Fact]
        public void ConditionalFormattingLifecycle_PreservesImportedUnknownMarkupDuringUpdateAndClone() {
            string path = Path.Combine(_directoryWithFiles, "ConditionalFormattingUnknownRoundTrip.xlsx");
            const string vendorNamespace = "urn:officeimo:test:conditional-formatting";
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Imported");
                var unknown = new OpenXmlUnknownElement("v", "state", vendorNamespace);
                unknown.SetAttribute(new OpenXmlAttribute("mode", string.Empty, "retained"));
                var importedRule = new X14.ConditionalFormattingRule(
                    new Xm.Formula("A1>0"),
                    new X14.ExtensionList(unknown)) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 7,
                    Id = "{E98C471C-34A0-47B6-B139-7623F7B48D76}"
                };
                importedRule.SetAttribute(new OpenXmlAttribute("v", "owner", vendorNamespace, "producer"));
                var importedContainer = new X14.ConditionalFormatting(
                    importedRule,
                    new Xm.ReferenceSequence("A1:A2"));
                importedContainer.SetAttribute(new OpenXmlAttribute("v", "container", vendorNamespace, "retained-owner"));
                sheet.WorksheetPart.Worksheet.Append(
                    new WorksheetExtensionList(
                        new WorksheetExtension(new X14.ConditionalFormattings(importedContainer)) {
                            Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                        }));

                ExcelConditionalFormattingInfo imported = Assert.Single(sheet.GetConditionalFormattingRules());
                Assert.True(imported.HasPreservedUnknownMarkup);
                imported.Range = "B1:B2";
                imported.Priority = 1;
                imported.Formulas = new[] { "B1>0" };
                sheet.UpdateConditionalFormattingRule(imported);
                ExcelConditionalFormattingInfo clone = sheet.CloneConditionalFormattingRule(imported, "C1:C2", priority: 2);
                Assert.True(clone.HasPreservedUnknownMarkup);
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
            Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            X14.ConditionalFormattingRule[] rules = worksheet.Descendants<X14.ConditionalFormattingRule>().ToArray();
            Assert.Equal(2, rules.Length);
            Assert.Equal(new[] { "B1:B2", "C1:C2" }, worksheet.Descendants<X14.ConditionalFormatting>()
                .Select(item => item.GetFirstChild<Xm.ReferenceSequence>()!.Text));
            Assert.Equal(new[] { "B1>0", "B1>0" }, rules.Select(item => item.GetFirstChild<Xm.Formula>()!.Text));
            Assert.All(rules, item => Assert.Equal("producer", item.GetAttribute("owner", vendorNamespace).Value));
            Assert.All(worksheet.Descendants<X14.ConditionalFormatting>(), item =>
                Assert.Equal("retained-owner", item.GetAttribute("container", vendorNamespace).Value));
            Assert.All(rules, item => Assert.Equal("retained", item.Descendants<OpenXmlUnknownElement>()
                .Single(element => element.LocalName == "state")
                .GetAttribute("mode", string.Empty).Value));
        }

        [Fact]
        public void ConditionalFormattingLifecycle_RejectsUnknownAuthoredTypesButEditsImportedTypes() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Imported");

            Assert.Throws<ArgumentException>(() => sheet.AddConditionalFormattingRule(
                new ExcelConditionalFormattingInfo {
                    Range = "A1",
                    Type = "futureRule"
                }));
            Assert.Empty(sheet.GetConditionalFormattingRules());

            var importedRule = new X14.ConditionalFormattingRule {
                Priority = 1,
                Id = "{9AD8DD4D-C1D1-4978-93D7-A92C6B3C20A1}"
            };
            importedRule.SetAttribute(new OpenXmlAttribute("type", string.Empty, "futureRule"));
            var importedContainer = new X14.ConditionalFormatting(
                importedRule,
                new Xm.ReferenceSequence("A1"));
            sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                new WorksheetExtension(new X14.ConditionalFormattings(importedContainer)) {
                    Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                }));

            ExcelConditionalFormattingInfo imported = Assert.Single(sheet.GetConditionalFormattingRules());
            Assert.Equal("futureRule", imported.Type);
            imported.Range = "B2";
            imported.Priority = 2;
            sheet.UpdateConditionalFormattingRule(imported);

            Assert.Equal("futureRule", importedRule.GetAttribute("type", string.Empty).Value);
            Assert.Equal("B2", importedContainer.GetFirstChild<Xm.ReferenceSequence>()!.Text);
            Assert.Equal(2, importedRule.Priority!.Value);
        }

        [Fact]
        public void ConditionalFormattingLifecycle_ClearPreservesUnrelatedWorksheetExtensions() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Extensions");
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Office2010Extension,
                Range = "A1",
                Type = "Expression",
                Formulas = new[] { "A1>0" }
            });
            WorksheetExtensionList extensions = sheet.WorksheetPart.Worksheet
                .GetFirstChild<WorksheetExtensionList>()!;
            var unrelated = new WorksheetExtension { Uri = "{FA4B2191-6B46-40F0-8D7F-1039A67097C9}" };
            unrelated.SetAttribute(new OpenXmlAttribute("v", "state", "urn:officeimo:test:extension", "retained"));
            extensions.Append(unrelated);

            sheet.ClearConditionalFormatting();

            WorksheetExtension retained = Assert.Single(
                sheet.WorksheetPart.Worksheet.GetFirstChild<WorksheetExtensionList>()!
                    .Elements<WorksheetExtension>());
            Assert.Same(unrelated, retained);
            Assert.Equal("retained", retained.GetAttribute("state", "urn:officeimo:test:extension").Value);
        }

        [Fact]
        public void ConditionalFormattingLifecycle_LegacyXlsRejectsOfficeExtensionRulesInsteadOfDroppingThem() {
            AssertNativeXlsSaveNotSupported("Office 2010 conditional formatting extension rules", (_, sheet) => {
                sheet.CellAt(1, 1).SetValue(1);
                sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                    Source = ExcelConditionalFormattingSource.Office2010Extension,
                    Range = "A1",
                    Type = "Expression",
                    Formulas = new[] { "A1>0" }
                });
            });
        }

        [Fact]
        public void ConditionalFormattingLifecycle_AuthorsOfficeExtensionColorScalesAndCustomIcons() {
            string path = Path.Combine(_directoryWithFiles, "ConditionalFormattingExtensionVisuals.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Visuals");
                for (int row = 1; row <= 4; row++) sheet.CellAt(row, 1).SetValue(row);

                sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                    Source = ExcelConditionalFormattingSource.Office2010Extension,
                    Range = "A1:A4",
                    Type = "ColorScale",
                    ColorScaleThresholds = new[] {
                        new ExcelConditionalFormatThreshold { Type = "Min" },
                        new ExcelConditionalFormatThreshold { Type = "Percentile", Value = "50" },
                        new ExcelConditionalFormatThreshold { Type = "Max" }
                    },
                    ColorScaleColors = new[] { "FFF8696B", "FFFFEB84", "FF63BE7B" }
                });
                sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                    Source = ExcelConditionalFormattingSource.Office2010Extension,
                    Range = "A1:A4",
                    Type = "IconSet",
                    IconSet = "ThreeStars",
                    IconSetCustom = true,
                    IconSetShowValue = false,
                    IconSetThresholds = new[] {
                        new ExcelConditionalIconSetThreshold { Type = "Percent", Value = "0" },
                        new ExcelConditionalIconSetThreshold { Type = "Percent", Value = "33" },
                        new ExcelConditionalIconSetThreshold { Type = "Percent", Value = "67" }
                    },
                    CustomIcons = new[] {
                        new ExcelConditionalFormatIcon { IconSet = "ThreeSymbols", IconId = 0 },
                        new ExcelConditionalFormatIcon { IconSet = "ThreeSymbols", IconId = 1 },
                        new ExcelConditionalFormatIcon { IconSet = "ThreeSymbols", IconId = 2 }
                    }
                });
                document.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                IReadOnlyList<ExcelConditionalFormattingInfo> rules = document.Sheets[0].GetConditionalFormattingRules();
                ExcelConditionalFormattingInfo scale = Assert.Single(rules, item => item.Type == "ColorScale");
                Assert.Equal(3, scale.ColorScaleColors.Count);
                Assert.Equal("50", scale.ColorScaleThresholds[1].Value);
                ExcelConditionalFormattingInfo icons = Assert.Single(rules, item => item.Type == "IconSet");
                Assert.True(icons.IconSetCustom);
                Assert.False(icons.IconSetShowValue);
                Assert.Equal(3, icons.CustomIcons.Count);
                IReadOnlyList<string> validation = document.ValidateOpenXml();
                Assert.True(validation.Count == 0, string.Join(Environment.NewLine, validation));
            }
        }

        [Fact]
        public void ConditionalFormattingLifecycle_CommonEditsPreserveUnprojectedVisualsAndValidateBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Imported");
            var firstThreshold = new X14.ConditionalFormattingValueObject();
            firstThreshold.SetAttribute(new OpenXmlAttribute("type", string.Empty, "autoMin"));
            var secondThreshold = new X14.ConditionalFormattingValueObject();
            secondThreshold.SetAttribute(new OpenXmlAttribute("type", string.Empty, "autoMax"));
            var themeFill = new X14.FillColor { Theme = 4U, Tint = 0.25D };
            themeFill.SetAttribute(new OpenXmlAttribute("v", "profile", "urn:officeimo:test:conditional-formatting", "theme"));
            var dataBar = new X14.DataBar(firstThreshold, secondThreshold, themeFill) {
                MinLength = 10,
                MaxLength = 90,
                ShowValue = true
            };
            string importedVisual = dataBar.OuterXml;
            var importedRule = new X14.ConditionalFormattingRule(dataBar) {
                Type = ConditionalFormatValues.DataBar,
                Priority = 1,
                Id = "{55ED4312-5798-4510-89B8-2D917154423B}"
            };
            var importedContainer = new X14.ConditionalFormatting(
                importedRule,
                new Xm.ReferenceSequence("A1:A2"));
            sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                new WorksheetExtension(new X14.ConditionalFormattings(importedContainer)) {
                    Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                }));

            ExcelConditionalFormattingInfo imported = Assert.Single(sheet.GetConditionalFormattingRules());
            Assert.Null(imported.DataBarColor);
            imported.Range = "B1:B2";
            sheet.UpdateConditionalFormattingRule(imported);
            Assert.Equal(importedVisual, dataBar.OuterXml);
            Assert.Equal("B1:B2", importedContainer.GetFirstChild<Xm.ReferenceSequence>()!.Text);

            imported.Range = "C1:C2";
            imported.DataBarMaximumLength = 101;
            Assert.ThrowsAny<ArgumentException>(() => sheet.UpdateConditionalFormattingRule(imported));
            Assert.Equal("B1:B2", importedContainer.GetFirstChild<Xm.ReferenceSequence>()!.Text);
            Assert.Equal((uint)90, dataBar.MaxLength!.Value);
        }

        [Fact]
        public void ConditionalFormattingLifecycle_RangeUpdateSplitsSharedImportedOwners() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Imported");
            var first = new ConditionalFormattingRule(new Formula("A1>0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 1
            };
            var second = new ConditionalFormattingRule(new Formula("A1<0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 2
            };
            sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(first, second) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3" }
            });

            ExcelConditionalFormattingInfo moved = sheet.GetConditionalFormattingRules()
                .Single(rule => rule.Priority == 1);
            moved.Range = "B1:B3";
            sheet.UpdateConditionalFormattingRule(moved);

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules();
            Assert.Equal("B1:B3", rules.Single(rule => rule.Priority == 1).Range);
            Assert.Equal("A1:A3", rules.Single(rule => rule.Priority == 2).Range);
            Assert.Equal(2, sheet.WorksheetPart.Worksheet.Elements<ConditionalFormatting>().Count());
        }

        [Fact]
        public void ConditionalFormattingLifecycle_AuthorsStandardAndInlineExtensionDifferentialStyles() {
            string path = Path.Combine(_directoryWithFiles, "ConditionalFormattingDifferentialStyles.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Styles");
                sheet.CellAt(1, 1).SetValue(1);
                sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                    Range = "A1",
                    Type = "Expression",
                    Formulas = new[] { "A1>0" },
                    DifferentialFillColorArgb = "FFFFC000",
                    DifferentialFontColorArgb = "FF9C0006",
                    DifferentialFontBold = true
                });
                sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                    Source = ExcelConditionalFormattingSource.Office2010Extension,
                    Range = "A1",
                    Type = "Expression",
                    Formulas = new[] { "A1<0" },
                    DifferentialFillColorArgb = "FFC6EFCE",
                    DifferentialFontColorArgb = "FF006100",
                    DifferentialFontItalic = true
                });
                document.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                IReadOnlyList<ExcelConditionalFormattingInfo> rules = document.Sheets[0].GetConditionalFormattingRules();
                ExcelConditionalFormattingInfo standard = Assert.Single(rules, rule => rule.Source == ExcelConditionalFormattingSource.Standard);
                Assert.NotNull(standard.DifferentialFormatId);
                Assert.Equal("FFFFC000", standard.DifferentialFillColorArgb);
                Assert.Equal("FF9C0006", standard.DifferentialFontColorArgb);
                Assert.True(standard.DifferentialFontBold);
                ExcelConditionalFormattingInfo extension = Assert.Single(rules, rule => rule.Source == ExcelConditionalFormattingSource.Office2010Extension);
                Assert.Null(extension.DifferentialFormatId);
                Assert.Equal("FFC6EFCE", extension.DifferentialFillColorArgb);
                Assert.Equal("FF006100", extension.DifferentialFontColorArgb);
                Assert.True(extension.DifferentialFontItalic);
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
