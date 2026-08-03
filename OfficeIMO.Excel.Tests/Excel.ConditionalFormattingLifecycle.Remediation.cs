using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ConditionalFormattingLifecycle_AuthorsSchemaLexicalNamesAndChildOrder() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Schema");
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Office2010Extension,
                Range = "A1:A3",
                Type = "Expression",
                Formulas = new[] { "A1>0" },
                DifferentialFillColorArgb = "FFFFC000"
            });
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Office2010Extension,
                Range = "B1:B3",
                Type = "IconSet",
                IconSet = "ThreeStars",
                IconSetCustom = true,
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

            X14.ConditionalFormattingRule expression = sheet.WorksheetPart.Worksheet
                .Descendants<X14.ConditionalFormattingRule>()
                .Single(rule => rule.Type?.InnerText == "expression");
            Assert.True(expression.ChildElements.ToList().FindIndex(child => child is Xm.Formula) <
                expression.ChildElements.ToList().FindIndex(child => child is X14.DifferentialType));
            X14.IconSet iconSet = sheet.WorksheetPart.Worksheet.Descendants<X14.IconSet>().Single();
            Assert.Equal("3Stars", iconSet.GetAttribute("iconSet", string.Empty).Value);
            Assert.All(iconSet.Elements<X14.ConditionalFormattingIcon>(), icon =>
                Assert.Equal("3Symbols", icon.GetAttribute("iconSet", string.Empty).Value));
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void ConditionalFormattingLifecycle_RejectsStaleSnapshotsAndInvalidTypeConversions() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Rules");
            var firstRule = new ConditionalFormattingRule(new Formula("A1>0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 1
            };
            var secondRule = new ConditionalFormattingRule(new Formula("A1<0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 2
            };
            sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(firstRule, secondRule) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3" }
            });
            ExcelConditionalFormattingInfo current = sheet.GetConditionalFormattingRules().Single(rule => rule.Priority == 1);
            ExcelConditionalFormattingInfo stale = sheet.GetConditionalFormattingRules().Single(rule => rule.Priority == 1);
            current.Range = "B1:B3";
            sheet.UpdateConditionalFormattingRule(current);
            stale.Range = "C1:C3";
            Assert.Throws<InvalidOperationException>(() => sheet.UpdateConditionalFormattingRule(stale));
            Assert.Equal("A1:A3", sheet.GetConditionalFormattingRules().Single(rule => rule.Priority == 2).Range);

            ExcelConditionalFormattingInfo detached = sheet.GetConditionalFormattingRules().Single(rule => rule.Priority == 2);
            sheet.ClearConditionalFormatting();
            Assert.Throws<InvalidOperationException>(() => sheet.RemoveConditionalFormattingRule(detached));

            ExcelConditionalFormattingInfo top = sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Range = "D1:D3",
                Type = "Top10",
                TopBottomRank = 1
            });
            top.Type = "Expression";
            Assert.Throws<ArgumentException>(() => sheet.UpdateConditionalFormattingRule(top));
            top.Type = "futureRule";
            Assert.Throws<ArgumentException>(() => sheet.UpdateConditionalFormattingRule(top));
            Assert.Throws<ArgumentException>(() => sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Range = "E1:E3",
                Type = "CellIs",
                Operator = "Between",
                Formulas = new[] { "1" }
            }));
        }

        [Fact]
        public void ConditionalFormattingLifecycle_VisualEditsPreserveImportedThemeAndUnknownMarkup() {
            const string vendorNamespace = "urn:officeimo:test:conditional-visual";
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Imported");
            var minimum = new X14.ConditionalFormattingValueObject();
            minimum.SetAttribute(new OpenXmlAttribute("type", string.Empty, "autoMin"));
            var maximum = new X14.ConditionalFormattingValueObject();
            maximum.SetAttribute(new OpenXmlAttribute("type", string.Empty, "autoMax"));
            var themeColor = new X14.FillColor { Theme = 4U, Tint = 0.25D };
            var dataBar = new X14.DataBar(minimum, maximum, themeColor) {
                MinLength = 10,
                MaxLength = 90,
                ShowValue = true
            };
            dataBar.SetAttribute(new OpenXmlAttribute("v", "profile", vendorNamespace, "retained"));
            var importedRule = new X14.ConditionalFormattingRule(dataBar) {
                Type = ConditionalFormatValues.DataBar,
                Priority = 1,
                Id = "{8C23A8FB-C99C-4607-BA26-4363664ACED3}"
            };
            sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                new WorksheetExtension(new X14.ConditionalFormattings(
                    new X14.ConditionalFormatting(importedRule, new Xm.ReferenceSequence("A1:A3")))) {
                    Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                }));
            var standardColor = new DocumentFormat.OpenXml.Spreadsheet.Color { Theme = 5U, Tint = -0.15D };
            var standardDataBar = new DataBar(
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                standardColor) { ShowValue = true };
            standardDataBar.SetAttribute(new OpenXmlAttribute("v", "profile", vendorNamespace, "standard-retained"));
            var standardRule = new ConditionalFormattingRule(standardDataBar) {
                Type = ConditionalFormatValues.DataBar,
                Priority = 2
            };
            sheet.WorksheetPart.Worksheet.InsertBefore(new ConditionalFormatting(
                standardRule) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1:B3" }
            }, sheet.WorksheetPart.Worksheet.GetFirstChild<WorksheetExtensionList>());

            ExcelConditionalFormattingInfo imported = sheet.GetConditionalFormattingRules()
                .Single(rule => rule.Source == ExcelConditionalFormattingSource.Office2010Extension);
            Assert.Null(imported.DataBarColor);
            imported.DataBarMaximumLength = 80;
            sheet.UpdateConditionalFormattingRule(imported);
            ExcelConditionalFormattingInfo importedStandard = sheet.GetConditionalFormattingRules()
                .Single(rule => rule.Source == ExcelConditionalFormattingSource.Standard);
            Assert.Null(importedStandard.DataBarColor);
            importedStandard.DataBarShowValue = false;
            sheet.UpdateConditionalFormattingRule(importedStandard);

            Assert.Same(dataBar, importedRule.GetFirstChild<X14.DataBar>());
            Assert.Equal((uint)80, dataBar.MaxLength!.Value);
            Assert.Equal("retained", dataBar.GetAttribute("profile", vendorNamespace).Value);
            Assert.Equal((uint)4, themeColor.Theme!.Value);
            Assert.Equal(0.25D, themeColor.Tint!.Value);
            Assert.Same(standardDataBar, standardRule.GetFirstChild<DataBar>());
            Assert.False(standardDataBar.ShowValue!.Value);
            Assert.Equal("standard-retained", standardDataBar.GetAttribute("profile", vendorNamespace).Value);
            Assert.Equal((uint)5, standardColor.Theme!.Value);
            Assert.Equal(-0.15D, standardColor.Tint!.Value);
        }

        [Fact]
        public void ConditionalFormattingLifecycle_VisualEditsPreserveTintedAndMixedColorKinds() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("MixedColors");

            var standardBarColor = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF112233", Tint = 0.2D };
            var standardBar = new DataBar(
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                standardBarColor) { ShowValue = true };
            var standardBarRule = new ConditionalFormattingRule(standardBar) {
                Type = ConditionalFormatValues.DataBar,
                Priority = 1
            };

            var standardScaleRgb = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FFFF0000", Tint = -0.1D };
            var standardScaleTheme = new DocumentFormat.OpenXml.Spreadsheet.Color { Theme = 4U, Tint = 0.3D };
            var standardScaleEnd = new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF00FF00" };
            var standardScale = new ColorScale(
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min },
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = "50" },
                new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                standardScaleRgb,
                standardScaleTheme,
                standardScaleEnd);
            var standardScaleRule = new ConditionalFormattingRule(standardScale) {
                Type = ConditionalFormatValues.ColorScale,
                Priority = 2
            };
            sheet.WorksheetPart.Worksheet.Append(
                new ConditionalFormatting(standardBarRule) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3" }
                },
                new ConditionalFormatting(standardScaleRule) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1:B3" }
                });

            X14.ConditionalFormattingValueObject NewExtensionThreshold(string type, string value) {
                var threshold = new X14.ConditionalFormattingValueObject(new Xm.Formula(value));
                threshold.SetAttribute(new OpenXmlAttribute("type", string.Empty, type));
                return threshold;
            }
            var extensionBarColor = new X14.FillColor { Rgb = "FF445566", Tint = -0.2D };
            var extensionBar = new X14.DataBar(
                NewExtensionThreshold("autoMin", "0"),
                NewExtensionThreshold("autoMax", "0"),
                extensionBarColor) { MinLength = 10U, MaxLength = 90U, ShowValue = true };
            var extensionBarRule = new X14.ConditionalFormattingRule(extensionBar) {
                Type = ConditionalFormatValues.DataBar,
                Priority = 3,
                Id = "{15AD5EA4-67B6-475D-BF93-3B57814E4BB0}"
            };
            var extensionScaleRgb = new X14.Color { Rgb = "FF0000FF", Tint = 0.15D };
            var extensionScaleTheme = new X14.Color { Theme = 5U, Tint = -0.25D };
            var extensionScaleEnd = new X14.Color { Rgb = "FFFFFF00" };
            var extensionScale = new X14.ColorScale(
                NewExtensionThreshold("autoMin", "0"),
                NewExtensionThreshold("percentile", "50"),
                NewExtensionThreshold("autoMax", "0"),
                extensionScaleRgb,
                extensionScaleTheme,
                extensionScaleEnd);
            var extensionScaleRule = new X14.ConditionalFormattingRule(extensionScale) {
                Type = ConditionalFormatValues.ColorScale,
                Priority = 4,
                Id = "{7838BFCD-70E6-4D78-B426-9D1223287AE1}"
            };
            sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                new WorksheetExtension(new X14.ConditionalFormattings(
                    new X14.ConditionalFormatting(extensionBarRule, new Xm.ReferenceSequence("C1:C3")),
                    new X14.ConditionalFormatting(extensionScaleRule, new Xm.ReferenceSequence("D1:D3")))) {
                        Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                    }));

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules();
            ExcelConditionalFormattingInfo standardBarInfo = rules.Single(rule => rule.Range == "A1:A3");
            ExcelConditionalFormattingInfo standardScaleInfo = rules.Single(rule => rule.Range == "B1:B3");
            ExcelConditionalFormattingInfo extensionBarInfo = rules.Single(rule => rule.Range == "C1:C3");
            ExcelConditionalFormattingInfo extensionScaleInfo = rules.Single(rule => rule.Range == "D1:D3");
            Assert.Equal(new[] { "FFFF0000", string.Empty, "FF00FF00" }, standardScaleInfo.ColorScaleColors);
            Assert.Equal(new[] { "FF0000FF", string.Empty, "FFFFFF00" }, extensionScaleInfo.ColorScaleColors);

            standardBarInfo.DataBarShowValue = false;
            standardScaleInfo.ColorScaleThresholds[1].Value = "60";
            extensionBarInfo.DataBarMaximumLength = 80U;
            extensionScaleInfo.ColorScaleThresholds[1].Value = "60";
            sheet.UpdateConditionalFormattingRule(standardBarInfo);
            sheet.UpdateConditionalFormattingRule(standardScaleInfo);
            sheet.UpdateConditionalFormattingRule(extensionBarInfo);
            sheet.UpdateConditionalFormattingRule(extensionScaleInfo);

            Assert.Equal(0.2D, standardBarColor.Tint!.Value);
            Assert.Equal(-0.1D, standardScaleRgb.Tint!.Value);
            Assert.Equal((uint)4, standardScaleTheme.Theme!.Value);
            Assert.Equal(0.3D, standardScaleTheme.Tint!.Value);
            Assert.Equal(-0.2D, extensionBarColor.Tint!.Value);
            Assert.Equal(0.15D, extensionScaleRgb.Tint!.Value);
            Assert.Equal((uint)5, extensionScaleTheme.Theme!.Value);
            Assert.Equal(-0.25D, extensionScaleTheme.Tint!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void ConditionalFormattingLifecycle_StylesApplyBordersAndCanBeCleared() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Styles");
            ExcelConditionalFormattingInfo rule = sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Range = "A1",
                Type = "Expression",
                Formulas = new[] { "A1>0" },
                DifferentialFillColorArgb = "FFFFC000"
            });
            rule.DifferentialFillColorArgb = null;
            rule.DifferentialBorder = new ExcelCellBorderSnapshot(
                left: new ExcelBorderSideSnapshot("Thin", "FFFF0000"));
            sheet.UpdateConditionalFormattingRule(rule);

            ExcelConditionalFormattingInfo bordered = Assert.Single(sheet.GetConditionalFormattingRules());
            Assert.Null(bordered.DifferentialFillColorArgb);
            Assert.Equal("thin", bordered.DifferentialBorder!.Left!.Style, ignoreCase: true);
            Assert.Equal("FFFF0000", bordered.DifferentialBorder.Left.ColorArgb);
            bordered.DifferentialBorder = null;
            sheet.UpdateConditionalFormattingRule(bordered);
            ExcelConditionalFormattingInfo cleared = Assert.Single(sheet.GetConditionalFormattingRules());
            Assert.Null(cleared.DifferentialFormatId);
            Assert.Null(cleared.DifferentialBorder);

            ExcelConditionalFormattingInfo invalid = sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Range = "B1",
                Type = "Expression",
                Formulas = new[] { "B1>0" }
            });
            invalid.DifferentialBorder = new ExcelCellBorderSnapshot(
                left: new ExcelBorderSideSnapshot("unsupported-border"));
            Assert.Throws<ArgumentException>(() => sheet.UpdateConditionalFormattingRule(invalid));
            Assert.Null(sheet.GetConditionalFormattingRules().Single(item => item.Range == "B1").DifferentialBorder);
        }

        [Fact]
        public void ConditionalFormattingLifecycle_PartialClearTranslatesFormulaAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Anchors");
            var standard = new ConditionalFormattingRule(new Formula("A1>0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 1
            };
            sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(standard) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3 C1:C3" }
            });
            var extension = new X14.ConditionalFormattingRule(new Xm.Formula("A1<0")) {
                Type = ConditionalFormatValues.Expression,
                Priority = 2,
                Id = "{21649510-B00F-47B0-88FB-34E65C996A32}"
            };
            sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                new WorksheetExtension(new X14.ConditionalFormattings(
                    new X14.ConditionalFormatting(extension, new Xm.ReferenceSequence("A1:A3 C1:C3")))) {
                    Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                }));

            sheet.ClearConditionalFormatting("A1:A3");

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules();
            Assert.Equal(2, rules.Count);
            Assert.All(rules, rule => Assert.Equal("C1:C3", rule.Range));
            Assert.Equal("C1>0", rules.Single(rule => rule.Source == ExcelConditionalFormattingSource.Standard).Formulas.Single());
            Assert.Equal("C1<0", rules.Single(rule => rule.Source == ExcelConditionalFormattingSource.Office2010Extension).Formulas.Single());
        }

        [Fact]
        public void ConditionalFormattingLifecycle_FeatureReportCountsRulesInsideSharedOwners() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Imported");
            sheet.WorksheetPart.Worksheet.Append(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("A1>0")) { Type = ConditionalFormatValues.Expression, Priority = 1 },
                new ConditionalFormattingRule(new Formula("A1<0")) { Type = ConditionalFormatValues.Expression, Priority = 2 }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A3" }
            });
            var extensionOwner = new X14.ConditionalFormatting(
                new X14.ConditionalFormattingRule(new Xm.Formula("A1=1")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 3,
                    Id = "{569D66B8-A1BC-4D07-9E68-86C5D7EC5B8A}"
                },
                new X14.ConditionalFormattingRule(new Xm.Formula("A1=2")) {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 4,
                    Id = "{FB80FE64-771C-45F8-842A-C449CBEDC7CE}"
                },
                new Xm.ReferenceSequence("A1:A3"));
            sheet.WorksheetPart.Worksheet.Append(new WorksheetExtensionList(
                new WorksheetExtension(new X14.ConditionalFormattings(extensionOwner)) {
                    Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                }));

            ExcelFeatureFinding feature = Assert.Single(document.InspectFeatures().FindFeatures("Conditional formatting"));
            Assert.Equal(4, feature.Count);
            Assert.Equal(4, sheet.GetConditionalFormattingRules().Count);
        }
    }
}
