using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MailMerge_KeepComplexFieldCreatesSeparatorBeforeMissingResult() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Name ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(WordMailMergeFieldStatus.Merged, Assert.Single(report.Fields).Status);
            List<Run> runs = paragraph.Elements<Run>().ToList();
            int separatorIndex = runs.FindIndex(run =>
                run.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate);
            int resultIndex = runs.FindIndex(run => run.Elements<Text>().Any(text => text.Text == "Ada"));
            int endIndex = runs.FindIndex(run =>
                run.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End);
            Assert.True(separatorIndex >= 0);
            Assert.True(separatorIndex < resultIndex);
            Assert.True(resultIndex < endIndex);
        }

        [Fact]
        public void Test_MailMerge_KeepComplexFieldReusesEmptySeparator() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Name ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            List<Run> runs = paragraph.Elements<Run>().ToList();
            Assert.Single(runs, run =>
                run.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate);
            Assert.Contains(runs, run => run.Elements<Text>().Any(text => text.Text == "Ada"));
        }

        [Fact]
        public void Test_MailMerge_KeepComplexFieldReplacesTextSharingMarkerRuns() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Name ")),
                new Run(
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    new Text("old prefix")),
                new Run(
                    new Text("old suffix"),
                    new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(WordMailMergeFieldStatus.Merged, Assert.Single(report.Fields).Status);
            Assert.Equal(
                new[] { "Ada", string.Empty },
                paragraph.Descendants<Text>().Select(text => text.Text).ToArray());
        }

        [Fact]
        public void Test_MailMerge_ProcessesMultipleFieldMarkersWithinOneRun() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Name ")),
                new Run(
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    new Text("old result"),
                    new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(WordMailMergeFieldStatus.Merged, Assert.Single(report.Fields).Status);
            Assert.Equal("Ada", Assert.Single(paragraph.Descendants<Text>()).Text);
        }

        [Fact]
        public void Test_MailMerge_InsertsMissingResultBetweenMarkersWithinOneRun() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Name ")),
                new Run(
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(WordMailMergeFieldStatus.Merged, Assert.Single(report.Fields).Status);
            Run markerRun = paragraph.Elements<Run>().Last();
            Assert.Collection(
                markerRun.ChildElements,
                element => Assert.Equal(FieldCharValues.Separate, Assert.IsType<FieldChar>(element).FieldCharType!.Value),
                element => Assert.Equal("Ada", Assert.IsType<Text>(element).Text),
                element => Assert.Equal(FieldCharValues.End, Assert.IsType<FieldChar>(element).FieldCharType!.Value));
        }

        [Fact]
        public void Test_MailMerge_InsertsResultAroundVisibleNonTextChildWithoutChangingSurroundingText() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(new Run(
                new Text("Before "),
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD Name "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Break(),
                new FieldChar { FieldCharType = FieldCharValues.End },
                new Text(" after")));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(new[] { "Before ", "Ada", " after" }, paragraph.Descendants<Text>().Select(text => text.Text).ToArray());
            Assert.Single(paragraph.Descendants<Break>());
        }

        [Fact]
        public void Test_MailMerge_CompletesEntireComplexFieldWithinOneRun() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(new Run(
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD Name "),
                new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Equal(WordMailMergeFieldStatus.Merged, Assert.Single(report.Fields).Status);
            Run fieldRun = Assert.Single(paragraph.Elements<Run>());
            Assert.Collection(
                fieldRun.ChildElements,
                element => Assert.Equal(FieldCharValues.Begin, Assert.IsType<FieldChar>(element).FieldCharType!.Value),
                element => Assert.Equal(" MERGEFIELD Name ", Assert.IsType<FieldCode>(element).Text),
                element => Assert.Equal(FieldCharValues.Separate, Assert.IsType<FieldChar>(element).FieldCharType!.Value),
                element => Assert.Equal("Ada", Assert.IsType<Text>(element).Text),
                element => Assert.Equal(FieldCharValues.End, Assert.IsType<FieldChar>(element).FieldCharType!.Value));
        }

        [Fact]
        public void Test_MailMerge_RemovingSameRunFieldPreservesSurroundingText() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(new Run(
                new Text("Before "),
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD Name "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Text("old"),
                new FieldChar { FieldCharType = FieldCharValues.End },
                new Text(" after")));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: true);

            Assert.True(report.IsComplete);
            Assert.Equal("Before Ada after", paragraph.InnerText);
            Assert.Empty(paragraph.Descendants<FieldChar>());
            Assert.Empty(paragraph.Descendants<FieldCode>());
        }

        [Fact]
        public void Test_MailMerge_RemovingSplitBoundaryFieldPreservesSurroundingText() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(
                new Run(new Text("Before "), new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" MERGEFIELD Name ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("old")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }, new Text(" after")));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> { ["Name"] = "Ada" },
                removeFields: true);

            Assert.True(report.IsComplete);
            Assert.Equal("Before Ada after", paragraph.InnerText);
            Assert.Empty(paragraph.Descendants<FieldChar>());
            Assert.Empty(paragraph.Descendants<FieldCode>());
        }

        [Fact]
        public void Test_MailMerge_ProcessesAdjacentComplexFieldsWithinOneRun() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(new Run(
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD First "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Text("old first"),
                new FieldChar { FieldCharType = FieldCharValues.End },
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD Last "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Text("old last"),
                new FieldChar { FieldCharType = FieldCharValues.End }));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> {
                    ["First"] = "Ada",
                    ["Last"] = "Lovelace"
                },
                removeFields: false);

            Assert.True(report.IsComplete);
            Assert.Collection(
                report.Fields,
                field => {
                    Assert.Equal("First", field.Name);
                    Assert.Equal(WordMailMergeFieldStatus.Merged, field.Status);
                },
                field => {
                    Assert.Equal("Last", field.Name);
                    Assert.Equal(WordMailMergeFieldStatus.Merged, field.Status);
                });
            Assert.Equal(new[] { "Ada", "Lovelace" }, paragraph.Descendants<Text>().Select(text => text.Text).ToArray());
        }

        [Fact]
        public void Test_MailMerge_RemovesAdjacentComplexFieldsWithinOneRun() {
            using WordDocument document = WordDocument.Create();
            Paragraph paragraph = document.AddParagraph()._paragraph;
            paragraph.Append(new Run(
                new Text("Before "),
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD First "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Text("old first"),
                new FieldChar { FieldCharType = FieldCharValues.End },
                new Text(" "),
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode(" MERGEFIELD Last "),
                new FieldChar { FieldCharType = FieldCharValues.Separate },
                new Text("old last"),
                new FieldChar { FieldCharType = FieldCharValues.End },
                new Text(" after")));

            WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
                document,
                new Dictionary<string, string> {
                    ["First"] = "Ada",
                    ["Last"] = "Lovelace"
                },
                removeFields: true);

            Assert.True(report.IsComplete);
            Assert.Equal("Before Ada Lovelace after", paragraph.InnerText);
            Assert.Empty(paragraph.Descendants<FieldChar>());
            Assert.Empty(paragraph.Descendants<FieldCode>());
        }
    }
}
