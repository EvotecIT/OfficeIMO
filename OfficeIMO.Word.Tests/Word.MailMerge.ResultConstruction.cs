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
    }
}
