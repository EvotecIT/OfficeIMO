using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlTimeTag {
        [Fact]
        public void TimeRoundsTripWithDateTime() {
            const string html = "<p>On <time datetime=\"2023-01-01\">2023-01-01</time> we met.</p>";
            using var doc = html.ToWordDocument();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlTime", runs[1].CharacterStyleId);
            Assert.Equal("2023-01-01", runs[1].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<time", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("datetime=\"2023-01-01", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2023-01-01</time>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void TimePreservesDateTimeWhenVisibleTextDiffers() {
            const string html = "<p>On <time datetime=\"2023-01-01\">New Year's Day</time> we met.</p>";
            using var doc = html.ToWordDocument();

            string roundTrip = doc.ToHtml();

            Assert.Contains("<time", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("datetime=\"2023-01-01\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">New Year's Day</time>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("id=\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void TimePreservesDateTimeAfterSaveAndReload() {
            const string html = "<p>On <time datetime=\"2023-01-01\">New Year's Day</time> we met.</p>";

            using var doc = html.ToWordDocument();
            using MemoryStream stream = doc.ToStream();
            byte[] packageBytes = stream.ToArray();

            using (var validationStream = new MemoryStream(packageBytes))
            using (var package = WordprocessingDocument.Open(validationStream, false)) {
                var errors = new OpenXmlValidator().Validate(package).ToList();
                Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
            }

            using var reloadStream = new MemoryStream(packageBytes);
            using var reloaded = WordDocument.Load(reloadStream, new WordLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly });

            string roundTrip = reloaded.ToHtml();

            Assert.Contains("<time", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("datetime=\"2023-01-01\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">New Year's Day</time>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}

