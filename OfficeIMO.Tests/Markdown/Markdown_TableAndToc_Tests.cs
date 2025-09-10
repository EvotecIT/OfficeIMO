using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_TableAndToc_Tests {
        [Fact]
        public void HeaderTransform_Pretty_Uppercases_Acronyms() {
            System.Func<string,string> transform = HeaderTransforms.Pretty;
            Assert.Equal("DMARC Policy", transform("DmarcPolicy"));
            Assert.Equal("TLS Grade", transform("TlsGrade"));
            Assert.Equal("SPF Aligned", transform("SpfAligned"));
        }

        [Fact]
        public void HeaderTransform_CustomAcronyms() {
            var transform = HeaderTransforms.PrettyWithAcronyms(new [] { "ABC" });
            Assert.Equal("ABC Count", transform("AbcCount"));
        }

        [Fact]
        public void Table_FromAny_WithAlignmentRow() {
            var rows = new [] {
                new { Name = "Alice", Score = 10 },
                new { Name = "Bob", Score = 15 }
            };
            var md = MarkdownDoc.Create()
                .Table(t => t.FromAny(rows).Align(ColumnAlignment.Left, ColumnAlignment.Right));
            var text = md.ToMarkdown();
            // Alignment row should contain :--- and ---:
            Assert.Contains("| :--- | ---: |", text.Replace("\r", ""));
        }

        [Fact]
        public void Table_FromSequence_WithSelectors() {
            var rows = new [] {
                new { Host = "example.com", SPF = true, Score = 91.2 },
                new { Host = "evotec.xyz", SPF = false, Score = 88.0 }
            };
            var md = MarkdownDoc.Create()
                .Table(t => t.FromSequence(rows,
                        ("Host", x => x.Host),
                        ("SPF", x => x.SPF ? "yes" : "no"),
                        ("Score", x => x.Score))
                    .AlignNumericRight());
            var text = md.ToMarkdown();
            Assert.Contains("| Host | SPF | Score |", text);
            Assert.Matches(new Regex(@"\|\s*:\-\-\-\s*\|\s*:\-\-\-\s*\|\s*\-\-\-:\s*\|"), text);
        }

        [Fact]
        public void Table_AlignDatesCenter_And_CurrencyRight() {
            var prev = System.Globalization.CultureInfo.CurrentCulture;
            try {
                System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                var rows = new [] {
                    new { Date = "01/02/2025", Amount = "$12.50" },
                    new { Date = "02/03/2025", Amount = "$100.00" },
                };
                var md = MarkdownDoc.Create()
                    .Table(t => t.FromAny(rows).AlignDatesCenter().AlignNumericRight());
                var text = md.ToMarkdown().Replace("\r", "");
                // Center for Date, Right for Amount
                Assert.Contains("| :---: | ---: |", text);
            } finally {
                System.Globalization.CultureInfo.CurrentCulture = prev;
            }
        }

        [Fact]
        public void Toc_Generates_On_Render() {
            var md = MarkdownDoc.Create()
                .H1("Doc")
                .H2("Install").P("...")
                .H2("Usage").P("...")
                .TocAtTop("Contents", min: 2, max: 2);
            var text = md.ToMarkdown();
            Assert.Contains("- [Install](#install)", text);
            Assert.Contains("- [Usage](#usage)", text);
        }

        [Fact]
        public void TableFromAny_With_Options_Order_Rename_Formatters() {
            var obj = new { DmarcPolicy = "p=none", Score = 91.2345, Notes = "ok" };
            var md = MarkdownDoc.Create()
                .Table(t => t.FromAny(obj, o => {
                    o.Include.UnionWith(new[] { "Score", "DmarcPolicy" });
                    o.Order.AddRange(new[] { "DmarcPolicy", "Score" });
                    o.HeaderRenames["DmarcPolicy"] = "DMARC Policy";
                    o.Formatters["Score"] = v => string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.00}", v);
                }));
            var text = md.ToMarkdown();
            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            Assert.Contains("| DMARC Policy | Score |", text);
            Assert.Contains("| p=none | 91.23 |", text);
        }

        [Fact]
        public void ToHtml_Adds_Heading_Ids() {
            var md = MarkdownDoc.Create().H2("Install");
            var html = md.ToHtml();
            Assert.Contains("<h2 id=\"install\">Install</h2>", html);
        }

        [Fact]
        public void TocHere_And_TocForSection_Work() {
            var md = MarkdownDoc.Create()
                .H1("Doc")
                .H2("Intro").P("...")
                .TocHere(o => { o.MinLevel = 2; o.MaxLevel = 2; })
                .H2("Appendix").H3("Extra")
                .TocForPreviousHeading("Appendix Contents", min: 3, max: 3);
            var text = md.ToMarkdown();
            Assert.Contains("- [Intro](#intro)", text);
            Assert.Contains("- [Extra](#extra)", text);
        }

        [Fact]
        public void TableFromAuto_Applies_Heuristics() {
            var rows = new [] {
                new { Date = "2025-01-02", Amount = "12.50" },
                new { Date = "2025-01-03", Amount = "100.00" },
            };
            var md = MarkdownDoc.Create().TableFromAuto(rows);
            var text = md.ToMarkdown().Replace("\r", "");
            Assert.Contains("| :---: | ---: |", text);
        }
    }
}
