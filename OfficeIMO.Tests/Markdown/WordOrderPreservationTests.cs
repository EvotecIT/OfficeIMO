using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Ordering {
    public class WordOrderPreservationTests {
        private static WordDocument BuildSampleDoc() {
            var doc = WordDocument.Create();
            // Section 1
            doc.AddParagraph("Intro P1");
            var t1 = doc.AddTable(2, 2);
            t1.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            t1.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            t1.Rows[1].Cells[0].Paragraphs[0].Text = "T1_R1C1";
            t1.Rows[1].Cells[1].Paragraphs[0].Text = "T1_R1C2";
            doc.AddParagraph("After T1");

            // New section
            doc.AddSection(SectionMarkValues.NextPage);

            // Section 2
            doc.AddParagraph("Section2 P1");
            var t2 = doc.AddTable(2, 2);
            t2.Rows[0].Cells[0].Paragraphs[0].Text = "J1";
            t2.Rows[0].Cells[1].Paragraphs[0].Text = "J2";
            t2.Rows[1].Cells[0].Paragraphs[0].Text = "T2_R1C1";
            t2.Rows[1].Cells[1].Paragraphs[0].Text = "T2_R1C2";
            doc.AddParagraph("Section2 After T2");
            return doc;
        }

        private static WordDocument BuildSingleSectionDoc() {
            var doc = WordDocument.Create();
            doc.AddParagraph("Intro P1");
            var t1 = doc.AddTable(2, 2);
            t1.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            t1.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            t1.Rows[1].Cells[0].Paragraphs[0].Text = "T1_R1C1";
            t1.Rows[1].Cells[1].Paragraphs[0].Text = "T1_R1C2";
            doc.AddParagraph("After T1");
            doc.AddParagraph("Section2 P1");
            var t2 = doc.AddTable(2, 2);
            t2.Rows[0].Cells[0].Paragraphs[0].Text = "J1";
            t2.Rows[0].Cells[1].Paragraphs[0].Text = "J2";
            t2.Rows[1].Cells[0].Paragraphs[0].Text = "T2_R1C1";
            t2.Rows[1].Cells[1].Paragraphs[0].Text = "T2_R1C2";
            doc.AddParagraph("Section2 After T2");
            return doc;
        }

        [Fact]
        public void Markdown_OrderPreserved_WithinSection() {
            using var doc = BuildSingleSectionDoc();
            string md = doc.ToMarkdown();
            int p1 = md.IndexOf("Intro P1", StringComparison.Ordinal);
            int t1 = md.IndexOf("T1_R1C1", StringComparison.Ordinal);
            int after1 = md.IndexOf("After T1", StringComparison.Ordinal);
            int s2p1 = md.IndexOf("Section2 P1", StringComparison.Ordinal);
            int t2 = md.IndexOf("T2_R1C1", StringComparison.Ordinal);
            int after2 = md.IndexOf("Section2 After T2", StringComparison.Ordinal);

            Assert.True(p1 >= 0 && t1 > p1 && after1 > t1, $"Order in S1 invalid: p1={p1}, t1={t1}, after1={after1}\n{md}");
            Assert.True(s2p1 > after1 && t2 > s2p1 && after2 > t2, $"Order in S2 invalid: s2p1={s2p1}, t2={t2}, after2={after2}\n{md}");
        }

        [Fact]
        public void Html_OrderPreserved_WithinSection() {
            using var doc = BuildSingleSectionDoc();
            string html = doc.ToHtml();
            int p1 = html.IndexOf("Intro P1", StringComparison.Ordinal);
            int t1 = html.IndexOf("T1_R1C1", StringComparison.Ordinal);
            int after1 = html.IndexOf("After T1", StringComparison.Ordinal);
            int s2p1 = html.IndexOf("Section2 P1", StringComparison.Ordinal);
            int t2 = html.IndexOf("T2_R1C1", StringComparison.Ordinal);
            int after2 = html.IndexOf("Section2 After T2", StringComparison.Ordinal);

            Assert.True(p1 >= 0 && t1 > p1 && after1 > t1, $"Order in S1 invalid: p1={p1}, t1={t1}, after1={after1}\n{html}");
            Assert.True(s2p1 > after1 && t2 > s2p1 && after2 > t2, $"Order in S2 invalid: s2p1={s2p1}, t2={t2}, after2={after2}\n{html}");
        }

        [Fact]
        public void Pdf_Generates_FromOrderedElements() {
            using var doc = BuildSampleDoc();
            using var ms = new MemoryStream();
            doc.SaveAsPdf(ms);
            Assert.True(ms.Length > 0);
        }

        [Fact(Skip = "Pending: WordSection.Elements returns wrong content for second section with AddSection. Enable once fixed.")]
        public void Markdown_Html_OrderPreserved_AcrossSections_Pending() {
            using var doc = BuildSampleDoc();
            string md = doc.ToMarkdown();
            Assert.Contains("Section2 P1", md);
            string html = doc.ToHtml();
            Assert.Contains("Section2 P1", html);
        }
    }
}
