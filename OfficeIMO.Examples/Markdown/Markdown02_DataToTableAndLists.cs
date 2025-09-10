using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown02_DataToTableAndLists {
        public static void Example_TablesAndLists(string folderPath, bool open) {
            Console.WriteLine("[*] Tables/Lists: From data");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "TablesAndLists.md");

            var people = new[] {
                new { Name = "Alice", Role = "Dev", Score = 98 },
                new { Name = "Bob", Role = "Ops", Score = 91 },
                new { Name = "Cara", Role = "QA", Score = 95 }
            };
            var features = new[] { "SPF", "DKIM", "DMARC", "BIMI" };
            var meta = new { Project = "DomainDetective", Version = "0.9.0", Enforced = true };

            var md = MarkdownDoc.Create()
                .H1("Data → Markdown Examples")
                .H2("Object → Table (Property/Value)")
                .TableFrom(meta)
                .H2("Auto header transform")
                .Table(t => t.Columns(HeaderTransforms.Pretty).FromAny(new { DmarcPolicy = "p=none", TlsGrade = "A", SpfAligned = true }))
                .H2("Sequence of Objects → Table (with alignment)")
                .Table(t => t.FromAny(people).Align(ColumnAlignment.Left, ColumnAlignment.Left, ColumnAlignment.Right))
                .H2("Array → Unordered List")
                .Ul(features)
                .H2("Array → Ordered List")
                .Ol(features, start: 1)
                .H2("FromSequence with selectors")
                .Table(t => t.FromSequence(people,
                    ("Name", x => x.Name),
                    ("Role", x => x.Role),
                    ("Score", x => x.Score)).AlignNumericRight());

            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_Toc(string folderPath, bool open) {
            Console.WriteLine("[*] TOC Generation example");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "WithTOC.md");

            var md = MarkdownDoc.Create()
                .H1("Report")
                .H2("Install")
                .P("dotnet add package OfficeIMO.Markdown")
                .H2("Usage")
                .H3("Tables")
                .P("Create tables from objects or sequences.")
                .H3("Lists")
                .P("Create lists from arrays or sequences.")
                .H2("FAQ");

            // Insert TOC at the top, including H2..H3
            md.TocAtTop("Contents", min: 2, max: 3, ordered: false, titleLevel: 2)
              .H2("Appendix")
              .H3("Extra")
              // Insert TOC here for the previous section (Appendix)
              .TocForPreviousHeading("Appendix Contents", min: 3, max: 3, ordered: false, titleLevel: 3);
            
            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_Table_FromAny_WithOptions(string folderPath, bool open) {
            Console.WriteLine("[*] Table FromAny: include/exclude/order");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "TablesAdvanced.md");

            var rows = new[] {
                new { Host = "evotec.pl", SPF = true, DMARC = "p=none", Score = 88 },
                new { Host = "evotec.xyz", SPF = false, DMARC = "p=quarantine", Score = 92 }
            };
            var opts = new TableFromOptions();
            opts.Include.UnionWith(new[] { "Host", "Score", "DMARC" });
            opts.Order.AddRange(new[] { "Host", "DMARC", "Score" });
            opts.HeaderRenames["DMARC"] = "DMARC Policy";

            var md = MarkdownDoc.Create()
                .H1("Advanced Table FromAny")
                .Table(t => t.FromAny(rows, opts).Align(ColumnAlignment.Left, ColumnAlignment.Center, ColumnAlignment.Right));

            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_Table_FromSequence_WithSelectors(string folderPath, bool open) {
            Console.WriteLine("[*] Table FromSequence: selector columns");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "TablesSelectors.md");

            var rows = new[] {
                new { Host = "evotec.pl", SPF = true, DMARC = "p=none", Score = 88.25 },
                new { Host = "evotec.xyz", SPF = false, DMARC = "p=quarantine", Score = 92.0 }
            };

            var md = MarkdownDoc.Create()
                .H1("Table FromSequence with selectors")
                .Table(t => t.FromSequence(rows,
                    ("Host", x => x.Host),
                    ("SPF", x => x.SPF ? "yes" : "no"),
                    ("DMARC", x => x.DMARC),
                    ("Score", x => x.Score.ToString("0.0")))
                    .AlignLeft(0, 2).AlignCenter(1).AlignRight(3));

            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_HeaderTransform_CustomAcronyms(string folderPath, bool open) {
            Console.WriteLine("[*] Header transform with custom acronyms");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "HeaderTransformCustom.md");

            var tx = HeaderTransforms.PrettyWithAcronyms(new[] { "ID", "DMARC", "SPF" });
            var obj = new { DmarcPolicy = "p=none", SpfAligned = true, Id = 25 };

            var md = MarkdownDoc.Create()
                .H1("Header Transform (Custom Acronyms)")
                .Table(t => t.Columns(tx).FromAny(obj));

            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_Table_AutoAligners(string folderPath, bool open) {
            Console.WriteLine("[*] Table auto aligners");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "TablesAutoAlign.md");

            var rows = new[] {
                new { Date = "2025-01-02", Amount = "$12.50", Notes = "ok" },
                new { Date = "2025-02-03", Amount = "$100.00", Notes = "ok" },
            };

            var md = MarkdownDoc.Create()
                .H1("Auto Alignment Helpers")
                .H2("TableFromAuto")
                .TableFromAuto(rows)
                .H2("TableAuto with builder")
                .TableAuto(t => t.FromAny(rows));

            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_TocForSection(string folderPath, bool open) {
            Console.WriteLine("[*] TOC for a named section");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "WithTOCSection.md");

            var md = MarkdownDoc.Create()
                .H1("Report")
                .H2("Intro").P("...")
                .H2("Usage").H3("Tables").P("...").H3("Lists").P("...")
                .H2("Appendix").H3("Extra").P("...")
                .TocForSection("Usage", "Usage Contents", min: 3, max: 3);

            File.WriteAllText(path, md.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }
    }
}
