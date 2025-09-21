using System;
using System.IO;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown01_Builder_Basics {
        public static void Example_Builder_Readme(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown builder: README sample");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "BuilderReadme.md");

            MarkdownDoc md = MarkdownDoc
                .Create()
                .FrontMatter(new { title = "DomainDetective", tags = new[] { "dns", "email", "security" } })
                .H1("DomainDetective")
                .P("All-in-one DNS, email, and TLS analyzer with rich reports.")
                .Callout("info", "Early access", "APIs may change before 1.0.")
                .H2("Install")
                .Code("bash", "dotnet tool install -g DomainDetective")
                .H2("Quick start")
                .Code("powershell",
                    "Test-DDMailDomainClassification -DomainName 'evotec.pl','evotec.xyz' -ExportFormat Word")
                .H2("Features")
                .Ul(ul => ul
                    .Item("SPF/DKIM/DMARC scoring")
                    .Item("TLS/SSL tests and cipher hints")
                    .Item("WHOIS, MX, PTR, DNSSEC, BIMI")
                    .Item("Exports: Word, HTML, PDF, Markdown"))
                .H2("Testimonials")
                .Quote(q => q
                    .Line("OfficeIMO.Markdown keeps our release notes consistent.")
                    .Quote(inner => inner
                        .Line("Nested quotes help us show conversation context.")
                        .Line("Great for highlighting support threads."))
                    .P(p => p.Text("— Docs Team")))
                .Hr()
                .H2("Links")
                .Ul(ul => ul
                    .ItemLink("Docs", "https://evotec.xyz/hub/")
                    .ItemLink("Issues", "https://github.com/EvotecIT/DomainDetective/issues"));

            string markdown = md.ToMarkdown();
            File.WriteAllText(path, markdown, Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }

        public static void Example_Scaffold_Readme(string folderPath, bool open) {
            Console.WriteLine("[*] Markdown scaffold: README");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string path = Path.Combine(mdFolder, "ScaffoldReadme.md");

            MarkdownDoc readme = Scaffold.Readme("OfficeIMO.Markdown", opts => opts
                .NuGet("OfficeIMO.Markdown")
                .Description("Fluent Markdown builder for C# with CommonMark+GFM support.")
                .Badges(b => b
                    .NuGet("OfficeIMO.Markdown")
                    .BuildForGitHub("EvotecIT","OfficeIMO", workflow: null) // set workflow for dynamic build badge
                    .CoverageCodecov("EvotecIT","OfficeIMO"))
                .GettingStarted("dotnet add package OfficeIMO.Markdown",
                    "var md = MarkdownDoc.Create().H1(\"Hello\").P(\"World\");\nFile.WriteAllText(\"README.md\", md.ToMarkdown());")
                .LicenseMIT()
                .Links(("Docs","https://evotec.xyz/hub/"), ("Issues","https://github.com/EvotecIT/OfficeIMO/issues")));

            File.WriteAllText(path, readme.ToMarkdown(), Encoding.UTF8);
            Console.WriteLine($"✓ Markdown saved: {path}");
        }
    }
}
