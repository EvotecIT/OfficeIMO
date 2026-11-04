using System;
using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown03_Html_FragmentsAndStyles {
        public static void Example(string folderPath, bool open = false) {
            Console.WriteLine("[*] Markdown → HTML fragments and styles");

            var doc = MarkdownDocument.Create()
                .H1("HTML Rendering")
                .P("This demonstrates fragment vs document and style presets.")
                .H2("List")
                .Ul(u => u.Item("One").Item("Two").Item("Three"))
                .H2("Table")
                .Table(t => t.Headers("Name", "Score").Row("Alice", "98").Row("Bob", "91"));

            string outDir = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(outDir);

            // Fragment (inline CSS by default)
            var fragment = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.GithubAuto, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });
            var fragmentPath = Path.Combine(outDir, "Sample.Fragment.html");
            File.WriteAllText(fragmentPath, fragment);
            Console.WriteLine($"✓ HTML fragment: {fragmentPath}");

            // Full document with inline GitHub auto style
            var fullDoc = doc.ToHtml(new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Document, Title = "Sample", Style = MarkdownHtmlStyle.GithubAuto, CssDelivery = MarkdownCssDelivery.Inline });
            var fullDocPath = Path.Combine(outDir, "Sample.Document.Inline.html");
            File.WriteAllText(fullDocPath, fullDoc);
            Console.WriteLine($"✓ HTML document (inline CSS): {fullDocPath}");

            // Full document with external CSS sidecar
            var externalPath = Path.Combine(outDir, "Sample.Document.External.html");
            doc.SaveAsHtml(externalPath, new MarkdownHtmlExportOptions { Title = "Sample External", Style = MarkdownHtmlStyle.Clean, CssDelivery = MarkdownCssDelivery.ExternalFile });
            Console.WriteLine($"✓ HTML document (external CSS): {externalPath}");

            // Online mode referencing a CDN CSS, and Offline mode that downloads and inlines it
            var cdnUrl = "https://cdn.jsdelivr.net/npm/github-markdown-css@5.5.1/github-markdown.min.css";
            var online = doc.ToHtml(new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Document, Title = "CDN Online", CssDelivery = MarkdownCssDelivery.LinkHref, CssHref = cdnUrl, AssetMode = MarkdownAssetMode.Online, BodyClass = "markdown-body" });
            File.WriteAllText(Path.Combine(outDir, "Sample.Document.CdnOnline.html"), online);
            Console.WriteLine("✓ HTML document (CDN link)");

            var offline = doc.ToHtml(new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Document, Title = "CDN Offline", CssDelivery = MarkdownCssDelivery.LinkHref, CssHref = cdnUrl, AssetMode = MarkdownAssetMode.Offline, BodyClass = "markdown-body" });
            File.WriteAllText(Path.Combine(outDir, "Sample.Document.CdnOfflineInline.html"), offline);
            Console.WriteLine("✓ HTML document (CDN downloaded + inlined)");

            // Global CSS example when hosts provide their own container and want to opt out of scoping
            var globalOptions = new MarkdownHtmlExportOptions
            {
                Kind = MarkdownHtmlKind.Document,
                Title = "Global Scope",
                Style = MarkdownHtmlStyle.Clean,
                BodyClass = null,
                CssScopeSelector = null
            };
            var globalPath = Path.Combine(outDir, "Sample.Document.GlobalScope.html");
            File.WriteAllText(globalPath, doc.ToHtml(globalOptions));
            Console.WriteLine($"✓ HTML document (global CSS scope): {globalPath}");

            // Prism highlighting with manifest only (host can dedupe) + plugin example
            var opts = new MarkdownHtmlExportOptions
            {
                Kind = MarkdownHtmlKind.Fragment,
                EmitMode = MarkdownAssetEmitMode.ManifestOnly,
                Prism = new MarkdownPrismOptions { Enabled = true, Theme = MarkdownPrismTheme.GithubAuto, Languages = { "csharp", "bash" }, Plugins = { "line-numbers", "copy-to-clipboard" } }
            };
            var parts = doc.ToHtmlParts(opts);
            var manifestPath = Path.Combine(outDir, "Sample.Prism.Manifest.json");
            File.WriteAllText(manifestPath, System.Text.Json.JsonSerializer.Serialize(parts.Assets));
            Console.WriteLine("✓ Prism manifest-only output (no tags emitted)");

            // Merge manifests from multiple fragments (simulate host behavior)
            var more = MarkdownDocument.Create().H2("Code").Code("bash", "echo 'hi'");
            var parts2 = more.ToHtmlParts(opts);
            var merged = OfficeIMO.Markdown.MarkdownHtmlAssetMerger.Build(new[] { parts.Assets, parts2.Assets });
            File.WriteAllText(Path.Combine(outDir, "Sample.Prism.MergedHead.html"), "<!-- head links -->\n" + merged.headLinks);
            File.WriteAllText(Path.Combine(outDir, "Sample.Prism.MergedInline.css"), merged.inlineCss);
            File.WriteAllText(Path.Combine(outDir, "Sample.Prism.MergedInline.js"), merged.inlineJs);
            Console.WriteLine("✓ Prism assets merged without duplicates");

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fullDocPath) { UseShellExecute = true });
            }
        }
    }
}
