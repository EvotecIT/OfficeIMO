using System;
using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown03_Html_FragmentsAndStyles {
        public static void Example(string folderPath, bool open = false) {
            Console.WriteLine("[*] Markdown → HTML fragments and styles");

            var doc = MarkdownDoc.Create()
                .H1("HTML Rendering")
                .P("This demonstrates fragment vs document and style presets.")
                .H2("List")
                .Ul(u => u.Item("One").Item("Two").Item("Three"))
                .H2("Table")
                .Table(t => t.Headers("Name", "Score").Row("Alice", "98").Row("Bob", "91"));

            string outDir = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(outDir);

            // Fragment (inline CSS by default)
            var fragment = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.GithubAuto });
            var fragmentPath = Path.Combine(outDir, "Sample.Fragment.html");
            File.WriteAllText(fragmentPath, fragment);
            Console.WriteLine($"✓ HTML fragment: {fragmentPath}");

            // Full document with inline GitHub auto style
            var fullDoc = doc.ToHtmlDocument(new HtmlOptions { Title = "Sample", Style = HtmlStyle.GithubAuto, CssDelivery = CssDelivery.Inline });
            var fullDocPath = Path.Combine(outDir, "Sample.Document.Inline.html");
            File.WriteAllText(fullDocPath, fullDoc);
            Console.WriteLine($"✓ HTML document (inline CSS): {fullDocPath}");

            // Full document with external CSS sidecar
            var externalPath = Path.Combine(outDir, "Sample.Document.External.html");
            doc.SaveHtml(externalPath, new HtmlOptions { Title = "Sample External", Style = HtmlStyle.Clean, CssDelivery = CssDelivery.ExternalFile });
            Console.WriteLine($"✓ HTML document (external CSS): {externalPath}");

            // Online mode referencing a CDN CSS, and Offline mode that downloads and inlines it
            var cdnUrl = "https://cdn.jsdelivr.net/npm/github-markdown-css@5.5.1/github-markdown.min.css";
            var online = doc.ToHtmlDocument(new HtmlOptions { Title = "CDN Online", CssDelivery = CssDelivery.LinkHref, CssHref = cdnUrl, AssetMode = AssetMode.Online, BodyClass = "markdown-body" });
            File.WriteAllText(Path.Combine(outDir, "Sample.Document.CdnOnline.html"), online);
            Console.WriteLine("✓ HTML document (CDN link)");

            var offline = doc.ToHtmlDocument(new HtmlOptions { Title = "CDN Offline", CssDelivery = CssDelivery.LinkHref, CssHref = cdnUrl, AssetMode = AssetMode.Offline, BodyClass = "markdown-body" });
            File.WriteAllText(Path.Combine(outDir, "Sample.Document.CdnOfflineInline.html"), offline);
            Console.WriteLine("✓ HTML document (CDN downloaded + inlined)");

            // Global CSS example when hosts provide their own container and want to opt out of scoping
            var globalOptions = new HtmlOptions {
                Title = "Global Scope",
                Style = HtmlStyle.Clean,
                BodyClass = null,
                CssScopeSelector = null
            };
            var globalPath = Path.Combine(outDir, "Sample.Document.GlobalScope.html");
            File.WriteAllText(globalPath, doc.ToHtmlDocument(globalOptions));
            Console.WriteLine($"✓ HTML document (global CSS scope): {globalPath}");

            // Prism highlighting with manifest only (host can dedupe) + plugin example
            var opts = new HtmlOptions {
                Kind = HtmlKind.Fragment,
                EmitMode = AssetEmitMode.ManifestOnly,
                Prism = new PrismOptions { Enabled = true, Theme = PrismTheme.GithubAuto, Languages = { "csharp", "bash" }, Plugins = { "line-numbers", "copy-to-clipboard" } }
            };
            var parts = doc.ToHtmlParts(opts);
            var manifestPath = Path.Combine(outDir, "Sample.Prism.Manifest.json");
            File.WriteAllText(manifestPath, System.Text.Json.JsonSerializer.Serialize(parts.Assets));
            Console.WriteLine("✓ Prism manifest-only output (no tags emitted)");

            // Merge manifests from multiple fragments (simulate host behavior)
            var more = MarkdownDoc.Create().H2("Code").Code("bash", "echo 'hi'");
            var parts2 = more.ToHtmlParts(opts);
            var merged = OfficeIMO.Markdown.HtmlAssetMerger.Build(new[] { parts.Assets, parts2.Assets });
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
