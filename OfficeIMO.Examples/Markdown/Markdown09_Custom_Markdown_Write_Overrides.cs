using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown09_Custom_Markdown_Write_Overrides {
        public static void Example_Custom_Markdown_Write_Overrides(string folderPath, bool open = false) {
            Console.WriteLine("[*] Markdown custom markdown write overrides");

            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);

            var doc = CreateDocument();
            var customOptions = CreateWriteOptions();

            string defaultMarkdown = doc.ToMarkdown();
            string customMarkdown = doc.ToMarkdown(customOptions);

            string defaultPath = Path.Combine(mdFolder, "CustomMarkdownWriteOverrides.Default.md");
            string customPath = Path.Combine(mdFolder, "CustomMarkdownWriteOverrides.Custom.md");

            File.WriteAllText(defaultPath, defaultMarkdown, Encoding.UTF8);
            File.WriteAllText(customPath, customMarkdown, Encoding.UTF8);

            Console.WriteLine($"✓ Default markdown saved: {defaultPath}");
            Console.WriteLine($"✓ Custom markdown saved: {customPath}");

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(customPath) { UseShellExecute = true });
            }
        }

        private static MarkdownDoc CreateDocument() {
            return MarkdownDoc.Create()
                .H1("Extension Output Profiles")
                .Callout("note", "Heads up", "This sample customizes markdown writing without changing the underlying document model.")
                .H2("Reader Hooks")
                .P("Inline, fenced, and block parser extensions plug in during parsing.")
                .H2("Writer Hooks")
                .P("Writer overrides can emit portable or host-specific markdown output.")
                .H3("HTML Overrides")
                .P("HTML block overrides already receive MarkdownBodyRenderContext.")
                .H3("Markdown Overrides")
                .P("Markdown block overrides now receive MarkdownWriteContext.")
                .TocForPreviousHeading("Writer hooks contents", min: 3, max: 3, ordered: false, titleLevel: 3);
        }

        private static MarkdownWriteOptions CreateWriteOptions() {
            var options = new MarkdownWriteOptions();
            options.BlockRenderExtensions.Add(MarkdownBlockMarkdownRenderExtension.CreateContextual(
                "section-toc",
                typeof(TocBlock),
                static (block, context) => {
                    if (block is not TocBlock toc) {
                        return null;
                    }

                    var blockIndex = context.GetBlockIndex(toc);
                    var tocOptions = new TocOptions {
                        Scope = TocScope.PreviousHeading,
                        IncludeTitle = true,
                        Title = "Writer hooks contents",
                        TitleLevel = 3,
                        MinLevel = 3,
                        MaxLevel = 3
                    };
                    var titleAnchor = context.GetPrecedingHeadingAnchor(blockIndex, tocOptions);
                    var entries = context.BuildTocEntries(blockIndex, tocOptions, titleAnchor);

                    var sb = new StringBuilder();
                    sb.AppendLine($"<!-- toc-scope:{titleAnchor ?? "(none)"}; block-index:{blockIndex} -->");
                    if (tocOptions.IncludeTitle && !string.IsNullOrWhiteSpace(tocOptions.Title)) {
                        sb.AppendLine($"{new string('#', tocOptions.TitleLevel)} {tocOptions.Title}");
                    }

                    foreach (var entry in entries) {
                        sb.Append("- [")
                            .Append(entry.Text)
                            .Append("](#")
                            .Append(entry.Anchor)
                            .AppendLine(")");
                    }

                    return sb.ToString().TrimEnd();
                }));
            options.BlockRenderExtensions.Add(new MarkdownBlockMarkdownRenderExtension(
                "portable-callout",
                typeof(CalloutBlock),
                static (block, writerOptions) => {
                    if (block is not CalloutBlock callout) {
                        return null;
                    }

                    return RenderPortableCallout(callout, writerOptions);
                }));
            return options;
        }

        private static string RenderPortableCallout(CalloutBlock callout, MarkdownWriteOptions writerOptions) {
            var lines = new List<string>();
            if (!string.IsNullOrWhiteSpace(callout.Title)) {
                lines.Add($"> **{callout.Title}**");
            }

            var bodyLines = (callout.Body ?? string.Empty).Replace("\r\n", "\n").Split('\n');
            foreach (var bodyLine in bodyLines) {
                lines.Add(bodyLine.Length == 0 ? ">" : $"> {bodyLine}");
            }

            lines.Add(">");
            lines.Add($"> _rendered via legacy override; image mode: {writerOptions.ImageRenderingMode}_");
            return string.Join("\n", lines).TrimEnd();
        }
    }
}
