using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class VisioShowcase {
        public static void Example_VisioShowcase(string folderPath, bool openVisio, bool exportPreviews) {
            Console.WriteLine("[*] Visio - Showcase examples");
            string showcasePath = Path.Combine(folderPath, "Visio Showcase");
            PrepareShowcaseDirectory(showcasePath);

            List<VisioShowcaseExample> examples = new() {
                new VisioShowcaseExample("01 Basic fluent shapes", () => FluentBasicVisio.Example_FluentBasicVisio(showcasePath, false)),
                new VisioShowcaseExample("02 Flowchart builder", () => FlowchartBuilder.Example_FlowchartBuilder(showcasePath, false)),
                new VisioShowcaseExample("03 Block diagram builder", () => BlockDiagramBuilder.Example_BlockDiagramBuilder(showcasePath, false)),
                new VisioShowcaseExample("04 Swimlane builder", () => SwimlaneDiagramBuilder.Example_SwimlaneDiagramBuilder(showcasePath, false)),
                new VisioShowcaseExample("05 Sequence builder", () => SequenceDiagramBuilder.Example_SequenceDiagramBuilder(showcasePath, false)),
                new VisioShowcaseExample("06 Network topology builder", () => NetworkTopologyDiagramBuilder.Example_NetworkTopologyDiagramBuilder(showcasePath, false)),
                new VisioShowcaseExample("07 Azure architecture builder", () => ArchitectureDiagramBuilder.Example_ArchitectureDiagramBuilder(showcasePath, false)),
                new VisioShowcaseExample("08 Generic graph builder", () => GraphDiagramBuilder.Example_GraphDiagramBuilder(showcasePath, false)),
                new VisioShowcaseExample("09 Editing and data", () => ShapeDataEditing.Example_ShapeDataEditing(showcasePath, false)),
                new VisioShowcaseExample("10 Containers and routing", () => ContainerEditing.Example_ContainerEditing(showcasePath, false)),
                new VisioShowcaseExample("11 Visual quality gallery", () => VisualQualityGallery.Example_VisualQualityGallery(showcasePath, false))
            };
            string? externalStencilPack = Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_STENCIL_PACK");
            if (!string.IsNullOrWhiteSpace(externalStencilPack) && File.Exists(externalStencilPack)) {
                examples.Add(new VisioShowcaseExample("12 External VSSX stencil pack", () => ExternalStencilPack.Example_ExternalStencilPack(showcasePath, false, externalStencilPack)));
            }
            string? integrationStencilPack = MicrosoftIntegrationAzureStencils.ResolveConfiguredPackPath(Array.Empty<string>());
            if (MicrosoftIntegrationAzureStencils.IsConfigured(integrationStencilPack)) {
                examples.Add(new VisioShowcaseExample("13 Microsoft Integration/Azure stencil graph", () => MicrosoftIntegrationAzureStencils.Example_MicrosoftIntegrationAzureStencils(showcasePath, false, integrationStencilPack!)));
            }
            if (global::OfficeIMO.Visio.Stencils.VisioStencilPackageCatalog.DiscoverInstalledVisioPackages().Count > 0) {
                examples.Add(new VisioShowcaseExample("14 Installed Visio stencil packages", () => InstalledVisioStencils.Example_InstalledVisioStencils(showcasePath, false)));
            }

            foreach (VisioShowcaseExample example in examples) {
                Console.WriteLine($"  - {example.Name}");
                example.Create();
            }

            List<string> generatedFiles = Directory
                .EnumerateFiles(showcasePath, "*.vsdx", SearchOption.AllDirectories)
                .Where(file => !file.EndsWith(".visio-roundtrip.vsdx", StringComparison.OrdinalIgnoreCase))
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase)
                .ToList();

            ValidateGeneratedPackages(generatedFiles);

            if (exportPreviews || string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_DESKTOP_SHOWCASE"), "1", StringComparison.OrdinalIgnoreCase)) {
                ExportPreviewFiles(showcasePath, generatedFiles);
            }

            if (openVisio && generatedFiles.Count > 0) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(generatedFiles[0]) { UseShellExecute = true });
            }
        }

        private static void PrepareShowcaseDirectory(string showcasePath) {
            Directory.CreateDirectory(showcasePath);
            foreach (string filePath in Directory.EnumerateFiles(showcasePath, "*.vsdx", SearchOption.AllDirectories)) {
                File.Delete(filePath);
            }

            string previewPath = Path.Combine(showcasePath, "Preview");
            if (Directory.Exists(previewPath)) {
                Directory.Delete(previewPath, recursive: true);
            }
        }

        private static void ValidateGeneratedPackages(IEnumerable<string> generatedFiles) {
            foreach (string filePath in generatedFiles) {
                IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
                if (issues.Count == 0) {
                    Console.WriteLine($"    package ok: {Path.GetFileName(filePath)}");
                    continue;
                }

                string message = string.Join(Environment.NewLine, issues.Select(issue => "      " + issue));
                throw new InvalidOperationException($"Visio example failed package validation: {filePath}{Environment.NewLine}{message}");
            }
        }

        private static void ExportPreviewFiles(string showcasePath, IReadOnlyList<string> generatedFiles) {
            if (!VisioDesktopValidator.IsAvailable()) {
                Console.WriteLine("    Visio desktop preview export skipped: Microsoft Visio automation is not available.");
                return;
            }

            string previewPath = Path.Combine(showcasePath, "Preview");
            Directory.CreateDirectory(previewPath);
            List<string> previewFiles = new();

            foreach (string filePath in generatedFiles) {
                VisioDesktopValidationOptions options = new() {
                    ExportDirectory = previewPath,
                    ExportFileNamePrefix = CreatePreviewPrefix(filePath, showcasePath)
                };
                options.ExportFormats.Add(VisioDesktopExportFormat.Png);
                options.ExportFormats.Add(VisioDesktopExportFormat.Svg);

                VisioDesktopValidationResult result = VisioDesktopValidator.Validate(filePath, options);
                if (!result.IsValid) {
                    string message = string.Join(Environment.NewLine, result.Issues.Select(issue => "      " + issue));
                    throw new InvalidOperationException($"Visio desktop preview export failed: {filePath}{Environment.NewLine}{message}");
                }

                foreach (string outputFile in result.OutputFiles) {
                    previewFiles.Add(outputFile);
                    Console.WriteLine($"    preview: {outputFile}");
                }
            }

            string galleryPath = WritePreviewGallery(previewPath, previewFiles);
            Console.WriteLine($"    gallery: {galleryPath}");
        }

        private static string WritePreviewGallery(string previewPath, IEnumerable<string> previewFiles) {
            string galleryPath = Path.Combine(previewPath, "index.html");
            IEnumerable<string> images = previewFiles
                .Where(file => string.Equals(Path.GetExtension(file), ".png", StringComparison.OrdinalIgnoreCase))
                .OrderBy(file => file, StringComparer.OrdinalIgnoreCase);

            using StreamWriter writer = new(galleryPath, false);
            writer.WriteLine("<!doctype html>");
            writer.WriteLine("<html lang=\"en\">");
            writer.WriteLine("<head>");
            writer.WriteLine("  <meta charset=\"utf-8\">");
            writer.WriteLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
            writer.WriteLine("  <title>OfficeIMO Visio Showcase</title>");
            writer.WriteLine("  <style>");
            writer.WriteLine("    body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:#f6f7f9;color:#20242a}");
            writer.WriteLine("    header{padding:28px 36px;background:#ffffff;border-bottom:1px solid #dfe3e8}");
            writer.WriteLine("    h1{margin:0;font-size:28px;font-weight:650}");
            writer.WriteLine("    main{display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:24px;padding:24px}");
            writer.WriteLine("    figure{margin:0;background:#ffffff;border:1px solid #dfe3e8;border-radius:8px;overflow:hidden}");
            writer.WriteLine("    figcaption{padding:12px 14px;font-weight:600;border-bottom:1px solid #edf0f3}");
            writer.WriteLine("    img{display:block;width:100%;height:auto;background:#ffffff}");
            writer.WriteLine("  </style>");
            writer.WriteLine("</head>");
            writer.WriteLine("<body>");
            writer.WriteLine("  <header><h1>OfficeIMO Visio Showcase</h1></header>");
            writer.WriteLine("  <main>");

            foreach (string image in images) {
                string fileName = Path.GetFileName(image);
                string label = WebUtility.HtmlEncode(Path.GetFileNameWithoutExtension(image).Replace("-page1", string.Empty));
                string src = WebUtility.HtmlEncode(fileName);
                writer.WriteLine("    <figure>");
                writer.WriteLine($"      <figcaption>{label}</figcaption>");
                writer.WriteLine($"      <img src=\"{src}\" alt=\"{label}\">");
                writer.WriteLine("    </figure>");
            }

            writer.WriteLine("  </main>");
            writer.WriteLine("</body>");
            writer.WriteLine("</html>");
            return galleryPath;
        }

        private static string CreatePreviewPrefix(string filePath, string showcasePath) {
            string relative = Path.GetRelativePath(showcasePath, filePath);
            string withoutExtension = Path.Combine(
                Path.GetDirectoryName(relative) ?? string.Empty,
                Path.GetFileNameWithoutExtension(relative));

            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Concat(withoutExtension.Select(ch =>
                invalidChars.Contains(ch) || ch == Path.DirectorySeparatorChar || ch == Path.AltDirectorySeparatorChar
                    ? '-'
                    : ch));
        }

        private sealed class VisioShowcaseExample {
            public VisioShowcaseExample(string name, Action create) {
                Name = name;
                Create = create;
            }

            public string Name { get; }

            public Action Create { get; }
        }
    }
}
