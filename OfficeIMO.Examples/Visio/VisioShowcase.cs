using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class VisioShowcase {
        public static void Example_VisioShowcase(string folderPath, bool openVisio, bool exportNativePreviews = false) {
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
                new VisioShowcaseExample("11 Visual quality gallery", () => VisualQualityGallery.Example_VisualQualityGallery(showcasePath, false)),
                new VisioShowcaseExample("12 Data-driven CI/CD inventory graph", () => DataDrivenInventoryGraph.Example_DataDrivenInventoryGraph(showcasePath, false)),
                new VisioShowcaseExample("13 Data-driven identity authentication graph", () => DataDrivenIdentityGraph.Example_DataDrivenIdentityGraph(showcasePath, false)),
                new VisioShowcaseExample("14 Data-driven Kubernetes service-mesh graph", () => DataDrivenKubernetesGraph.Example_DataDrivenKubernetesGraph(showcasePath, false)),
                new VisioShowcaseExample("15 Data-driven application dependency graph", () => DataDrivenApplicationGraph.Example_DataDrivenApplicationGraph(showcasePath, false)),
                new VisioShowcaseExample("16 Data-driven incident runbook sequence", () => DataDrivenIncidentRunbookSequence.Example_DataDrivenIncidentRunbookSequence(showcasePath, false)),
                new VisioShowcaseExample("17 Data-driven network segmentation diagram", () => DataDrivenNetworkSegmentation.Example_DataDrivenNetworkSegmentation(showcasePath, false)),
                new VisioShowcaseExample("18 Premium scenario showcase", () => PremiumVisioShowcase.Example_PremiumVisioShowcase(showcasePath, false))
            };
            string? externalStencilPack = Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_STENCIL_PACK");
            if (!string.IsNullOrWhiteSpace(externalStencilPack) && File.Exists(externalStencilPack)) {
                examples.Add(new VisioShowcaseExample("19 External VSSX stencil pack", () => ExternalStencilPack.Example_ExternalStencilPack(showcasePath, false, externalStencilPack)));
            }
            string? integrationStencilPack = MicrosoftIntegrationAzureStencils.ResolveConfiguredPackPath(Array.Empty<string>());
            if (MicrosoftIntegrationAzureStencils.IsConfigured(integrationStencilPack)) {
                examples.Add(new VisioShowcaseExample("20 Microsoft Integration/Azure stencil graph", () => MicrosoftIntegrationAzureStencils.Example_MicrosoftIntegrationAzureStencils(showcasePath, false, integrationStencilPack!)));
            }
            string? stencilGalleryPath = ExternalStencilGallery.ResolveConfiguredGalleryPath(Array.Empty<string>());
            if (ExternalStencilGallery.IsConfigured(stencilGalleryPath)) {
                examples.Add(new VisioShowcaseExample("21 External stencil catalog gallery", () => ExternalStencilGallery.Example_ExternalStencilGallery(showcasePath, false, stencilGalleryPath!)));
            }
            if (global::OfficeIMO.Visio.Stencils.VisioStencilPackageCatalog.DiscoverInstalledVisioPackages().Count > 0) {
                examples.Add(new VisioShowcaseExample("22 Installed Visio stencil packages", () => InstalledVisioStencils.Example_InstalledVisioStencils(showcasePath, false)));
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

            if (generatedFiles.Count == 0) {
                throw new InvalidOperationException("Visio showcase did not generate any .vsdx files.");
            }

            ValidateGeneratedPackages(generatedFiles);
            IReadOnlyList<string> proofFiles = ExportStructuralProofFiles(showcasePath, generatedFiles);
            IReadOnlyList<string> previewFiles = Array.Empty<string>();
            if (exportNativePreviews || string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_NATIVE_SHOWCASE"), "1", StringComparison.OrdinalIgnoreCase)) {
                previewFiles = previewFiles.Concat(ExportNativePreviewFiles(showcasePath, generatedFiles)).ToList();
            }

            VisioShowcaseSummary summary = VisioShowcaseSummary.Create(
                showcasePath,
                generatedFiles,
                previewFiles,
                proofFiles: proofFiles);
            summary.EnsureArtifactsValid(requirePreviewsPerDiagram: previewFiles.Count > 0, requireProofsPerDiagram: true);
            summary.SaveArtifacts();
            Console.WriteLine($"    summary: {Path.Combine(showcasePath, VisioShowcaseSummary.MarkdownFileName)}");
            Console.WriteLine($"    summary json: {Path.Combine(showcasePath, VisioShowcaseSummary.JsonFileName)}");
            Console.WriteLine($"    gallery: {Path.Combine(showcasePath, VisioShowcaseSummary.HtmlFileName)}");

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

            string nativePreviewPath = Path.Combine(showcasePath, "Native Preview");
            if (Directory.Exists(nativePreviewPath)) {
                Directory.Delete(nativePreviewPath, recursive: true);
            }

            string structuralProofPath = Path.Combine(showcasePath, "Structural Proof");
            if (Directory.Exists(structuralProofPath)) {
                Directory.Delete(structuralProofPath, recursive: true);
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

        private static IReadOnlyList<string> ExportNativePreviewFiles(string showcasePath, IReadOnlyList<string> generatedFiles) {
            string previewPath = Path.Combine(showcasePath, "Native Preview");
            Directory.CreateDirectory(previewPath);
            List<string> previewFiles = new();

            foreach (string filePath in generatedFiles) {
                VisioDocument document = VisioDocument.Load(filePath);
                string prefix = CreatePreviewPrefix(filePath, showcasePath);
                string svgPath = Path.Combine(previewPath, prefix + "-page1.native.svg");
                string pngPath = Path.Combine(previewPath, prefix + "-page1.native.png");

                document.SaveAsSvg(svgPath, new VisioSvgSaveOptions { PixelsPerInch = 96 });
                document.SaveAsPng(pngPath, new VisioPngSaveOptions { PixelsPerInch = 96 });

                VerifyNonEmptyPreview(svgPath, "native SVG");
                VerifyNonEmptyPreview(pngPath, "native PNG");
                previewFiles.Add(svgPath);
                previewFiles.Add(pngPath);
                Console.WriteLine($"    native preview: {svgPath}");
                Console.WriteLine($"    native preview: {pngPath}");
            }

            string galleryPath = WritePreviewGallery(previewPath, previewFiles);
            Console.WriteLine($"    native gallery: {galleryPath}");
            return previewFiles;
        }

        private static IReadOnlyList<string> ExportStructuralProofFiles(string showcasePath, IReadOnlyList<string> generatedFiles) {
            string proofPath = Path.Combine(showcasePath, "Structural Proof");
            Directory.CreateDirectory(proofPath);
            List<string> proofFiles = new();

            foreach (string filePath in generatedFiles) {
                VisioDocument document = VisioDocument.Load(filePath);
                string prefix = CreatePreviewPrefix(filePath, showcasePath);
                string inspectionPath = Path.Combine(proofPath, prefix + ".inspection.txt");
                string stencilProfilePath = Path.Combine(proofPath, prefix + ".stencil-profile.txt");
                string visualQualityPath = Path.Combine(proofPath, prefix + ".visual-quality.txt");

                VisioInspectionSnapshot inspection = document.CreateInspectionSnapshot();
                File.WriteAllText(inspectionPath, inspection.ToText(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                File.WriteAllText(stencilProfilePath, inspection.CreateStencilProfile().ToText(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                File.WriteAllText(visualQualityPath, document.GetVisualQualityReport().ToText(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                VerifyNonEmptyPreview(inspectionPath, "inspection proof");
                VerifyNonEmptyPreview(stencilProfilePath, "stencil profile proof");
                VerifyNonEmptyPreview(visualQualityPath, "visual quality proof");
                proofFiles.Add(inspectionPath);
                proofFiles.Add(stencilProfilePath);
                proofFiles.Add(visualQualityPath);
                Console.WriteLine($"    structural proof: {inspectionPath}");
                Console.WriteLine($"    structural proof: {stencilProfilePath}");
                Console.WriteLine($"    visual quality proof: {visualQualityPath}");
            }

            return proofFiles;
        }

        private static void VerifyNonEmptyPreview(string path, string description) {
            FileInfo file = new(path);
            if (!file.Exists || file.Length == 0) {
                throw new InvalidOperationException($"OfficeIMO native export created an empty or missing {description}: {path}");
            }
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
