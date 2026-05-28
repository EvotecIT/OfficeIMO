using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ExternalStencilGallery {
        public static void Example_ExternalStencilGallery(string folderPath, bool openVisio, string stencilPackPathOrDirectory) {
            if (string.IsNullOrWhiteSpace(stencilPackPathOrDirectory)) throw new ArgumentException("Stencil pack path cannot be null or whitespace.", nameof(stencilPackPathOrDirectory));

            Console.WriteLine("[*] Visio - External stencil catalog gallery");
            string filePath = Path.Combine(folderPath, "External Stencil Catalog Gallery.vsdx");
            string[] packagePaths = ResolvePackagePaths(stencilPackPathOrDirectory).ToArray();
            if (packagePaths.Length == 0) {
                throw new FileNotFoundException("No supported Visio stencil packages were found.", stencilPackPathOrDirectory);
            }

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.LoadMany(packagePaths, new VisioStencilPackageLoadOptions {
                CatalogName = "External Stencil Catalog",
                IncludeUnsupportedMasters = true
            });
            IReadOnlyList<VisioStencilShape> selected = SelectGalleryShapes(catalog);
            VisioStencilCatalog selectedCatalog = new(catalog.Name, selected);

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Stencil Gallery", 11, 8.5);
            page.Grid(visible: false, snap: true);
            IReadOnlyList<VisioShape> placed = page.AddStencilGallery(selectedCatalog, new VisioStencilGalleryOptions {
                Title = "External stencil catalog gallery",
                Columns = 4,
                MaxShapes = 16,
                IdPrefix = "external-stencils",
                CellWidth = 2.35D,
                CellHeight = 1.38D,
                IconMaxWidth = 0.86D,
                IconMaxHeight = 0.66D
            });

            VisioShape note = page.AddTextBox(
                "source-packages",
                page.Width / 2D,
                0.36D,
                page.Width - 1.2D,
                0.28D,
                "Source packages: " + string.Join(", ", packagePaths.Select(Path.GetFileName)),
                VisioMeasurementUnit.Inches);
            note.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 8.2D,
                Color = Color.FromRgb(91, 101, 112),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };

            document.Save();
            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated external stencil gallery failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            Console.WriteLine("    Source packages: " + string.Join(", ", packagePaths.Select(Path.GetFileName)));
            Console.WriteLine("    Rendered shapes: " + placed.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            Console.WriteLine("    Output: " + filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        public static string? ResolveConfiguredGalleryPath(string[] args) {
            string? argument = GetArgumentValue(args, "--visio-stencil-gallery")
                ?? GetArgumentValue(args, "--visio-stencil-gallery-pack");
            if (!string.IsNullOrWhiteSpace(argument)) {
                return argument;
            }

            return Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_STENCIL_GALLERY");
        }

        public static bool IsConfigured(string? pathOrDirectory) {
            if (string.IsNullOrWhiteSpace(pathOrDirectory)) {
                return false;
            }

            if (File.Exists(pathOrDirectory)) {
                string fullPath = Path.GetFullPath(pathOrDirectory);
                string? directory = Path.GetDirectoryName(fullPath);
                return !string.IsNullOrWhiteSpace(directory) &&
                    VisioStencilPackageCatalog.EnumeratePackageFiles(directory, recursive: false)
                        .Contains(fullPath, StringComparer.OrdinalIgnoreCase);
            }

            return Directory.Exists(pathOrDirectory) &&
                VisioStencilPackageCatalog.EnumeratePackageFiles(pathOrDirectory, recursive: true).Count > 0;
        }

        private static IEnumerable<string> ResolvePackagePaths(string pathOrDirectory) {
            if (File.Exists(pathOrDirectory)) {
                yield return Path.GetFullPath(pathOrDirectory);
                yield break;
            }

            if (!Directory.Exists(pathOrDirectory)) {
                throw new DirectoryNotFoundException("Visio stencil pack directory was not found: " + pathOrDirectory);
            }

            foreach (string package in VisioStencilPackageCatalog.EnumeratePackageFiles(pathOrDirectory, recursive: true).Take(12)) {
                yield return package;
            }
        }

        private static IReadOnlyList<VisioStencilShape> SelectGalleryShapes(VisioStencilCatalog catalog) {
            string[] preferred = {
                "API Connections",
                "3rd Party",
                "3rd Party Integration",
                "Adapter/Connector",
                "Application",
                "AWS",
                "AWS: Logo",
                "Apache HBASE",
                ".NET",
                "AccuWeather",
                "Android",
                "API Management",
                "APIM Gateway",
                "API Apps #1",
                "API Apps"
            };

            List<VisioStencilShape> selected = new();
            foreach (string query in preferred) {
                VisioStencilShape? match = catalog.TryGet(query, out VisioStencilShape? exact)
                    ? exact
                    : catalog.Search(query).FirstOrDefault();
                if (match != null && !selected.Any(shape => string.Equals(shape.MasterNameU, match.MasterNameU, StringComparison.OrdinalIgnoreCase))) {
                    selected.Add(match);
                }
            }

            if (selected.Count >= 8) {
                return selected.Take(16).ToList();
            }

            foreach (VisioStencilShape fallback in catalog.Shapes
                         .GroupBy(shape => shape.Category, StringComparer.OrdinalIgnoreCase)
                         .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                         .SelectMany(group => group.OrderBy(shape => shape.Name, StringComparer.OrdinalIgnoreCase).Take(4))) {
                if (selected.Count >= 16) {
                    break;
                }

                if (!selected.Any(shape => string.Equals(shape.MasterNameU, fallback.MasterNameU, StringComparison.OrdinalIgnoreCase))) {
                    selected.Add(fallback);
                }
            }

            return selected;
        }

        private static string? GetArgumentValue(string[] args, string name) {
            for (int i = 0; i < args.Length; i++) {
                if (!string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal)) {
                    return args[i + 1];
                }

                return null;
            }

            string prefix = name + "=";
            return args.FirstOrDefault(arg => arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))?.Substring(prefix.Length);
        }
    }
}
