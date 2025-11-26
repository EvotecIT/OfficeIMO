#if !NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Examples.PowerPoint;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Generates each shipped PowerPoint example and validates the produced PPTX with OpenXmlValidator.
    /// This mirrors the paths users execute and helps catch packages that would trigger PowerPoint's
    /// repair prompt. The test runs entirely headless inside WSL.
    /// </summary>
    public class PowerPointExamplesValidation {
        private readonly string _repoRoot;
        private readonly string _examplesDir;

        public PowerPointExamplesValidation() {
            _repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
            _examplesDir = Path.Combine(_repoRoot, "OfficeIMO.Examples");
        }

        [Fact]
        public void AllExamples_ValidateWithoutErrors() {
            // Point current directory at the Examples project so image/template relative paths resolve.
            Directory.SetCurrentDirectory(_examplesDir);

            string temp = Path.Combine(Path.GetTempPath(), "officeimo-ppt-examples", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);

            var generators = new List<(string Name, Action<string> Generate)> {
                ("BasicPowerPointDocument", fp => BasicPowerPointDocument.Example_BasicPowerPoint(fp, false)),
                ("AdvancedPowerPoint", fp => AdvancedPowerPoint.Example_AdvancedPowerPoint(fp, false)),
                ("FluentPowerPoint", fp => FluentPowerPoint.Example_FluentPowerPoint(fp, false)),
                ("ShapesPowerPoint", fp => ShapesPowerPoint.Example_PowerPointShapes(fp, false)),
                ("SlidesManagementPowerPoint", fp => SlidesManagementPowerPoint.Example_SlidesManagement(fp, false)),
                ("TablesPowerPoint", fp => TablesPowerPoint.Example_PowerPointTables(fp, false)),
                ("TextFormattingPowerPoint", fp => TextFormattingPowerPoint.Example_TextFormattingPowerPoint(fp, false)),
                ("ThemeAndLayoutPowerPoint", fp => ThemeAndLayoutPowerPoint.Example_PowerPointThemeAndLayout(fp, false)),
                ("UpdatePicturePowerPoint", fp => UpdatePicturePowerPoint.Example_PowerPointUpdatePicture(fp, false)),
                ("ValidateDocument", fp => ValidateDocument.Example(fp, false)),
                ("TestLazyInit", fp => TestLazyInit.Example_TestLazyInit(fp, false)),
                ("InitializeDefaultsPowerPoint", fp => InitializeDefaultsPowerPoint.Example_PowerPointInitializeDefaults(fp, false)),
                ("NotesMasterPowerPoint", fp => NotesMasterPowerPoint.Example_PowerPointNotesMaster(fp, false))
            };

            var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
            var failures = new List<string>();

            foreach (var (name, generate) in generators) {
                string folder = Path.Combine(temp, name);
                Directory.CreateDirectory(folder);
                try {
                    generate(folder);
                } catch (Exception ex) {
                    failures.Add($"{name}: generation threw {ex.GetType().Name}: {ex.Message}");
                    continue;
                }

                foreach (string pptx in Directory.EnumerateFiles(folder, "*.pptx", SearchOption.TopDirectoryOnly)) {
                    using var doc = PresentationDocument.Open(pptx, false);
                    var errors = validator.Validate(doc).ToList();

                    // Emit a manifest to help compare package shape when debugging failures.
                    PowerPointPackageInspector.WriteManifest(pptx, folder);

                    if (errors.Count > 0) {
                        var details = string.Join(Environment.NewLine, errors.Select(e => $"  - {e.Description} @ {e.Path?.XPath} (Part: {e.Part?.Uri})"));
                        failures.Add($"{name}: {pptx} has {errors.Count} validation errors:\n{details}");
                    }
                }
            }

            if (failures.Count > 0) {
                var message = "PowerPoint example validation failed:\n" + string.Join("\n\n", failures);
                throw new Xunit.Sdk.XunitException(message);
            }
        }
    }
}
#endif
