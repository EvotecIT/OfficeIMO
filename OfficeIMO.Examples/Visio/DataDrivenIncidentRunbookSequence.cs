using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class DataDrivenIncidentRunbookSequence {
        public static void Example_DataDrivenIncidentRunbookSequence(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Data-driven incident runbook sequence");
            string filePath = Path.Combine(folderPath, "Data Driven Incident Runbook Sequence.vsdx");

            VisioDocument document = VisioGallery.CreateIncidentRunbookSequence(filePath);
            document.Save();

            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated data-driven incident runbook sequence failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            document.EnsureVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = true,
                CheckConnectorShapeIntersections = true,
                CheckConnectorLabels = true
            });

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
