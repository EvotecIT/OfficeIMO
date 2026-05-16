using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Optional slide preview/export boundary backed by installed Microsoft PowerPoint automation.
    /// </summary>
    public static class PowerPointPreviewExporter {
        /// <summary>
        ///     Returns true when this process can locate Microsoft PowerPoint automation.
        /// </summary>
        public static bool IsPowerPointAutomationAvailable() {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
                   Type.GetTypeFromProgID("PowerPoint.Application") != null;
        }

        /// <summary>
        ///     Attempts to export slides from a saved presentation to image files.
        /// </summary>
        public static bool TryExportSlides(
            string presentationPath,
            string outputDirectory,
            out PowerPointPreviewExportResult result,
            PowerPointPreviewExportFormat format = PowerPointPreviewExportFormat.Png,
            int width = 0,
            int height = 0) {
            if (presentationPath == null) {
                throw new ArgumentNullException(nameof(presentationPath));
            }
            if (outputDirectory == null) {
                throw new ArgumentNullException(nameof(outputDirectory));
            }
            if (width < 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (height < 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }
            if (!File.Exists(presentationPath)) {
                throw new FileNotFoundException("Presentation file not found.", presentationPath);
            }

            Type? powerPointType = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? Type.GetTypeFromProgID("PowerPoint.Application")
                : null;
            if (powerPointType == null) {
                result = new PowerPointPreviewExportResult(false, Array.Empty<string>(),
                    "Microsoft PowerPoint automation is not available on this host.", null);
                return false;
            }

            Directory.CreateDirectory(outputDirectory);
            string fullPresentationPath = Path.GetFullPath(presentationPath);
            string fullOutputDirectory = Path.GetFullPath(outputDirectory);
            object? application = null;
            object? presentation = null;

            try {
                application = Activator.CreateInstance(powerPointType);
                if (application == null) {
                    throw new InvalidOperationException("PowerPoint automation could not be started.");
                }

                object presentations = InvokeGet(application, "Presentations");
                presentation = Invoke(presentations, "Open", fullPresentationPath, -1, -1, 0);
                if (presentation == null) {
                    throw new InvalidOperationException("PowerPoint automation did not return an open presentation.");
                }

                Invoke(presentation, "Export", fullOutputDirectory, GetAutomationFormat(format), width, height);

                IReadOnlyList<string> files = Directory
                    .EnumerateFiles(fullOutputDirectory, "*." + GetFileExtension(format), SearchOption.TopDirectoryOnly)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                result = new PowerPointPreviewExportResult(true, files, "Slides exported successfully.", null);
                return true;
            } catch (Exception ex) {
                result = new PowerPointPreviewExportResult(false, Array.Empty<string>(),
                    "Slide export failed through Microsoft PowerPoint automation.", ex);
                return false;
            } finally {
                if (presentation != null) {
                    TryInvoke(presentation, "Close");
                    ReleaseComObject(presentation);
                }

                if (application != null) {
                    TryInvoke(application, "Quit");
                    ReleaseComObject(application);
                }
            }
        }

        private static string GetAutomationFormat(PowerPointPreviewExportFormat format) {
            return format switch {
                PowerPointPreviewExportFormat.Jpeg => "JPG",
                _ => "PNG"
            };
        }

        private static string GetFileExtension(PowerPointPreviewExportFormat format) {
            return format switch {
                PowerPointPreviewExportFormat.Jpeg => "JPG",
                _ => "PNG"
            };
        }

        private static object InvokeGet(object target, string name) {
            return target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null)
                ?? throw new InvalidOperationException("PowerPoint automation returned null for " + name + ".");
        }

        private static object? Invoke(object target, string name, params object[] args) {
            return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args);
        }

        private static void TryInvoke(object target, string name) {
            try {
                Invoke(target, name);
            } catch {
                // Best effort cleanup for optional automation.
            }
        }

        private static void ReleaseComObject(object target) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && Marshal.IsComObject(target)) {
                Marshal.FinalReleaseComObject(target);
            }
        }
    }
}
