using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Status of one presentation compatibility lane.</summary>
    public enum PowerPointCompatibilityStatus {
        /// <summary>The lane was intentionally not run.</summary>
        NotRun,
        /// <summary>The lane is unavailable in the current environment.</summary>
        Unavailable,
        /// <summary>The lane accepted or converted the presentation.</summary>
        Passed,
        /// <summary>The lane ran and rejected or failed to convert the presentation.</summary>
        Failed
    }

    /// <summary>One named compatibility result.</summary>
    public sealed class PowerPointCompatibilityLaneResult {
        internal PowerPointCompatibilityLaneResult(string lane, PowerPointCompatibilityStatus status,
            string message, IEnumerable<string>? artifactPaths = null) {
            Lane = lane;
            Status = status;
            Message = message ?? string.Empty;
            ArtifactPaths = new ReadOnlyCollection<string>((artifactPaths ?? Array.Empty<string>()).ToList());
        }

        /// <summary>Stable lane name.</summary>
        public string Lane { get; }
        /// <summary>Lane outcome.</summary>
        public PowerPointCompatibilityStatus Status { get; }
        /// <summary>Human-readable evidence.</summary>
        public string Message { get; }
        /// <summary>Produced evidence paths, if any.</summary>
        public IReadOnlyList<string> ArtifactPaths { get; }
    }

    /// <summary>Options for local and external presentation compatibility validation.</summary>
    public sealed class PowerPointCompatibilityOptions {
        /// <summary>Explicitly enables PowerPoint Desktop reference rendering.</summary>
        public bool EnablePowerPointDesktop { get; set; }
        /// <summary>Explicitly enables a headless LibreOffice PDF import/conversion check.</summary>
        public bool EnableLibreOffice { get; set; }
        /// <summary>Directory where compatibility artifacts are written.</summary>
        public string? OutputDirectory { get; set; }
        /// <summary>Maximum LibreOffice conversion time.</summary>
        public TimeSpan LibreOfficeTimeout { get; set; } = TimeSpan.FromSeconds(60);
    }

    /// <summary>Structured compatibility evidence across local and externally recorded lanes.</summary>
    public sealed class PowerPointCompatibilityReport {
        private readonly List<PowerPointCompatibilityLaneResult> _lanes;

        internal PowerPointCompatibilityReport(IEnumerable<PowerPointCompatibilityLaneResult> lanes) {
            _lanes = lanes.ToList();
        }

        /// <summary>Report schema version.</summary>
        public int SchemaVersion => 1;
        /// <summary>Compatibility lane results.</summary>
        public IReadOnlyList<PowerPointCompatibilityLaneResult> Lanes => _lanes.AsReadOnly();
        /// <summary>Whether every lane that ran passed.</summary>
        public bool IsSuccessful => _lanes.All(lane => lane.Status != PowerPointCompatibilityStatus.Failed);

        /// <summary>Records evidence from an external Google Slides, Keynote, or other import lane.</summary>
        public PowerPointCompatibilityReport RecordExternal(string lane, PowerPointCompatibilityStatus status,
            string message, params string[] artifactPaths) {
            if (string.IsNullOrWhiteSpace(lane)) throw new ArgumentException("Lane cannot be empty.", nameof(lane));
            _lanes.RemoveAll(item => string.Equals(item.Lane, lane, StringComparison.OrdinalIgnoreCase));
            _lanes.Add(new PowerPointCompatibilityLaneResult(lane.Trim(), status, message, artifactPaths));
            return this;
        }

        /// <summary>Serializes the compatibility report as deterministic JSON.</summary>
        public string ToJson(bool indented = true) {
            string nl = indented ? Environment.NewLine : string.Empty;
            string i1 = indented ? "  " : string.Empty;
            string i2 = indented ? "    " : string.Empty;
            var json = new StringBuilder();
            json.Append('{').Append(nl)
                .Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(nl)
                .Append(i1).Append("\"isSuccessful\":").Append(IsSuccessful ? "true" : "false").Append(',').Append(nl)
                .Append(i1).Append("\"lanes\": [").Append(nl);
            for (int index = 0; index < _lanes.Count; index++) {
                PowerPointCompatibilityLaneResult lane = _lanes[index];
                json.Append(i2).Append('{')
                    .Append("\"lane\":\"").Append(Escape(lane.Lane)).Append("\",")
                    .Append("\"status\":\"").Append(lane.Status).Append("\",")
                    .Append("\"message\":\"").Append(Escape(lane.Message)).Append("\",")
                    .Append("\"artifactCount\":").Append(lane.ArtifactPaths.Count).Append('}');
                if (index < _lanes.Count - 1) json.Append(',');
                json.Append(nl);
            }
            json.Append(i1).Append(']').Append(nl).Append('}');
            return json.ToString();
        }

        private static string Escape(string value) => (value ?? string.Empty).Replace("\\", "\\\\")
            .Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    /// <summary>Local compatibility validation with explicit placeholders for external import evidence.</summary>
    public static class PowerPointCompatibilityInspector {
        /// <summary>Inspects a saved presentation through selected compatibility lanes.</summary>
        public static PowerPointCompatibilityReport Inspect(string presentationPath,
            PowerPointCompatibilityOptions? options = null) {
            if (string.IsNullOrWhiteSpace(presentationPath)) {
                throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));
            }
            string fullPath = Path.GetFullPath(presentationPath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Presentation was not found.", fullPath);
            options ??= new PowerPointCompatibilityOptions();
            string output = string.IsNullOrWhiteSpace(options.OutputDirectory)
                ? Path.Combine(Path.GetTempPath(), "OfficeIMO.PowerPointCompatibility",
                    Path.GetFileNameWithoutExtension(fullPath) + "-" + Guid.NewGuid().ToString("N"))
                : Path.GetFullPath(options.OutputDirectory!);
            Directory.CreateDirectory(output);

            var lanes = new List<PowerPointCompatibilityLaneResult>();
            using (PowerPointPresentation presentation = PowerPointPresentation.Load(fullPath,
                       new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                var validation = presentation.ValidateDocument();
                lanes.Add(new PowerPointCompatibilityLaneResult("OpenXml", validation.Count == 0
                    ? PowerPointCompatibilityStatus.Passed : PowerPointCompatibilityStatus.Failed,
                    validation.Count == 0 ? "The package passed Open XML validation."
                        : validation.Count + " Open XML validation error(s) were found."));
            }

            string desktopOutput = Path.Combine(output, "powerpoint-desktop");
            PowerPointReferenceRenderResult desktop = TryRenderDesktopWhenSupported(
                fullPath, desktopOutput, options.EnablePowerPointDesktop);
            lanes.Add(new PowerPointCompatibilityLaneResult("PowerPointDesktop", MapStatus(desktop.Status),
                desktop.Message, desktop.ImagePaths));
            lanes.Add(InspectLibreOffice(fullPath, Path.Combine(output, "libreoffice"), options));
            lanes.Add(new PowerPointCompatibilityLaneResult("Keynote", PowerPointCompatibilityStatus.NotRun,
                "Record a macOS Keynote import/export result with RecordExternal when that environment is available."));
            lanes.Add(new PowerPointCompatibilityLaneResult("GoogleSlides", PowerPointCompatibilityStatus.NotRun,
                "Record a Google Slides import result with RecordExternal when an authenticated external lane is available."));
            return new PowerPointCompatibilityReport(lanes);
        }

        [UnconditionalSuppressMessage("Trimming", "IL2026",
            Justification = "The late-bound PowerPoint Desktop lane is called only when dynamic code is supported. NativeAOT reports the lane as unavailable and uses the in-process renderers.")]
        [UnconditionalSuppressMessage("AOT", "IL3050",
            Justification = "The late-bound PowerPoint Desktop lane is called only when dynamic code is supported. NativeAOT reports the lane as unavailable and uses the in-process renderers.")]
        private static PowerPointReferenceRenderResult TryRenderDesktopWhenSupported(string presentationPath,
            string outputDirectory, bool enabled) {
#if NET5_0_OR_GREATER
            if (!RuntimeFeature.IsDynamicCodeSupported) {
                return new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Unavailable,
                    "PowerPoint Desktop reference rendering is unavailable in NativeAOT applications; use the in-process renderers.");
            }
#endif

            return PowerPointDesktopReferenceRenderer.TryRender(presentationPath, outputDirectory, enabled);
        }

        private static PowerPointCompatibilityLaneResult InspectLibreOffice(string presentationPath,
            string outputDirectory, PowerPointCompatibilityOptions options) {
            if (!options.EnableLibreOffice) {
                return new PowerPointCompatibilityLaneResult("LibreOffice", PowerPointCompatibilityStatus.NotRun,
                    "LibreOffice compatibility validation was not enabled.");
            }
            string? executable = FindLibreOffice();
            if (executable == null) {
                return new PowerPointCompatibilityLaneResult("LibreOffice", PowerPointCompatibilityStatus.Unavailable,
                    "LibreOffice was not found on this machine.");
            }
            Directory.CreateDirectory(outputDirectory);
            var startInfo = new ProcessStartInfo {
                FileName = executable,
                Arguments = "--headless --convert-to pdf --outdir \"" + outputDirectory + "\" \"" + presentationPath + "\"",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            try {
                using Process? process = Process.Start(startInfo);
                if (process == null) throw new InvalidOperationException("LibreOffice process could not be started.");
                Task<string> standardOutput = process.StandardOutput.ReadToEndAsync();
                Task<string> standardError = process.StandardError.ReadToEndAsync();
                int timeoutMilliseconds = (int)Math.Min(int.MaxValue,
                    Math.Max(1000D, options.LibreOfficeTimeout.TotalMilliseconds));
                if (!process.WaitForExit(timeoutMilliseconds)) {
                    try {
                        process.Kill();
                        if (process.WaitForExit(5000)) {
                            _ = ReadProcessOutput(standardOutput, standardError);
                        }
                    } catch { }
                    return new PowerPointCompatibilityLaneResult("LibreOffice", PowerPointCompatibilityStatus.Failed,
                        "LibreOffice conversion exceeded " + options.LibreOfficeTimeout.TotalSeconds
                            .ToString("0.#", CultureInfo.InvariantCulture) + " seconds.");
                }
                process.WaitForExit();
                string expectedPdf = Path.Combine(outputDirectory,
                    Path.GetFileNameWithoutExtension(presentationPath) + ".pdf");
                bool passed = process.ExitCode == 0 && File.Exists(expectedPdf) && new FileInfo(expectedPdf).Length > 0;
                string details = ReadProcessOutput(standardOutput, standardError);
                return new PowerPointCompatibilityLaneResult("LibreOffice", passed
                    ? PowerPointCompatibilityStatus.Passed : PowerPointCompatibilityStatus.Failed,
                    passed ? "LibreOffice imported the deck and produced a non-empty PDF."
                        : "LibreOffice did not produce the expected PDF. " + details,
                    passed ? new[] { expectedPdf } : Array.Empty<string>());
            } catch (Exception ex) {
                return new PowerPointCompatibilityLaneResult("LibreOffice", PowerPointCompatibilityStatus.Failed,
                    "LibreOffice compatibility validation failed: " + ex.Message);
            }
        }

        private static string ReadProcessOutput(Task<string> standardOutput, Task<string> standardError) {
            try {
                return (standardOutput.GetAwaiter().GetResult() + " " +
                        standardError.GetAwaiter().GetResult()).Trim();
            } catch {
                return string.Empty;
            }
        }

        private static string? FindLibreOffice() {
            var candidates = new List<string>();
            string? programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!string.IsNullOrWhiteSpace(programFiles)) {
                candidates.Add(Path.Combine(programFiles, "LibreOffice", "program", "soffice.exe"));
            }
            string? path = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrWhiteSpace(path)) {
                foreach (string directory in path!.Split(Path.PathSeparator)) {
                    if (string.IsNullOrWhiteSpace(directory)) continue;
                    candidates.Add(Path.Combine(directory.Trim(), "soffice.exe"));
                    candidates.Add(Path.Combine(directory.Trim(), "soffice"));
                }
            }
            return candidates.FirstOrDefault(File.Exists);
        }

        private static PowerPointCompatibilityStatus MapStatus(PowerPointReferenceRenderStatus status) {
            switch (status) {
                case PowerPointReferenceRenderStatus.Succeeded: return PowerPointCompatibilityStatus.Passed;
                case PowerPointReferenceRenderStatus.Failed: return PowerPointCompatibilityStatus.Failed;
                case PowerPointReferenceRenderStatus.Unavailable: return PowerPointCompatibilityStatus.Unavailable;
                default: return PowerPointCompatibilityStatus.NotRun;
            }
        }
    }
}
