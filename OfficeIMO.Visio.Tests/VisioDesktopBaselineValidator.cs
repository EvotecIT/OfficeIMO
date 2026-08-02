using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests {
    internal static class VisioDesktopBaselineValidator {
        private const int VisOpenHidden = 64;
        private const int VisOpenMacrosDisabled = 128;

        internal static bool IsAvailable() {
            return TryGetApplicationType(out _);
        }

        internal static VisioDesktopValidationResult Validate(string vsdxPath) {
            return Validate(vsdxPath, null);
        }

        internal static VisioDesktopValidationResult Validate(string vsdxPath, VisioDesktopValidationOptions? options) {
            if (string.IsNullOrWhiteSpace(vsdxPath)) {
                throw new ArgumentException("VSDX path cannot be null or whitespace.", nameof(vsdxPath));
            }

            string fullPath = Path.GetFullPath(vsdxPath);
            if (!File.Exists(fullPath)) {
                throw new FileNotFoundException("VSDX file was not found.", fullPath);
            }

            if (!TryGetApplicationType(out Type? applicationType)) {
                return new VisioDesktopValidationResult(
                    isAvailable: false,
                    isValid: false,
                    version: null,
                    issues: new[] { "Microsoft Visio desktop automation is not available on this machine." });
            }

            object? application = null;
            object? documents = null;
            object? document = null;
            object? pages = null;
            string? version = null;
            List<string> issues = new();
            List<string> outputFiles = new();
            List<string> deferredPackageOutputs = new();

            try {
                application = Activator.CreateInstance(applicationType!);
                if (application == null) {
                    return new VisioDesktopValidationResult(
                        isAvailable: false,
                        isValid: false,
                        version: null,
                        issues: new[] { "Microsoft Visio desktop automation could not be created." });
                }

                version = Convert.ToString(TryGetProperty(application, "Version"));
                TrySetProperty(application, "Visible", false);
                TrySetProperty(application, "AlertResponse", 7);

                documents = GetProperty(application, "Documents");
                document = InvokeMethod(documents, "OpenEx", fullPath, VisOpenHidden | VisOpenMacrosDisabled);
                pages = GetProperty(document, "Pages");

                int pageCount = Convert.ToInt32(GetProperty(pages, "Count"));
                if (pageCount < 1) {
                    return new VisioDesktopValidationResult(
                        isAvailable: true,
                        isValid: false,
                        version: version,
                        issues: new[] { "Microsoft Visio opened the file, but the document has no pages." });
                }

                if (options != null) {
                    RunOptionalValidationSteps(fullPath, document, pages,
                        pageCount, options, issues, outputFiles,
                        deferredPackageOutputs);
                }

                if (deferredPackageOutputs.Count > 0) {
                    TryInvokeMethod(document, "Close");
                    ReleaseComObject(pages);
                    pages = null;
                    ReleaseComObject(document);
                    document = null;
                    TryInvokeMethod(application, "Quit");
                    ReleaseComObject(documents);
                    documents = null;
                    ReleaseComObject(application);
                    application = null;

                    foreach (string path in deferredPackageOutputs) {
                        AddVerifiedOutputFile(path, "round-tripped VSDX",
                            issues, outputFiles);
                    }
                }

                return new VisioDesktopValidationResult(
                    isAvailable: true,
                    isValid: issues.Count == 0,
                    version: version,
                    issues: issues,
                    outputFiles: outputFiles);
            } catch (Exception exception) {
                Exception root = exception is TargetInvocationException tie && tie.InnerException != null
                    ? tie.InnerException
                    : exception;

                return new VisioDesktopValidationResult(
                    isAvailable: true,
                    isValid: false,
                    version: version,
                    issues: new[] { $"Microsoft Visio could not open the file: {root.Message}" });
            } finally {
                TryInvokeMethod(document, "Close");
                TryInvokeMethod(application, "Quit");
                ReleaseComObject(pages);
                ReleaseComObject(document);
                ReleaseComObject(documents);
                ReleaseComObject(application);
            }
        }

        private static void RunOptionalValidationSteps(
            string inputPath,
            object document,
            object pages,
            int pageCount,
            VisioDesktopValidationOptions options,
            IList<string> issues,
            IList<string> outputFiles,
            IList<string> deferredPackageOutputs) {
            if (options.SaveCopy) {
                string saveCopyPath = GetSaveCopyPath(inputPath, options);
                try {
                    PrepareOutputFile(saveCopyPath);
                    InvokeMethod(document, "SaveAs", saveCopyPath);
                    deferredPackageOutputs.Add(saveCopyPath);
                } catch (Exception exception) {
                    issues.Add($"Microsoft Visio could not save a round-tripped VSDX copy: {GetRootMessage(exception)}");
                }
            }

            if (options.ExportFormats.Count == 0) {
                return;
            }

            object? firstPage = null;
            try {
                firstPage = GetProperty(pages, "Item", 1);
                foreach (VisioDesktopExportFormat format in options.ExportFormats) {
                    string exportPath = GetExportPath(inputPath, options, format);
                    try {
                        PrepareOutputFile(exportPath);
                        Export(document, firstPage, pageCount, format, exportPath);
                        AddVerifiedOutputFile(exportPath, format + " export", issues,
                            outputFiles, format == VisioDesktopExportFormat.Pdf
                                ? pageCount
                                : null);
                    } catch (Exception exception) {
                        issues.Add($"Microsoft Visio could not export {format}: {GetRootMessage(exception)}");
                    }
                }
            } finally {
                ReleaseComObject(firstPage);
            }
        }

        private static void Export(object document, object firstPage, int pageCount, VisioDesktopExportFormat format, string exportPath) {
            switch (format) {
                case VisioDesktopExportFormat.Svg:
                case VisioDesktopExportFormat.Png:
                    InvokeMethod(firstPage, "Export", exportPath);
                    break;
                case VisioDesktopExportFormat.Pdf:
                    InvokeMethod(
                        document,
                        "ExportAsFixedFormat",
                        1,
                        exportPath,
                        0,
                        0,
                        1,
                        pageCount,
                        false,
                        true,
                        true,
                        true,
                        false,
                        Type.Missing);
                    break;
                default:
                    throw new NotSupportedException($"Unsupported Visio desktop export format: {format}.");
            }
        }

        private static string GetSaveCopyPath(string inputPath, VisioDesktopValidationOptions options) {
            string path = !string.IsNullOrWhiteSpace(options.SaveCopyPath)
                ? options.SaveCopyPath!
                : Path.ChangeExtension(inputPath, ".visio-roundtrip.vsdx");

            string fullPath = Path.GetFullPath(path);
            if (string.Equals(fullPath, Path.GetFullPath(inputPath), StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException("SaveCopyPath must not be the same as the input VSDX path.");
            }

            return fullPath;
        }

        private static string GetExportPath(string inputPath, VisioDesktopValidationOptions options, VisioDesktopExportFormat format) {
            string directory = !string.IsNullOrWhiteSpace(options.ExportDirectory)
                ? options.ExportDirectory!
                : Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
            string prefix = !string.IsNullOrWhiteSpace(options.ExportFileNamePrefix)
                ? options.ExportFileNamePrefix!
                : Path.GetFileNameWithoutExtension(inputPath);
            string extension = format.ToString().ToLowerInvariant();
            return Path.GetFullPath(Path.Combine(directory, prefix + "-page1." + extension));
        }

        private static void PrepareOutputFile(string path) {
            string? directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            if (File.Exists(path)) {
                File.Delete(path);
            }
        }

        private static void AddVerifiedOutputFile(string path, string description,
            IList<string> issues, IList<string> outputFiles,
            int? expectedPdfPageCount = null) {
            FileInfo file = new(path);
            if (!file.Exists || file.Length == 0) {
                issues.Add($"Microsoft Visio created an empty or missing {description}: {path}");
                return;
            }

            if (!WaitForReadableFile(file.FullName, out string readinessIssue)) {
                issues.Add($"Microsoft Visio created an inaccessible {description}: {readinessIssue}");
                return;
            }

            if (!ValidateOutputFile(file.FullName, out string validationIssue,
                    expectedPdfPageCount)) {
                issues.Add($"Microsoft Visio created an invalid {description}: {validationIssue}");
                return;
            }

            outputFiles.Add(file.FullName);
        }

        private static bool WaitForReadableFile(string path, out string issue) {
            const int attempts = 40;
            for (int attempt = 0; attempt < attempts; attempt++) {
                try {
                    using (new FileStream(path, FileMode.Open, FileAccess.Read,
                               FileShare.Read)) {
                    }
                    issue = string.Empty;
                    return true;
                } catch (IOException exception) when (attempt + 1 < attempts) {
                    issue = exception.Message;
                    Thread.Sleep(50);
                } catch (UnauthorizedAccessException exception) when (
                    attempt + 1 < attempts) {
                    issue = exception.Message;
                    Thread.Sleep(50);
                } catch (Exception exception) {
                    issue = exception.Message;
                    return false;
                }
            }

            issue = $"The file remained locked after {attempts * 50} ms: {path}";
            return false;
        }

        internal static bool ValidateOutputFile(string path, out string issue,
            int? expectedPdfPageCount = null) {
            try {
                string extension = Path.GetExtension(path).ToLowerInvariant();
                switch (extension) {
                    case ".vsdx":
                        IReadOnlyList<string> packageIssues = VisioValidator.Validate(path);
                        if (packageIssues.Count > 0) {
                            issue = string.Join(" | ", packageIssues.Take(5));
                            return false;
                        }
                        break;
                    case ".png":
                        if (!OfficePngReader.TryDecode(File.ReadAllBytes(path),
                                out OfficeRasterImage? image)
                            || image == null || image.Width <= 0 || image.Height <= 0) {
                            issue = "PNG data is missing, corrupt, or dimensionless: " + path;
                            return false;
                        }
                        break;
                    case ".svg":
                        using (FileStream stream = File.OpenRead(path)) {
                            XDocument svg = XDocument.Load(stream,
                                LoadOptions.PreserveWhitespace);
                            if (!string.Equals(svg.Root?.Name.LocalName, "svg",
                                    StringComparison.OrdinalIgnoreCase)) {
                                issue = "SVG root element was not found: " + path;
                                return false;
                            }
                            if (!HasVisibleSvgContent(svg.Root)) {
                                issue = "SVG contains no visible graphical content with usable bounds: "
                                    + path;
                                return false;
                            }
                        }
                        break;
                    case ".pdf":
                        byte[] pdf = File.ReadAllBytes(path);
                        PdfCore.PdfReadDocument parsed =
                            PdfCore.PdfReadDocument.Open(pdf);
                        if (parsed.Pages.Count < 1) {
                            issue = "PDF page tree is empty: " + path;
                            return false;
                        }
                        if (expectedPdfPageCount.HasValue
                            && parsed.Pages.Count != expectedPdfPageCount.Value) {
                            issue = $"PDF contains {parsed.Pages.Count} pages; expected "
                                + expectedPdfPageCount.Value + ": " + path;
                            return false;
                        }
                        break;
                    default:
                        issue = "Unsupported desktop validation output extension: " + extension;
                        return false;
                }
                issue = string.Empty;
                return true;
            } catch (Exception exception) {
                issue = GetRootMessage(exception);
                return false;
            }
        }

        private static bool HasVisibleSvgContent(XElement root) {
            foreach (XElement element in root.Descendants()) {
                if (!IsVisibleSvgElement(element)) continue;
                switch (element.Name.LocalName.ToLowerInvariant()) {
                    case "path":
                        string? data = (string?)element.Attribute("d");
                        if (!string.IsNullOrWhiteSpace(data)
                            && Regex.IsMatch(data, "[LlHhVvCcSsQqTtAaZz]")) {
                            return true;
                        }
                        break;
                    case "rect":
                    case "image":
                    case "foreignobject":
                        if (HasPositiveSvgLength(element, "width")
                            && HasPositiveSvgLength(element, "height")) {
                            return true;
                        }
                        break;
                    case "circle":
                        if (HasPositiveSvgLength(element, "r")) return true;
                        break;
                    case "ellipse":
                        if (HasPositiveSvgLength(element, "rx")
                            && HasPositiveSvgLength(element, "ry")) {
                            return true;
                        }
                        break;
                    case "line":
                        if (TryReadSvgNumber(element, "x1", out double x1)
                            && TryReadSvgNumber(element, "x2", out double x2)
                            && TryReadSvgNumber(element, "y1", out double y1)
                            && TryReadSvgNumber(element, "y2", out double y2)
                            && (x1 != x2 || y1 != y2)) {
                            return true;
                        }
                        break;
                    case "polygon":
                    case "polyline":
                        string? points = (string?)element.Attribute("points");
                        if (!string.IsNullOrWhiteSpace(points)
                            && Regex.Matches(points,
                                @"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
                                .Count >= 4) {
                            return true;
                        }
                        break;
                    case "text":
                        if (!string.IsNullOrWhiteSpace(element.Value)) return true;
                        break;
                }
            }
            return false;
        }

        private static bool IsVisibleSvgElement(XElement element) {
            foreach (XElement current in element.AncestorsAndSelf()) {
                string? display = (string?)current.Attribute("display");
                string? visibility = (string?)current.Attribute("visibility");
                string style = ((string?)current.Attribute("style")
                    ?? string.Empty).Replace(" ", string.Empty);
                if (string.Equals(display, "none",
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(visibility, "hidden",
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(visibility, "collapse",
                        StringComparison.OrdinalIgnoreCase)
                    || style.IndexOf("display:none",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    || style.IndexOf("visibility:hidden",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    || style.IndexOf("visibility:collapse",
                        StringComparison.OrdinalIgnoreCase) >= 0) {
                    return false;
                }
                if (TryReadSvgNumber((string?)current.Attribute("opacity"),
                        out double opacity) && opacity <= 0D) {
                    return false;
                }
            }
            return true;
        }

        private static bool HasPositiveSvgLength(XElement element,
            string attributeName) =>
            TryReadSvgNumber(element, attributeName, out double value)
            && value > 0D;

        private static bool TryReadSvgNumber(XElement element,
            string attributeName, out double value) =>
            TryReadSvgNumber((string?)element.Attribute(attributeName),
                out value);

        private static bool TryReadSvgNumber(string? text,
            out double value) {
            value = 0D;
            if (string.IsNullOrWhiteSpace(text)) return false;
            Match match = Regex.Match(text,
                @"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)");
            return match.Success && double.TryParse(match.Groups[1].Value,
                NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static bool TryGetApplicationType(out Type? applicationType) {
            applicationType = null;
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return false;
            }

            applicationType = Type.GetTypeFromProgID("Visio.Application");
            return applicationType != null;
        }

        private static object GetProperty(object target, string name) {
            return target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, Array.Empty<object>())!;
        }

        private static object GetProperty(object target, string name, params object[] args) {
            return target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, args)!;
        }

        private static object? TryGetProperty(object target, string name) {
            try {
                return GetProperty(target, name);
            } catch {
                return null;
            }
        }

        private static void TrySetProperty(object target, string name, object value) {
            try {
                target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] { value });
            } catch {
                // Older Visio versions may not expose every automation property used for quiet validation.
            }
        }

        private static object InvokeMethod(object target, string name, params object[] args) {
            return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args)!;
        }

        private static string GetRootMessage(Exception exception) {
            Exception root = exception is TargetInvocationException tie && tie.InnerException != null
                ? tie.InnerException
                : exception;
            return root.Message;
        }

        private static void TryInvokeMethod(object? target, string name) {
            if (target == null) {
                return;
            }

            try {
                InvokeMethod(target, name);
            } catch {
                // Best effort cleanup only.
            }
        }

        private static void ReleaseComObject(object? value) {
            if (value == null) {
                return;
            }

            try {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && Marshal.IsComObject(value)) {
#pragma warning disable CA1416
                    Marshal.FinalReleaseComObject(value);
#pragma warning restore CA1416
                }
            } catch {
                // Best effort cleanup only.
            }
        }
    }

    internal sealed class VisioDesktopValidationOptions {
        internal bool SaveCopy { get; set; }

        internal string? SaveCopyPath { get; set; }

        internal IList<VisioDesktopExportFormat> ExportFormats { get; } = new List<VisioDesktopExportFormat>();

        internal string? ExportDirectory { get; set; }

        internal string? ExportFileNamePrefix { get; set; }

        internal static VisioDesktopValidationOptions RoundTripWithSvg() {
            VisioDesktopValidationOptions options = new() {
                SaveCopy = true
            };
            options.ExportFormats.Add(VisioDesktopExportFormat.Svg);
            return options;
        }
    }

    internal sealed class VisioDesktopValidationResult {
        internal VisioDesktopValidationResult(bool isAvailable, bool isValid, string? version, IEnumerable<string> issues)
            : this(isAvailable, isValid, version, issues, Array.Empty<string>()) {
        }

        internal VisioDesktopValidationResult(bool isAvailable, bool isValid, string? version, IEnumerable<string> issues, IEnumerable<string> outputFiles) {
            IsAvailable = isAvailable;
            IsValid = isValid;
            Version = version;
            Issues = issues.ToList().AsReadOnly();
            OutputFiles = outputFiles.ToList().AsReadOnly();
        }

        internal bool IsAvailable { get; }

        internal bool IsValid { get; }

        internal string? Version { get; }

        internal IReadOnlyList<string> Issues { get; }

        internal IReadOnlyList<string> OutputFiles { get; }
    }

    internal enum VisioDesktopExportFormat {
        Svg,
        Png,
        Pdf
    }
}
