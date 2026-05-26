using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Optional Microsoft Visio desktop validation using late-bound COM automation.
    /// </summary>
    public static class VisioDesktopValidator {
        private const int VisOpenHidden = 64;
        private const int VisOpenMacrosDisabled = 128;

        /// <summary>
        /// Gets whether Microsoft Visio desktop automation is registered on this machine.
        /// </summary>
        public static bool IsAvailable() {
            return TryGetApplicationType(out _);
        }

        /// <summary>
        /// Opens a VSDX with Microsoft Visio desktop when available and reports whether Visio accepts it.
        /// This is an optional Windows-only validation path; it does not add a compile-time Visio dependency.
        /// </summary>
        /// <param name="vsdxPath">Path to the VSDX package.</param>
        /// <returns>Desktop validation result.</returns>
        public static VisioDesktopValidationResult Validate(string vsdxPath) {
            return Validate(vsdxPath, null);
        }

        /// <summary>
        /// Opens a VSDX with Microsoft Visio desktop when available, then optionally asks Visio
        /// to save a round-tripped copy and export proof files.
        /// </summary>
        /// <param name="vsdxPath">Path to the VSDX package.</param>
        /// <param name="options">Optional extra validation steps.</param>
        /// <returns>Desktop validation result.</returns>
        public static VisioDesktopValidationResult Validate(string vsdxPath, VisioDesktopValidationOptions? options) {
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
            object? firstPage = null;
            string? version = null;
            List<string> issues = new();
            List<string> outputFiles = new();

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
                    RunOptionalValidationSteps(fullPath, document, pages, pageCount, options, issues, outputFiles);
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
                ReleaseComObject(firstPage);
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
            IList<string> outputFiles) {
            if (options.SaveCopy) {
                string saveCopyPath = GetSaveCopyPath(inputPath, options);
                try {
                    PrepareOutputFile(saveCopyPath);
                    InvokeMethod(document, "SaveAs", saveCopyPath);
                    AddVerifiedOutputFile(saveCopyPath, "round-tripped VSDX", issues, outputFiles);
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
                        AddVerifiedOutputFile(exportPath, format + " export", issues, outputFiles);
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

        private static void AddVerifiedOutputFile(string path, string description, IList<string> issues, IList<string> outputFiles) {
            FileInfo file = new(path);
            if (!file.Exists || file.Length == 0) {
                issues.Add($"Microsoft Visio created an empty or missing {description}: {path}");
                return;
            }

            outputFiles.Add(file.FullName);
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
}
