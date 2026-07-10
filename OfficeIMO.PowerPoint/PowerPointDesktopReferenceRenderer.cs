using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeIMO.PowerPoint {
    /// <summary>Outcome of an opt-in PowerPoint Desktop reference render.</summary>
    public enum PowerPointReferenceRenderStatus {
        /// <summary>The caller did not opt in.</summary>
        Disabled,
        /// <summary>The requested renderer is unavailable on this machine.</summary>
        Unavailable,
        /// <summary>The renderer produced reference images.</summary>
        Succeeded,
        /// <summary>The renderer was available but failed.</summary>
        Failed
    }

    /// <summary>Result returned by the opt-in PowerPoint Desktop reference renderer.</summary>
    public sealed class PowerPointReferenceRenderResult {
        internal PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus status, string message,
            IEnumerable<string>? imagePaths = null) {
            Status = status;
            Message = message ?? string.Empty;
            ImagePaths = new ReadOnlyCollection<string>((imagePaths ?? Array.Empty<string>()).ToList());
        }

        /// <summary>Reference-render status.</summary>
        public PowerPointReferenceRenderStatus Status { get; }
        /// <summary>Human-readable renderer outcome.</summary>
        public string Message { get; }
        /// <summary>PowerPoint-generated slide image paths.</summary>
        public IReadOnlyList<string> ImagePaths { get; }
        /// <summary>Whether reference images were generated.</summary>
        public bool IsSuccessful => Status == PowerPointReferenceRenderStatus.Succeeded;
    }

    /// <summary>
    /// Explicit PowerPoint Desktop reference-render lane. Office automation is never invoked unless
    /// the method's <c>enabled</c> argument is true and is not used by ordinary image, HTML, PDF, or save operations.
    /// </summary>
    public static class PowerPointDesktopReferenceRenderer {
        /// <summary>Attempts to export each slide to PNG through locally installed PowerPoint Desktop.</summary>
        public static PowerPointReferenceRenderResult TryRender(string presentationPath, string outputDirectory,
            bool enabled = false) {
            if (!enabled) {
                return new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Disabled,
                    "PowerPoint Desktop reference rendering is opt-in and was not enabled.");
            }
            if (string.IsNullOrWhiteSpace(presentationPath)) {
                throw new ArgumentException("Presentation path cannot be empty.", nameof(presentationPath));
            }
            if (string.IsNullOrWhiteSpace(outputDirectory)) {
                throw new ArgumentException("Output directory cannot be empty.", nameof(outputDirectory));
            }
            string fullPath = Path.GetFullPath(presentationPath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Presentation was not found.", fullPath);
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Unavailable,
                    "PowerPoint Desktop COM rendering is available only on Windows.");
            }

            Type? powerPointType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (powerPointType == null) {
                return new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Unavailable,
                    "PowerPoint Desktop is not registered on this machine.");
            }

            string fullOutput = Path.GetFullPath(outputDirectory);
            Directory.CreateDirectory(fullOutput);
            object? application = null;
            object? presentations = null;
            object? presentation = null;
            try {
                application = Activator.CreateInstance(powerPointType);
                if (application == null) throw new InvalidOperationException("PowerPoint application could not be created.");
                presentations = GetProperty(application, "Presentations");
                presentation = InvokeMethod(presentations, "Open", fullPath, -1, 0, 0);
                InvokeMethod(presentation, "Export", fullOutput, "PNG", 0, 0);
                InvokeMethod(presentation, "Close");
                presentation = null;
                InvokeMethod(application, "Quit");
                application = null;
                string[] images = Directory.GetFiles(fullOutput, "*.png")
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
                return images.Length == 0
                    ? new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Failed,
                        "PowerPoint Desktop completed without producing PNG slide images.")
                    : new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Succeeded,
                        "PowerPoint Desktop exported " + images.Length + " slide image(s).", images);
            } catch (Exception ex) {
                return new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Failed,
                    "PowerPoint Desktop reference rendering failed: " + ex.Message);
            } finally {
                TryClosePresentation(presentation);
                TryQuitApplication(application);
                ReleaseComObject(presentation);
                ReleaseComObject(presentations);
                ReleaseComObject(application);
            }
        }

        private static void TryClosePresentation(object? presentation) {
            if (presentation == null) return;
            try { InvokeMethod(presentation, "Close"); } catch { }
        }

        private static void TryQuitApplication(object? application) {
            if (application == null) return;
            try { InvokeMethod(application, "Quit"); } catch { }
        }

        private static object GetProperty(object target, string name) =>
            target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null)
            ?? throw new MissingMemberException("PowerPoint COM property '" + name + "' returned null.");

        private static object InvokeMethod(object target, string name, params object[] arguments) =>
            target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, arguments)
            ?? target;

        private static void ReleaseComObject(object? value) {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
                value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.FinalReleaseComObject(value); } catch { }
        }
    }
}
