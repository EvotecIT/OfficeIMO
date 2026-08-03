using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
        [RequiresUnreferencedCode("PowerPoint Desktop rendering uses late-bound COM automation. Use OfficeIMO's in-process renderers for trimmed deployments.")]
        [RequiresDynamicCode("PowerPoint Desktop rendering uses late-bound COM automation and is not a NativeAOT deployment path.")]
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
            ClearExistingSlideImages(fullOutput);
            int expectedSlideCount;
            bool[] slidesWithVisibleContent;
            try {
                using PowerPointPresentation source = PowerPointPresentation.Load(
                    fullPath, new PowerPointLoadOptions {
                        AccessMode = DocumentAccessMode.ReadOnly
                    });
                expectedSlideCount = source.Slides.Count;
                slidesWithVisibleContent = source.Slides
                    .Select(HasExpectedVisibleContent)
                    .ToArray();
            } catch (Exception ex) {
                return new PowerPointReferenceRenderResult(
                    PowerPointReferenceRenderStatus.Failed,
                    "PowerPoint Desktop reference rendering could not inspect the source package: "
                    + GetRootMessage(ex));
            }
            object? application = null;
            object? presentations = null;
            object? presentation = null;
            try {
                application = Activator.CreateInstance(powerPointType);
                if (application == null) throw new InvalidOperationException("PowerPoint application could not be created.");
                if (!TryConfigureApplicationSecurity(application,
                        out string securityMessage)) {
                    return new PowerPointReferenceRenderResult(
                        PowerPointReferenceRenderStatus.Failed,
                        securityMessage);
                }
                presentations = GetProperty(application, "Presentations");
                presentation = InvokeMethod(presentations, "Open", fullPath, -1, 0, 0);
                InvokeMethod(presentation, "Export", fullOutput, "PNG", 0, 0);
                InvokeMethod(presentation, "Close");
                InvokeMethod(application, "Quit");
                string[] images = GetSlideImagesInOrder(fullOutput);
                return !ValidateSlideImages(images, expectedSlideCount,
                        slidesWithVisibleContent,
                        out string validationMessage)
                    ? new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Failed,
                        validationMessage)
                    : new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Succeeded,
                        "PowerPoint Desktop exported " + images.Length + " slide image(s).", images);
            } catch (Exception ex) {
                return new PowerPointReferenceRenderResult(PowerPointReferenceRenderStatus.Failed,
                    "PowerPoint Desktop reference rendering failed: " + GetRootMessage(ex));
            } finally {
                TryClosePresentation(presentation);
                TryQuitApplication(application);
                ReleaseComObject(presentation);
                ReleaseComObject(presentations);
                ReleaseComObject(application);
            }
        }

        internal static void ClearExistingSlideImages(string outputDirectory) {
            foreach (string path in Directory.EnumerateFiles(outputDirectory)) {
                if (!string.Equals(Path.GetExtension(path), ".png", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                if (TryGetSlideNumber(path, out _)) File.Delete(path);
            }
        }

        internal static string[] GetSlideImagesInOrder(string outputDirectory) =>
            Directory.EnumerateFiles(outputDirectory)
                .Where(path => string.Equals(Path.GetExtension(path), ".png", StringComparison.OrdinalIgnoreCase))
                .Where(path => TryGetSlideNumber(path, out _))
                .OrderBy(path => {
                    TryGetSlideNumber(path, out int number);
                    return number;
                })
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();

        internal static bool ValidateSlideImages(IReadOnlyList<string> imagePaths,
            int expectedSlideCount, out string message) =>
            ValidateSlideImages(imagePaths, expectedSlideCount,
                Enumerable.Repeat(true, Math.Max(0, expectedSlideCount)).ToArray(),
                out message);

        internal static bool ValidateSlideImages(IReadOnlyList<string> imagePaths,
            int expectedSlideCount, IReadOnlyList<bool> slidesWithVisibleContent,
            out string message) {
            if (expectedSlideCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(expectedSlideCount));
            }
            if (slidesWithVisibleContent == null) {
                throw new ArgumentNullException(nameof(slidesWithVisibleContent));
            }
            if (slidesWithVisibleContent.Count != expectedSlideCount) {
                throw new ArgumentException(
                    "Visible-content flags must match the expected slide count.",
                    nameof(slidesWithVisibleContent));
            }
            if (expectedSlideCount == 0 || imagePaths.Count == 0) {
                message = "PowerPoint Desktop exported no slide images.";
                return false;
            }
            if (imagePaths.Count != expectedSlideCount) {
                message = "PowerPoint Desktop exported " + imagePaths.Count
                    + " PNG slide image(s); expected " + expectedSlideCount + ".";
                return false;
            }
            for (int index = 0; index < imagePaths.Count; index++) {
                string path = imagePaths[index];
                if (!TryGetSlideNumber(path, out int slideNumber)
                    || slideNumber != index + 1) {
                    message = "PowerPoint Desktop did not export the contiguous slide image Slide"
                        + (index + 1) + ".png; found " + Path.GetFileName(path) + ".";
                    return false;
                }
                if (!File.Exists(path)) {
                    message = "PowerPoint Desktop did not create the expected slide image: " + path;
                    return false;
                }
                byte[] bytes = File.ReadAllBytes(path);
                if (!OfficePngReader.TryDecode(bytes, out OfficeRasterImage? image)
                    || image == null || image.Width <= 0 || image.Height <= 0) {
                    message = "PowerPoint Desktop created an invalid PNG slide image: " + path;
                    return false;
                }
                if (slidesWithVisibleContent[index]
                    && !HasMeaningfulNonWhiteContent(image)) {
                    message = "PowerPoint Desktop created a blank PNG for slide "
                        + (index + 1) + ": " + path;
                    return false;
                }
            }
            message = string.Empty;
            return true;
        }

        internal static bool HasExpectedVisibleContent(PowerPointSlide slide) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            if (slide.SmartArts.Any(smartArt => !smartArt.Hidden
                    && IsPotentiallyVisibleOnSlide(slide, smartArt)
                    && !smartArt.TryGetOfficeDiagramSnapshot(out _))) {
                return true;
            }
            OfficeImageExportResult rendered = slide.ExportImage(
                OfficeImageExportFormat.Png);
            if (HasVisibleOmittedContent(slide, rendered.Diagnostics)) {
                return true;
            }
            return OfficePngReader.TryDecode(rendered.Bytes,
                       out OfficeRasterImage? image)
                   && image != null
                   && HasMeaningfulNonWhiteContent(image);
        }

        private static bool HasVisibleOmittedContent(PowerPointSlide slide,
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            bool HasOmissionContaining(string text) => diagnostics.Any(
                diagnostic => diagnostic.Code
                        == PowerPointImageExportDiagnosticCodes.UnsupportedShape
                    && diagnostic.LossKind == OfficeImageExportLossKind.Omission
                    && diagnostic.Message.IndexOf(text,
                        StringComparison.OrdinalIgnoreCase) >= 0);

            if (HasOmissionContaining("PowerPoint chart")
                && slide.Charts.Any(chart => IsPotentiallyVisibleOnSlide(
                    slide, chart))) {
                return true;
            }

            return HasOmissionContaining("custom geometry")
                && slide.Shapes.Any(shape =>
                    IsPotentiallyVisibleOnSlide(slide, shape)
                    && shape.Element.Descendants<A.CustomGeometry>().Any());
        }

        private static bool IsPotentiallyVisibleOnSlide(PowerPointSlide slide,
            PowerPointShape shape) {
            if (shape.Hidden || shape.Width <= 0L || shape.Height <= 0L) {
                return false;
            }
            P.SlideSize? size = slide.SlidePart.GetParentParts()
                .OfType<DocumentFormat.OpenXml.Packaging.PresentationPart>()
                .FirstOrDefault()?.Presentation?.SlideSize;
            long width = size?.Cx?.Value ?? 0L;
            long height = size?.Cy?.Value ?? 0L;
            if (width <= 0L || height <= 0L) return true;
            return shape.Left < width && shape.Top < height
                && shape.Right > 0L && shape.Bottom > 0L;
        }

        private static bool HasMeaningfulNonWhiteContent(OfficeRasterImage image) {
            int requiredPixels = Math.Min(16, checked(image.Width * image.Height));
            int visiblePixels = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A == 0) continue;
                    int alpha = pixel.A;
                    int red = (pixel.R * alpha + 255 * (255 - alpha) + 127)
                        / 255;
                    int green = (pixel.G * alpha + 255 * (255 - alpha) + 127)
                        / 255;
                    int blue = (pixel.B * alpha + 255 * (255 - alpha) + 127)
                        / 255;
                    if (red < 245 || green < 245 || blue < 245) {
                        visiblePixels++;
                        if (visiblePixels >= requiredPixels) return true;
                    }
                }
            }
            return false;
        }

        private static bool TryGetSlideNumber(string path, out int number) {
            number = 0;
            string name = Path.GetFileNameWithoutExtension(path);
            return name.Length > 5 && name.StartsWith("Slide", StringComparison.OrdinalIgnoreCase) &&
                   int.TryParse(name.Substring(5), out number) && number > 0;
        }

        private static void TryClosePresentation(object? presentation) {
            if (presentation == null) return;
            try { InvokeMethod(presentation, "Close"); } catch { }
        }

        private static void TryQuitApplication(object? application) {
            if (application == null) return;
            try { InvokeMethod(application, "Quit"); } catch { }
        }

        [UnconditionalSuppressMessage("Trimming", "IL2075", Justification = "Late-bound COM members are supplied by installed PowerPoint Desktop and are outside the managed trimming graph.")]
        internal static void SetRequiredProperty(object target, string name,
            object value) {
            target.GetType().InvokeMember(name, BindingFlags.SetProperty,
                null, target, new[] { value });
        }

        internal static bool TryConfigureApplicationSecurity(object application,
            out string message) {
            try {
                SetRequiredProperty(application, "AutomationSecurity", 3);
                message = string.Empty;
                return true;
            } catch (Exception exception) {
                message = "PowerPoint Desktop reference rendering could not force-disable macros: "
                    + GetRootMessage(exception);
                return false;
            }
        }

        private static string GetRootMessage(Exception exception) {
            Exception root = exception is TargetInvocationException invocation
                && invocation.InnerException != null
                ? invocation.InnerException
                : exception;
            return root.Message;
        }

        [UnconditionalSuppressMessage("Trimming", "IL2075", Justification = "Late-bound COM members are supplied by installed PowerPoint Desktop and are outside the managed trimming graph.")]
        private static object GetProperty(object target, string name) =>
            target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null)
            ?? throw new MissingMemberException("PowerPoint COM property '" + name + "' returned null.");

        [UnconditionalSuppressMessage("Trimming", "IL2075", Justification = "Late-bound COM members are supplied by installed PowerPoint Desktop and are outside the managed trimming graph.")]
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
