using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Default maximum VBA project size accepted by the public mutation API.</summary>
        public const long DefaultMaximumVbaProjectBytes = 64L * 1024L * 1024L;

        /// <summary>Gets whether the presentation contains a VBA project.</summary>
        public bool HasVbaProject {
            get {
                ThrowIfDisposed();
                if (TryReadPackageVbaProject(maximumBytes: null, inspectSignature: false, out var project, out _)) return project.HasValue;
                return SdkPresentationPart.VbaProjectPart != null;
            }
        }

        /// <summary>
        ///     Reads the exact embedded VBA compound storage without changing the caller-visible package.
        /// </summary>
        public byte[]? GetVbaProjectBytes(
            long maximumBytes = DefaultMaximumVbaProjectBytes) {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumBytes);
            if (TryReadPackageVbaProject(maximumBytes, inspectSignature: false, out _, out byte[]? bytes)) return bytes;
            VbaProjectPart? part = SdkPresentationPart.VbaProjectPart;
            if (part == null) return null;
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            return OfficeStreamReader.ReadAllBytes(input, maximumBytes);
        }

        /// <summary>
        ///     Adds or replaces the embedded VBA project from a complete, valid compound storage.
        /// </summary>
        /// <remarks>
        ///     This API treats the VBA project as an opaque signed-capable artifact. It does not edit modules.
        ///     Save to a macro-enabled destination such as <c>.pptm</c>, <c>.potm</c>, or <c>.ppsm</c> to retain it.
        ///     Replacing a project removes related signature or cache parts because they no longer describe the new bytes.
        /// </remarks>
        public void SetVbaProject(Stream project,
            long maximumBytes = DefaultMaximumVbaProjectBytes) {
            ThrowIfDisposed();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumBytes);
            byte[] bytes = OfficeStreamReader.ReadAllBytes(project, maximumBytes);
            SetVbaProject(bytes, maximumBytes);
        }

        /// <summary>
        ///     Adds or replaces the embedded VBA project from a complete, valid compound storage.
        /// </summary>
        public void SetVbaProject(byte[] project,
            long maximumBytes = DefaultMaximumVbaProjectBytes) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(project);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumBytes);
            if (project.LongLength > maximumBytes) {
                throw new InvalidDataException(
                    $"VBA project exceeds the configured maximum size ({maximumBytes} bytes).");
            }
            if (!LegacyPptVbaProjectCodec.IsValidProject(project,
                    out string? reason)) {
                throw new InvalidDataException(
                    "VBA project is not a valid compound storage: " + reason);
            }

            if (TryMutatePackageVbaProject(project, out _)) return;

            VbaProjectPart part = SdkPresentationPart.VbaProjectPart ?? AddSdkVbaProject();
            foreach (IdPartPair child in part.Parts.ToArray()) {
                part.DeletePart(child.OpenXmlPart);
            }
            using var input = new MemoryStream(project, writable: false);
            part.FeedData(input);
        }

        private VbaProjectPart AddSdkVbaProject() {
            PresentationPart presentation = SdkPresentationPart;
            VbaProjectPart part = presentation.AddNewPart<VbaProjectPart>();
            if (_packagePartSequence == null || !_packagePartSequence.TryGetValue(VbaProjectContentType, out int previous)) return part;

            // After the SDK opens, its allocator starts with the parts still in the package. Consume any earlier VBA
            // names that were already allocated and deleted while the document used the owned package.
            const string stem = "vbaProject";
            int last = 0;
            while (true) {
                string name = Path.GetFileNameWithoutExtension(part.Uri.OriginalString);
                if (!name.StartsWith(stem, StringComparison.Ordinal)) return part;
                int number = 1;
                if (name.Length > stem.Length && !int.TryParse(name.Substring(stem.Length), NumberStyles.None,
                    CultureInfo.InvariantCulture, out number)) return part;
                if (number <= last || number > previous) return part;
                last = number;
                presentation.DeletePart(part);
                part = presentation.AddNewPart<VbaProjectPart>();
            }
        }

        /// <summary>Removes the embedded VBA project and its related signature or cache parts.</summary>
        public bool RemoveVbaProject() {
            ThrowIfDisposed();
            if (TryMutatePackageVbaProject(null, out bool removed)) return removed;
            VbaProjectPart? part = SdkPresentationPart.VbaProjectPart;
            if (part == null) return false;
            SdkPresentationPart.DeletePart(part);
            return true;
        }
    }
}
