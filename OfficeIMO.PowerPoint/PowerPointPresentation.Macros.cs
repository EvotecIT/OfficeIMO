using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>Default maximum VBA project size accepted by the public mutation API.</summary>
        public const long DefaultMaximumVbaProjectBytes = 64L * 1024L * 1024L;

        /// <summary>Gets whether the presentation contains a VBA project.</summary>
        public bool HasVbaProject {
            get {
                ThrowIfDisposed();
                return _presentationPart.VbaProjectPart != null;
            }
        }

        /// <summary>
        ///     Reads the exact embedded VBA compound storage without changing the caller-visible package.
        /// </summary>
        public byte[]? GetVbaProjectBytes(
            long maximumBytes = DefaultMaximumVbaProjectBytes) {
            ThrowIfDisposed();
            if (maximumBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            }
            VbaProjectPart? part = _presentationPart.VbaProjectPart;
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
            if (maximumBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            }
            byte[] bytes = OfficeStreamReader.ReadAllBytes(project, maximumBytes);
            SetVbaProject(bytes, maximumBytes);
        }

        /// <summary>
        ///     Adds or replaces the embedded VBA project from a complete, valid compound storage.
        /// </summary>
        public void SetVbaProject(byte[] project,
            long maximumBytes = DefaultMaximumVbaProjectBytes) {
            ThrowIfDisposed();
            if (project == null) throw new ArgumentNullException(nameof(project));
            if (maximumBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumBytes));
            }
            if (project.LongLength > maximumBytes) {
                throw new InvalidDataException(
                    $"VBA project exceeds the configured maximum size ({maximumBytes} bytes).");
            }
            if (!LegacyPptVbaProjectCodec.IsValidProject(project,
                    out string? reason)) {
                throw new InvalidDataException(
                    "VBA project is not a valid compound storage: " + reason);
            }

            VbaProjectPart part = _presentationPart.VbaProjectPart
                ?? _presentationPart.AddNewPart<VbaProjectPart>();
            foreach (IdPartPair child in part.Parts.ToArray()) {
                part.DeletePart(child.OpenXmlPart);
            }
            using var input = new MemoryStream(project, writable: false);
            part.FeedData(input);
        }

        /// <summary>Removes the embedded VBA project and its related signature or cache parts.</summary>
        public bool RemoveVbaProject() {
            ThrowIfDisposed();
            VbaProjectPart? part = _presentationPart.VbaProjectPart;
            if (part == null) return false;
            _presentationPart.DeletePart(part);
            return true;
        }
    }
}
