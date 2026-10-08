using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Shared bounded VBA part operations for Word, Excel, and PowerPoint.</summary>
    internal static class OfficeVbaProjectPartEditor {
        internal static byte[] Read(VbaProjectPart part, long maximumBytes) {
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            return OfficeStreamReader.ReadAllBytes(input, maximumBytes);
        }

        /// <summary>Removes a newly created relationship and payload if initialization fails.</summary>
        internal static bool Apply(OpenXmlPart owner, VbaProjectPart? existing, byte[] bytes,
            OfficeVbaWriteOptions options, out VbaProjectPart part, Action<VbaProjectPart>? initialize = null) {
            ValidateRecoveryLimit(options);
            part = existing ?? owner.AddNewPart<VbaProjectPart>();
            try { return ApplyCore(part, bytes, options, initialize, existing != null); }
            catch {
                if (existing == null) owner.DeletePart(part);
                throw;
            }
        }

        internal static bool Apply(VbaProjectPart part, byte[] bytes, OfficeVbaWriteOptions options) =>
            ApplyCore(part, bytes, options, null, true);

        /// <summary>Checks signature authorization before a host changes its own package metadata.</summary>
        internal static void EnsureCanApply(VbaProjectPart? part, byte[] bytes, OfficeVbaWriteOptions options) {
            ValidateRecoveryLimit(options);
            if (part != null && !IsUnchanged(part, bytes)) CheckSignatureRemoval(GetSignatures(part), options);
        }

        private static bool ApplyCore(VbaProjectPart part, byte[] bytes, OfficeVbaWriteOptions options,
            Action<VbaProjectPart>? initialize, bool restoreExisting) {
            ValidateRecoveryLimit(options);
            bool changed = !IsUnchanged(part, bytes);
            if (!changed && initialize == null) return false;
            OpenXmlPart[] signatures = changed ? GetSignatures(part) : Array.Empty<OpenXmlPart>();
            CheckSignatureRemoval(signatures, options);
            using RollbackSnapshot? snapshot = restoreExisting ? RollbackSnapshot.Capture(part, options.MaximumRecoveryBytes) : null;
            try {
                initialize?.Invoke(part);
                if (changed) {
                    using MemoryStream input = new MemoryStream(bytes, writable: false);
                    part.FeedData(input);
                    foreach (OpenXmlPart signature in signatures) part.DeletePart(signature);
                }
                return changed;
            } catch (Exception failure) {
                if (snapshot != null) {
                    try { snapshot.Restore(part); }
                    catch (Exception restorationFailure) {
                        throw new AggregateException("The VBA update failed and its original package data could not be restored. Discard this document instance.", failure, restorationFailure);
                    }
                }
                throw;
            }
        }

        private static OpenXmlPart[] GetSignatures(VbaProjectPart part) => part.Parts.Where(pair =>
            pair.OpenXmlPart.RelationshipType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0
            || pair.OpenXmlPart.ContentType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0)
            .Select(pair => pair.OpenXmlPart).ToArray();

        private static void ValidateRecoveryLimit(OfficeVbaWriteOptions options) {
            if (options.MaximumRecoveryBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaximumRecoveryBytes));
        }

        private static void CheckSignatureRemoval(OpenXmlPart[] signatures, OfficeVbaWriteOptions options) {
            if (signatures.Length > 0 && !options.AllowSignatureRemoval) {
                throw new InvalidOperationException("Changing the VBA project invalidates its signatures. Explicit signature removal is required.");
            }
        }

        private static bool IsUnchanged(VbaProjectPart part, byte[] bytes) {
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            if (input.CanSeek && input.Length != bytes.LongLength) return false;
            try {
                return OfficeStreamReader.ReadAllBytes(input, Math.Max(1L, bytes.LongLength)).SequenceEqual(bytes);
            } catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
                return false;
            }
        }

        // The SDK copies the bounded VBA subgraph into an independent in-memory carrier,
        // including unknown signature profiles and opaque child parts. No document is saved.
        private sealed class RollbackSnapshot : IDisposable {
            private readonly MemoryStream _storage;
            private readonly SpreadsheetDocument _document;
            private readonly VbaProjectPart _project;
            private RollbackSnapshot(MemoryStream storage, SpreadsheetDocument document, VbaProjectPart project) {
                _storage = storage; _document = document; _project = project;
            }
            internal static RollbackSnapshot Capture(VbaProjectPart part, long maximumBytes) {
                long remaining = maximumBytes;
                HashSet<OpenXmlPart> visited = new HashSet<OpenXmlPart>();
                Queue<OpenXmlPart> pending = new Queue<OpenXmlPart>(); pending.Enqueue(part);
                while (pending.Count > 0) {
                    OpenXmlPart current = pending.Dequeue();
                    if (!visited.Add(current)) continue;
                    if (visited.Count > 1024) throw new InvalidDataException("The VBA recovery subgraph exceeds the part limit.");
                    using (Stream input = current.GetStream(FileMode.Open, FileAccess.Read)) {
                        remaining -= OfficeStreamReader.ReadAllBytes(input, Math.Max(1L, remaining)).LongLength;
                        if (remaining < 0) throw new InvalidDataException("The VBA recovery subgraph exceeds the configured recovery byte limit.");
                    }
                    foreach (IdPartPair child in current.Parts) pending.Enqueue(child.OpenXmlPart);
                }
                MemoryStream storage = new MemoryStream();
                SpreadsheetDocument? document = null;
                try {
                    document = SpreadsheetDocument.Create(storage, SpreadsheetDocumentType.MacroEnabledWorkbook);
                    VbaProjectPart copy = document.AddWorkbookPart().AddPart(part);
                    return new RollbackSnapshot(storage, document, copy);
                } catch { document?.Dispose(); storage.Dispose(); throw; }
            }
            internal void Restore(VbaProjectPart part) {
                foreach (IdPartPair child in part.Parts.ToArray()) part.DeletePart(child.OpenXmlPart);
                using (Stream input = _project.GetStream(FileMode.Open, FileAccess.Read)) part.FeedData(input);
                foreach (IdPartPair child in _project.Parts) part.AddPart(child.OpenXmlPart, child.RelationshipId);
            }
            public void Dispose() { _document.Dispose(); _storage.Dispose(); }
        }
    }
}
