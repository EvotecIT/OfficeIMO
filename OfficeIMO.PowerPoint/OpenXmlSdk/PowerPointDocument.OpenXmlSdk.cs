using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The one place where a presentation held by the engine's own package meets the Open XML SDK. The engine hands over
    /// the package as it stands; the SDK opens over those bytes and is the document from then on, and slides and shapes
    /// read from the package bind to the SDK elements at the same places. Everything here moves into an opt-in package
    /// later, with the SDK dependency.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private PowerPointDocument(string filePath, DocumentAccessMode accessMode) {
            _filePath = filePath;
            _accessMode = accessMode;
            _loadedDocumentType = PresentationDocumentType.Presentation;
            _sdkDeferred = true;
        }

        private PresentationDocument OpenDeferredSdk() {
            // The SDK opens the package with any edits made without it.
            InvalidateRemovedPackageObjects(sdkOpening: true);
            CommitPendingEdits();
            ReleaseLoadedPackage();
            Stream packageStream = _packageStream ?? throw new InvalidOperationException("The presentation has no Open XML package to open.");
            packageStream.Position = 0;
            OpenSettings settings = CreateOpenSettings(_deferredOpenSettings);
            PresentationDocument document = _accessMode == DocumentAccessMode.ReadWrite
                ? PresentationDocument.Open(packageStream, true, settings)
                : OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenPresentation(packageStream, settings);
            try {
                InitializeFromSdk(document, isNewDocument: false);
            } catch {
                _sdkDocument = null;
                _sdkPresentationPart = null;
                _sdkSlides.Clear();
                document.Dispose();
                throw;
            }
            _sdkDeferred = false;
            _deferredOpenSettings = null;
            _packageSlides = null;
            _packageSlideOrder = null;
            _packageSlideEntries = null;
            _packageSlideEntryIndex = null;
            BindPackageObjects();
            DeferredSdkOpenCount++;
            RecordSdkFallback();
            return document;
        }

        /// <summary>Opens the SDK for a package-backed slide and returns the part it is bound to.</summary>
        internal SlidePart BindSdkSlide(PowerPointSlide slide) {
            ThrowIfDisposed();
            EnsureSdk();
            // Opening the SDK binds every slide still listed; one that is not was removed.
            return slide.IsPackageBacked
                ? throw new InvalidOperationException("The slide is no longer part of the presentation.")
                : slide.SlidePart;
        }

        /// <summary>The package-backed slide for an SDK slide part, which the SDK load binds instead of creating a new one.</summary>
        private PowerPointSlide? TakePackageSlide(SlidePart slidePart) {
            if (_packageSlides == null) return null;
            string identity = PartIdentity(slidePart.Uri.OriginalString);
            if (!_packageSlides.TryGetValue(identity, out PowerPointSlide? slide)) return null;
            slide.BindSdk(slidePart);
            return slide;
        }

        /// <summary>The presentation part through the SDK, opening it.</summary>
        private OpenXml.PowerPointPartRef SdkPresentationPartRef => OpenXml.PowerPointPartRef.FromSdk(SdkPresentationPart);

        /// <summary>Opens the SDK for a part too large for the package to read.</summary>
        internal void OpenSdkForOversizedPart() => EnsureSdk();

        // The SDK document's table styles part, added with the built-in styles when it has none.
        private OpenXml.PowerPointPartRef SdkEnsureTableStylesPart() {
            PowerPointUtils.CreateTableStylesPart(SdkPresentationPart);
            return OpenXml.PowerPointPartRef.FromSdk(SdkPresentationPart.TableStylesPart);
        }
    }
}
