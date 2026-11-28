using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// A plain, unsigned Open XML presentation, template, slideshow or add-in loads without opening the Open XML SDK.
    /// The owned package reader checks the input. A later SDK operation opens the same stream and binds objects
    /// that were read from the owned package.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private const string PresentationMainContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
        private const string DigitalSignatureOriginRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        private const string SlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
        // The byte limit the signature check reads docProps/app.xml under when the SDK opens at load.
        private const long MaximumApplicationPropertiesBytes = 64L * 1024 * 1024;
        private const string ExtendedPropertiesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "extended-properties";

        // The settings a deferred load opens the SDK package with; set while the SDK has not been opened.
        private OfficeOpenXmlLoadSettings? _deferredOpenSettings;
        private bool _sdkDeferred;
        // The number of slide list entries the load checked; the slides the SDK would build from them.
        private int _deferredSlideCount;


        /// <summary>
        /// A new presentation over the default package bytes that works without the SDK: its creation and modification
        /// dates are set in the package, as the SDK sets them; null when the package doesn't have them to set.
        /// </summary>
        private static PowerPointDocument? TryCreateWithoutSdk(byte[] packageBytes, string? filePath, Stream? sourceStream,
            PowerPointCreateOptions options, OfficeTheme? theme) {
            var packageStream = new MemoryStream(packageBytes.Length + StreamBufferSize);
            packageStream.Write(packageBytes, 0, packageBytes.Length);
            packageStream.Position = 0;
            var presentation = new PowerPointDocument(filePath ?? string.Empty, DocumentAccessMode.ReadWrite) {
                _packageStream = packageStream,
                _sourceStream = sourceStream,
                _persistenceMode = options.PersistenceMode,
                _maximumSlideXmlCharacters = PowerPointXmlReader.MaximumPackageXmlCharacters
            };
            try {
                if (presentation.GetPackage(edit: true) is not { } package || PackageSlideList.Read(package) is not { } slides) {
                    presentation.Dispose();
                    return null;
                }
                string? core = FirstRelated(package, null, CorePropertiesRelationshipType, out _);
                OfficeXmlTree? properties = core == null ? null : package.GetXml(core);
                int createdDate = properties?.FindChild(properties.Root, OfficeXmlNamespaces.DublinCoreTerms, "created") ?? OfficeXmlTree.NoNode;
                int modifiedDate = properties?.FindChild(properties.Root, OfficeXmlNamespaces.DublinCoreTerms, "modified") ?? OfficeXmlTree.NoNode;
                if (createdDate == OfficeXmlTree.NoNode || modifiedDate == OfficeXmlTree.NoNode) {
                    presentation.Dispose();
                    return null;
                }
                string timestamp = FormatPackageDate(DateTime.UtcNow);
                properties!.SetInnerText(createdDate, timestamp);
                properties.SetInnerText(modifiedDate, timestamp);
                presentation._pendingEdits = true;
                presentation._deferredSlideCount = slides.Entries.Count;
                if (theme != null) {
                    presentation._baseTheme = theme;
                    presentation._layoutTheme = PowerPointLayoutTheme.From(theme);
                }
                return presentation;
            } catch {
                presentation.Dispose();
                throw;
            }
        }

        // The read-only package over the bytes as loaded, which the load checked; null once the document's bytes change.
        private OfficeOpcPackage? _loadedPackage;

        /// <summary>The read-only package over the document's current bytes, when it is still the one the load opened.</summary>
        private OfficeOpcPackage? LoadedPackageForCurrentBytes => _loadedPackage != null && !_pendingEdits ? _loadedPackage : null;

        private void ReleaseLoadedPackage() {
            _loadedPackage?.Dispose();
            _loadedPackage = null;
        }

        /// <summary>How many times the SDK package was opened after a deferred load; zero while it was not needed.</summary>
        internal int DeferredSdkOpenCount { get; private set; }


        // Settings beyond the defaults change how the SDK reads parts, so those loads open it right away. A part size
        // limit is kept: the SDK gets it when it opens, and the load checks the one part the SDK would read.
        private static bool IsPlainOpenSettings(OfficeOpenXmlLoadSettings? settings) =>
            settings == null
            || settings.MaxCharactersInPart >= 0
            && settings.CompatibilityLevel == new OfficeOpenXmlLoadSettings().CompatibilityLevel
            && settings.MarkupCompatibilityMode == new OfficeOpenXmlLoadSettings().MarkupCompatibilityMode
            && settings.MarkupCompatibilityTargetVersion == new OfficeOpenXmlLoadSettings().MarkupCompatibilityTargetVersion;

        /// <summary>
        /// Checks for an unsigned Transitional package and resolves every slide entry to a slide part.
        /// Other packages use the existing SDK load.
        /// </summary>
        private static bool CanDeferSdk(Stream packageStream, PowerPointLoadOptions options, out int slideCount,
            out OfficeOpcPackage? loaded, out string? sourceExtension) {
            slideCount = 0;
            loaded = null;
            sourceExtension = null;
            if (!IsPlainOpenSettings(options.OpenSettings) || !packageStream.CanSeek) return false;
            long position = packageStream.Position;
            OfficeOpcPackage? package = null;
            try {
                packageStream.Position = 0;
                package = OfficeOpcPackage.Open(packageStream, leaveOpen: true);
                sourceExtension = GetUnsignedPresentationExtension(package, options.OpenSettings?.MaxCharactersInPart ?? 0, out slideCount);
                if (sourceExtension == null) {
                    package.Dispose();
                    return false;
                }
                // The document keeps the package it checked, so reads and saves of the bytes as loaded don't open it again.
                loaded = package;
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException
                or ArgumentException or FormatException or NotSupportedException or InvalidOperationException or OverflowException) {
                package?.Dispose();
                return false;
            } finally {
                packageStream.Position = position;
            }
        }

        private static string? GetUnsignedPresentationExtension(OfficeOpcPackage package, long maximumCharactersInPart, out int slideCount) {
            slideCount = 0;
            // A signature part, even one nothing refers to, is reported by the signature check; such a package opens
            // the SDK at load, so a deferred package has none.
            if (package.DeclaresContentType("application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml")
                || package.DeclaresContentType("application/vnd.openxmlformats-package.digital-signature-origin")) {
                return null;
            }
            IReadOnlyList<OfficeOpcRelationship> relationships = package.GetRelationships(null);
            foreach (OfficeOpcRelationship relationship in relationships) {
                if (string.Equals(relationship.Type, DigitalSignatureOriginRelationshipType, StringComparison.Ordinal)) return null;
            }
            string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
            if (main == null || !package.PartExists(main)) return null;

            string? contentType = package.GetContentType(main);
            string? extension = string.Equals(contentType, PresentationMainContentType, StringComparison.OrdinalIgnoreCase)
                ? ".pptx"
                : PresentationExtension(contentType);
            if (extension == null) return null;
            bool validateRoot = extension != ".pptx";
            if (validateRoot && (package.HasInterleavedParts || !CanDeferSubtypeRelationships(package, relationships))) return null;
            // A UTF-8 part has no more characters than bytes, so one within the limit in bytes is within it in characters.
            if (maximumCharactersInPart > 0 && package.GetPartLength(main) > maximumCharactersInPart) return null;
            string? application = package.GetRelatedPartName(null, ExtendedPropertiesRelationshipType);
            // The signature check reads the conventional properties part even when nothing relates to it.
            bool validateApplicationRoot = validateRoot && application != null;
            if (application == null && package.PartExists("/docProps/app.xml")) application = "/docProps/app.xml";
            if (application != null && package.PartExists(application)) {
                // The signature check at an SDK load reads this part under the part size limit too.
                long length = package.GetPartLength(application);
                if (length > MaximumApplicationPropertiesBytes || maximumCharactersInPart > 0 && length > maximumCharactersInPart) return null;
                if (HasDigitalSignatureFlag(package, application, validateApplicationRoot)) return null;
            }
            return SlideListResolves(package, main, out slideCount) ? extension : null;
        }

        // These subtypes previously validated their typed root parts during the SDK load.
        private static bool CanDeferSubtypeRelationships(OfficeOpcPackage package, IReadOnlyList<OfficeOpcRelationship> relationships) {
            int mainParts = 0;
            foreach (OfficeOpcRelationship relationship in relationships) {
                if (relationship.IsExternal) continue;
                if (relationship.TargetPartName is not { } part || !package.PartExists(part)) return false;
                string? expectedType;
                switch (relationship.Type) {
                    case OfficeOpcPackage.OfficeDocumentRelationshipType:
                        mainParts++;
                        continue;
                    case OfficeOpcPackage.CorePropertiesRelationshipType:
                        expectedType = PowerPointPackageModel.CorePropertiesContentType;
                        break;
                    case ExtendedPropertiesRelationshipType:
                        expectedType = PowerPointPackageModel.ExtendedPropertiesContentType;
                        break;
                    case OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "custom-properties":
                        expectedType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
                        break;
                    case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail":
                        if (package.GetContentType(part) == null) return false;
                        continue;
                    default:
                        return false;
                }
                if (!string.Equals(expectedType, package.GetContentType(part), StringComparison.Ordinal)) return false;
            }
            return mainParts == 1;
        }

        // These reads leave markup compatibility as written, as the SDK's default NoProcess mode reads it.
        private static bool HasDigitalSignatureFlag(OfficeOpcPackage package, string application, bool validateRoot) {
            OfficeXmlPullReader reader = package.OpenXml(application, out OfficePooledBytes content, processMarkupCompatibility: false);
            using (content) {
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (validateRoot && reader.Depth == 1 && !reader.IsElement(OfficeXmlNamespaces.ExtendedProperties, "Properties")) {
                        throw new InvalidDataException("The application properties part has no Properties root element.");
                    }
                    if (reader.Depth == 2 && reader.IsElement(OfficeXmlNamespaces.ExtendedProperties, "DigSig")) {
                        return true;
                    }
                }
            }
            return false;
        }

        // The SDK reads the slide list at load and fails on an entry it cannot resolve or whose part is not a slide; such
        // a package keeps doing so.
        private static bool SlideListResolves(OfficeOpcPackage package, string main, out int slideCount) {
            slideCount = 0;
            OfficeXmlPullReader reader = package.OpenXml(main, out OfficePooledBytes content, processMarkupCompatibility: false);
            using (content) {
                bool root = false;
                // The SDK reads the first slide list under the root; entries anywhere else are not slides to it.
                bool inSlideList = false, slideListSeen = false;
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (reader.Depth == 2) {
                        inSlideList = !slideListSeen && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldIdLst");
                        slideListSeen |= inSlideList;
                    }
                    if (reader.Depth == 1) {
                        // Strict markup goes through the SDK, which maps it.
                        if (!reader.IsElement(OfficeXmlNamespaces.PresentationML, "presentation")
                            || reader.ElementUriId != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.PresentationML)) {
                            return false;
                        }
                        root = true;
                        continue;
                    }
                    if (reader.Depth != 3 || !inSlideList || !reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldId")) continue;
                    if (!reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) return false;
                    OfficeOpcRelationship? slide = package.FindRelationship(main, id);
                    if (slide == null || slide.IsExternal
                        || !string.Equals(slide.Type, OfficeOpcPackage.SlideRelationshipType, StringComparison.Ordinal)
                        || slide.TargetPartName is not { } target || !package.PartExists(target)
                        || !string.Equals(package.GetContentType(target), SlideContentType, StringComparison.OrdinalIgnoreCase)) {
                        return false;
                    }
                    slideCount++;
                }
                return root;
            }
        }

        /// <summary>
        /// Extracts chunks with the streamed slide reader while the SDK package is not open, or returns null for the
        /// model to extract them. The chunks are read in full first, so a package the reader declines partway through
        /// goes to the model with nothing returned.
        /// </summary>
        internal List<PowerPointExtractChunk>? TryExtractChunksWithoutSdk(PowerPointExtractionExtensions.PowerPointExtractOptions extract,
            int maxChars, string? sourcePath, CancellationToken cancellationToken) {
            if (_disposed || !_sdkDeferred || IsSdkLoaded) return null;
            // The package the load opened still reads the bytes as they are, so the reader shares it.
            if (LoadedPackageForCurrentBytes is { } loadedPackage) {
                try {
                    using PowerPointSlideReader reader = PowerPointSlideReader.OpenLoaded(loadedPackage, _maximumSlideXmlCharacters, cancellationToken);
                    var chunking = new PowerPointExtractChunkingOptions { MaxChars = maxChars };
                    return PowerPointExtractionExtensions.ExtractMarkdownChunks(reader, extract, chunking, sourcePath, cancellationToken).ToList();
                } catch (OperationCanceledException) {
                    throw;
                } catch (Exception exception) when (exception is not OutOfMemoryException) {
                    return null;
                }
            }
            if (!TryGetLoadedPackage(out MemoryStream? loaded) || loaded is not { } package) return null;
            try {
                using (package) {
                    var options = new PowerPointLoadOptions { MaxInputBytes = null, OpenSettings = _deferredOpenSettings };
                    using PowerPointSlideReader reader = PowerPointSlideReader.Open(package, options, cancellationToken);
                    var chunking = new PowerPointExtractChunkingOptions { MaxChars = maxChars };
                    return PowerPointExtractionExtensions.ExtractMarkdownChunks(reader, extract, chunking, sourcePath, cancellationToken).ToList();
                }
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
                return null;
            }
        }

        // A read-only view over the loaded bytes, without copying them; null when the package stream holds none.
        private bool TryGetLoadedPackage(out MemoryStream? package) {
            CommitPendingEdits(keepPackage: true);
            package = null;
            ArraySegment<byte> buffer = default;
            bool found = _packageStream switch {
                OfficeIMO.Core.Internal.OfficeCopyOnWriteStream copyOnWrite => copyOnWrite.TryGetBuffer(out buffer),
                MemoryStream memory => memory.TryGetBuffer(out buffer),
                _ => false
            };
            if (!found || buffer.Array == null) return false;
            package = new MemoryStream(buffer.Array, buffer.Offset, buffer.Count, writable: false);
            return true;
        }
    }
}
