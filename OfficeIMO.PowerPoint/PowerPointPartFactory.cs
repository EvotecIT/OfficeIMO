using System.Globalization;
using System.IO;
using System.Linq;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointPartFactory {
        private static readonly MethodInfo CreateInternalMethod = ResolveCreateInternal();
        private static readonly bool CreateInternalUsesUri =
            CreateInternalMethod.GetParameters()[3].ParameterType == typeof(Uri);

#if NET5_0_OR_GREATER
        [DynamicDependency("CreateInternal", typeof(OpenXmlPart))]
#endif
        private static MethodInfo ResolveCreateInternal() {
            var methods = typeof(OpenXmlPart)
                .GetMethods(BindingFlags.Instance | BindingFlags.NonPublic)
                .Where(m => string.Equals(m.Name, "CreateInternal", StringComparison.Ordinal))
                .ToList();

            var uriOverload = methods.FirstOrDefault(m => {
                var parameters = m.GetParameters();
                return parameters.Length == 4
                    && parameters[0].ParameterType == typeof(OpenXmlPackage)
                    && parameters[1].ParameterType == typeof(OpenXmlPart)
                    && parameters[2].ParameterType == typeof(string)
                    && parameters[3].ParameterType == typeof(Uri);
            });
            if (uriOverload != null) {
                return uriOverload;
            }

            var stringOverload = methods.FirstOrDefault(m => {
                var parameters = m.GetParameters();
                return parameters.Length == 4
                    && parameters[0].ParameterType == typeof(OpenXmlPackage)
                    && parameters[1].ParameterType == typeof(OpenXmlPart)
                    && parameters[2].ParameterType == typeof(string)
                    && parameters[3].ParameterType == typeof(string);
            });
            if (stringOverload != null) {
                return stringOverload;
            }

            throw new NotSupportedException("OpenXmlPart.CreateInternal overload was not found.");
        }

#if NET5_0_OR_GREATER
        internal static TPart CreatePart<
            [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor | DynamicallyAccessedMemberTypes.NonPublicConstructors)] TPart>(
#else
        internal static TPart CreatePart<TPart>(
#endif
            OpenXmlPart parent,
            string? contentType,
            string partUri,
            string? relationshipId = null) where TPart : OpenXmlPart {
            if (parent == null) {
                throw new ArgumentNullException(nameof(parent));
            }
            if (parent.OpenXmlPackage == null) {
                throw new InvalidOperationException("Parent part is not attached to a package.");
            }
            if (string.IsNullOrWhiteSpace(partUri)) {
                throw new ArgumentException("Part URI is required.", nameof(partUri));
            }

            object? instance = Activator.CreateInstance(typeof(TPart), nonPublic: true);
            if (instance is not TPart part) {
                throw new InvalidOperationException($"Unable to create instance of {typeof(TPart).Name}.");
            }
            string? resolvedContentType = !string.IsNullOrWhiteSpace(contentType)
                ? contentType
                : part.ContentType;
            if (string.IsNullOrWhiteSpace(resolvedContentType)) {
                throw new InvalidOperationException($"Content type for {typeof(TPart).Name} is required.");
            }

            string normalizedUri = NormalizePartUri(partUri);
            if (CreateInternalUsesUri) {
                Uri partUriValue = new Uri(normalizedUri, UriKind.Relative);
                CreateInternalMethod.Invoke(part, new object?[] { parent.OpenXmlPackage, parent, resolvedContentType!, partUriValue });
            } else {
                CreateInternalMethod.Invoke(part, new object?[] { parent.OpenXmlPackage, parent, resolvedContentType!, normalizedUri });
            }

            if (string.IsNullOrWhiteSpace(relationshipId)) {
                parent.AddPart(part);
            } else {
                parent.AddPart(part, relationshipId!);
            }

            return part;
        }

        internal static OpenXmlPart AddNewPartLike(
            OpenXmlPartContainer parent,
            OpenXmlPart sourcePart,
            string relationshipId) {
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (sourcePart == null) throw new ArgumentNullException(nameof(sourcePart));
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                throw new ArgumentException("Relationship identifier is required.", nameof(relationshipId));
            }

            string contentType = sourcePart.ContentType;
            return sourcePart switch {
                AlternativeFormatImportPart => parent.AddNewPart<AlternativeFormatImportPart>(contentType, relationshipId),
                ChartColorStylePart => parent.AddNewPart<ChartColorStylePart>(contentType, relationshipId),
                ChartDrawingPart => parent.AddNewPart<ChartDrawingPart>(contentType, relationshipId),
                ChartPart => parent.AddNewPart<ChartPart>(contentType, relationshipId),
                ChartStylePart => parent.AddNewPart<ChartStylePart>(contentType, relationshipId),
                CommentAuthorsPart => parent.AddNewPart<CommentAuthorsPart>(contentType, relationshipId),
                ControlPropertiesPart => parent.AddNewPart<ControlPropertiesPart>(contentType, relationshipId),
                CoreFilePropertiesPart => parent.AddNewPart<CoreFilePropertiesPart>(contentType, relationshipId),
                CustomDataPart => parent.AddNewPart<CustomDataPart>(contentType, relationshipId),
                CustomDataPropertiesPart => parent.AddNewPart<CustomDataPropertiesPart>(contentType, relationshipId),
                CustomFilePropertiesPart => parent.AddNewPart<CustomFilePropertiesPart>(contentType, relationshipId),
                CustomPropertyPart => parent.AddNewPart<CustomPropertyPart>(contentType, relationshipId),
                CustomXmlMappingsPart => parent.AddNewPart<CustomXmlMappingsPart>(contentType, relationshipId),
                CustomXmlPart => parent.AddNewPart<CustomXmlPart>(contentType, relationshipId),
                CustomXmlPropertiesPart => parent.AddNewPart<CustomXmlPropertiesPart>(contentType, relationshipId),
                DiagramColorsPart => parent.AddNewPart<DiagramColorsPart>(contentType, relationshipId),
                DiagramDataPart => parent.AddNewPart<DiagramDataPart>(contentType, relationshipId),
                DiagramLayoutDefinitionPart => parent.AddNewPart<DiagramLayoutDefinitionPart>(contentType, relationshipId),
                DiagramPersistLayoutPart => parent.AddNewPart<DiagramPersistLayoutPart>(contentType, relationshipId),
                DiagramStylePart => parent.AddNewPart<DiagramStylePart>(contentType, relationshipId),
                DigitalSignatureOriginPart => parent.AddNewPart<DigitalSignatureOriginPart>(contentType, relationshipId),
                EmbeddedControlPersistenceBinaryDataPart => parent.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>(contentType, relationshipId),
                EmbeddedControlPersistencePart => parent.AddNewPart<EmbeddedControlPersistencePart>(contentType, relationshipId),
                EmbeddedObjectPart => parent.AddNewPart<EmbeddedObjectPart>(contentType, relationshipId),
                EmbeddedPackagePart => parent.AddNewPart<EmbeddedPackagePart>(contentType, relationshipId),
                ExtendedChartPart => parent.AddNewPart<ExtendedChartPart>(contentType, relationshipId),
                ExtendedFilePropertiesPart => parent.AddNewPart<ExtendedFilePropertiesPart>(contentType, relationshipId),
                ExternalWorkbookPart => parent.AddNewPart<ExternalWorkbookPart>(contentType, relationshipId),
                FontPart => parent.AddNewPart<FontPart>(contentType, relationshipId),
                HandoutMasterPart => parent.AddNewPart<HandoutMasterPart>(contentType, relationshipId),
                ImagePart => parent.AddNewPart<ImagePart>(contentType, relationshipId),
                LabelInfoPart => parent.AddNewPart<LabelInfoPart>(contentType, relationshipId),
                LegacyDiagramTextInfoPart => parent.AddNewPart<LegacyDiagramTextInfoPart>(contentType, relationshipId),
                LegacyDiagramTextPart => parent.AddNewPart<LegacyDiagramTextPart>(contentType, relationshipId),
                Model3DReferenceRelationshipPart => parent.AddNewPart<Model3DReferenceRelationshipPart>(contentType, relationshipId),
                NotesMasterPart => parent.AddNewPart<NotesMasterPart>(contentType, relationshipId),
                NotesSlidePart => parent.AddNewPart<NotesSlidePart>(contentType, relationshipId),
                PowerPointAuthorsPart => parent.AddNewPart<PowerPointAuthorsPart>(contentType, relationshipId),
                PowerPointCommentPart => parent.AddNewPart<PowerPointCommentPart>(contentType, relationshipId),
                PresentationPropertiesPart => parent.AddNewPart<PresentationPropertiesPart>(contentType, relationshipId),
                QuickAccessToolbarCustomizationsPart => parent.AddNewPart<QuickAccessToolbarCustomizationsPart>(contentType, relationshipId),
                RibbonAndBackstageCustomizationsPart => parent.AddNewPart<RibbonAndBackstageCustomizationsPart>(contentType, relationshipId),
                RibbonExtensibilityPart => parent.AddNewPart<RibbonExtensibilityPart>(contentType, relationshipId),
                SlideCommentsPart => parent.AddNewPart<SlideCommentsPart>(contentType, relationshipId),
                SlideLayoutPart => parent.AddNewPart<SlideLayoutPart>(contentType, relationshipId),
                SlideMasterPart => parent.AddNewPart<SlideMasterPart>(contentType, relationshipId),
                SlidePart => parent.AddNewPart<SlidePart>(contentType, relationshipId),
                SlideSyncDataPart => parent.AddNewPart<SlideSyncDataPart>(contentType, relationshipId),
                TableStylesPart => parent.AddNewPart<TableStylesPart>(contentType, relationshipId),
                ThemeOverridePart => parent.AddNewPart<ThemeOverridePart>(contentType, relationshipId),
                ThemePart => parent.AddNewPart<ThemePart>(contentType, relationshipId),
                ThumbnailPart => parent.AddNewPart<ThumbnailPart>(contentType, relationshipId),
                UserDefinedTagsPart => parent.AddNewPart<UserDefinedTagsPart>(contentType, relationshipId),
                VbaDataPart => parent.AddNewPart<VbaDataPart>(contentType, relationshipId),
                VbaProjectPart => parent.AddNewPart<VbaProjectPart>(contentType, relationshipId),
                ViewPropertiesPart => parent.AddNewPart<ViewPropertiesPart>(contentType, relationshipId),
                VmlDrawingPart => parent.AddNewPart<VmlDrawingPart>(contentType, relationshipId),
                WebExTaskpanesPart => parent.AddNewPart<WebExTaskpanesPart>(contentType, relationshipId),
                WebExtensionPart => parent.AddNewPart<WebExtensionPart>(contentType, relationshipId),
                XmlSignaturePart => parent.AddNewPart<XmlSignaturePart>(contentType, relationshipId),
                _ => throw new NotSupportedException($"Open XML part type '{sourcePart.GetType().FullName}' is not supported by the NativeAOT-safe PowerPoint clone path.")
            };
        }

        internal static string GetIndexedPartUri(
            OpenXmlPackage package,
            string folder,
            string baseName,
            string extension,
            bool allowBaseWithoutIndex) {
            if (package == null) {
                throw new ArgumentNullException(nameof(package));
            }
            if (string.IsNullOrWhiteSpace(folder)) {
                throw new ArgumentException("Folder is required.", nameof(folder));
            }
            if (string.IsNullOrWhiteSpace(baseName)) {
                throw new ArgumentException("Base name is required.", nameof(baseName));
            }
            if (string.IsNullOrWhiteSpace(extension)) {
                throw new ArgumentException("Extension is required.", nameof(extension));
            }

            string normalizedFolder = NormalizeFolder(folder);
            string normalizedExtension = extension.StartsWith(".", StringComparison.Ordinal)
                ? extension
                : "." + extension;

            int maxIndex = 0;
            bool hasBase = false;

            foreach (var pair in package.Parts) {
                string path = NormalizePartUri(pair.OpenXmlPart.Uri.OriginalString);
                if (!path.StartsWith(normalizedFolder + "/", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string fileName = Path.GetFileName(path);
                if (!fileName.StartsWith(baseName, StringComparison.OrdinalIgnoreCase) ||
                    !fileName.EndsWith(normalizedExtension, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string numberPart = fileName.Substring(
                    baseName.Length,
                    fileName.Length - baseName.Length - normalizedExtension.Length);

                if (string.IsNullOrEmpty(numberPart)) {
                    hasBase = true;
                    continue;
                }

                if (int.TryParse(numberPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx)) {
                    if (idx > maxIndex) {
                        maxIndex = idx;
                    }
                }
            }

            if (allowBaseWithoutIndex && !hasBase && maxIndex == 0) {
                return CombinePartUri(normalizedFolder, baseName + normalizedExtension);
            }

            int nextIndex = Math.Max(1, maxIndex + 1);
            return CombinePartUri(normalizedFolder, baseName + nextIndex.ToString(CultureInfo.InvariantCulture) + normalizedExtension);
        }

        internal static string GetImageExtension(ImagePartType type, string? sourcePath = null) {
            if (!string.IsNullOrWhiteSpace(sourcePath)) {
                string extension = Path.GetExtension(sourcePath);
                if (!string.IsNullOrWhiteSpace(extension)) {
                    return extension.ToLowerInvariant();
                }
            }

            OfficeImageFormat format = OfficeImageInfo.FromMimeType(type.ToPartTypeInfo().ContentType);
            return OfficeImageInfo.GetDefaultExtension(format == OfficeImageFormat.Unknown ? OfficeImageFormat.Png : format);
        }

        private static string NormalizePartUri(string partUri) {
            string normalized = partUri.Replace('\\', '/');
            if (!normalized.StartsWith("/", StringComparison.Ordinal)) {
                normalized = "/" + normalized;
            }
            return normalized;
        }

        private static string NormalizeFolder(string folder) {
            string normalized = folder.Replace('\\', '/').TrimEnd('/');
            if (!normalized.StartsWith("/", StringComparison.Ordinal)) {
                normalized = "/" + normalized;
            }
            return normalized;
        }

        private static string CombinePartUri(string folder, string fileName) {
            string normalizedFolder = NormalizeFolder(folder);
            return normalizedFolder + "/" + fileName;
        }
    }
}
