using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointPartFactory {
        private static readonly MethodInfo CreateInternalMethod = ResolveCreateInternal();
        private static readonly bool CreateInternalUsesUri =
            CreateInternalMethod.GetParameters()[3].ParameterType == typeof(Uri);

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

        internal static TPart CreatePart<TPart>(
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

            return type switch {
                ImagePartType.Jpeg => ".jpeg",
                ImagePartType.Gif => ".gif",
                ImagePartType.Bmp => ".bmp",
                _ => ".png"
            };
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
