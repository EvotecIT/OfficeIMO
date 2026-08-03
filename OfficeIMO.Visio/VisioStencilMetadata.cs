using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    internal static class VisioStencilMetadata {
        private const string ListSeparator = ";";
        private const string EncodedListPrefix = "OfficeIMO.List.v1:";
        private const string ListEncodingVersion = "length-prefixed-v1";

        internal static void Apply(VisioShape shape, VisioStencilShape stencil, string? catalogName) {
            Clear(shape);
            bool encodeLists = RequiresEncodedLists(stencil.Keywords,
                stencil.Aliases, stencil.Tags);
            Set(shape, VisioSemanticUserCells.StencilId, stencil.Id);
            Set(shape, VisioSemanticUserCells.StencilName, stencil.Name);
            Set(shape, VisioSemanticUserCells.StencilCategory, stencil.Category);
            Set(shape, VisioSemanticUserCells.StencilCatalog, catalogName);
            Set(shape, VisioSemanticUserCells.StencilSourcePackagePath, NormalizePath(stencil.SourcePackagePath));
            Set(shape, VisioSemanticUserCells.StencilIsSupported,
                stencil.IsSupported ? "true" : "false");
            Set(shape, VisioSemanticUserCells.StencilSourceLicense,
                stencil.SourceLicense);
            Set(shape, VisioSemanticUserCells.StencilSourceAttribution,
                stencil.SourceAttribution);
            Set(shape, VisioSemanticUserCells.StencilListEncoding,
                encodeLists ? ListEncodingVersion : null);
            Set(shape, VisioSemanticUserCells.StencilKeywords,
                Join(stencil.Keywords, encodeLists));
            Set(shape, VisioSemanticUserCells.StencilAliases,
                Join(stencil.Aliases, encodeLists));
            Set(shape, VisioSemanticUserCells.StencilTags,
                Join(stencil.Tags, encodeLists));
            Set(shape, VisioSemanticUserCells.StencilIconNameU, stencil.IconNameU);
            Set(shape, VisioSemanticUserCells.StencilDefaultWidth, FormatDouble(stencil.DefaultWidth));
            Set(shape, VisioSemanticUserCells.StencilDefaultHeight, FormatDouble(stencil.DefaultHeight));
            Set(shape, VisioSemanticUserCells.StencilDefaultUnit, stencil.DefaultUnit?.ToString());
            Set(shape, VisioSemanticUserCells.StencilPreviewImageRelationshipId, stencil.PreviewImage?.RelationshipId);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageTarget, stencil.PreviewImage?.Target);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageContentType, stencil.PreviewImage?.ContentType);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageExtension, stencil.PreviewImage?.Extension);
            Set(shape, VisioSemanticUserCells.StencilPreviewImageByteLength, FormatLong(stencil.PreviewImage?.ByteLength));
            ApplyConnectionPoints(shape, stencil);
        }

        internal static void Apply(VisioMaster master, VisioStencilShape stencil, string? catalogName) {
            if (master.IsPackageBacked) {
                if (string.IsNullOrWhiteSpace(stencil.SourcePackagePath)) {
                    throw new InvalidOperationException(
                        $"Visio master '{master.NameU}' is package-backed and cannot be reused " +
                        "with source-less stencil metadata. Preserve the trusted source package " +
                        "path when loading or applying the stencil.");
                }
                EnsureSourcePackageMatches(master, stencil.SourcePackagePath!);
            }
            master.StencilId = stencil.Id;
            master.StencilName = stencil.Name;
            master.StencilCategory = stencil.Category;
            master.StencilCatalogName = string.IsNullOrWhiteSpace(catalogName) ? master.StencilCatalogName : catalogName;
            master.StencilSourcePackagePath = NormalizePath(stencil.SourcePackagePath) ?? master.StencilSourcePackagePath;
            master.StencilIsSupported = stencil.IsSupported;
            master.StencilSourceLicense = stencil.SourceLicense;
            master.StencilSourceAttribution = stencil.SourceAttribution;
            master.StencilKeywords = Normalize(stencil.Keywords);
            master.StencilAliases = Normalize(stencil.Aliases);
            master.StencilTags = Normalize(stencil.Tags);
            master.StencilIconNameU = stencil.IconNameU;
            master.StencilDefaultWidth = stencil.DefaultWidth;
            master.StencilDefaultHeight = stencil.DefaultHeight;
            master.StencilDefaultUnit = stencil.DefaultUnit;
            master.StencilPreviewImageRelationshipId = stencil.PreviewImage?.RelationshipId;
            master.StencilPreviewImageTarget = stencil.PreviewImage?.Target;
            master.StencilPreviewImageContentType = stencil.PreviewImage?.ContentType;
            master.StencilPreviewImageExtension = stencil.PreviewImage?.Extension;
            master.StencilPreviewImageByteLength = stencil.PreviewImage?.ByteLength;
        }

        internal static void EnsureSourcePackageMatches(VisioMaster master,
            string sourcePackagePath) {
            string? registeredPath = NormalizePath(
                master.StencilSourcePackagePath);
            string? requestedPath = NormalizePath(sourcePackagePath);
            if (SourcePackagePathsMatch(registeredPath, requestedPath)) {
                return;
            }

            throw new InvalidOperationException(
                $"Visio master '{master.NameU}' is already bound to source package " +
                $"'{registeredPath ?? "<unknown>"}' and cannot be reused for " +
                $"'{requestedPath ?? "<unknown>"}'. Import package masters with unique NameU values.");
        }

        internal static bool SourcePackagePathsMatch(string? firstPath,
            string? secondPath) {
            if (firstPath == null || secondPath == null) {
                return firstPath == secondPath;
            }
            if (string.Equals(firstPath, secondPath, StringComparison.Ordinal)) {
                return true;
            }
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return string.Equals(firstPath, secondPath,
                    StringComparison.OrdinalIgnoreCase);
            }
            if (!File.Exists(firstPath) || !File.Exists(secondPath)) {
                return false;
            }
            string firstCanonical = ResolveExistingPathCasing(firstPath);
            string secondCanonical = ResolveExistingPathCasing(secondPath);
            firstCanonical = ResolveSymbolicLinkPath(firstCanonical);
            secondCanonical = ResolveSymbolicLinkPath(secondCanonical);
            return string.Equals(firstCanonical, secondCanonical,
                StringComparison.Ordinal);
        }

        private static string ResolveSymbolicLinkPath(string path) {
            IntPtr resolved = IntPtr.Zero;
            try {
                resolved = RealPath(path, IntPtr.Zero);
                return resolved == IntPtr.Zero
                    ? path
                    : Marshal.PtrToStringAnsi(resolved) ?? path;
            } catch (DllNotFoundException) {
                return path;
            } catch (EntryPointNotFoundException) {
                return path;
            } finally {
                if (resolved != IntPtr.Zero) Free(resolved);
            }
        }

        [DllImport("libc", EntryPoint = "realpath", CharSet = CharSet.Ansi,
            SetLastError = true)]
        private static extern IntPtr RealPath(string path, IntPtr resolvedPath);

        [DllImport("libc", EntryPoint = "free")]
        private static extern void Free(IntPtr pointer);

        private static string ResolveExistingPathCasing(string path) {
            string fullPath = Path.GetFullPath(path);
            string? root = Path.GetPathRoot(fullPath);
            if (string.IsNullOrEmpty(root)) return fullPath;
            string current = root!;
            string[] segments = fullPath.Substring(root!.Length)
                .Split(new[] { Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar },
                    StringSplitOptions.RemoveEmptyEntries);
            try {
                foreach (string segment in segments) {
                    string? match = Directory.EnumerateFileSystemEntries(current)
                        .FirstOrDefault(entry => string.Equals(
                            Path.GetFileName(entry), segment,
                            StringComparison.Ordinal));
                    match ??= Directory.EnumerateFileSystemEntries(current)
                        .FirstOrDefault(entry => string.Equals(
                            Path.GetFileName(entry), segment,
                            StringComparison.OrdinalIgnoreCase));
                    current = match ?? Path.Combine(current, segment);
                }
                return Path.GetFullPath(current);
            } catch (IOException) {
                return fullPath;
            } catch (UnauthorizedAccessException) {
                return fullPath;
            }
        }

        internal static void Apply(VisioMaster master, IEnumerable<VisioUserCell> userCells) {
            Dictionary<string, string?> values = userCells
                .GroupBy(cell => cell.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First().Value, StringComparer.OrdinalIgnoreCase);

            master.StencilId = Get(values, VisioSemanticUserCells.StencilId) ?? master.StencilId;
            master.StencilName = Get(values, VisioSemanticUserCells.StencilName) ?? master.StencilName;
            master.StencilCategory = Get(values, VisioSemanticUserCells.StencilCategory) ?? master.StencilCategory;
            master.StencilCatalogName = Get(values, VisioSemanticUserCells.StencilCatalog) ?? master.StencilCatalogName;
            master.StencilSourcePackagePath = Get(values, VisioSemanticUserCells.StencilSourcePackagePath) ?? master.StencilSourcePackagePath;
            master.StencilIsSupported = GetBool(values,
                VisioSemanticUserCells.StencilIsSupported)
                ?? master.StencilIsSupported;
            master.StencilSourceLicense = Get(values,
                VisioSemanticUserCells.StencilSourceLicense)
                ?? master.StencilSourceLicense;
            master.StencilSourceAttribution = Get(values,
                VisioSemanticUserCells.StencilSourceAttribution)
                ?? master.StencilSourceAttribution;
            bool encodedLists = string.Equals(Get(values,
                VisioSemanticUserCells.StencilListEncoding),
                ListEncodingVersion, StringComparison.Ordinal);
            master.StencilKeywords = Coalesce(SplitList(Get(values, VisioSemanticUserCells.StencilKeywords), encodedLists), master.StencilKeywords);
            master.StencilAliases = Coalesce(SplitList(Get(values, VisioSemanticUserCells.StencilAliases), encodedLists), master.StencilAliases);
            master.StencilTags = Coalesce(SplitList(Get(values, VisioSemanticUserCells.StencilTags), encodedLists), master.StencilTags);
            master.StencilIconNameU = Get(values, VisioSemanticUserCells.StencilIconNameU) ?? master.StencilIconNameU;
            master.StencilDefaultWidth = GetDouble(values, VisioSemanticUserCells.StencilDefaultWidth) ?? master.StencilDefaultWidth;
            master.StencilDefaultHeight = GetDouble(values, VisioSemanticUserCells.StencilDefaultHeight) ?? master.StencilDefaultHeight;
            master.StencilDefaultUnit = GetUnit(values, VisioSemanticUserCells.StencilDefaultUnit) ?? master.StencilDefaultUnit;
            master.StencilPreviewImageRelationshipId = Get(values, VisioSemanticUserCells.StencilPreviewImageRelationshipId) ?? master.StencilPreviewImageRelationshipId;
            master.StencilPreviewImageTarget = Get(values, VisioSemanticUserCells.StencilPreviewImageTarget) ?? master.StencilPreviewImageTarget;
            master.StencilPreviewImageContentType = Get(values, VisioSemanticUserCells.StencilPreviewImageContentType) ?? master.StencilPreviewImageContentType;
            master.StencilPreviewImageExtension = Get(values, VisioSemanticUserCells.StencilPreviewImageExtension) ?? master.StencilPreviewImageExtension;
            master.StencilPreviewImageByteLength = GetLong(values, VisioSemanticUserCells.StencilPreviewImageByteLength) ?? master.StencilPreviewImageByteLength;
        }

        internal static IReadOnlyList<VisioUserCell> CreateMasterUserCells(VisioMaster master) {
            List<VisioUserCell> cells = new();
            bool encodeLists = RequiresEncodedLists(master.StencilKeywords,
                master.StencilAliases, master.StencilTags);
            Add(cells, VisioSemanticUserCells.StencilId, master.StencilId);
            Add(cells, VisioSemanticUserCells.StencilName, master.StencilName);
            Add(cells, VisioSemanticUserCells.StencilCategory, master.StencilCategory);
            Add(cells, VisioSemanticUserCells.StencilCatalog, master.StencilCatalogName);
            Add(cells, VisioSemanticUserCells.StencilSourcePackagePath, master.StencilSourcePackagePath);
            Add(cells, VisioSemanticUserCells.StencilIsSupported,
                FormatBool(master.StencilIsSupported));
            Add(cells, VisioSemanticUserCells.StencilSourceLicense,
                master.StencilSourceLicense);
            Add(cells, VisioSemanticUserCells.StencilSourceAttribution,
                master.StencilSourceAttribution);
            Add(cells, VisioSemanticUserCells.StencilListEncoding,
                encodeLists ? ListEncodingVersion : null);
            Add(cells, VisioSemanticUserCells.StencilKeywords,
                Join(master.StencilKeywords, encodeLists));
            Add(cells, VisioSemanticUserCells.StencilAliases,
                Join(master.StencilAliases, encodeLists));
            Add(cells, VisioSemanticUserCells.StencilTags,
                Join(master.StencilTags, encodeLists));
            Add(cells, VisioSemanticUserCells.StencilIconNameU, master.StencilIconNameU);
            Add(cells, VisioSemanticUserCells.StencilDefaultWidth, FormatDouble(master.StencilDefaultWidth));
            Add(cells, VisioSemanticUserCells.StencilDefaultHeight, FormatDouble(master.StencilDefaultHeight));
            Add(cells, VisioSemanticUserCells.StencilDefaultUnit, master.StencilDefaultUnit?.ToString());
            Add(cells, VisioSemanticUserCells.StencilPreviewImageRelationshipId, master.StencilPreviewImageRelationshipId);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageTarget, master.StencilPreviewImageTarget);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageContentType, master.StencilPreviewImageContentType);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageExtension, master.StencilPreviewImageExtension);
            Add(cells, VisioSemanticUserCells.StencilPreviewImageByteLength, FormatLong(master.StencilPreviewImageByteLength));
            return cells.AsReadOnly();
        }

        internal static bool HasStencilMetadata(VisioShape shape) {
            if (shape == null) {
                return false;
            }

            return shape.UserCells.Any(cell =>
                string.Equals(cell.Name, VisioSemanticUserCells.StencilId, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(cell.Name, VisioSemanticUserCells.StencilName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(cell.Name, VisioSemanticUserCells.StencilCatalog, StringComparison.OrdinalIgnoreCase));
        }

        internal static void Clear(VisioShape shape) {
            if (shape == null) {
                return;
            }

            for (int i = shape.UserCells.Count - 1; i >= 0; i--) {
                string name = shape.UserCells[i].Name;
                if (string.Equals(name, VisioSemanticUserCells.StencilId, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilCategory, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilCatalog, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilSourcePackagePath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilIsSupported, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilSourceLicense, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilSourceAttribution, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilListEncoding, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilKeywords, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilAliases, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilTags, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilIconNameU, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilDefaultWidth, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilDefaultHeight, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilDefaultUnit, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageRelationshipId, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageTarget, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageContentType, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageExtension, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, VisioSemanticUserCells.StencilPreviewImageByteLength, StringComparison.OrdinalIgnoreCase)) {
                    shape.UserCells.RemoveAt(i);
                }
            }
        }

        internal static string? GetUserCellValue(IEnumerable<VisioInspectionUserCellSnapshot> userCells, string name) {
            return userCells
                .FirstOrDefault(cell => string.Equals(cell.Name, name, StringComparison.OrdinalIgnoreCase))
                ?.Value;
        }

        internal static IReadOnlyList<string> GetUserCellList(IEnumerable<VisioInspectionUserCellSnapshot> userCells, string name) {
            bool encodedLists = string.Equals(GetUserCellValue(userCells,
                VisioSemanticUserCells.StencilListEncoding),
                ListEncodingVersion, StringComparison.Ordinal);
            return SplitList(GetUserCellValue(userCells, name), encodedLists);
        }

        internal static IReadOnlyList<string> Normalize(IEnumerable<string>? values) {
            return (values ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        internal static string Join(IEnumerable<string>? values) {
            return Join(values, encode: false);
        }

        private static string Join(IEnumerable<string>? values, bool encode) {
            IReadOnlyList<string> normalized = Normalize(values);
            if (normalized.Count == 0) return string.Empty;
            if (!encode) return string.Join(ListSeparator, normalized);
            var builder = new StringBuilder(EncodedListPrefix);
            foreach (string item in normalized) {
                builder.Append(item.Length.ToString(CultureInfo.InvariantCulture));
                builder.Append(':');
                builder.Append(item);
            }
            return builder.ToString();
        }

        internal static IReadOnlyList<string> Split(string? value) {
            return SplitList(value, encoded: false);
        }

        private static IReadOnlyList<string> SplitList(string? value,
            bool encoded) {
            if (string.IsNullOrWhiteSpace(value)) {
                return Array.Empty<string>();
            }

            if (encoded && value!.StartsWith(EncodedListPrefix,
                    StringComparison.Ordinal)
                && TrySplitEncodedList(value, out IReadOnlyList<string> decoded)) {
                return Normalize(decoded);
            }

            return Normalize(value!.Split(new[] { ListSeparator }, StringSplitOptions.RemoveEmptyEntries));
        }

        private static bool RequiresEncodedLists(
            params IEnumerable<string>?[] lists) => lists
            .Where(list => list != null)
            .SelectMany(list => list!)
            .Any(value => value?.IndexOf(';') >= 0);

        private static bool TrySplitEncodedList(string value,
            out IReadOnlyList<string> items) {
            var result = new List<string>();
            int position = EncodedListPrefix.Length;
            while (position < value.Length) {
                int separator = value.IndexOf(':', position);
                if (separator <= position
                    || !int.TryParse(value.Substring(position,
                            separator - position), NumberStyles.None,
                        CultureInfo.InvariantCulture, out int length)
                    || length < 0
                    || length > value.Length - separator - 1) {
                    items = Array.Empty<string>();
                    return false;
                }
                position = separator + 1;
                result.Add(value.Substring(position, length));
                position += length;
            }
            items = result.AsReadOnly();
            return true;
        }

        internal static string? NormalizePath(string? path) {
            if (string.IsNullOrWhiteSpace(path)) {
                return null;
            }

            try {
                return Path.GetFullPath(path!);
            } catch (Exception) {
                return path!.Trim();
            }
        }

        private static void Set(VisioShape shape, string name, string? value) {
            if (!string.IsNullOrWhiteSpace(value)) {
                shape.SetUserCell(name, value, "STR", prompt: "OfficeIMO stencil metadata");
            }
        }

        private static void Add(ICollection<VisioUserCell> cells, string name, string? value) {
            if (!string.IsNullOrWhiteSpace(value)) {
                cells.Add(new VisioUserCell(name, value) {
                    Unit = "STR",
                    Prompt = "OfficeIMO stencil metadata"
                });
            }
        }

        private static string? Get(IReadOnlyDictionary<string, string?> values, string key) {
            return values.TryGetValue(key, out string? value) && !string.IsNullOrWhiteSpace(value)
                ? value
                : null;
        }

        private static double? GetDouble(IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : null;
        }

        private static VisioMeasurementUnit? GetUnit(IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return Enum.TryParse(value, ignoreCase: true, out VisioMeasurementUnit unit)
                ? unit
                : null;
        }

        private static long? GetLong(IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsed)
                ? parsed
                : null;
        }

        private static bool? GetBool(
            IReadOnlyDictionary<string, string?> values, string key) {
            string? value = Get(values, key);
            return bool.TryParse(value, out bool parsed) ? parsed : null;
        }

        private static string? FormatDouble(double? value) {
            return value.HasValue
                ? value.Value.ToString("0.######", CultureInfo.InvariantCulture)
                : null;
        }

        private static string? FormatLong(long? value) {
            return value.HasValue
                ? value.Value.ToString(CultureInfo.InvariantCulture)
                : null;
        }

        private static string? FormatBool(bool? value) => value.HasValue
            ? (value.Value ? "true" : "false")
            : null;

        private static void ApplyConnectionPoints(VisioShape shape, VisioStencilShape stencil) {
            if (shape.ConnectionPoints.Count > 0 ||
                stencil.SourceConnectionPoints.Count == 0) {
                return;
            }

            foreach (VisioStencilConnectionPoint point in stencil.SourceConnectionPoints) {
                double baseWidth = point.SourceWidth ?? GetDefaultSizeInInches(stencil.DefaultWidth, stencil.DefaultUnit);
                double baseHeight = point.SourceHeight ?? GetDefaultSizeInInches(stencil.DefaultHeight, stencil.DefaultUnit);
                double scaleX = baseWidth > 0 ? shape.Width / baseWidth : 1D;
                double scaleY = baseHeight > 0 ? shape.Height / baseHeight : 1D;
                shape.ConnectionPoints.Add(new VisioConnectionPoint(point.X * scaleX, point.Y * scaleY, point.DirX, point.DirY) {
                    SectionIndex = point.SectionIndex
                });
            }
        }

        private static double GetDefaultSizeInInches(double value, VisioMeasurementUnit? unit) {
            return unit.HasValue
                ? value.ToInches(unit.Value)
                : value;
        }

        private static IReadOnlyList<string> Coalesce(IReadOnlyList<string> candidate, IReadOnlyList<string> fallback) {
            return candidate.Count > 0 ? candidate : fallback;
        }
    }
}
