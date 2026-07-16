namespace OfficeIMO.Epub;

using System.Globalization;

internal static partial class EpubReader {
    private static EpubNavigationResult ReadNavigation(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubPackage? package,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        var result = new EpubNavigationResult();
        if (package == null) return result;

        var limits = new NavigationLimitState();
        if (!string.IsNullOrWhiteSpace(package.NavDocumentPath)) {
            TryReadNavigationDocument(entryIndex, package.NavDocumentPath!, result, options, limits, diagnostics);
        }

        bool needNcxToc = result.TableOfContents.Count == 0;
        bool needNcxPageList = result.PageList.Count == 0;
        bool needNcxTitleFallback = !needNcxToc;
        if (needNcxTitleFallback) {
            AddNavigationTitles(result.TableOfContents, result.TitleMap);
        }
        if (!string.IsNullOrWhiteSpace(package.NcxPath) && (needNcxToc || needNcxPageList || needNcxTitleFallback)) {
            TryReadNcxDocument(
                entryIndex,
                package.NcxPath!,
                result,
                includeTableOfContents: needNcxToc,
                includePageList: needNcxPageList,
                includeTitleFallback: needNcxTitleFallback,
                options,
                limits,
                diagnostics);
        }

        if (result.Landmarks.Count == 0) {
            foreach (EpubNavigationItem guideItem in package.Guide) {
                if (!TryReserveNavigationItem(1, package.OpfPath, options, limits, diagnostics)) break;
                result.Landmarks.Add(guideItem);
            }
        }

        ValidateNavigationTargets(entryIndex, result, diagnostics);
        AddNavigationTitles(result.TableOfContents, result.TitleMap);
        return result;
    }

    private static void TryReadNavigationDocument(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        string navPath,
        EpubNavigationResult result,
        EpubReadOptions options,
        NavigationLimitState limits,
        EpubDiagnosticCollector diagnostics) {
        if (!entryIndex.TryGetValue(navPath, out ZipArchiveEntry? navEntry)) {
            string code = Uri.TryCreate(navPath, UriKind.Absolute, out _)
                ? "epub.navigation.remote"
                : "epub.navigation.missing";
            diagnostics.Warning(
                code,
                Uri.TryCreate(navPath, UriKind.Absolute, out _)
                    ? $"Remote EPUB navigation document '{navPath}' is not fetched."
                    : $"EPUB navigation document '{navPath}' was not found in archive.",
                navPath);
            return;
        }

        if (navEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.navigation.metadata-size-limit",
                $"EPUB navigation document '{navPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                navPath);
            return;
        }

        string navContent = ReadEntryText(navEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(navContent, out XDocument? navDocument) || navDocument == null) {
            diagnostics.Warning(
                "epub.navigation.invalid-xml",
                $"EPUB navigation document '{navPath}' could not be parsed.",
                navPath);
            return;
        }

        foreach (XElement nav in navDocument.Descendants().Where(element => IsName(element, "nav"))) {
            string type = GetAttribute(nav, "type");
            List<EpubNavigationItem>? destination = ContainsSpaceSeparatedToken(type, "toc")
                ? result.TableOfContents
                : ContainsSpaceSeparatedToken(type, "page-list")
                    ? result.PageList
                    : ContainsSpaceSeparatedToken(type, "landmarks")
                        ? result.Landmarks
                        : null;
            if (destination == null) continue;

            XElement? list = nav.Elements().FirstOrDefault(element => IsName(element, "ol"))
                ?? nav.Descendants().FirstOrDefault(element => IsName(element, "ol"));
            if (list == null) continue;
            destination.AddRange(ParseHtmlNavigationList(list, navPath, 1, options, limits, diagnostics));
        }
    }

    private static IReadOnlyList<EpubNavigationItem> ParseHtmlNavigationList(
        XElement list,
        string navPath,
        int depth,
        EpubReadOptions options,
        NavigationLimitState limits,
        EpubDiagnosticCollector diagnostics) {
        var items = new List<EpubNavigationItem>();
        foreach (XElement listItem in list.Elements().Where(element => IsName(element, "li"))) {
            if (!TryReserveNavigationItem(depth, navPath, options, limits, diagnostics)) break;

            XElement? anchor = listItem.Elements().FirstOrDefault(element => IsName(element, "a"));
            XElement? labelElement = anchor ?? listItem.Elements().FirstOrDefault(element => IsName(element, "span"));
            string label = labelElement == null
                ? NormalizeWhitespace(string.Concat(listItem.Nodes().OfType<XText>().Select(static text => text.Value)))
                : NormalizeWhitespace(labelElement.Value);
            string? href = anchor == null ? null : NullIfWhiteSpace(GetAttribute(anchor, "href"));
            string? target = null;
            string? fragment = null;
            bool isRemote = false;
            if (href != null && !TryResolveNavigationTarget(navPath, href, out target, out fragment, out isRemote)) {
                diagnostics.Warning(
                    "epub.navigation.target-invalid",
                    $"Navigation item '{label}' has invalid href '{href}'.",
                    navPath);
            }

            XElement? childList = listItem.Elements().FirstOrDefault(element => IsName(element, "ol"));
            IReadOnlyList<EpubNavigationItem> children = childList == null
                ? Array.Empty<EpubNavigationItem>()
                : ParseHtmlNavigationList(childList, navPath, depth + 1, options, limits, diagnostics);
            if (label.Length == 0) label = target ?? string.Empty;
            items.Add(new EpubNavigationItem {
                Source = EpubNavigationSource.Epub3Navigation,
                Label = label,
                Href = href,
                Target = target,
                Fragment = fragment,
                SemanticType = NullIfWhiteSpace(GetAttribute(anchor ?? listItem, "type")),
                IsRemote = isRemote,
                Children = children
            });
        }
        return items;
    }

    private static void TryReadNcxDocument(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        string ncxPath,
        EpubNavigationResult result,
        bool includeTableOfContents,
        bool includePageList,
        bool includeTitleFallback,
        EpubReadOptions options,
        NavigationLimitState limits,
        EpubDiagnosticCollector diagnostics) {
        if (!entryIndex.TryGetValue(ncxPath, out ZipArchiveEntry? ncxEntry)) {
            string code = Uri.TryCreate(ncxPath, UriKind.Absolute, out _) ? "epub.ncx.remote" : "epub.ncx.missing";
            diagnostics.Warning(
                code,
                Uri.TryCreate(ncxPath, UriKind.Absolute, out _)
                    ? $"Remote EPUB NCX document '{ncxPath}' is not fetched."
                    : $"EPUB NCX document '{ncxPath}' was not found in archive.",
                ncxPath);
            return;
        }

        if (ncxEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.ncx.metadata-size-limit",
                $"EPUB NCX document '{ncxPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                ncxPath);
            return;
        }

        string ncxContent = ReadEntryText(ncxEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(ncxContent, out XDocument? ncxDocument) || ncxDocument == null) {
            diagnostics.Warning(
                "epub.ncx.invalid-xml",
                $"EPUB NCX document '{ncxPath}' could not be parsed.",
                ncxPath);
            return;
        }

        if (includeTableOfContents || includeTitleFallback) {
            XElement? navMap = ncxDocument.Descendants().FirstOrDefault(element => IsName(element, "navMap"));
            if (navMap != null) {
                NavigationLimitState ncxLimits = includeTableOfContents ? limits : new NavigationLimitState();
                IReadOnlyList<EpubNavigationItem> ncxItems = ParseNcxItems(
                    navMap.Elements().Where(element => IsName(element, "navPoint")),
                    ncxPath,
                    1,
                    options,
                    ncxLimits,
                    diagnostics);
                if (includeTableOfContents) {
                    result.TableOfContents.AddRange(ncxItems);
                } else {
                    AddNavigationTitles(ncxItems, result.TitleMap);
                }
            }
        }

        if (includePageList) {
            XElement? pageList = ncxDocument.Descendants().FirstOrDefault(element => IsName(element, "pageList"));
            if (pageList != null) {
                result.PageList.AddRange(ParseNcxItems(
                    pageList.Elements().Where(element => IsName(element, "pageTarget")),
                    ncxPath,
                    1,
                    options,
                    limits,
                    diagnostics));
            }
        }
    }

    private static IReadOnlyList<EpubNavigationItem> ParseNcxItems(
        IEnumerable<XElement> source,
        string ncxPath,
        int depth,
        EpubReadOptions options,
        NavigationLimitState limits,
        EpubDiagnosticCollector diagnostics) {
        var items = new List<EpubNavigationItem>();
        foreach (XElement element in source) {
            if (!TryReserveNavigationItem(depth, ncxPath, options, limits, diagnostics)) break;

            XElement? textElement = element.Elements()
                .FirstOrDefault(child => IsName(child, "navLabel"))?
                .Descendants()
                .FirstOrDefault(child => IsName(child, "text"));
            string label = textElement == null ? string.Empty : NormalizeWhitespace(textElement.Value);
            XElement? content = element.Elements().FirstOrDefault(child => IsName(child, "content"));
            string? href = content == null ? null : NullIfWhiteSpace(GetAttribute(content, "src"));
            string? target = null;
            string? fragment = null;
            bool isRemote = false;
            if (href != null && !TryResolveNavigationTarget(ncxPath, href, out target, out fragment, out isRemote)) {
                diagnostics.Warning(
                    "epub.ncx.target-invalid",
                    $"NCX item '{label}' has invalid src '{href}'.",
                    ncxPath);
            }

            int? playOrder = null;
            if (int.TryParse(GetAttribute(element, "playOrder"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedOrder)) {
                playOrder = parsedOrder;
            }
            IReadOnlyList<EpubNavigationItem> children = ParseNcxItems(
                element.Elements().Where(child => IsName(child, "navPoint")),
                ncxPath,
                depth + 1,
                options,
                limits,
                diagnostics);
            if (label.Length == 0) label = NullIfWhiteSpace(GetAttribute(element, "value")) ?? target ?? string.Empty;
            items.Add(new EpubNavigationItem {
                Source = EpubNavigationSource.Ncx,
                Label = label,
                Href = href,
                Target = target,
                Fragment = fragment,
                SemanticType = NullIfWhiteSpace(GetAttribute(element, "type")),
                PlayOrder = playOrder,
                IsRemote = isRemote,
                Children = children
            });
        }
        return items;
    }

    private static bool TryResolveNavigationTarget(
        string basePath,
        string href,
        out string? target,
        out string? fragment,
        out bool isRemote) {
        target = null;
        fragment = null;
        isRemote = false;
        if (string.IsNullOrWhiteSpace(href)) return false;

        string candidate = href.Trim();
        int fragmentIndex = candidate.IndexOf('#');
        string reference = fragmentIndex < 0 ? candidate : candidate.Substring(0, fragmentIndex);
        if (fragmentIndex >= 0 && fragmentIndex + 1 < candidate.Length) {
            string encodedFragment = candidate.Substring(fragmentIndex + 1);
            try {
                fragment = Uri.UnescapeDataString(encodedFragment);
            } catch {
                fragment = encodedFragment;
            }
        }

        if (reference.Length == 0) {
            target = NormalizePath(basePath);
            return target.Length > 0;
        }
        if (TryGetRemoteResourceUri(reference, out string? remoteUri)) {
            target = remoteUri;
            isRemote = true;
            return true;
        }

        target = ResolveRelativePath(basePath, reference);
        return target.Length > 0;
    }

    private static bool TryReserveNavigationItem(
        int depth,
        string sourcePath,
        EpubReadOptions options,
        NavigationLimitState limits,
        EpubDiagnosticCollector diagnostics) {
        if (depth > options.MaxNavigationDepth) {
            if (!limits.DepthLimitReported) {
                limits.DepthLimitReported = true;
                diagnostics.Warning(
                    "epub.navigation.depth-limit",
                    $"EPUB navigation nesting was truncated at MaxNavigationDepth ({options.MaxNavigationDepth}).",
                    sourcePath);
            }
            return false;
        }
        if (limits.Count >= options.MaxNavigationItems) {
            if (!limits.CountLimitReported) {
                limits.CountLimitReported = true;
                diagnostics.Warning(
                    "epub.navigation.item-count-limit",
                    $"EPUB navigation was truncated at MaxNavigationItems ({options.MaxNavigationItems}).",
                    sourcePath);
            }
            return false;
        }
        limits.Count++;
        return true;
    }

    private static void AddNavigationTitles(
        IEnumerable<EpubNavigationItem> items,
        Dictionary<string, string> titleMap) {
        foreach (EpubNavigationItem item in items) {
            if (!item.IsRemote && !string.IsNullOrWhiteSpace(item.Target) && !string.IsNullOrWhiteSpace(item.Label) &&
                !titleMap.ContainsKey(item.Target!)) {
                titleMap[item.Target!] = item.Label;
            }
            AddNavigationTitles(item.Children, titleMap);
        }
    }

    private static void ValidateNavigationTargets(
        IReadOnlyDictionary<string, ZipArchiveEntry> entryIndex,
        EpubNavigationResult navigation,
        EpubDiagnosticCollector diagnostics) {
        var reported = new HashSet<string>(StringComparer.Ordinal);
        ValidateNavigationTargets(entryIndex, navigation.TableOfContents, reported, diagnostics);
        ValidateNavigationTargets(entryIndex, navigation.PageList, reported, diagnostics);
        ValidateNavigationTargets(entryIndex, navigation.Landmarks, reported, diagnostics);
    }

    private static void ValidateNavigationTargets(
        IReadOnlyDictionary<string, ZipArchiveEntry> entryIndex,
        IEnumerable<EpubNavigationItem> items,
        HashSet<string> reported,
        EpubDiagnosticCollector diagnostics) {
        foreach (EpubNavigationItem item in items) {
            if (!string.IsNullOrWhiteSpace(item.Target)) {
                string target = item.Target!;
                if (item.IsRemote && reported.Add("remote\n" + target)) {
                    diagnostics.Info(
                        "epub.navigation.remote-target",
                        $"EPUB navigation references remote target '{target}'. Remote content is not fetched.",
                        target);
                } else if (!item.IsRemote && !entryIndex.ContainsKey(target) && reported.Add("missing\n" + target)) {
                    diagnostics.Warning(
                        "epub.navigation.target-missing",
                        $"EPUB navigation references missing archive resource '{target}'.",
                        target);
                }
            }
            ValidateNavigationTargets(entryIndex, item.Children, reported, diagnostics);
        }
    }
}
