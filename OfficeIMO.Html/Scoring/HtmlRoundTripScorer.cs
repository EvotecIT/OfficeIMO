using System.Security.Cryptography;
using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>
/// Scores structural HTML round-trip fidelity for gallery manifests and regression tests.
/// </summary>
public static class HtmlRoundTripScorer {
    private static readonly char[] WhitespaceSeparators = { ' ', '\t', '\r', '\n', '\f' };
    private static readonly string[] FormControlStateAttributes = {
        "type",
        "name",
        "value",
        "checked",
        "selected",
        "disabled",
        "multiple",
        "placeholder",
        "form",
        "formaction",
        "formmethod",
        "formenctype",
        "formtarget",
        "formnovalidate",
        "data-fieldset-disabled",
        "src",
        "data-src",
        "alt",
        "required",
        "readonly",
        "min",
        "max",
        "minlength",
        "maxlength",
        "pattern",
        "step",
        "autocomplete",
        "inputmode"
    };
    private static readonly string[] FormStateAttributes = {
        "id",
        "action",
        "method",
        "enctype",
        "target",
        "novalidate",
        "accept-charset",
        "type",
        "name",
        "value",
        "checked",
        "selected",
        "disabled",
        "multiple",
        "placeholder",
        "form",
        "formaction",
        "formmethod",
        "formenctype",
        "formtarget",
        "formnovalidate",
        "data-fieldset-disabled",
        "src",
        "data-src",
        "alt",
        "required",
        "readonly",
        "min",
        "max",
        "minlength",
        "maxlength",
        "pattern",
        "step",
        "autocomplete",
        "inputmode"
    };
    private static readonly HashSet<string> SubmitterOverrideAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "formaction",
        "formmethod",
        "formenctype",
        "formtarget",
        "formnovalidate"
    };
    private static readonly HashSet<string> ImageSubmitterAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "src",
        "data-src",
        "alt"
    };
    private static readonly HashSet<string> BooleanSignatureAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "checked",
        "selected",
        "disabled",
        "multiple",
        "formnovalidate",
        "data-fieldset-disabled",
        "novalidate",
        "required",
        "readonly",
        "controls",
        "autoplay",
        "loop",
        "muted",
        "default"
    };

    /// <summary>
    /// Compares source HTML with target HTML and returns a structural score.
    /// </summary>
    public static HtmlRoundTripScore Compare(string sourceHtml, string targetHtml) {
        return Compare(HtmlConversionDocument.Parse(sourceHtml), HtmlConversionDocument.Parse(targetHtml));
    }

    /// <summary>
    /// Compares retained conversion documents without reparsing either HTML source.
    /// </summary>
    public static HtmlRoundTripScore Compare(HtmlConversionDocument sourceDocument, HtmlConversionDocument targetDocument) {
        if (sourceDocument == null) throw new ArgumentNullException(nameof(sourceDocument));
        if (targetDocument == null) throw new ArgumentNullException(nameof(targetDocument));

        AngleSharp.Html.Dom.IHtmlDocument sourceScoringDocument = PrepareDocumentForScoring(sourceDocument);
        AngleSharp.Html.Dom.IHtmlDocument targetScoringDocument = PrepareDocumentForScoring(targetDocument);
        HtmlLogicalDocument source = HtmlLogicalDocumentBuilder.FromDocument(sourceScoringDocument);
        HtmlLogicalDocument target = HtmlLogicalDocumentBuilder.FromDocument(targetScoringDocument);
        IReadOnlyList<string> sourceFormOwners = ExtractFormOwnerSignatures(sourceScoringDocument);
        IReadOnlyList<string> targetFormOwners = ExtractFormOwnerSignatures(targetScoringDocument);
        double? formOwnerSimilarity = sourceFormOwners.Count == 0 && targetFormOwners.Count == 0
            ? (double?)null
            : SignatureSimilarity(targetFormOwners, sourceFormOwners);
        return Compare(
            source,
            target,
            TextSimilarityFromText(
                ExtractVisibleTextFromHtml(sourceDocument.CreateSourceDocumentForConversion()),
                ExtractVisibleTextFromHtml(targetDocument.CreateSourceDocumentForConversion())),
            formOwnerSimilarity);
    }

    /// <summary>
    /// Compares logical documents and returns a structural score.
    /// </summary>
    public static HtmlRoundTripScore Compare(HtmlLogicalDocument source, HtmlLogicalDocument target) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }

        return Compare(source, target, TextSimilarityFromText(ExtractLogicalText(source), ExtractLogicalText(target)));
    }

    private static HtmlRoundTripScore Compare(HtmlLogicalDocument source, HtmlLogicalDocument target, double textSimilarity, double? formOwnerSimilarity = null) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }

        var metrics = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        AddMetric(metrics, "nodes", Ratio(SumCounts(target), SumCounts(source)));
        AddCountMetric(metrics, "headings", target.Count(HtmlLogicalNodeKind.Heading), source.Count(HtmlLogicalNodeKind.Heading));
        AddSignatureMetric(metrics, "heading-levels", ExtractSignatures(target, HtmlLogicalNodeKind.Heading, CreateTextualNodeSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Heading, CreateTextualNodeSignature));
        AddCountMetric(metrics, "paragraphs", target.Count(HtmlLogicalNodeKind.Paragraph), source.Count(HtmlLogicalNodeKind.Paragraph));
        AddCountMetric(metrics, "tables", target.Count(HtmlLogicalNodeKind.Table), source.Count(HtmlLogicalNodeKind.Table));
        AddCountMetric(metrics, "table-rows", target.Count(HtmlLogicalNodeKind.TableRow), source.Count(HtmlLogicalNodeKind.TableRow));
        AddCountMetric(metrics, "table-cells", target.Count(HtmlLogicalNodeKind.TableCell), source.Count(HtmlLogicalNodeKind.TableCell));
        AddSignatureMetric(metrics, "table-grid", ExtractTableGridSignatures(target), ExtractTableGridSignatures(source));
        AddSignatureMetric(metrics, "table-captions", ExtractSignatures(target, HtmlLogicalNodeKind.TableCaption, CreateTextualNodeSignature), ExtractSignatures(source, HtmlLogicalNodeKind.TableCaption, CreateTextualNodeSignature));
        AddCountMetric(metrics, "figures", target.Count(HtmlLogicalNodeKind.Figure), source.Count(HtmlLogicalNodeKind.Figure));
        AddSignatureMetric(metrics, "figure-signatures", ExtractSignatures(target, HtmlLogicalNodeKind.Figure, CreateFigureSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Figure, CreateFigureSignature));
        AddCountMetric(metrics, "images", target.Count(HtmlLogicalNodeKind.Image), source.Count(HtmlLogicalNodeKind.Image));
        AddSignatureMetric(metrics, "image-sources", ExtractSignatures(target, HtmlLogicalNodeKind.Image, CreateImageSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Image, CreateImageSignature));
        AddCountMetric(metrics, "media", target.Count(HtmlLogicalNodeKind.Media), source.Count(HtmlLogicalNodeKind.Media));
        AddSignatureMetric(metrics, "media-sources", ExtractSignatures(target, HtmlLogicalNodeKind.Media, CreateMediaSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Media, CreateMediaSignature));
        AddCountMetric(metrics, "lists", target.Count(HtmlLogicalNodeKind.List), source.Count(HtmlLogicalNodeKind.List));
        AddSignatureMetric(metrics, "list-kinds", ExtractSignatures(target, HtmlLogicalNodeKind.List, CreateElementNameSignature), ExtractSignatures(source, HtmlLogicalNodeKind.List, CreateElementNameSignature));
        AddCountMetric(metrics, "list-items", target.Count(HtmlLogicalNodeKind.ListItem), source.Count(HtmlLogicalNodeKind.ListItem));
        AddCountMetric(metrics, "forms", target.Count(HtmlLogicalNodeKind.FormControl) + target.Count(HtmlLogicalNodeKind.Form), source.Count(HtmlLogicalNodeKind.FormControl) + source.Count(HtmlLogicalNodeKind.Form));
        AddSignatureMetric(metrics, "form-state", ExtractFormSignatures(target), ExtractFormSignatures(source));
        if (formOwnerSimilarity.HasValue) {
            double existingFormState;
            metrics["form-state"] = metrics.TryGetValue("form-state", out existingFormState)
                ? Math.Min(existingFormState, formOwnerSimilarity.Value)
                : formOwnerSimilarity.Value;
        }

        AddCountMetric(metrics, "links", target.Count(HtmlLogicalNodeKind.Link), source.Count(HtmlLogicalNodeKind.Link));
        AddSignatureMetric(metrics, "link-targets", ExtractSignatures(target, HtmlLogicalNodeKind.Link, CreateLinkSignature), ExtractSignatures(source, HtmlLogicalNodeKind.Link, CreateLinkSignature));
        AddMetric(metrics, "text", textSimilarity);

        int compared = metrics.Count;
        int matched = metrics.Values.Count(value => value >= 0.95D);
        double score = compared == 0 ? 1D : metrics.Values.Average();
        return new HtmlRoundTripScore(score, SumCounts(source), SumCounts(target), matched, compared, metrics);
    }

    private static void AddMetric(IDictionary<string, double> metrics, string name, double value) {
        metrics[name] = Math.Max(0D, Math.Min(1D, value));
    }

    private static void AddCountMetric(IDictionary<string, double> metrics, string name, int actual, int expected) {
        if (actual == 0 && expected == 0) {
            return;
        }

        AddMetric(metrics, name, Ratio(actual, expected));
    }

    private static void AddSignatureMetric(IDictionary<string, double> metrics, string name, IReadOnlyList<string> actual, IReadOnlyList<string> expected) {
        if (actual.Count == 0 && expected.Count == 0) {
            return;
        }

        AddMetric(metrics, name, SignatureSimilarity(actual, expected));
    }

    private static int SumCounts(HtmlLogicalDocument document) {
        return document.GetCounts().Values.Sum();
    }

    private static double Ratio(int actual, int expected) {
        if (expected == 0) {
            return actual == 0 ? 1D : 0D;
        }

        return Math.Min(actual, expected) / (double)Math.Max(actual, expected);
    }

    private static double SignatureSimilarity(IReadOnlyList<string> actual, IReadOnlyList<string> expected) {
        if (expected.Count == 0) {
            return actual.Count == 0 ? 1D : 0D;
        }

        var remaining = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string signature in expected) {
            if (!remaining.ContainsKey(signature)) {
                remaining[signature] = 0;
            }

            remaining[signature]++;
        }

        int matched = 0;
        foreach (string signature in actual) {
            int count;
            if (remaining.TryGetValue(signature, out count) && count > 0) {
                remaining[signature] = count - 1;
                matched++;
            }
        }

        return matched / (double)Math.Max(actual.Count, expected.Count);
    }

    private static IReadOnlyList<string> ExtractFormSignatures(HtmlLogicalDocument document) {
        var signatures = new List<string>();
        AppendFormSignatures(document.Root, signatures);
        return signatures;
    }

    private static AngleSharp.Html.Dom.IHtmlDocument PrepareDocumentForScoring(HtmlConversionDocument source) {
        AngleSharp.Html.Dom.IHtmlDocument document = source.CreateSourceDocumentForConversion();
        ResolveResourceSourceAttributes(document);
        SynthesizeImplicitSelectedOptions(document);
        PropagateFieldsetDisabledState(document);
        PruneHiddenStructure(document);
        return document;
    }

    private static void PruneHiddenStructure(AngleSharp.Html.Dom.IHtmlDocument document) {
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);
        foreach (IElement element in document.QuerySelectorAll("*").Reverse().ToList()) {
            if (element.ParentElement == null || !IsPrunableHiddenElement(element, styles)) {
                continue;
            }

            element.ParentElement.RemoveChild(element);
        }
    }

    private static void ResolveResourceSourceAttributes(AngleSharp.Html.Dom.IHtmlDocument document) {
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, null);
        if (baseUri == null) {
            return;
        }

        var policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
        foreach (var element in document.QuerySelectorAll("a,area,form,input,button,img,image,source,video,audio,track")) {
            ResolveUrlAttribute(element, "href", baseUri, policy);
            ResolveUrlAttribute(element, "action", baseUri, policy);
            ResolveUrlAttribute(element, "formaction", baseUri, policy);
            ResolveUrlAttribute(element, "src", baseUri, policy);
            ResolveUrlAttribute(element, "data-src", baseUri, policy);
            ResolveUrlAttribute(element, "data-original", baseUri, policy);
            ResolveUrlAttribute(element, "data-original-src", baseUri, policy);
            ResolveUrlAttribute(element, "data-lazy-src", baseUri, policy);
            ResolveUrlAttribute(element, "poster", baseUri, policy);
            ResolveUrlAttribute(element, "data-poster", baseUri, policy);
            ResolveUrlAttribute(element, "xlink:href", baseUri, policy);
            ResolveSrcSetAttribute(element, "srcset", baseUri, policy);
            ResolveSrcSetAttribute(element, "data-srcset", baseUri, policy);
            ResolveSrcSetAttribute(element, "data-original-srcset", baseUri, policy);
            ResolveSrcSetAttribute(element, "data-lazy-srcset", baseUri, policy);
        }
    }

    private static void ResolveUrlAttribute(AngleSharp.Dom.IElement element, string attributeName, Uri baseUri, HtmlUrlPolicy policy) {
        string? raw = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return;
        }

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(raw, baseUri, policy);
        if (!string.IsNullOrWhiteSpace(resolved)) {
            element.SetAttribute(attributeName, resolved);
        }
    }

    private static void ResolveSrcSetAttribute(AngleSharp.Dom.IElement element, string attributeName, Uri baseUri, HtmlUrlPolicy policy) {
        string? raw = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return;
        }

        var parts = new List<string>();
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(raw)) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, policy);
            if (string.IsNullOrWhiteSpace(resolved)) {
                resolved = candidate.Url;
            }

            parts.Add(string.IsNullOrWhiteSpace(candidate.Descriptor)
                ? resolved
                : resolved + " " + candidate.Descriptor);
        }

        if (parts.Count > 0) {
            element.SetAttribute(attributeName, string.Join(", ", parts));
        }
    }

    private static IReadOnlyList<string> ExtractFormOwnerSignatures(AngleSharp.Html.Dom.IHtmlDocument document) {
        var signatures = new List<string>();
        foreach (var control in document.QuerySelectorAll("input,select,textarea,button")) {
            var parts = new List<string> {
                control.TagName.ToLowerInvariant()
            };

            foreach (string attributeName in FormControlStateAttributes) {
                if (!ShouldIncludeFormControlAttribute(control, attributeName)) {
                    continue;
                }

                if (string.Equals(attributeName, "type", StringComparison.OrdinalIgnoreCase)) {
                    string defaultType = GetEffectiveFormControlType(control.TagName.ToLowerInvariant(), control.GetAttribute(attributeName));
                    if (!string.IsNullOrWhiteSpace(defaultType) && (!control.HasAttribute(attributeName) || !IsValidFormControlType(control.TagName, control.GetAttribute(attributeName)))) {
                        parts.Add("type=" + defaultType);
                        continue;
                    }
                }

                if (string.Equals(attributeName, "value", StringComparison.OrdinalIgnoreCase) && !control.HasAttribute(attributeName)) {
                    string defaultValue = GetDefaultFormControlValue(control.TagName, control.GetAttribute("type"), control.TextContent);
                    if (!string.IsNullOrWhiteSpace(defaultValue)) {
                        parts.Add("value=" + defaultValue);
                        continue;
                    }
                }

                if (control.HasAttribute(attributeName)) {
                    parts.Add(FormatAttributePart(attributeName, control.GetAttribute(attributeName)));
                }
            }

            string owner = ResolveFormOwnerSignature(document, control);
            if (!string.IsNullOrWhiteSpace(owner)) {
                parts.Add("owner=" + owner);
            }

            signatures.Add(string.Join("|", parts));
        }

        return signatures;
    }

    private static void SynthesizeImplicitSelectedOptions(AngleSharp.Html.Dom.IHtmlDocument document) {
        foreach (var select in document.QuerySelectorAll("select:not([multiple])")) {
            if (select.QuerySelector("option[selected]") != null) {
                continue;
            }

            AngleSharp.Dom.IElement? firstOption = select.QuerySelector("option");
            firstOption?.SetAttribute("selected", string.Empty);
        }
    }

    private static void PropagateFieldsetDisabledState(AngleSharp.Html.Dom.IHtmlDocument document) {
        foreach (var fieldset in document.QuerySelectorAll("fieldset[disabled]")) {
            AngleSharp.Dom.IElement? firstLegend = fieldset.Children.FirstOrDefault(child => string.Equals(child.TagName, "legend", StringComparison.OrdinalIgnoreCase));
            foreach (var control in fieldset.QuerySelectorAll("input,select,textarea,button")) {
                if (firstLegend != null && IsDescendantOf(control, firstLegend)) {
                    continue;
                }

                control.SetAttribute("data-fieldset-disabled", "true");
            }
        }
    }

    private static bool IsDescendantOf(AngleSharp.Dom.IElement element, AngleSharp.Dom.IElement ancestor) {
        AngleSharp.Dom.IElement? current = element;
        while (current != null) {
            if (ReferenceEquals(current, ancestor)) {
                return true;
            }

            current = current.ParentElement;
        }

        return false;
    }

    private static string ResolveFormOwnerSignature(AngleSharp.Html.Dom.IHtmlDocument document, AngleSharp.Dom.IElement control) {
        string? explicitOwner = control.GetAttribute("form");
        if (!string.IsNullOrWhiteSpace(explicitOwner)) {
            string owner = explicitOwner!.Trim();
            return HasFormWithId(document, owner) ? owner : string.Empty;
        }

        AngleSharp.Dom.IElement? current = control.ParentElement;
        while (current != null) {
            if (string.Equals(current.TagName, "form", StringComparison.OrdinalIgnoreCase)) {
                string? id = current.GetAttribute("id");
                if (!string.IsNullOrWhiteSpace(id)) {
                    return id!.Trim();
                }

                string? action = current.GetAttribute("action");
                return string.IsNullOrWhiteSpace(action) ? "ancestor-form" : action!.Trim();
            }

            current = current.ParentElement;
        }

        return string.Empty;
    }

    private static IReadOnlyList<string> ExtractSignatures(HtmlLogicalDocument document, HtmlLogicalNodeKind kind, Func<HtmlLogicalNode, string> createSignature) {
        var signatures = new List<string>();
        AppendSignatures(document.Root, kind, createSignature, signatures);
        return signatures;
    }

    private static void AppendSignatures(HtmlLogicalNode node, HtmlLogicalNodeKind kind, Func<HtmlLogicalNode, string> createSignature, ICollection<string> signatures) {
        if (node.Kind == kind) {
            signatures.Add(createSignature(node));
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendSignatures(child, kind, createSignature, signatures);
        }
    }

    private static IReadOnlyList<string> ExtractTableGridSignatures(HtmlLogicalDocument document) {
        var signatures = new List<string>();
        AppendTableGridSignatures(document.Root, signatures);
        return signatures;
    }

    private static void AppendTableGridSignatures(HtmlLogicalNode node, ICollection<string> signatures) {
        if (node.Kind == HtmlLogicalNodeKind.Table) {
            var rowSignatures = new List<string>();
            foreach (HtmlLogicalNode row in DirectTableRows(node)) {
                var cellSignatures = new List<string>();
                foreach (HtmlLogicalNode cell in row.Children.Where(child => child.Kind == HtmlLogicalNodeKind.TableCell)) {
                    cellSignatures.Add(CreateTableCellGridSignature(cell));
                }

                rowSignatures.Add(string.Join("+", cellSignatures));
            }

            signatures.Add("table|" + string.Join(",", rowSignatures));
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendTableGridSignatures(child, signatures);
        }
    }

    private static IEnumerable<HtmlLogicalNode> DirectTableRows(HtmlLogicalNode table) {
        foreach (HtmlLogicalNode child in table.Children) {
            if (child.Kind == HtmlLogicalNodeKind.TableRow) {
                yield return child;
                continue;
            }

            if (IsTableRowGroup(child.Name)) {
                foreach (HtmlLogicalNode row in child.Children.Where(row => row.Kind == HtmlLogicalNodeKind.TableRow)) {
                    yield return row;
                }
            }
        }
    }

    private static bool IsTableRowGroup(string name) {
        return string.Equals(name, "thead", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "tbody", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "tfoot", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<HtmlLogicalNode> Descendants(HtmlLogicalNode node, HtmlLogicalNodeKind kind) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (child.Kind == kind) {
                yield return child;
            }

            foreach (HtmlLogicalNode descendant in Descendants(child, kind)) {
                yield return descendant;
            }
        }
    }

    private static void AppendFormSignatures(HtmlLogicalNode node, ICollection<string> signatures) {
        if (node.Kind == HtmlLogicalNodeKind.Form || node.Kind == HtmlLogicalNodeKind.FormControl) {
            signatures.Add(CreateFormSignature(node));
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendFormSignatures(child, signatures);
        }
    }

    private static string CreateFormSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };

        foreach (string attributeName in FormStateAttributes) {
            if (node.Kind == HtmlLogicalNodeKind.FormControl) {
                AddFormControlAttributePart(parts, node, attributeName);
            } else {
                AddFormAttributePart(parts, node, attributeName);
            }
        }

        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        if (node.Kind == HtmlLogicalNodeKind.Form) {
            foreach (HtmlLogicalNode control in Descendants(node, HtmlLogicalNodeKind.FormControl)) {
                parts.Add("control=" + CreateFormControlSignature(control));
            }
        }

        return string.Join("|", parts);
    }

    private static string CreateFormControlSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };

        foreach (string attributeName in FormControlStateAttributes) {
            AddFormControlAttributePart(parts, node, attributeName);
        }

        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        return string.Join("|", parts);
    }

    private static void AddFormAttributePart(ICollection<string> parts, HtmlLogicalNode node, string attributeName) {
        if (string.Equals(attributeName, "action", StringComparison.OrdinalIgnoreCase)) {
            if (node.Attributes.TryGetValue(attributeName, out string? value) && !string.IsNullOrWhiteSpace(value)) {
                parts.Add("action=" + value);
            }

            return;
        }

        if (string.Equals(attributeName, "method", StringComparison.OrdinalIgnoreCase)) {
            parts.Add("method=" + GetEffectiveFormMethod(node.Attributes.TryGetValue(attributeName, out string? value) ? value : null));
            return;
        }

        if (string.Equals(attributeName, "enctype", StringComparison.OrdinalIgnoreCase)) {
            parts.Add("enctype=" + GetEffectiveFormEncoding(node.Attributes.TryGetValue(attributeName, out string? value) ? value : null));
            return;
        }

        AddAttributePart(parts, node, attributeName);
    }

    private static void AddFormControlAttributePart(ICollection<string> parts, HtmlLogicalNode node, string attributeName) {
        if (string.Equals(attributeName, "form", StringComparison.OrdinalIgnoreCase) && !IsFormAssociatedControl(node.Name)) {
            return;
        }

        if (SubmitterOverrideAttributes.Contains(attributeName) && !IsSubmitterControl(node)) {
            return;
        }

        if (ImageSubmitterAttributes.Contains(attributeName) && !IsImageSubmitterControl(node)) {
            return;
        }

        if (string.Equals(attributeName, "type", StringComparison.OrdinalIgnoreCase)) {
            node.Attributes.TryGetValue(attributeName, out string? rawType);
            string defaultType = GetEffectiveFormControlType(node.Name, rawType);
            if (!string.IsNullOrWhiteSpace(defaultType) && (!node.Attributes.ContainsKey(attributeName) || !IsValidFormControlType(node.Name, rawType))) {
                parts.Add("type=" + defaultType);
                return;
            }
        }

        if (string.Equals(attributeName, "value", StringComparison.OrdinalIgnoreCase) && !node.Attributes.ContainsKey(attributeName)) {
            node.Attributes.TryGetValue("type", out string? rawType);
            string defaultValue = GetDefaultFormControlValue(node.Name, rawType, ExtractLogicalNodeText(node));
            if (!string.IsNullOrWhiteSpace(defaultValue)) {
                parts.Add("value=" + defaultValue);
                return;
            }
        }

        AddAttributePart(parts, node, attributeName);
    }

    private static bool ShouldIncludeFormControlAttribute(AngleSharp.Dom.IElement control, string attributeName) {
        if (string.Equals(attributeName, "form", StringComparison.OrdinalIgnoreCase) && !IsFormAssociatedControl(control.TagName)) {
            return false;
        }

        if (SubmitterOverrideAttributes.Contains(attributeName) && !IsSubmitterControl(control)) {
            return false;
        }

        if (ImageSubmitterAttributes.Contains(attributeName) && !IsImageSubmitterControl(control)) {
            return false;
        }

        return true;
    }

    private static bool IsFormAssociatedControl(string name) {
        return string.Equals(name, "input", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "select", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "textarea", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "button", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSubmitterControl(AngleSharp.Dom.IElement control) {
        string name = control.TagName.ToLowerInvariant();
        string type = GetEffectiveFormControlType(name, control.GetAttribute("type"));
        if (string.Equals(name, "button", StringComparison.OrdinalIgnoreCase)) {
            return !string.Equals(type, "button", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(type, "reset", StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(name, "input", StringComparison.OrdinalIgnoreCase)
            && (string.Equals(type, "submit", StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, "image", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsSubmitterControl(HtmlLogicalNode node) {
        string type = GetEffectiveFormControlType(node.Name, node.Attributes.TryGetValue("type", out string? value) ? value : null);
        if (string.Equals(node.Name, "button", StringComparison.OrdinalIgnoreCase)) {
            return !string.Equals(type, "button", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(type, "reset", StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(node.Name, "input", StringComparison.OrdinalIgnoreCase)
            && (string.Equals(type, "submit", StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, "image", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsImageSubmitterControl(AngleSharp.Dom.IElement control) {
        string name = control.TagName.ToLowerInvariant();
        string type = GetEffectiveFormControlType(name, control.GetAttribute("type"));
        return string.Equals(name, "input", StringComparison.OrdinalIgnoreCase)
            && string.Equals(type, "image", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsImageSubmitterControl(HtmlLogicalNode node) {
        string type = GetEffectiveFormControlType(node.Name, node.Attributes.TryGetValue("type", out string? value) ? value : null);
        return string.Equals(node.Name, "input", StringComparison.OrdinalIgnoreCase)
            && string.Equals(type, "image", StringComparison.OrdinalIgnoreCase);
    }

    private static string CreateElementNameSignature(HtmlLogicalNode node) {
        return node.Name;
    }

    private static string CreateTextualNodeSignature(HtmlLogicalNode node) {
        string text = ExtractLogicalNodeText(node);
        return string.IsNullOrWhiteSpace(text)
            ? node.Name
            : node.Name + "|text=" + NormalizeText(text);
    }

    private static string CreateLinkSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        AddAttributePart(parts, node, "href");
        AddAttributePart(parts, node, "shape");
        AddAttributePart(parts, node, "coords");
        AddAttributePart(parts, node, "alt");
        AddAttributePart(parts, node, "aria-label");
        AddAttributePart(parts, node, "title");
        AddAttributePart(parts, node, "target");
        AddAttributePart(parts, node, "rel");
        AddAttributePart(parts, node, "download");
        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        return string.Join("|", parts);
    }

    private static string CreateFigureSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        string text = ExtractLogicalNodeText(node);
        if (!string.IsNullOrWhiteSpace(text)) {
            parts.Add("text=" + NormalizeText(text));
        }

        foreach (HtmlLogicalNode image in Descendants(node, HtmlLogicalNodeKind.Image)) {
            parts.Add("image=" + CreateImageSignature(image));
        }

        return string.Join("|", parts);
    }

    private static string CreateImageSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        AddAttributePart(parts, node, "src");
        AddAttributePart(parts, node, "href");
        AddAttributePart(parts, node, "xlink:href");
        AddAttributePart(parts, node, "srcset");
        AddAttributePart(parts, node, "data-original-srcset");
        AddAttributePart(parts, node, "data-lazy-srcset");
        AddAttributePart(parts, node, "media");
        AddAttributePart(parts, node, "type");
        AddAttributePart(parts, node, "sizes");
        AddAttributePart(parts, node, "data-src");
        AddAttributePart(parts, node, "data-original");
        AddAttributePart(parts, node, "data-original-src");
        AddAttributePart(parts, node, "data-lazy-src");
        AddAttributePart(parts, node, "data-srcset");
        AddAttributePart(parts, node, "alt");
        return string.Join("|", parts);
    }

    private static string CreateMediaSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        AddAttributePart(parts, node, "src");
        AddAttributePart(parts, node, "data-src");
        if (!string.Equals(node.Name, "source", StringComparison.OrdinalIgnoreCase)) {
            AddAttributePart(parts, node, "srcset");
            AddAttributePart(parts, node, "data-srcset");
        }

        AddAttributePart(parts, node, "poster");
        AddAttributePart(parts, node, "data-poster");
        AddAttributePart(parts, node, "kind");
        AddAttributePart(parts, node, "srclang");
        AddAttributePart(parts, node, "type");
        AddAttributePart(parts, node, "controls");
        AddAttributePart(parts, node, "autoplay");
        AddAttributePart(parts, node, "loop");
        AddAttributePart(parts, node, "muted");
        AddAttributePart(parts, node, "preload");
        AddAttributePart(parts, node, "default");
        AddAttributePart(parts, node, "label");
        return string.Join("|", parts);
    }

    private static string CreateTableCellGridSignature(HtmlLogicalNode node) {
        var parts = new List<string> {
            node.Name
        };
        AddAttributePart(parts, node, "colspan");
        AddAttributePart(parts, node, "rowspan");
        return string.Join("|", parts);
    }

    private static void AddAttributePart(ICollection<string> parts, HtmlLogicalNode node, string attributeName) {
        string? value;
        if (node.Attributes.TryGetValue(attributeName, out value)) {
            parts.Add(FormatAttributePart(attributeName, value));
        }
    }

    private static string FormatAttributePart(string attributeName, string? value) {
        if (string.Equals(attributeName, "download", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(value)) {
            return attributeName + "=present";
        }

        if (BooleanSignatureAttributes.Contains(attributeName)) {
            return attributeName + "=present";
        }

        if (string.Equals(attributeName, "rel", StringComparison.OrdinalIgnoreCase)) {
            return attributeName + "=" + NormalizeTokenList(value);
        }

        return attributeName + "=" + (value ?? string.Empty);
    }

    private static string GetEffectiveFormControlType(string name, string? type) {
        string normalized = (type ?? string.Empty).Trim().ToLowerInvariant();
        if (string.Equals(name, "input", StringComparison.OrdinalIgnoreCase)) {
            return IsValidInputType(normalized) ? normalized : "text";
        }

        if (string.Equals(name, "button", StringComparison.OrdinalIgnoreCase)) {
            return IsValidButtonType(normalized) ? normalized : "submit";
        }

        return string.Empty;
    }

    private static bool HasFormWithId(AngleSharp.Html.Dom.IHtmlDocument document, string id) {
        foreach (AngleSharp.Dom.IElement form in document.QuerySelectorAll("form[id]")) {
            if (string.Equals(form.GetAttribute("id"), id, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsValidFormControlType(string name, string? type) {
        string normalized = (type ?? string.Empty).Trim().ToLowerInvariant();
        if (string.Equals(name, "input", StringComparison.OrdinalIgnoreCase)) {
            return IsValidInputType(normalized);
        }

        if (string.Equals(name, "button", StringComparison.OrdinalIgnoreCase)) {
            return IsValidButtonType(normalized);
        }

        return false;
    }

    private static bool IsValidInputType(string type) {
        switch (type) {
            case "button":
            case "checkbox":
            case "color":
            case "date":
            case "datetime-local":
            case "email":
            case "file":
            case "hidden":
            case "image":
            case "month":
            case "number":
            case "password":
            case "radio":
            case "range":
            case "reset":
            case "search":
            case "submit":
            case "tel":
            case "text":
            case "time":
            case "url":
            case "week":
                return true;
            default:
                return false;
        }
    }

    private static bool IsValidButtonType(string type) {
        return string.Equals(type, "submit", StringComparison.Ordinal)
            || string.Equals(type, "reset", StringComparison.Ordinal)
            || string.Equals(type, "button", StringComparison.Ordinal);
    }

    private static string GetEffectiveFormMethod(string? method) {
        string normalized = (method ?? string.Empty).Trim().ToLowerInvariant();
        switch (normalized) {
            case "get":
            case "post":
            case "dialog":
                return normalized;
            default:
                return "get";
        }
    }

    private static string GetEffectiveFormEncoding(string? enctype) {
        string normalized = (enctype ?? string.Empty).Trim().ToLowerInvariant();
        switch (normalized) {
            case "application/x-www-form-urlencoded":
            case "multipart/form-data":
            case "text/plain":
                return normalized;
            default:
                return "application/x-www-form-urlencoded";
        }
    }

    private static string GetDefaultFormControlValue(string name, string? type, string textContent) {
        if (string.Equals(name, "option", StringComparison.OrdinalIgnoreCase)) {
            return NormalizeText(textContent);
        }

        if (string.Equals(name, "input", StringComparison.OrdinalIgnoreCase)) {
            string effectiveType = GetEffectiveFormControlType(name, type);
            if (string.Equals(effectiveType, "checkbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(effectiveType, "radio", StringComparison.OrdinalIgnoreCase)) {
                return "on";
            }
        }

        return string.Empty;
    }

    private static string NormalizeTokenList(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        return string.Join(" ", value!
            .Split(WhitespaceSeparators, StringSplitOptions.RemoveEmptyEntries)
            .Select(token => token.Trim().ToLowerInvariant())
            .Distinct(StringComparer.Ordinal)
            .OrderBy(token => token, StringComparer.Ordinal));
    }

    private static string ExtractLogicalNodeText(HtmlLogicalNode node) {
        var parts = new List<string>();
        AppendLogicalText(node, parts);
        return string.Join(" ", parts);
    }

    private static double TextSimilarityFromText(string sourceText, string targetText) {
        sourceText = NormalizeText(sourceText);
        targetText = NormalizeText(targetText);
        if (sourceText.Length == 0 && targetText.Length == 0) {
            return 1D;
        }

        if (string.Equals(sourceText, targetText, StringComparison.Ordinal)) {
            return 1D;
        }

        Dictionary<string, int> sourceWindows = HashWindows(sourceText);
        Dictionary<string, int> targetWindows = HashWindows(targetText);
        int unionCount = CountWindowUnion(sourceWindows, targetWindows);
        if (unionCount == 0) {
            return 1D;
        }

        return CountWindowIntersection(sourceWindows, targetWindows) / (double)unionCount;
    }

    private static string ExtractVisibleTextFromHtml(AngleSharp.Html.Dom.IHtmlDocument document) {
        var parts = new List<string>();
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);
        INode? root = document.Body ?? (INode?)document.DocumentElement;
        if (root != null) {
            AppendVisibleText(root, parts, styles);
        }

        return string.Join(" ", parts);
    }

    private static void AppendVisibleText(INode node, ICollection<string> parts, IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        AppendVisibleText(node, parts, styles, true);
    }

    private static void AppendVisibleText(INode node, ICollection<string> parts, IReadOnlyDictionary<IElement, HtmlComputedStyle> styles, bool inheritedVisibility) {
        bool currentVisibility = inheritedVisibility;
        if (node is IElement element) {
            if (IsNonVisibleTextElement(element.TagName) || IsPrunableHiddenElement(element, styles)) {
                return;
            }

            HtmlComputedStyle? computedStyle;
            if (styles.TryGetValue(element, out computedStyle) && computedStyle != null) {
                string visibility = computedStyle.GetValue("visibility");
                if (string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase)) {
                    currentVisibility = false;
                } else if (string.Equals(visibility, "visible", StringComparison.OrdinalIgnoreCase)) {
                    currentVisibility = true;
                }
            }
        }

        if (currentVisibility && node.NodeType == NodeType.Text && !string.IsNullOrWhiteSpace(node.TextContent)) {
            parts.Add(node.TextContent);
            return;
        }

        foreach (INode child in node.ChildNodes) {
            AppendVisibleText(child, parts, styles, currentVisibility);
        }
    }

    private static string ExtractLogicalText(HtmlLogicalDocument document) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var parts = new List<string>();
        AppendLogicalText(document.Root, parts);
        return string.Join(" ", parts);
    }

    private static void AppendLogicalText(HtmlLogicalNode node, ICollection<string> parts) {
        if (IsNonVisibleTextElement(node.Name) || IsHiddenLogicalNode(node)) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(node.Text) && (node.Kind == HtmlLogicalNodeKind.Text || !HasTextChild(node))) {
            parts.Add(node.Text);
        }

        foreach (HtmlLogicalNode child in node.Children) {
            AppendLogicalText(child, parts);
        }
    }

    private static bool HasTextChild(HtmlLogicalNode node) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (!string.IsNullOrWhiteSpace(child.Text) || HasTextChild(child)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsNonVisibleTextElement(string name) {
        return string.Equals(name, "script", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "style", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "template", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPrunableHiddenElement(IElement element, IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        if (element.HasAttribute("hidden")) {
            return true;
        }

        string? ariaHidden = element.GetAttribute("aria-hidden");
        if (string.Equals(ariaHidden, "true", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        HtmlComputedStyle? computedStyle;
        if (styles.TryGetValue(element, out computedStyle) && computedStyle != null) {
            if (string.Equals(computedStyle.GetValue("display"), "none", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string visibility = computedStyle.GetValue("visibility");
            if ((string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase))
                && !HasVisibleVisibilityDescendant(element, styles)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleVisibilityDescendant(IElement element, IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        foreach (IElement descendant in element.QuerySelectorAll("*")) {
            HtmlComputedStyle? descendantStyle;
            if (!styles.TryGetValue(descendant, out descendantStyle) || descendantStyle == null) {
                continue;
            }

            string visibility = descendantStyle.GetValue("visibility");
            if (!string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsHiddenLogicalNode(HtmlLogicalNode node) {
        if (node.Attributes.ContainsKey("hidden")) {
            return true;
        }

        string? ariaHidden;
        if (node.Attributes.TryGetValue("aria-hidden", out ariaHidden) && string.Equals(ariaHidden, "true", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        string? style;
        return node.Attributes.TryGetValue("style", out style) && ContainsHiddenStyle(style);
    }

    private static bool ContainsHiddenStyle(string? style) {
        if (string.IsNullOrWhiteSpace(style)) {
            return false;
        }

        string styleText = style!;
        return styleText.IndexOf("display:none", StringComparison.OrdinalIgnoreCase) >= 0
            || styleText.IndexOf("display: none", StringComparison.OrdinalIgnoreCase) >= 0
            || styleText.IndexOf("visibility:hidden", StringComparison.OrdinalIgnoreCase) >= 0
            || styleText.IndexOf("visibility: hidden", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static Dictionary<string, int> HashWindows(string text) {
        var windows = new Dictionary<string, int>(StringComparer.Ordinal);
        if (text.Length <= 32) {
            AddWindow(windows, Hash(text));
            return windows;
        }

        for (int i = 0; i <= text.Length - 32; i += 16) {
            AddWindow(windows, Hash(text.Substring(i, 32)));
        }

        AddWindow(windows, Hash(text.Substring(text.Length - 32, 32)));
        return windows;
    }

    private static void AddWindow(IDictionary<string, int> windows, string hash) {
        int count;
        windows.TryGetValue(hash, out count);
        windows[hash] = count + 1;
    }

    private static int CountWindowIntersection(IReadOnlyDictionary<string, int> source, IReadOnlyDictionary<string, int> target) {
        int count = 0;
        foreach (KeyValuePair<string, int> pair in source) {
            int targetCount;
            if (target.TryGetValue(pair.Key, out targetCount)) {
                count += Math.Min(pair.Value, targetCount);
            }
        }

        return count;
    }

    private static int CountWindowUnion(IReadOnlyDictionary<string, int> source, IReadOnlyDictionary<string, int> target) {
        int count = 0;
        var keys = new HashSet<string>(source.Keys, StringComparer.Ordinal);
        keys.UnionWith(target.Keys);
        foreach (string key in keys) {
            int sourceCount;
            int targetCount;
            source.TryGetValue(key, out sourceCount);
            target.TryGetValue(key, out targetCount);
            count += Math.Max(sourceCount, targetCount);
        }

        return count;
    }

    private static string NormalizeText(string text) {
        return string.IsNullOrWhiteSpace(text)
            ? string.Empty
            : string.Join(" ", text.Split(WhitespaceSeparators, StringSplitOptions.RemoveEmptyEntries));
    }

    private static string Hash(string value) {
        using (SHA256 sha = SHA256.Create()) {
            byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
            return Convert.ToBase64String(bytes);
        }
    }
}
