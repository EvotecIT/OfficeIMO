using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private const int InvisibleFlag = 1;
    private const int HiddenFlag = 2;
    private const int PrintFlag = 4;
    private const int NoViewFlag = 32;
    private const int ToggleNoViewFlag = 256;

    private static readonly HashSet<string> KnownAnnotationTypes = new(StringComparer.Ordinal) {
        "Text", "Link", "FreeText", "Line", "Square", "Circle", "Polygon", "PolyLine", "Highlight", "Underline", "Squiggly",
        "StrikeOut", "Stamp", "Caret", "Ink", "Popup", "FileAttachment", "Widget", "PrinterMark", "TrapNet", "Watermark", "Redact"
    };

    private static readonly HashSet<string> ForbiddenAnnotationTypes = new(StringComparer.Ordinal) {
        "3D", "Sound", "Screen", "Movie", "RichMedia"
    };

    private void ProcessPages() {
        IReadOnlyList<PdfReadPage> pages = _document.Pages;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            PdfDictionary page = pages[pageIndex].PageDictionary;
            string location = "page " + (pageIndex + 1).ToString(CultureInfo.InvariantCulture);
            if (page.Items.ContainsKey("AA") &&
                Change(PdfAConversionChangeKind.ActionRemoved, location, "Removed page additional-actions (AA), which PDF/A forbids.", page, OfficeConversionLossKind.Omission)) {
                page.Items.Remove("AA");
            }
            if (page.Items.ContainsKey("PresSteps") &&
                Change(PdfAConversionChangeKind.StructureUpdated, location, "Removed presentation steps (PresSteps), which PDF/A forbids.", page, OfficeConversionLossKind.Omission)) {
                page.Items.Remove("PresSteps");
            }
            if (Resolve(page.Items.TryGetValue("Annots", out PdfObject? annots) ? annots : null) is not PdfArray annotations) continue;
            var kept = new List<PdfObject>(annotations.Items.Count);
            bool removed = false;
            for (int index = 0; index < annotations.Items.Count; index++) {
                PdfObject item = annotations.Items[index];
                string annotationLocation = location + ", annotation " + (index + 1).ToString(CultureInfo.InvariantCulture);
                if (ResolveDictionary(item) is not PdfDictionary annotation) {
                    kept.Add(item);
                    continue;
                }
                if (ProcessAnnotation(annotation, annotationLocation)) {
                    kept.Add(item);
                } else {
                    removed = true;
                }
            }
            if (removed && !_verifyOnly) {
                annotations.Items.Clear();
                annotations.Items.AddRange(kept);
            }
        }
    }

    /// <summary>Fixes one annotation; returns <see langword="false"/> when it is removed.</summary>
    private bool ProcessAnnotation(PdfDictionary annotation, string location) {
        string subtype = GetName(annotation, "Subtype") ?? string.Empty;
        string described = location + " (" + (subtype.Length == 0 ? "untyped" : subtype) + ")";
        int flags = Resolve(annotation.Items.TryGetValue("F", out PdfObject? flagValue) ? flagValue : null) is PdfNumber number ? (int)number.Value : 0;
        bool hidden = (flags & (HiddenFlag | NoViewFlag)) != 0;

        if (ForbiddenAnnotationTypes.Contains(subtype) || !KnownAnnotationTypes.Contains(subtype)) {
            if (hidden || !HasVisibleAppearance(annotation)) {
                return !Change(PdfAConversionChangeKind.AnnotationRemoved, described,
                    "Removed a " + (subtype.Length == 0 ? "untyped" : subtype) + " annotation, which PDF/A forbids; it draws nothing on the page.",
                    annotation, OfficeConversionLossKind.Omission);
            }
            AddViolation(PdfAConversionRefusalReason.ForbiddenAnnotation, described,
                "A " + (subtype.Length == 0 ? "untyped" : subtype) + " annotation draws on the page, and PDF/A forbids this annotation type.", annotation);
            return true;
        }

        if (subtype == "Popup") {
            if (annotation.Items.ContainsKey("F") && ((flags & PrintFlag) == 0 || (flags & (InvisibleFlag | HiddenFlag | NoViewFlag | ToggleNoViewFlag)) != 0) &&
                Change(PdfAConversionChangeKind.AnnotationFlagsUpdated, described, "Removed pop-up window flags that PDF/A does not allow; the pop-up still opens from its parent.", annotation)) {
                annotation.Items.Remove("F");
            }
        } else if (hidden) {
            return !Change(PdfAConversionChangeKind.AnnotationRemoved, described,
                "Removed a hidden annotation. PDF/A forbids hidden annotations, and showing it would change the page.", annotation, OfficeConversionLossKind.Omission);
        } else {
            int fixedFlags = (flags | PrintFlag) & ~(InvisibleFlag | ToggleNoViewFlag);
            if ((!annotation.Items.ContainsKey("F") || fixedFlags != flags) &&
                Change(PdfAConversionChangeKind.AnnotationFlagsUpdated, described,
                    (flags & PrintFlag) == 0 ? "Set the Print flag, as PDF/A requires; the annotation now prints as it displays." : "Cleared annotation flags PDF/A forbids.",
                    annotation)) {
                annotation.Items["F"] = new PdfNumber(fixedFlags);
            }
        }

        if (annotation.Items.ContainsKey("AA") &&
            Change(PdfAConversionChangeKind.ActionRemoved, described, "Removed annotation additional-actions (AA), which PDF/A forbids.", annotation, OfficeConversionLossKind.Omission)) {
            annotation.Items.Remove("AA");
        }
        if (annotation.Items.TryGetValue("A", out PdfObject? action)) {
            PdfObject? keptAction = SanitizeAction(action, described, 0);
            if (keptAction is null) annotation.Items.Remove("A");
            else annotation.Items["A"] = keptAction;
        }

        ProcessAppearance(annotation, subtype, described);
        return true;
    }

    private void ProcessAppearance(PdfDictionary annotation, string subtype, string location) {
        PdfDictionary? appearances = ResolveDictionary(annotation.Items.TryGetValue("AP", out PdfObject? ap) ? ap : null);
        bool exempt = subtype is "Popup" or "Link" || HasZeroAreaRect(annotation);
        bool regenerate = subtype == "Widget" && _needAppearances && PdfFormFiller.IsVariableTextWidget(_objects, annotation);
        if ((appearances is null && !exempt) || regenerate) {
            if (TrySynthesizeAppearance(annotation, subtype, out PdfObject? synthesized)) {
                if (Change(PdfAConversionChangeKind.AnnotationAppearanceUpdated, location,
                        regenerate
                            ? "Stored the field appearance viewers generate from its value, because NeedAppearances asked them to regenerate it."
                            : "Stored the appearance viewers draw for this annotation, which had none; PDF/A requires an appearance stream.",
                        annotation)) {
                    var dictionary = new PdfDictionary();
                    dictionary.Items["N"] = synthesized!;
                    annotation.Items["AP"] = dictionary;
                    appearances = dictionary;
                }
            } else if (appearances is null) {
                AddViolation(PdfAConversionRefusalReason.AnnotationAppearanceMissing, location,
                    "The annotation has no appearance stream, and none can be built from its properties. PDF/A requires one.", annotation);
                return;
            }
        }
        if (appearances is null) return;
        if ((appearances.Items.ContainsKey("D") || appearances.Items.ContainsKey("R")) &&
            Change(PdfAConversionChangeKind.AnnotationAppearanceUpdated, location,
                "Removed the rollover and down appearances, which PDF/A forbids; the normal appearance is kept.", annotation)) {
            appearances.Items.Remove("D");
            appearances.Items.Remove("R");
        }
        PdfObject? normal = Resolve(appearances.Items.TryGetValue("N", out PdfObject? n) ? n : null);
        if (normal is null) {
            if (!exempt) {
                AddViolation(PdfAConversionRefusalReason.AnnotationAppearanceMissing, location, "The annotation appearance has no normal (N) appearance.", annotation);
            }
            return;
        }
        bool buttonWidget = subtype == "Widget" && GetInheritedFieldType(annotation) == "Btn";
        if (buttonWidget || normal is not PdfDictionary states) return;
        string? selected = GetName(annotation, "AS");
        if (selected is not null && states.Items.TryGetValue(selected, out PdfObject? selectedAppearance) && Resolve(selectedAppearance) is PdfStream) {
            if (Change(PdfAConversionChangeKind.AnnotationAppearanceUpdated, location,
                    "Replaced the appearance state dictionary with the selected state's appearance, as PDF/A requires for this annotation type.", annotation)) {
                appearances.Items["N"] = selectedAppearance;
                annotation.Items.Remove("AS");
            }
            return;
        }
        AddViolation(PdfAConversionRefusalReason.AnnotationAppearanceMissing, location,
            "The annotation's normal appearance is a state dictionary with no selected state; PDF/A requires a single appearance stream.", annotation);
    }

    /// <summary>Builds the appearance the renderer and viewers draw for an annotation without one.</summary>
    private bool TrySynthesizeAppearance(PdfDictionary annotation, string subtype, out PdfObject? normal) {
        normal = null;
        PdfStream stream;
        if (subtype == "Widget") {
            PdfReadLimits limits = _document.ReadOptions.Limits;
            if (!PdfFormFiller.TryCreateTransientWidgetAppearance(_objects, annotation, _acroForm, out stream,
                    limits.MaxDecodedStreamBytes, limits.MaxDecodedTextCharacters, _cancellationToken)) return false;
            if (GetInheritedFieldType(annotation) == "Btn") {
                string state = GetName(annotation, "AS") ?? GetInheritedValueName(annotation) ?? "Off";
                if (!_verifyOnly && GetName(annotation, "AS") is null) annotation.Items["AS"] = new PdfName(state);
                var states = new PdfDictionary();
                states.Items[state] = AddObject(stream);
                normal = states;
                return true;
            }
        } else if (!PdfAnnotationFlattener.TryCreateSyntheticAppearanceStream(_objects, annotation, out stream)) {
            return false;
        }
        normal = _verifyOnly ? stream : AddObject(stream);
        return true;
    }

    private string? GetInheritedValueName(PdfDictionary field) {
        PdfDictionary? current = field;
        for (int depth = 0; current is not null && depth < 64; depth++) {
            if (current.Items.TryGetValue("V", out PdfObject? value)) return NameOf(Resolve(value));
            current = ResolveDictionary(current.Items.TryGetValue("Parent", out PdfObject? parent) ? parent : null);
        }
        return null;
    }

    private string? GetInheritedFieldType(PdfDictionary field) {
        PdfDictionary? current = field;
        for (int depth = 0; current is not null && depth < 64; depth++) {
            string? type = GetName(current, "FT");
            if (type is not null) return type;
            current = ResolveDictionary(current.Items.TryGetValue("Parent", out PdfObject? parent) ? parent : null);
        }
        return null;
    }

    private bool HasZeroAreaRect(PdfDictionary annotation) {
        if (Resolve(annotation.Items.TryGetValue("Rect", out PdfObject? rect) ? rect : null) is not PdfArray array || array.Items.Count < 4) return false;
        double[] values = array.Items.Take(4).Select(item => Resolve(item) is PdfNumber value ? value.Value : 0D).ToArray();
        return Math.Abs(values[2] - values[0]) < 1e-9 || Math.Abs(values[3] - values[1]) < 1e-9;
    }

    private bool HasVisibleAppearance(PdfDictionary annotation) {
        if (HasZeroAreaRect(annotation)) return false;
        if (ResolveDictionary(annotation.Items.TryGetValue("AP", out PdfObject? ap) ? ap : null) is not PdfDictionary appearances) return false;
        PdfObject? normal = Resolve(appearances.Items.TryGetValue("N", out PdfObject? n) ? n : null);
        if (normal is PdfStream stream) return HasPaintingContent(stream);
        if (normal is PdfDictionary states) {
            foreach (PdfObject state in states.Items.Values) {
                if (Resolve(state) is PdfStream stateStream && HasPaintingContent(stateStream)) return true;
            }
        }
        return false;
    }

    private bool HasPaintingContent(PdfStream stream) {
        if (!StreamDecoder.TryDecode(stream, _document.ReadOptions.Limits.MaxDecodedStreamBytes, out byte[] data, _objects)) return true;
        foreach (byte value in data) {
            if (value is not ((byte)' ' or (byte)'\n' or (byte)'\r' or (byte)'\t' or (byte)'\f' or 0)) return true;
        }
        return false;
    }
}
