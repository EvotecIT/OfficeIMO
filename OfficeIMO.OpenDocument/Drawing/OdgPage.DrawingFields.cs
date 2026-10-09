namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Master text uses the page currently being rendered, never the master XML's position.
    private sealed class DrawingFieldContext {
        internal DrawingFieldContext(OdgPage page, int pageNumber, int pageCount, OdfDateTimeFieldProjection dateTimeFields,
            CancellationToken cancellationToken) {
            // A zero count denotes a standalone projection, which observes the current XML order.
            // Count native page elements directly without creating wrappers for unrelated pages.
            if (pageCount == 0) {
                foreach (XElement candidate in page._document.DrawingBody.Elements(OdfNamespaces.Draw + "page")) {
                    pageCount++;
                    if (ReferenceEquals(candidate, page.Element)) pageNumber = pageCount;
                }
            }
            PageCount = pageCount; PageNumber = pageNumber;
            XElement layout = page.LayoutProperties;
            NumberFormat = (string?)layout.Attribute(OdfNamespaces.Style + "num-format");
            NumberLetterSync = (string?)layout.Attribute(OdfNamespaces.Style + "num-letter-sync");
            DateTimeFields = dateTimeFields;
            TextFlowParticipants = FindTextFlowParticipants(page._document, cancellationToken);
        }
        internal int PageNumber { get; }
        internal int PageCount { get; }
        internal string? NumberFormat { get; }
        internal string? NumberLetterSync { get; }
        internal OdfDateTimeFieldProjection DateTimeFields { get; }
        internal HashSet<XElement> TextFlowParticipants { get; }
        internal Dictionary<XElement, TextFrameProjection> PreparedTextFrames { get; } = new();
        internal Dictionary<XElement, OdfRect> ProjectedFrameBounds { get; } = new();
    }

    private sealed class DrawingFieldResolver {
        private readonly DrawingFieldContext _context;
        private readonly OdfConversionReport _report;
        private readonly string _feature;
        private readonly HashSet<string> _reported = new(StringComparer.Ordinal);
        internal DrawingFieldResolver(DrawingFieldContext context, OdfConversionReport report, string feature) {
            _context = context; _report = report; _feature = feature;
        }

        internal string? Resolve(XElement element, string cache) {
            OdfTextField.TryGetKind(element.Name, out OdfTextFieldKind kind);
            string? fixedRaw = (string?)element.Attribute(OdfNamespaces.Text + "fixed");
            bool isFixed = false;
            if (fixedRaw != null && (!OdfBoolean.TryParseXml(fixedRaw, out isFixed) || kind == OdfTextFieldKind.PageCount))
                return Unsupported("field-fixed", "The invalid native fixed flag is preserved; cached text is retained.");
            if (kind is OdfTextFieldKind.Date or OdfTextFieldKind.Time && _context.DateTimeFields.EvaluatesValues) {
                try {
                    string formatted = _context.DateTimeFields.Resolve(element, kind, isFixed);
                    Approximate("field-date-time-value", "Date/time uses explicit civil values, Gregorian components and declared locale without host time-zone conversion. Native applications can normalize offsets or ignore styles and adjustments; saved XML is unchanged.");
                    return formatted;
                } catch (Exception exception) when (exception is NotSupportedException or InvalidDataException) {
                    return Unsupported("field-date-time-format", exception.Message + " Cached text is retained.");
                }
            }
            if (kind is OdfTextFieldKind.Date or OdfTextFieldKind.Time || isFixed) {
                if (cache.Length == 0) return Unsupported("field-cache-missing", "A fixed page number or date/time field without cached display text requires native formatting or refresh.");
                Approximate(isFixed ? "field-fixed-cache" : "field-date-time-cache",
                    isFixed ? "The fixed field's cached display text is used; native value/data-style formatting can differ." :
                    "Date/time uses the saved display snapshot. The system clock, adjustments and native data styles are not evaluated.");
                return cache;
            }
            if (element.Attributes().Any(attribute => !IsNumberAttribute(attribute, kind)))
                return Unsupported("field-attributes", "Unrecognized native numbering attributes are preserved; cached text is retained.");
            if (_context.PageNumber == 0)
                return Unsupported("field-page-context", "A detached page has no logical drawing page number or count context.");
            long number = _context.PageCount;
            if (kind == OdfTextFieldKind.PageNumber) {
                int delta;
                switch ((string?)element.Attribute(OdfNamespaces.Text + "select-page")) {
                    case null: case "current": delta = 0; break;
                    case "previous": delta = -1; break;
                    case "next": delta = 1; break;
                    default: return Unsupported("field-page-selection", "The native page selection is outside the projection profile; cached text is retained.");
                }
                string? raw = (string?)element.Attribute(OdfNamespaces.Text + "page-adjust");
                if (raw != null && !int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                    return Unsupported("field-page-adjustment", "The native page adjustment is not a supported integer; cached text is retained.");
                number = _context.PageNumber + delta;
                bool selectedExists = number >= 1 && number <= _context.PageCount;
                if (raw != null) number += int.Parse(raw, CultureInfo.InvariantCulture);
                if (!selectedExists || number < 1 || number > _context.PageCount) {
                    Approximate("field-page-number", "Page-number selection and adjustment use logical drawing page order; a missing selected or adjusted page displays no number.");
                    return string.Empty;
                }
            }
            string? format = (string?)element.Attribute(OdfNamespaces.Style + "num-format");
            string? syncRaw = (string?)element.Attribute(OdfNamespaces.Style + "num-letter-sync");
            if (kind == OdfTextFieldKind.PageNumber) {
                format ??= _context.NumberFormat;
                syncRaw ??= _context.NumberLetterSync;
            }
            bool sync = false;
            if (syncRaw != null && !OdfBoolean.TryParseXml(syncRaw, out sync))
                return Unsupported("field-number-format", "Invalid number-letter synchronization is preserved; cached text is retained.");
            if (format == null) {
                format = "1";
                Approximate("field-number-default", "An omitted field/page numbering format uses decimal numbering; native format defaults can differ.");
            }
            string result;
            try { result = format.Length == 0 ? string.Empty : OdfNumberingFormatter.Format(number, format, sync); }
            catch (NotSupportedException exception) { return Unsupported("field-number-format", exception.Message + " Cached text is retained."); }
            Approximate(kind == OdfTextFieldKind.PageNumber ? "field-page-number" : "field-page-count",
                kind == OdfTextFieldKind.PageNumber ? "The field resolves from the rendered page's logical order and native numbering attributes. Native Draw implementations can ignore fixed values, formats, selection or adjustments." :
                "Page count uses the drawing's current logical pages rather than a cached metadata statistic or printer-dependent pagination.");
            return result;
        }

        private static bool IsNumberAttribute(XAttribute attribute, OdfTextFieldKind kind) => attribute.IsNamespaceDeclaration ||
            attribute.Name == XNamespace.Xml + "id" || attribute.Name == OdfNamespaces.Style + "num-format" ||
            attribute.Name == OdfNamespaces.Style + "num-letter-sync" || kind == OdfTextFieldKind.PageNumber &&
            (attribute.Name == OdfNamespaces.Text + "fixed" || attribute.Name == OdfNamespaces.Text + "select-page" ||
             attribute.Name == OdfNamespaces.Text + "page-adjust");

        private string? Unsupported(string name, string message) {
            if (_reported.Add(name)) _report.Add(_feature + ":" + name, OdfConversionMappingStatus.Unsupported, message: message);
            return null;
        }
        private void Approximate(string name, string message) {
            if (_reported.Add(name)) _report.Add(_feature + ":" + name, OdfConversionMappingStatus.Approximated, message: message);
        }
    }
}
