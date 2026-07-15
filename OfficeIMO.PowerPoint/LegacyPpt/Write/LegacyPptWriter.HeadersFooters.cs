using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordHeadersFooters = 0x0FD9;
        private const ushort RecordHeadersFootersAtom = 0x0FDA;
        private const ushort RecordCString = 0x0FBA;

        private static byte[] BuildDocumentHeaderFooterRecord(
            PowerPointPresentation presentation, ushort instance) {
            LegacyPptWriterHeaderFooter settings = instance == 4
                ? ReadNotesHeaderFooter(presentation)
                : ReadSlideHeaderFooterDefaults(presentation);
            return BuildHeaderFooterRecord(settings, instance,
                allowHeader: instance == 4);
        }

        internal static LegacyPptWriterHeaderFooter? ReadSlideHeaderFooter(
            PowerPointSlide slide) {
            P.SlideLayout? layout = slide.SlidePart.SlideLayoutPart?.SlideLayout;
            P.HeaderFooter? headerFooter = layout?.GetFirstChild<P.HeaderFooter>();
            return headerFooter == null ? null : ReadHeaderFooter(
                headerFooter, layout?.CommonSlideData, allowHeader: false);
        }

        private static LegacyPptWriterHeaderFooter ReadSlideHeaderFooterDefaults(
            PowerPointPresentation presentation) {
            P.SlideMaster? master = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.FirstOrDefault()?.SlideMaster;
            P.HeaderFooter? headerFooter = master?.GetFirstChild<P.HeaderFooter>();
            return headerFooter == null
                ? LegacyPptWriterHeaderFooter.Empty
                : ReadHeaderFooter(headerFooter, master?.CommonSlideData,
                    allowHeader: false);
        }

        private static LegacyPptWriterHeaderFooter ReadNotesHeaderFooter(
            PowerPointPresentation presentation) {
            P.NotesMaster? notesMaster = presentation.OpenXmlDocument.PresentationPart?
                .NotesMasterPart?.NotesMaster;
            P.HeaderFooter? notesHeaderFooter = notesMaster?
                .GetFirstChild<P.HeaderFooter>();
            if (notesHeaderFooter != null) {
                return ReadHeaderFooter(notesHeaderFooter,
                    notesMaster?.CommonSlideData, allowHeader: true);
            }

            P.HandoutMaster? handoutMaster = presentation.OpenXmlDocument.PresentationPart?
                .HandoutMasterPart?.HandoutMaster;
            P.HeaderFooter? handoutHeaderFooter = handoutMaster?
                .GetFirstChild<P.HeaderFooter>();
            return handoutHeaderFooter == null
                ? LegacyPptWriterHeaderFooter.Empty
                : ReadHeaderFooter(handoutHeaderFooter,
                    handoutMaster?.CommonSlideData, allowHeader: true);
        }

        private static LegacyPptWriterHeaderFooter ReadHeaderFooter(
            P.HeaderFooter headerFooter, P.CommonSlideData? commonSlideData,
            bool allowHeader) {
            string userDate = ReadPlaceholderText(commonSlideData,
                P.PlaceholderValues.DateAndTime);
            string header = allowHeader
                ? ReadPlaceholderText(commonSlideData, P.PlaceholderValues.Header)
                : string.Empty;
            string footer = ReadPlaceholderText(commonSlideData,
                P.PlaceholderValues.Footer);
            bool showDate = ReadHeaderFooterBoolean(headerFooter.DateTime);
            bool useUserDate = showDate && userDate.Length > 0
                && !ContainsDateTimeField(commonSlideData);
            return new LegacyPptWriterHeaderFooter(
                formatId: 0,
                showDate,
                useAutomaticDateTime: showDate && !useUserDate,
                useUserDate,
                showSlideNumber: ReadHeaderFooterBoolean(headerFooter.SlideNumber),
                showHeader: allowHeader && ReadHeaderFooterBoolean(headerFooter.Header),
                showFooter: ReadHeaderFooterBoolean(headerFooter.Footer),
                userDate, header, footer);
        }

        private static bool ReadHeaderFooterBoolean(BooleanValue? value) =>
            value?.Value != false;

        private static string ReadPlaceholderText(P.CommonSlideData? commonSlideData,
            P.PlaceholderValues type) {
            P.Shape? shape = commonSlideData?.ShapeTree?.Elements<P.Shape>()
                .FirstOrDefault(candidate => candidate.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == type);
            return shape == null ? string.Empty : string.Concat(
                shape.TextBody?.Descendants<A.Text>().Select(text => text.Text)
                ?? Enumerable.Empty<string>());
        }

        private static bool ContainsDateTimeField(P.CommonSlideData? commonSlideData) {
            P.Shape? shape = commonSlideData?.ShapeTree?.Elements<P.Shape>()
                .FirstOrDefault(candidate => candidate.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value
                    == P.PlaceholderValues.DateAndTime);
            return shape?.TextBody?.Descendants<A.Field>().Any() == true;
        }

        internal static byte[] BuildHeaderFooterRecord(
            LegacyPptWriterHeaderFooter settings, ushort instance,
            bool allowHeader) {
            ushort flags = 0;
            if (settings.ShowDate) flags |= 0x0001;
            if (settings.UseAutomaticDateTime) flags |= 0x0002;
            if (settings.UseUserDate) flags |= 0x0004;
            if (settings.ShowSlideNumber) flags |= 0x0008;
            if (allowHeader && settings.ShowHeader) flags |= 0x0010;
            if (settings.ShowFooter) flags |= 0x0020;
            var atomPayload = new byte[4];
            WriteInt16(atomPayload, 0, settings.FormatId);
            WriteUInt16(atomPayload, 2, flags);
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0, RecordHeadersFootersAtom,
                    atomPayload)
            };
            if (settings.UseUserDate && settings.UserDateText.Length > 0) {
                children.Add(BuildCString(instance: 0, settings.UserDateText));
            }
            if (allowHeader && settings.HeaderText.Length > 0) {
                children.Add(BuildCString(instance: 1, settings.HeaderText));
            }
            if (settings.FooterText.Length > 0) {
                children.Add(BuildCString(instance: 2, settings.FooterText));
            }
            return BuildContainer(RecordHeadersFooters, instance, children);
        }

        private static byte[] BuildCString(ushort instance, string value) =>
            BuildRecord(version: 0, instance, RecordCString,
                Encoding.Unicode.GetBytes(value ?? string.Empty));

        internal sealed class LegacyPptWriterHeaderFooter {
            internal static LegacyPptWriterHeaderFooter Empty { get; } = new(
                0, false, false, false, false, false, false,
                string.Empty, string.Empty, string.Empty);

            internal LegacyPptWriterHeaderFooter(short formatId, bool showDate,
                bool useAutomaticDateTime, bool useUserDate, bool showSlideNumber,
                bool showHeader, bool showFooter, string userDateText,
                string headerText, string footerText) {
                FormatId = formatId;
                ShowDate = showDate;
                UseAutomaticDateTime = useAutomaticDateTime;
                UseUserDate = useUserDate;
                ShowSlideNumber = showSlideNumber;
                ShowHeader = showHeader;
                ShowFooter = showFooter;
                UserDateText = userDateText ?? string.Empty;
                HeaderText = headerText ?? string.Empty;
                FooterText = footerText ?? string.Empty;
            }

            internal short FormatId { get; }
            internal bool ShowDate { get; }
            internal bool UseAutomaticDateTime { get; }
            internal bool UseUserDate { get; }
            internal bool ShowSlideNumber { get; }
            internal bool ShowHeader { get; }
            internal bool ShowFooter { get; }
            internal string UserDateText { get; }
            internal string HeaderText { get; }
            internal string FooterText { get; }

            internal static LegacyPptWriterHeaderFooter? FromLegacy(
                OfficeIMO.PowerPoint.LegacyPpt.Model.LegacyPptHeaderFooterSettings? source) {
                if (source == null) return null;
                return new LegacyPptWriterHeaderFooter(source.DateTimeFormatId,
                    source.ShowDate, source.UseAutomaticDateTime, source.UseUserDate,
                    source.ShowSlideNumber, source.ShowHeader, source.ShowFooter,
                    source.UserDateText, source.HeaderText, source.FooterText);
            }

            internal bool IsEquivalentTo(LegacyPptWriterHeaderFooter? other) =>
                other != null
                && FormatId == other.FormatId
                && ShowDate == other.ShowDate
                && UseAutomaticDateTime == other.UseAutomaticDateTime
                && UseUserDate == other.UseUserDate
                && ShowSlideNumber == other.ShowSlideNumber
                && ShowHeader == other.ShowHeader
                && ShowFooter == other.ShowFooter
                && string.Equals(UserDateText, other.UserDateText,
                    StringComparison.Ordinal)
                && string.Equals(HeaderText, other.HeaderText,
                    StringComparison.Ordinal)
                && string.Equals(FooterText, other.FooterText,
                    StringComparison.Ordinal);
        }
    }
}
