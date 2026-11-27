using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordHeadersFooters = 0x0FD9;
        private const ushort RecordHeadersFootersAtom = 0x0FDA;
        private const ushort RecordCString = 0x0FBA;

        private static byte[] BuildDocumentHeaderFooterRecord(
            LegacyPptPackageView view, ushort instance) {
            LegacyPptWriterHeaderFooter settings = instance == 4
                ? ReadNotesHeaderFooter(view)
                : ReadSlideHeaderFooterDefaults(view);
            return BuildHeaderFooterRecord(settings, instance,
                allowHeader: instance == 4);
        }

        // Caller in LegacyPptPreservingWriter.ChangePlan.cs not yet passing a view.
        internal static LegacyPptWriterHeaderFooter? ReadSlideHeaderFooter(
            PowerPointSlide slide) {
            ArgumentNullException.ThrowIfNull(slide);
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(OwnerDocumentForWrite(slide));
            return ReadSlideHeaderFooter(SlidePartView(view, slide));
        }

        internal static LegacyPptWriterHeaderFooter? ReadSlideHeaderFooter(
            LegacyPptPartView slidePart) {
            OfficeXmlNode layout = slidePart.SlideLayout.Root;
            OfficeXmlNode headerFooter = layout.Child(Pml, "hf");
            return headerFooter.IsNone ? null : ReadHeaderFooter(
                headerFooter, layout.Child(Pml, "cSld"), allowHeader: false);
        }

        private static LegacyPptWriterHeaderFooter ReadSlideHeaderFooterDefaults(
            LegacyPptPackageView view) {
            List<LegacyPptPartView> masters = SlideMasterPartsInRelationshipOrder(view);
            OfficeXmlNode master = masters.Count == 0 ? default : masters[0].Root;
            OfficeXmlNode headerFooter = master.Child(Pml, "hf");
            return headerFooter.IsNone
                ? LegacyPptWriterHeaderFooter.Empty
                : ReadHeaderFooter(headerFooter, master.Child(Pml, "cSld"),
                    allowHeader: false);
        }

        private static LegacyPptWriterHeaderFooter ReadNotesHeaderFooter(
            LegacyPptPackageView view) {
            OfficeXmlNode notesMaster = view.NotesMaster.Root;
            OfficeXmlNode notesHeaderFooter = notesMaster.Child(Pml, "hf");
            if (!notesHeaderFooter.IsNone) {
                return ReadHeaderFooter(notesHeaderFooter,
                    notesMaster.Child(Pml, "cSld"), allowHeader: true);
            }

            OfficeXmlNode handoutMaster = view.HandoutMaster.Root;
            OfficeXmlNode handoutHeaderFooter = handoutMaster.Child(Pml, "hf");
            return handoutHeaderFooter.IsNone
                ? LegacyPptWriterHeaderFooter.Empty
                : ReadHeaderFooter(handoutHeaderFooter,
                    handoutMaster.Child(Pml, "cSld"), allowHeader: true);
        }

        private static LegacyPptWriterHeaderFooter ReadHeaderFooter(
            OfficeXmlNode headerFooter, OfficeXmlNode commonSlideData,
            bool allowHeader) {
            string userDate = ReadPlaceholderText(commonSlideData, "dt");
            string header = allowHeader
                ? ReadPlaceholderText(commonSlideData, "hdr")
                : string.Empty;
            string footer = ReadPlaceholderText(commonSlideData, "ftr");
            bool showDate = headerFooter.Bool("dt") != false;
            bool useUserDate = showDate && userDate.Length > 0
                && !ContainsDateTimeField(commonSlideData);
            return new LegacyPptWriterHeaderFooter(
                formatId: 0,
                showDate,
                useAutomaticDateTime: showDate && !useUserDate,
                useUserDate,
                showSlideNumber: headerFooter.Bool("sldNum") != false,
                showHeader: allowHeader && headerFooter.Bool("hdr") != false,
                showFooter: headerFooter.Bool("ftr") != false,
                userDate, header, footer);
        }

        // The first p:sp of the shape tree whose placeholder has this type.
        private static OfficeXmlNode FindPlaceholderShape(OfficeXmlNode commonSlideData,
            string type) {
            foreach (OfficeXmlNode shape in commonSlideData.Child(Pml, "spTree").Children(Pml, "sp")) {
                if (string.Equals(shape.Path(Pml, "nvSpPr", "nvPr", "ph").Attribute("type"), type,
                        StringComparison.Ordinal)) return shape;
            }
            return default;
        }

        private static string ReadPlaceholderText(OfficeXmlNode commonSlideData,
            string type) {
            OfficeXmlNode shape = FindPlaceholderShape(commonSlideData, type);
            if (shape.IsNone) return string.Empty;
            OfficeXmlNode body = shape.Child(Pml, "txBody");
            return body.IsNone ? string.Empty : string.Concat(
                body.Descendants(Dml, "t").Select(text => text.InnerText));
        }

        private static bool ContainsDateTimeField(OfficeXmlNode commonSlideData) {
            OfficeXmlNode shape = FindPlaceholderShape(commonSlideData, "dt");
            return shape.Child(Pml, "txBody").Descendants(Dml, "fld").Count > 0;
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
