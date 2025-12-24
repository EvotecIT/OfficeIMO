using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// PowerPoint utility methods based on the working open-xml-sdk-snippets implementation.
    /// CRITICAL: This class contains the exact initialization pattern required to prevent
    /// PowerPoint from showing a "repair" dialog. The order and relationship IDs used here
    /// are very specific and must not be changed.
    /// </summary>
    internal static partial class PowerPointUtils {
        private const int DefaultRestoredLeftSize = 15989;
        private const int DefaultRestoredTopSize = 94660;
        private const string DefaultTableStyleGuid = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
        private const string DefaultDocumentAuthor = "OfficeIMO";

        public static PresentationDocument CreatePresentation(string filepath) {
            // Create a presentation at a specified file path. The presentation document type is pptx by default.
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreatePresentationParts(presentationDoc, presentationPart);

            return presentationDoc;
        }

        internal static void CreatePresentationParts(PresentationDocument presentationDocument, PresentationPart presentationPart) {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            // Match the common 16:9 widescreen default (same as the shipped blank template)
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000, Type = SlideSizeValues.Screen16x9 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);
            presentationPart.Presentation.SaveSubsetFonts = true;

            // Create master and layouts directly under the presentation part so they land at ppt/slideMasters and ppt/slideLayouts.
            SlideMasterPart slideMasterPart1 = presentationPart.AddNewPart<SlideMasterPart>("rId1");
            slideMasterPart1.SlideMaster = CreateSlideMasterSkeleton();

            // Initial layout (Title Slide)
            SlideLayoutPart titleLayout = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            titleLayout.SlideLayout = CreateTitleSlideLayout();
            titleLayout.AddPart(slideMasterPart1);

            // Additional default layouts (full set of 11 matching a blank PowerPoint)
            CreateAdditionalSlideLayouts(slideMasterPart1, titleLayout);

            // Theme stored under /ppt/theme/theme1.xml and linked from master
            ThemePart themePart1 = CreateTheme(presentationPart);
            slideMasterPart1.AddPart(themePart1, "rId12");

            // Create initial slide and link to the title layout
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = CreateBlankSlide();
            slidePart1.AddPart(titleLayout, "rId1");

            CreatePresentationPropertiesPart(presentationPart);
            CreateViewPropertiesPart(presentationPart);
            CreateTableStylesPart(presentationPart);
            EnsureNotesMasterPart(presentationPart);
            EnsureDocumentProperties(presentationPart);
            EnsureThumbnail(presentationDocument);
        }


        private static byte[] LoadEmbeddedResource(string resourceName) {
            var assembly = typeof(PowerPointUtils).Assembly;
            using Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) {
                throw new InvalidOperationException($"Missing embedded resource '{resourceName}'.");
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

    }
}
