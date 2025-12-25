using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using D = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        // Thumbnail extracted from Assets/PowerPointTemplates/PowerPointBlank.pptx (docProps/thumbnail.jpeg)
        private static readonly byte[] ThumbnailBytes = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUEBAUEBQUFBQUEBQUFBQUEBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQT/wAARCAAkAA4DAREAAhEBAxEB/8QAHQABAAEEAwEAAAAAAAAAAAAAAAYHAgUEAwIBCf/EADgQAAECAwQJBgUEAwAAAAAAAAECAwAEBRESBiExBxNBUWFxFCIyobHBFBZCksEVM1NiY//EABsBAQEAAwEBAQAAAAAAAAAAAAABAgMEAQUG/8QALREBAAIBBAECBQMEAwAAAAAAAAECAxEEEiExQQUTIlFhcZGh0fAUI0JSsf/aAAwDAQACEQMRAD8A9P8A0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARvbLcm6goc0JdU5pUW4RjjdM9i9Wxj6WutKLa9c7WJXu1L05PwnHA/0eFX3c/B/5UX5xvV+zY+q2cL6ZrGdOHFjT3yL1clOVruomDfaYy6USsvXxP6c8g3k3kDWK6byPXF0jHGU0kqdZSk0nFJJySgAAAAAAAAAAAAAAAAAAABK/pv3iQvqVps2fwXyrGnUsPW9jHTV6hBZ6wqn6UdZacK3WN+S2z87zuzp/8AflmWzNtZ0rJQ1pqSikopSnVKUqk1SlKSgAAbcRLRqy5F06ys3xcrKzJSmkoxVFLKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKU//Z");
        private static readonly Lazy<byte[]> ThemeBytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.theme1.xml"));
        private static readonly Lazy<byte[]> TableStylesBytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.tableStyles.xml"));

        private static void CreatePresentationPropertiesPart(PresentationPart presentationPart) {

            PresentationPropertiesPart part = presentationPart.PresentationPropertiesPart ?? presentationPart.AddNewPart<PresentationPropertiesPart>();

            part.PresentationProperties ??= new PresentationProperties();

            ShowProperties showProperties = part.PresentationProperties.ShowProperties ??= new ShowProperties();
            showProperties.ShowNarration = false;
            showProperties.ShowAnimation = true;
            showProperties.UseTimings = true;
        }

        private static void CreateViewPropertiesPart(PresentationPart presentationPart) {
            ViewPropertiesPart viewPart = presentationPart.ViewPropertiesPart ?? presentationPart.AddNewPart<ViewPropertiesPart>();

            NormalViewProperties normalViewProperties = new NormalViewProperties(
                new RestoredLeft() { Size = DefaultRestoredLeftSize, AutoAdjust = false },
                new RestoredTop() { Size = DefaultRestoredTopSize }
            );

            SlideViewProperties slideViewProperties = new SlideViewProperties();

            CommonSlideViewProperties commonSlideViewProperties = new CommonSlideViewProperties() { SnapToGrid = false };
            CommonViewProperties commonViewProperties = new CommonViewProperties() { VariableScale = true };

            ScaleFactor scaleFactor = new ScaleFactor();
            scaleFactor.Append(new D.ScaleX() { Numerator = 142, Denominator = 100 });
            scaleFactor.Append(new D.ScaleY() { Numerator = 142, Denominator = 100 });
            commonViewProperties.Append(scaleFactor);
            commonViewProperties.Append(new Origin() { X = 0L, Y = 0L });

            commonSlideViewProperties.Append(commonViewProperties);
            slideViewProperties.Append(commonSlideViewProperties);

            NotesTextViewProperties notesTextViewProperties = new NotesTextViewProperties();
            CommonViewProperties notesCommonViewProperties = new CommonViewProperties();
            ScaleFactor notesScaleFactor = new ScaleFactor();
            notesScaleFactor.Append(new D.ScaleX() { Numerator = 1, Denominator = 1 });
            notesScaleFactor.Append(new D.ScaleY() { Numerator = 1, Denominator = 1 });
            notesCommonViewProperties.Append(notesScaleFactor);
            notesCommonViewProperties.Append(new Origin() { X = 0L, Y = 0L });
            notesTextViewProperties.Append(notesCommonViewProperties);

            GridSpacing gridSpacing = new GridSpacing() { Cx = 72008L, Cy = 72008L };

            ViewProperties viewProperties = new ViewProperties();
            viewProperties.Append(normalViewProperties);
            viewProperties.Append(slideViewProperties);
            viewProperties.Append(notesTextViewProperties);
            viewProperties.Append(gridSpacing);

            viewPart.ViewProperties = viewProperties;
        }

        private static void CreateTableStylesPart(PresentationPart presentationPart) {
            TableStylesPart tableStylesPart = presentationPart.TableStylesPart ?? presentationPart.AddNewPart<TableStylesPart>();

            if (tableStylesPart.TableStyleList == null) {
                using var stream = new MemoryStream(TableStylesBytes.Value);
                tableStylesPart.FeedData(stream);
            }
        }

        private static void EnsureDocumentProperties(PresentationPart presentationPart) {
            if (presentationPart.OpenXmlPackage is not PresentationDocument presentationDocument) {
                return;
            }

            ExtendedFilePropertiesPart extendedPart = presentationDocument.ExtendedFilePropertiesPart ?? presentationDocument.AddExtendedFilePropertiesPart();
            if (extendedPart.Properties == null) {
                extendedPart.Properties = new Ap.Properties();
            }

            extendedPart.Properties.TotalTime ??= new Ap.TotalTime() { Text = "0" };
            extendedPart.Properties.Application ??= new Ap.Application() { Text = "Microsoft Office PowerPoint" };
            extendedPart.Properties.PresentationFormat ??= new Ap.PresentationFormat() { Text = "Widescreen" };
            extendedPart.Properties.Slides ??= new Ap.Slides() { Text = "1" };
            extendedPart.Properties.Notes ??= new Ap.Notes() { Text = "0" };
            extendedPart.Properties.HiddenSlides ??= new Ap.HiddenSlides() { Text = "0" };

            DateTime timestamp = DateTime.UtcNow;
            CoreFilePropertiesPart corePart = presentationDocument.CoreFilePropertiesPart ?? presentationDocument.AddCoreFilePropertiesPart();
            bool coreHasContent;

            using (Stream coreStream = corePart.GetStream(FileMode.OpenOrCreate, FileAccess.Read)) {
                coreHasContent = coreStream.Length > 0;
            }

            if (!coreHasContent) {
                InitializeCoreFilePropertiesPart(corePart, timestamp);
            }

            var packageProperties = presentationDocument.PackageProperties;

            if (string.IsNullOrEmpty(packageProperties.Creator)) {
                packageProperties.Creator = DefaultDocumentAuthor;
            }

            if (string.IsNullOrEmpty(packageProperties.LastModifiedBy)) {
                packageProperties.LastModifiedBy = DefaultDocumentAuthor;
            }

            if (packageProperties.Created == null) {
                packageProperties.Created = timestamp;
            }

            if (packageProperties.Modified == null) {
                packageProperties.Modified = timestamp;
            }
        }

        private static void EnsureThumbnail(PresentationDocument doc) {
            if (doc.ThumbnailPart != null) return;
            ThumbnailPart thumbnailPart = doc.AddThumbnailPart("image/jpeg");
            using Stream stream = thumbnailPart.GetStream(FileMode.Create, FileAccess.Write);
            stream.Write(ThumbnailBytes, 0, ThumbnailBytes.Length);
        }

        private static void InitializeCoreFilePropertiesPart(CoreFilePropertiesPart corePart, DateTime timestamp) {
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace dcterms = "http://purl.org/dc/terms/";
            XNamespace dcmitype = "http://purl.org/dc/dcmitype/";
            XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";

            string serializedTimestamp = timestamp.ToString("s", CultureInfo.InvariantCulture) + "Z";

            XDocument coreDocument = new XDocument(
                new XElement(cp + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", cp),
                    new XAttribute(XNamespace.Xmlns + "dc", dc),
                    new XAttribute(XNamespace.Xmlns + "dcterms", dcterms),
                    new XAttribute(XNamespace.Xmlns + "dcmitype", dcmitype),
                    new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                    new XElement(dc + "creator", DefaultDocumentAuthor),
                    new XElement(cp + "lastModifiedBy", DefaultDocumentAuthor),
                    new XElement(dcterms + "created",
                        new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                        serializedTimestamp),
                    new XElement(dcterms + "modified",
                        new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                        serializedTimestamp))
            );

            using Stream stream = corePart.GetStream(FileMode.Create, FileAccess.Write);
            coreDocument.Save(stream);
        }

        private static ThemePart CreateTheme(PresentationPart presentationPart) {
            ThemePart themePart1 = presentationPart.AddNewPart<ThemePart>();
            using var stream = new MemoryStream(ThemeBytes.Value);
            themePart1.FeedData(stream);
            return themePart1;
        }
    }
}

