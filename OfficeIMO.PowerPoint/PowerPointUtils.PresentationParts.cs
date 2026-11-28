using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using D = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        // Thumbnail extracted from Assets/PowerPointTemplates/PowerPointBlank.pptx (docProps/thumbnail.jpeg)
        private static readonly byte[] ThumbnailBytes = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUEBAUEBQUFBQUEBQUFBQUEBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQT/wAARCAAkAA4DAREAAhEBAxEB/8QAHQABAAEEAwEAAAAAAAAAAAAAAAYHAgUEAwIBCf/EADgQAAECAwQJBgUEAwAAAAAAAAECAwAEBRESBiExBxNBUWFxFCIyobHBFBZCksEVM1NiY//EABsBAQEAAwEBAQAAAAAAAAAAAAABAgMEAQUG/8QALREBAAIBBAECBQMEAwAAAAAAAAECAxEEEiExQQUTIlFhcZGh0fAUI0JSsf/aAAwDAQACEQMRAD8A9P8A0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARvbLcm6goc0JdU5pUW4RjjdM9i9Wxj6WutKLa9c7WJXu1L05PwnHA/0eFX3c/B/5UX5xvV+zY+q2cL6ZrGdOHFjT3yL1clOVruomDfaYy6USsvXxP6c8g3k3kDWK6byPXF0jHGU0kqdZSk0nFJJySgAAAAAAAAAAAAAAAAAAABK/pv3iQvqVps2fwXyrGnUsPW9jHTV6hBZ6wqn6UdZacK3WN+S2z87zuzp/8AflmWzNtZ0rJQ1pqSikopSnVKUqk1SlKSgAAbcRLRqy5F06ys3xcrKzJSmkoxVFLKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKUpSlKU//Z");
        private static readonly Lazy<byte[]> ThemeBytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.theme1.xml"));
        internal static readonly Lazy<byte[]> TableStylesBytes =
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

        internal static void CreateTableStylesPart(PresentationPart presentationPart) {
            TableStylesPart tableStylesPart = presentationPart.TableStylesPart ?? presentationPart.AddNewPart<TableStylesPart>();

            if (tableStylesPart.TableStyleList == null) {
                using var stream = new MemoryStream(TableStylesBytes.Value);
                tableStylesPart.FeedData(stream);
            }
        }

        private static Ap.Properties EnsureApplicationProperties(PresentationDocument presentationDocument) {
            Ap.Properties properties = GetSdkApplicationProperties(presentationDocument);
            properties.TotalTime ??= new Ap.TotalTime() { Text = "0" };
            properties.Application ??= new Ap.Application() { Text = "Microsoft Office PowerPoint" };
            properties.PresentationFormat ??= new Ap.PresentationFormat() { Text = "Widescreen" };
            properties.Slides ??= new Ap.Slides() { Text = "1" };
            properties.Notes ??= new Ap.Notes() { Text = "0" };
            properties.HiddenSlides ??= new Ap.HiddenSlides() { Text = "0" };
            return properties;
        }

        internal static string? ReadSdkBuiltinProperty(PowerPointDocument document, string name) {
            var properties = document.SdkDocument.PackageProperties;
            return name switch {
                "creator" => properties.Creator,
                "title" => properties.Title,
                "description" => properties.Description,
                "category" => properties.Category,
                "keywords" => properties.Keywords,
                "subject" => properties.Subject,
                "revision" => properties.Revision,
                "lastModifiedBy" => properties.LastModifiedBy,
                "version" => properties.Version,
                _ => throw new ArgumentOutOfRangeException(nameof(name))
            };
        }

        internal static void WriteSdkBuiltinProperty(PowerPointDocument document, string name, string? value) {
            var properties = document.SdkDocument.PackageProperties;
            switch (name) {
                case "creator": properties.Creator = value; break;
                case "title": properties.Title = value; break;
                case "description": properties.Description = value; break;
                case "category": properties.Category = value; break;
                case "keywords": properties.Keywords = value; break;
                case "subject": properties.Subject = value; break;
                case "revision": properties.Revision = value; break;
                case "lastModifiedBy": properties.LastModifiedBy = value; break;
                case "version": properties.Version = value; break;
                default: throw new ArgumentOutOfRangeException(nameof(name));
            }
        }

        internal static DateTime? ReadSdkBuiltinDateProperty(PowerPointDocument document, string name) {
            var properties = document.SdkDocument.PackageProperties;
            return name switch {
                "created" => properties.Created,
                "modified" => properties.Modified,
                "lastPrinted" => properties.LastPrinted,
                _ => throw new ArgumentOutOfRangeException(nameof(name))
            };
        }

        internal static void WriteSdkBuiltinDateProperty(PowerPointDocument document, string name, DateTime? value) {
            var properties = document.SdkDocument.PackageProperties;
            switch (name) {
                case "created": properties.Created = value; break;
                case "modified": properties.Modified = value; break;
                case "lastPrinted": properties.LastPrinted = value; break;
                default: throw new ArgumentOutOfRangeException(nameof(name));
            }
        }

        private static Ap.Properties GetSdkApplicationProperties(PresentationDocument document) {
            ExtendedFilePropertiesPart part = document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart();
            return part.Properties ??= new Ap.Properties();
        }

        internal static string ReadSdkApplicationProperty(PowerPointDocument document, string name) {
            Ap.Properties properties = GetSdkApplicationProperties(document.SdkDocument);
            return (name switch {
                "Application" => properties.Application?.Text,
                "AppVersion" => properties.ApplicationVersion?.Text,
                "Company" => properties.Company?.Text,
                "Manager" => properties.Manager?.Text,
                "PresentationFormat" => properties.PresentationFormat?.Text,
                "TotalTime" => properties.TotalTime?.Text,
                "Slides" => properties.Slides?.Text,
                "Notes" => properties.Notes?.Text,
                "HiddenSlides" => properties.HiddenSlides?.Text,
                _ => throw new ArgumentOutOfRangeException(nameof(name))
            }) ?? string.Empty;
        }

        internal static void WriteSdkApplicationProperty(PowerPointDocument document, string name, string? value) {
            Ap.Properties properties = GetSdkApplicationProperties(document.SdkDocument);
            OpenXmlLeafTextElement property = name switch {
                "Application" => properties.Application ??= new Ap.Application(),
                "AppVersion" => properties.ApplicationVersion ??= new Ap.ApplicationVersion(),
                "Company" => properties.Company ??= new Ap.Company(),
                "Manager" => properties.Manager ??= new Ap.Manager(),
                "PresentationFormat" => properties.PresentationFormat ??= new Ap.PresentationFormat(),
                "TotalTime" => properties.TotalTime ??= new Ap.TotalTime(),
                _ => throw new ArgumentOutOfRangeException(nameof(name))
            };
            property.Text = value ?? string.Empty;
        }

        private static void InitializeDocumentProperties(PresentationDocument presentationDocument) {
            EnsureApplicationProperties(presentationDocument);
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

        internal static void UpdateDocumentProperties(PresentationPart presentationPart, long maximumSlideXmlCharacters) {
            if (presentationPart.OpenXmlPackage is not PresentationDocument presentationDocument) {
                return;
            }

            Ap.Properties properties = EnsureApplicationProperties(presentationDocument);
            int slideCount = 0;
            int hiddenCount = 0;
            SlideIdList? slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList != null) {
                foreach (SlideId slideId in slideIdList.Elements<SlideId>()) {
                    slideCount++;
                    if (IsHiddenSlide(presentationPart, slideId, maximumSlideXmlCharacters)) hiddenCount++;
                }
            }
            int notesCount = 0;
            foreach (SlidePart slidePart in presentationPart.SlideParts) {
                if (slidePart.NotesSlidePart != null) notesCount++;
            }
            properties.Slides!.Text = slideCount.ToString(CultureInfo.InvariantCulture);
            properties.Notes!.Text = notesCount.ToString(CultureInfo.InvariantCulture);
            properties.HiddenSlides!.Text = hiddenCount.ToString(CultureInfo.InvariantCulture);
            properties.PresentationFormat!.Text = GetPresentationFormat(presentationPart.Presentation?.SlideSize);

            presentationDocument.PackageProperties.Modified = DateTime.UtcNow;
        }

        private static bool IsHiddenSlide(PresentationPart presentationPart, SlideId slideId, long maximumSlideXmlCharacters) {
            string? relationshipId = slideId.RelationshipId?.Value;
            if (relationshipId is { Length: > 0 } &&
                presentationPart.TryGetPartById(relationshipId, out OpenXmlPart? part) &&
                part is SlidePart slidePart) {
                bool? show;
                if (slidePart.IsRootElementLoaded) {
                    show = slidePart.Slide?.Show?.Value;
                } else {
                    using Stream stream = slidePart.GetStream(FileMode.Open,
                        FileAccess.Read);
                    show = PowerPointXmlReader.ReadSlideShow(stream, maximumSlideXmlCharacters);
                }
                if (show.HasValue) return !show.Value;
            }

            string? legacyShowValue = null;
            if (slideId.HasAttributes) {
                foreach (OpenXmlAttribute attribute in slideId.ExtendedAttributes) {
                    if (attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri)) {
                        legacyShowValue = attribute.Value;
                        break;
                    }
                }
            }
            if (string.IsNullOrEmpty(legacyShowValue)) {
                return false;
            }

            return string.Equals(legacyShowValue, "0", StringComparison.Ordinal) ||
                   string.Equals(legacyShowValue, "false", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetPresentationFormat(SlideSize? slideSize) {
            SlideSizeValues? type = slideSize?.Type?.Value;
            if (type == SlideSizeValues.Screen4x3) {
                return "Standard";
            }

            if (type == SlideSizeValues.Custom) {
                return "Custom";
            }

            if (type == SlideSizeValues.Screen16x9 || type == SlideSizeValues.Screen16x10) {
                return "Widescreen";
            }

            return "Widescreen";
        }

        private static void EnsureThumbnail(PresentationDocument doc) {
            if (doc.ThumbnailPart != null) return;
            ThumbnailPart thumbnailPart = doc.AddThumbnailPart(OfficeImageInfo.GetMimeType(OfficeImageFormat.Jpeg));
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
