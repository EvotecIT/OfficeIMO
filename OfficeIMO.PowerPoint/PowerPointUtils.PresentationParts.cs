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

            D.TableStyleList tableStyleList = new D.TableStyleList() { Default = DefaultTableStyleGuid };
            tableStyleList.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart.TableStyleList = tableStyleList;
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
            // Theme should live under /ppt/theme/theme1.xml; create it on the presentation part
            ThemePart themePart1 = presentationPart.AddNewPart<ThemePart>();
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid }) {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid }) {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid }) {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;
        }
    }
}
