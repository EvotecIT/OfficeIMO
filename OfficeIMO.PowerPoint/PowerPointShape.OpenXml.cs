using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        private NonVisualDrawingProperties? GetNonVisualDrawingProperties(bool create) {
            switch (Element) {
                case Shape s: {
                    if (create) {
                        s.NonVisualShapeProperties ??= new NonVisualShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties());
                        s.NonVisualShapeProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return s.NonVisualShapeProperties?.NonVisualDrawingProperties;
                }
                case Picture p: {
                    if (create) {
                        p.NonVisualPictureProperties ??= new NonVisualPictureProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualPictureDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        p.NonVisualPictureProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return p.NonVisualPictureProperties?.NonVisualDrawingProperties;
                }
                case GraphicFrame g: {
                    if (create) {
                        g.NonVisualGraphicFrameProperties ??= new NonVisualGraphicFrameProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGraphicFrameDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        g.NonVisualGraphicFrameProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                }
                case GroupShape g: {
                    if (create) {
                        g.NonVisualGroupShapeProperties ??= new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                        g.NonVisualGroupShapeProperties.NonVisualDrawingProperties ??= new NonVisualDrawingProperties();
                    }

                    return g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
                }
                default:
                    return null;
            }
        }

        private PlaceholderShape? GetPlaceholderShape() {
            return Element switch {
                Shape s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                Picture p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                _ => null
            };
        }

        private static bool IsMediaPicture(Picture picture) {
            ApplicationNonVisualDrawingProperties? properties =
                picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            return properties?.Descendants<A.AudioFromFile>().Any() == true ||
                   properties?.Descendants<A.VideoFromFile>().Any() == true;
        }

        private ShapeProperties? GetShapeProperties(bool create = false) {
            switch (Element) {
                case Shape s:
                    if (create) {
                        s.ShapeProperties ??= new ShapeProperties();
                    }
                    return s.ShapeProperties;
                case Picture p:
                    if (create) {
                        p.ShapeProperties ??= new ShapeProperties();
                    }
                    return p.ShapeProperties;
                default:
                    return null;
            }
        }

        private static void InsertShapePropertyChild(ShapeProperties properties, OpenXmlElement child) {
            int childOrder = GetShapePropertyChildOrder(child);
            OpenXmlElement? insertBefore = properties.ChildElements
                .FirstOrDefault(existing => GetShapePropertyChildOrder(existing) > childOrder);

            if (insertBefore != null) {
                properties.InsertBefore(child, insertBefore);
            } else {
                properties.Append(child);
            }
        }

        private static int GetShapePropertyChildOrder(OpenXmlElement child) {
            return child switch {
                A.Transform2D => 0,
                A.CustomGeometry => 1,
                A.PresetGeometry => 1,
                A.NoFill => 2,
                A.SolidFill => 2,
                A.GradientFill => 2,
                A.BlipFill => 2,
                A.PatternFill => 2,
                A.GroupFill => 2,
                A.Outline => 3,
                A.EffectList => 4,
                A.EffectDag => 4,
                _ => 100
            };
        }
    }
}
