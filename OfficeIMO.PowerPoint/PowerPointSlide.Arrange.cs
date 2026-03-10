using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private static readonly MethodInfo AddNewPartWithContentTypeMethod =
            typeof(OpenXmlPartContainer)
                .GetMethods()
                .Single(m => m.Name == "AddNewPart" &&
                             m.IsGenericMethodDefinition &&
                             m.GetParameters().Length == 2);

        /// <summary>
        ///     Creates a duplicate of the provided shape and optionally offsets its position (EMUs).
        /// </summary>
        public PowerPointShape DuplicateShape(PowerPointShape shape, long offsetX = 0L, long offsetY = 0L) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            OpenXmlElement clone = shape.Element.CloneNode(true);
            UpdateNonVisualDrawingProperties(clone, GetDuplicateBaseName(shape));
            if (shape is PowerPointChart && clone is GraphicFrame duplicatedChartFrame) {
                RebindDuplicatedChart(duplicatedChartFrame);
            }

            PowerPointShape? duplicate = CreateShapeFromElement(clone);
            if (duplicate == null) {
                throw new InvalidOperationException("Duplicated shape type is not supported.");
            }

            if (offsetX != 0 || offsetY != 0) {
                duplicate.Left += offsetX;
                duplicate.Top += offsetY;
            }

            parent.InsertAfter(clone, shape.Element);
            _shapes.Insert(index + 1, duplicate);
            return duplicate;
        }

        /// <summary>
        ///     Creates a duplicate of the provided shape and offsets its position (centimeters).
        /// </summary>
        public PowerPointShape DuplicateShapeCm(PowerPointShape shape, double offsetXCm, double offsetYCm) {
            return DuplicateShape(shape,
                PowerPointUnits.FromCentimeters(offsetXCm),
                PowerPointUnits.FromCentimeters(offsetYCm));
        }

        /// <summary>
        ///     Creates a duplicate of the provided shape and offsets its position (inches).
        /// </summary>
        public PowerPointShape DuplicateShapeInches(PowerPointShape shape, double offsetXInches, double offsetYInches) {
            return DuplicateShape(shape,
                PowerPointUnits.FromInches(offsetXInches),
                PowerPointUnits.FromInches(offsetYInches));
        }

        /// <summary>
        ///     Creates a duplicate of the provided shape and offsets its position (points).
        /// </summary>
        public PowerPointShape DuplicateShapePoints(PowerPointShape shape, double offsetXPoints, double offsetYPoints) {
            return DuplicateShape(shape,
                PowerPointUnits.FromPoints(offsetXPoints),
                PowerPointUnits.FromPoints(offsetYPoints));
        }

        /// <summary>
        ///     Moves the shape one step forward in z-order.
        /// </summary>
        public void BringForward(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index >= _shapes.Count - 1) {
                return;
            }

            PowerPointShape next = _shapes[index + 1];
            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            shape.Element.Remove();
            parent.InsertAfter(shape.Element, next.Element);

            _shapes[index] = next;
            _shapes[index + 1] = shape;
        }

        /// <summary>
        ///     Moves the shape one step backward in z-order.
        /// </summary>
        public void SendBackward(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index <= 0) {
                return;
            }

            PowerPointShape previous = _shapes[index - 1];
            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            shape.Element.Remove();
            parent.InsertBefore(shape.Element, previous.Element);

            _shapes[index] = previous;
            _shapes[index - 1] = shape;
        }

        /// <summary>
        ///     Moves the shape to the front (top) of the z-order.
        /// </summary>
        public void BringToFront(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index >= _shapes.Count - 1) {
                return;
            }

            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");

            shape.Element.Remove();
            parent.Append(shape.Element);

            _shapes.RemoveAt(index);
            _shapes.Add(shape);
        }

        /// <summary>
        ///     Moves the shape to the back (bottom) of the z-order.
        /// </summary>
        public void SendToBack(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            int index = EnsureShapeOnSlide(shape);
            if (index <= 0) {
                return;
            }

            OpenXmlElement parent = shape.Element.Parent ?? throw new InvalidOperationException("Shape is not attached to a slide.");
            OpenXmlElement insertBefore = _shapes[0].Element;

            shape.Element.Remove();
            parent.InsertBefore(shape.Element, insertBefore);

            _shapes.RemoveAt(index);
            _shapes.Insert(0, shape);
        }

        private int EnsureShapeOnSlide(PowerPointShape shape) {
            int index = _shapes.IndexOf(shape);
            if (index < 0) {
                throw new ArgumentException("Shape does not belong to this slide.", nameof(shape));
            }

            return index;
        }

        private string GetDuplicateBaseName(PowerPointShape shape) {
            switch (shape) {
                case PowerPointAutoShape auto when auto.ShapeType != null:
                    return auto.ShapeType.Value.ToString();
                case PowerPointTextBox textBox when textBox.PlaceholderType == PlaceholderValues.Title:
                    return "Title";
                case PowerPointTextBox:
                    return "TextBox";
                case PowerPointPicture:
                    return "Picture";
                case PowerPointTable:
                    return "Table";
                case PowerPointChart:
                    return "Chart";
                default:
                    return shape.Name ?? "Shape";
            }
        }

        private void UpdateNonVisualDrawingProperties(OpenXmlElement element, string baseName) {
            ApplyUniqueNonVisualDrawingProperties(element, baseName);

            if (element is GroupShape groupShape) {
                foreach (OpenXmlElement child in groupShape.ChildElements) {
                    UpdateDescendantNonVisualDrawingProperties(child);
                }
            }
        }

        private void UpdateDescendantNonVisualDrawingProperties(OpenXmlElement element) {
            string baseName = GetElementBaseName(element);
            ApplyUniqueNonVisualDrawingProperties(element, baseName);

            if (element is GroupShape groupShape) {
                foreach (OpenXmlElement child in groupShape.ChildElements) {
                    UpdateDescendantNonVisualDrawingProperties(child);
                }
            }
        }

        private void ApplyUniqueNonVisualDrawingProperties(OpenXmlElement element, string baseName) {
            string resolvedBaseName = string.IsNullOrWhiteSpace(baseName) ? "Shape" : baseName;
            string name = GenerateUniqueName(resolvedBaseName);
            uint id = _nextShapeId++;

            switch (element) {
                case Shape s: {
                    NonVisualShapeProperties nonVisual = s.NonVisualShapeProperties ??=
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??=
                        new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
                case Picture p: {
                    NonVisualPictureProperties nonVisual = p.NonVisualPictureProperties ??=
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualPictureDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??=
                        new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
                case GraphicFrame g: {
                    NonVisualGraphicFrameProperties nonVisual = g.NonVisualGraphicFrameProperties ??=
                        new NonVisualGraphicFrameProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGraphicFrameDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??=
                        new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
                case GroupShape g: {
                    NonVisualGroupShapeProperties nonVisual = g.NonVisualGroupShapeProperties ??=
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties(),
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties());
                    NonVisualDrawingProperties drawing = nonVisual.NonVisualDrawingProperties ??=
                        new NonVisualDrawingProperties();
                    drawing.Id = id;
                    drawing.Name = name;
                    break;
                }
            }
        }

        private static string GetElementBaseName(OpenXmlElement element) {
            return element switch {
                Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Shape",
                Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture",
                GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null =>
                    g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table",
                GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null =>
                    g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Chart",
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "GraphicFrame",
                GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group",
                _ => "Shape"
            };
        }

        private PowerPointShape? CreateShapeFromElement(OpenXmlElement element) {
            switch (element) {
                case Shape s:
                    return s.TextBody != null ? new PowerPointTextBox(s, _slidePart) : new PowerPointAutoShape(s);
                case Picture p:
                    return new PowerPointPicture(p, _slidePart);
                case GroupShape g:
                    return new PowerPointGroupShape(g);
                case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null:
                    return new PowerPointTable(g, _slidePart);
                case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null:
                    return new PowerPointChart(g, _slidePart);
                default:
                    return null;
            }
        }

        private void RebindDuplicatedChart(GraphicFrame frame) {
            C.ChartReference? chartReference = frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>();
            string originalRelationshipId = chartReference?.Id?.Value
                ?? throw new InvalidOperationException("Chart reference not found for duplicated shape.");
            ChartPart sourceChartPart = (ChartPart)_slidePart.GetPartById(originalRelationshipId);
            chartReference!.Id = CloneChartPart(sourceChartPart);
        }

        private string CloneChartPart(ChartPart sourceChartPart) {
            string relationshipId = GetNextRelationshipId(_slidePart);
            ChartPart targetChartPart = (ChartPart)AddNewPartWithContentType(_slidePart, sourceChartPart, relationshipId);
            if (sourceChartPart.ChartSpace != null) {
                targetChartPart.ChartSpace = (C.ChartSpace)sourceChartPart.ChartSpace.CloneNode(true);
                targetChartPart.ChartSpace.Save();
            } else {
                CopyPartData(sourceChartPart, targetChartPart);
            }
            CloneReferenceRelationships(sourceChartPart, targetChartPart);

            foreach (IdPartPair childPair in sourceChartPart.Parts) {
                ClonePartRecursive(childPair.OpenXmlPart, targetChartPart, childPair.RelationshipId);
            }

            return relationshipId;
        }

        private static void ClonePartRecursive(OpenXmlPart sourcePart, OpenXmlPartContainer targetContainer, string relationshipId) {
            OpenXmlPart newPart = sourcePart is ExtendedPart extendedPart
                ? targetContainer.AddExtendedPart(extendedPart.RelationshipType, extendedPart.ContentType, relationshipId)
                : AddNewPartWithContentType(targetContainer, sourcePart, relationshipId);

            CopyPartData(sourcePart, newPart);
            CloneReferenceRelationships(sourcePart, newPart);

            foreach (IdPartPair childPair in sourcePart.Parts) {
                ClonePartRecursive(childPair.OpenXmlPart, newPart, childPair.RelationshipId);
            }
        }

        private static OpenXmlPart AddNewPartWithContentType(OpenXmlPartContainer container, OpenXmlPart sourcePart, string relationshipId) {
            MethodInfo method = AddNewPartWithContentTypeMethod.MakeGenericMethod(sourcePart.GetType());
            return (OpenXmlPart)method.Invoke(container, new object[] { sourcePart.ContentType, relationshipId })!;
        }

        private static void CopyPartData(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            using Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
            using Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        private static void CloneReferenceRelationships(OpenXmlPartContainer source, OpenXmlPartContainer target) {
            foreach (ExternalRelationship rel in source.ExternalRelationships) {
                target.AddExternalRelationship(rel.RelationshipType, rel.Uri, rel.Id);
            }

            foreach (HyperlinkRelationship rel in source.HyperlinkRelationships) {
                target.AddHyperlinkRelationship(rel.Uri, rel.IsExternal, rel.Id);
            }
        }

        private static string GetNextRelationshipId(OpenXmlPartContainer container) {
            var existingRelationships = container.Parts.Select(part => part.RelationshipId)
                .Concat(container.ExternalRelationships.Select(rel => rel.Id))
                .Concat(container.HyperlinkRelationships.Select(rel => rel.Id))
                .Where(id => !string.IsNullOrEmpty(id))
                .ToHashSet(StringComparer.Ordinal);

            int nextId = 1;
            string relationshipId;
            do {
                relationshipId = "rId" + nextId;
                nextId++;
            } while (!existingRelationships.Add(relationshipId));

            return relationshipId;
        }
    }
}
