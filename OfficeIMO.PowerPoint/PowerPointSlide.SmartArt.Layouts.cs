using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        internal static Dgm.LayoutDefinition CreateSmartArtLayoutDefinition(
            PowerPointSmartArtType type, int nodeCount, double aspectRatio) {
            Dgm.LayoutDefinition layout = new() {
                UniqueId = GetSmartArtLayoutId(type)
            };
            layout.AddNamespaceDeclaration("dgm",
                "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            layout.AddNamespaceDeclaration("a",
                "http://schemas.openxmlformats.org/drawingml/2006/main");
            layout.Append(new Dgm.Title { Val = string.Empty });
            layout.Append(new Dgm.Description { Val = string.Empty });
            layout.Append(new Dgm.CategoryList(new Dgm.Category {
                Type = GetSmartArtCategory(type),
                Priority = 400U
            }));
            layout.Append(CreateSmartArtRootLayoutNode(type, nodeCount,
                aspectRatio));
            return layout;
        }

        private static Dgm.LayoutNode CreateSmartArtRootLayoutNode(
            PowerPointSmartArtType type, int nodeCount, double aspectRatio) {
            switch (type) {
                case PowerPointSmartArtType.BasicProcess:
                    return CreateLinearSmartArtLayout();
                case PowerPointSmartArtType.BasicHierarchy:
                    return CreateHierarchySmartArtLayout(nodeCount);
                case PowerPointSmartArtType.BasicCycle:
                    return CreateCycleSmartArtLayout();
                case PowerPointSmartArtType.BasicList:
                    return CreateListSmartArtLayout(nodeCount, aspectRatio);
                case PowerPointSmartArtType.BasicMatrix:
                    return CreateMatrixSmartArtLayout(nodeCount);
                case PowerPointSmartArtType.BasicPyramid:
                    return CreatePyramidSmartArtLayout(nodeCount);
                case PowerPointSmartArtType.BasicRelationship:
                    return CreateRelationshipSmartArtLayout();
                default:
                    throw new System.ArgumentOutOfRangeException(nameof(type), type,
                        "Unsupported SmartArt type.");
            }
        }

        private static Dgm.LayoutNode CreateLinearSmartArtLayout() {
            Dgm.LayoutNode root = CreateSmartArtCanvas("process",
                CreateAlgorithm(Dgm.AlgorithmValues.Linear,
                    (Dgm.ParameterIdValues.LinearDirection, "fromL")));
            root.Append(CreateRootNodeConstraints("node", heightToWidth: 0.6D,
                fontSize: 50D));
            root.Append(new Dgm.RuleList());
            Dgm.ForEach nodes = CreateNodeIterator("processNodes",
                CreateTextLayoutNode("node", "roundRect", square: false));
            nodes.Append(CreateSiblingSpacerIterator("processSpacing"));
            root.Append(nodes);
            return root;
        }

        private static Dgm.LayoutNode CreateListSmartArtLayout(int nodeCount,
            double aspectRatio) {
            Dgm.LayoutNode root = CreateSmartArtCanvas("list",
                CreateAlgorithm(Dgm.AlgorithmValues.Composite,
                    (Dgm.ParameterIdValues.AspectRatio,
                        aspectRatio.ToString("R",
                            CultureInfo.InvariantCulture))));
            Dgm.Constraints constraints = new();
            int count = Math.Max(1, nodeCount);
            double nodeHeight = 0.68D / count;
            for (int index = 0; index < count; index++) {
                string name = $"listNode{index + 1}";
                AppendPositionConstraints(constraints, name,
                    width: 0.94D, height: nodeHeight,
                    centerX: 0.5D, centerY: (index + 0.5D) / count);
            }
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                For = Dgm.ConstraintRelationshipValues.Descendant,
                PointType = Dgm.ElementValues.Node,
                Operator = Dgm.BoolOperatorValues.Equal,
                Val = 42D
            });
            root.Append(constraints);
            root.Append(new Dgm.RuleList());
            for (int index = 1; index <= count; index++) {
                root.Append(CreateIndexedTextLayoutNode($"listNode{index}",
                    "rect", (uint)index, square: false));
            }
            return root;
        }

        private static Dgm.LayoutNode CreateMatrixSmartArtLayout(int nodeCount) {
            Dgm.LayoutNode root = CreateSmartArtCanvas("matrix",
                CreateAlgorithm(Dgm.AlgorithmValues.Composite,
                    (Dgm.ParameterIdValues.AspectRatio, "1")));
            Dgm.Constraints constraints = new();
            int count = Math.Max(1, nodeCount);
            int columns = (int)Math.Ceiling(Math.Sqrt(count));
            int rows = (int)Math.Ceiling(count / (double)columns);
            double size = Math.Min(0.86D / columns, 0.86D / rows) * 0.9D;
            for (int index = 0; index < count; index++) {
                string name = $"matrixNode{index + 1}";
                int column = index % columns;
                int row = index / columns;
                AppendPositionConstraints(constraints, name, size, size,
                    centerX: (column + 0.5D) / columns,
                    centerY: (row + 0.5D) / rows);
            }
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                For = Dgm.ConstraintRelationshipValues.Descendant,
                PointType = Dgm.ElementValues.Node,
                Operator = Dgm.BoolOperatorValues.Equal,
                Val = 65D
            });
            root.Append(constraints);
            root.Append(new Dgm.RuleList());
            for (int index = 1; index <= count; index++) {
                root.Append(CreateIndexedTextLayoutNode($"matrixNode{index}",
                    "roundRect", (uint)index));
            }
            return root;
        }

        private static Dgm.LayoutNode CreateCycleSmartArtLayout() {
            Dgm.LayoutNode root = CreateSmartArtCanvas("cycle",
                CreateAlgorithm(Dgm.AlgorithmValues.Cycle,
                    (Dgm.ParameterIdValues.StartAngle, "-90"),
                    (Dgm.ParameterIdValues.SpanAngle, "360")));
            root.Append(CreateRootNodeConstraints("node", heightToWidth: 1D,
                widthFact: 0.72D, fontSize: 42D));
            root.Append(new Dgm.RuleList());
            Dgm.ForEach nodes = CreateNodeIterator("cycleNodes",
                CreateTextLayoutNode("node", "ellipse", square: true));
            nodes.Append(CreateCycleConnectorIterator());
            root.Append(nodes);
            return root;
        }

        private static Dgm.LayoutNode CreateHierarchySmartArtLayout(int nodeCount) {
            Dgm.LayoutNode root = CreateSmartArtCanvas("hierarchy",
                CreateAlgorithm(Dgm.AlgorithmValues.Composite));
            Dgm.Constraints hierarchyConstraints = new();
            AppendPositionConstraints(hierarchyConstraints, "hierarchyRootNode",
                width: 0.36D, height: 0.24D, centerX: 0.5D, centerY: 0.22D);
            int childCount = Math.Max(0, nodeCount - 1);
            int columns = Math.Max(1, Math.Min(4, childCount));
            int rows = Math.Max(1, (int)Math.Ceiling(childCount / (double)columns));
            double childWidth = 0.88D / columns * 0.88D;
            double childHeight = Math.Min(0.24D, 0.45D / rows * 0.82D);
            for (int index = 0; index < childCount; index++) {
                int column = index % columns;
                int row = index / columns;
                AppendPositionConstraints(hierarchyConstraints,
                    $"hierarchyChild{index + 1}", width: childWidth,
                    height: childHeight,
                    centerX: (column + 0.5D) / columns,
                    centerY: 0.5D + (row + 0.5D) * (0.45D / rows));
            }
            hierarchyConstraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                For = Dgm.ConstraintRelationshipValues.Descendant,
                PointType = Dgm.ElementValues.Node,
                Operator = Dgm.BoolOperatorValues.Equal,
                Val = 30D
            });
            root.Append(hierarchyConstraints);
            root.Append(new Dgm.RuleList());
            root.Append(CreateMappedTextLayoutNode("hierarchyRootNode", "roundRect",
                axis: "ch", pointType: "node",
                start: new[] { 1 }, count: new[] { 1U }));
            for (int index = 1; index <= childCount; index++) {
                root.Append(CreateMappedTextLayoutNode(
                    $"hierarchyChild{index}", "roundRect",
                    axis: "ch ch", pointType: "node node",
                    start: new[] { 1, index },
                    count: new[] { 1U, 1U }));
            }
            return root;
        }

        private static Dgm.LayoutNode CreatePyramidSmartArtLayout(int nodeCount) {
            Dgm.LayoutNode root = CreateSmartArtCanvas("pyramid",
                CreateAlgorithm(Dgm.AlgorithmValues.Composite));
            Dgm.Constraints constraints = new();
            int count = Math.Max(1, nodeCount);
            for (int index = 0; index < count; index++) {
                OfficeDiagramNodeBounds bounds =
                    OfficeDiagramLayoutGeometry.GetPyramidNodeBounds(
                        count, index, 1D, 1D);
                AppendPositionConstraints(constraints, $"level{index + 1}",
                    width: bounds.Width, height: bounds.Height,
                    centerX: bounds.X + bounds.Width / 2D,
                    centerY: bounds.Y + bounds.Height / 2D);
            }
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                For = Dgm.ConstraintRelationshipValues.Descendant,
                PointType = Dgm.ElementValues.Node,
                Operator = Dgm.BoolOperatorValues.Equal,
                Val = 34D
            });
            root.Append(constraints);
            root.Append(new Dgm.RuleList());
            for (int index = 1; index <= count; index++) {
                root.Append(CreateIndexedTextLayoutNode($"level{index}",
                    "trapezoid", (uint)index, square: false));
            }
            return root;
        }

        private static Dgm.LayoutNode CreateRelationshipSmartArtLayout() {
            Dgm.LayoutNode root = CreateSmartArtCanvas("relationship",
                CreateAlgorithm(Dgm.AlgorithmValues.Cycle,
                    (Dgm.ParameterIdValues.StartAngle, "-90"),
                    (Dgm.ParameterIdValues.SpanAngle, "360"),
                    (Dgm.ParameterIdValues.CenterShapeMapping, "fNode")));
            Dgm.Constraints relationshipConstraints = CreateRootNodeConstraints(
                "center", heightToWidth: 1D, fontSize: 38D);
            AppendConstraints(relationshipConstraints, CreateRootNodeConstraints(
                "node", heightToWidth: 1D, widthFact: 0.72D, fontSize: 38D));
            root.Append(relationshipConstraints);
            root.Append(new Dgm.RuleList());

            Dgm.ForEach centerBranch = CreateNodeIterator("relationshipBranch",
                count: 1U);
            centerBranch.Append(CreateTextLayoutNode("center", "ellipse", square: true,
                selfOnly: true));
            Dgm.ForEach surrounding = new() {
                Name = "relationshipChildren",
                Axis = AxisList("ch")
            };
            surrounding.Append(CreateSelfNodeIterator("relationshipNodeSelector",
                CreateTextLayoutNode("node", "ellipse", square: true, selfOnly: true)));
            centerBranch.Append(surrounding);
            root.Append(centerBranch);
            return root;
        }

        private static Dgm.LayoutNode CreateSmartArtCanvas(string name,
            Dgm.Algorithm algorithm) {
            Dgm.LayoutNode root = new() { Name = name };
            root.Append(algorithm);
            root.Append(CreateLayoutShape());
            root.Append(new Dgm.PresentationOf());
            return root;
        }

        private static Dgm.Algorithm CreateAlgorithm(Dgm.AlgorithmValues type,
            params (Dgm.ParameterIdValues Type, string Value)[] parameters) {
            Dgm.Algorithm algorithm = new() { Type = type };
            foreach ((Dgm.ParameterIdValues parameterType, string value) in parameters) {
                algorithm.Append(new Dgm.Parameter {
                    Type = parameterType,
                    Val = value
                });
            }
            return algorithm;
        }

        private static Dgm.Shape CreateLayoutShape(string? shapeType = null) {
            Dgm.Shape shape = new() { Blip = string.Empty };
            if (!string.IsNullOrWhiteSpace(shapeType)) {
                shape.Type = shapeType;
            }
            shape.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            shape.Append(new Dgm.AdjustList());
            return shape;
        }

        private static Dgm.Constraints CreateRootNodeConstraints(string nodeName,
            double heightToWidth, double widthFact = 1D, double fontSize = 65D) {
            Dgm.Constraints constraints = new();
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.Width,
                For = Dgm.ConstraintRelationshipValues.Child,
                ForName = nodeName,
                ReferenceType = Dgm.ConstraintValues.Width,
                Fact = widthFact
            });
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.Height,
                For = Dgm.ConstraintRelationshipValues.Child,
                ForName = nodeName,
                ReferenceType = Dgm.ConstraintValues.Width,
                ReferenceFor = Dgm.ConstraintRelationshipValues.Child,
                ReferenceForName = nodeName,
                Fact = heightToWidth
            });
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                For = Dgm.ConstraintRelationshipValues.Child,
                ForName = nodeName,
                Operator = Dgm.BoolOperatorValues.Equal,
                Val = fontSize
            });
            return constraints;
        }

        private static Dgm.LayoutNode CreateTextLayoutNode(string name,
            string shapeType, bool square, bool selfOnly = false) {
            Dgm.LayoutNode node = new() { Name = name, StyleLabel = "node" };
            node.Append(new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Text });
            node.Append(CreateLayoutShape(shapeType));
            node.Append(new Dgm.PresentationOf {
                Axis = AxisList(selfOnly ? "self" : "desOrSelf"),
                PointType = PointTypeList("node")
            });
            Dgm.Constraints constraints = new();
            if (square) {
                constraints.Append(new Dgm.Constraint {
                    Type = Dgm.ConstraintValues.Height,
                    ReferenceType = Dgm.ConstraintValues.Width
                });
            }
            constraints.Append(CreateMarginConstraint(Dgm.ConstraintValues.LeftMargin));
            constraints.Append(CreateMarginConstraint(Dgm.ConstraintValues.RightMargin));
            constraints.Append(CreateMarginConstraint(Dgm.ConstraintValues.TopMargin));
            constraints.Append(CreateMarginConstraint(Dgm.ConstraintValues.BottomMargin));
            node.Append(constraints);
            Dgm.RuleList rules = new();
            rules.Append(new Dgm.Rule {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                Val = 5D,
                Fact = new DoubleValue { InnerText = "NaN" },
                Max = new DoubleValue { InnerText = "NaN" }
            });
            node.Append(rules);
            return node;
        }

        private static Dgm.LayoutNode CreateIndexedTextLayoutNode(string name,
            string shapeType, uint index, bool square = true) =>
            CreateMappedTextLayoutNode(name, shapeType,
                axis: "ch desOrSelf", pointType: "node node",
                start: new[] { (int)index, 1 }, count: new[] { 1U, 0U },
                square: square);

        private static Dgm.LayoutNode CreateMappedTextLayoutNode(string name,
            string shapeType, string axis, string pointType,
            int[] start, uint[] count, bool square = false) {
            Dgm.LayoutNode node = CreateTextLayoutNode(name, shapeType, square,
                selfOnly: true);
            Dgm.PresentationOf presentation = node.GetFirstChild<Dgm.PresentationOf>()!;
            presentation.Axis = AxisList(axis);
            presentation.PointType = PointTypeList(pointType);
            presentation.Start = IntList(start);
            presentation.Count = UIntList(count);
            return node;
        }

        private static Dgm.Constraint CreateNamedConstraint(
            Dgm.ConstraintValues type, string name,
            Dgm.ConstraintValues referenceType, double fact) => new() {
                Type = type,
                For = Dgm.ConstraintRelationshipValues.Child,
                ForName = name,
                ReferenceType = referenceType,
                Fact = fact
            };

        private static void AppendPositionConstraints(Dgm.Constraints constraints,
            string name, double width, double height, double centerX,
            double centerY) {
            constraints.Append(CreateNamedConstraint(Dgm.ConstraintValues.Width,
                name, Dgm.ConstraintValues.Width, width));
            constraints.Append(CreateNamedConstraint(Dgm.ConstraintValues.Height,
                name, Dgm.ConstraintValues.Height, height));
            constraints.Append(CreateNamedConstraint(Dgm.ConstraintValues.CenterWidth,
                name, Dgm.ConstraintValues.Width, centerX));
            constraints.Append(CreateNamedConstraint(Dgm.ConstraintValues.CenterHeight,
                name, Dgm.ConstraintValues.Height, centerY));
        }

        private static void AppendConstraints(Dgm.Constraints target,
            Dgm.Constraints source) {
            foreach (Dgm.Constraint constraint in source.Elements<Dgm.Constraint>().ToList()) {
                target.Append(constraint.CloneNode(true));
            }
        }

        private static Dgm.ForEach CreateSiblingSpacerIterator(string name) {
            Dgm.ForEach iterator = new() {
                Name = name,
                Axis = AxisList("followSib"),
                PointType = PointTypeList("sibTrans"),
                Count = UIntList(1U)
            };
            Dgm.LayoutNode spacer = CreateSmartArtCanvas("sibTrans",
                CreateAlgorithm(Dgm.AlgorithmValues.Space));
            spacer.Append(new Dgm.Constraints());
            spacer.Append(new Dgm.RuleList());
            iterator.Append(spacer);
            return iterator;
        }

        private static Dgm.ForEach CreateCycleConnectorIterator() {
            Dgm.ForEach iterator = new() {
                Name = "cycleTransitions",
                Axis = AxisList("followSib"),
                PointType = PointTypeList("sibTrans"),
                HideLastTrans = new ListValue<BooleanValue> { InnerText = "0" },
                Count = UIntList(1U)
            };
            Dgm.LayoutNode connector = new() { Name = "cycleConnector" };
            connector.Append(CreateAlgorithm(Dgm.AlgorithmValues.Connector,
                (Dgm.ParameterIdValues.BeginningPoints, "radial"),
                (Dgm.ParameterIdValues.EndPoints, "radial")));
            connector.Append(CreateLayoutShape("conn"));
            connector.Append(new Dgm.PresentationOf { Axis = AxisList("self") });
            connector.Append(new Dgm.Constraints(
                new Dgm.Constraint {
                    Type = Dgm.ConstraintValues.Height,
                    ReferenceType = Dgm.ConstraintValues.Width,
                    Fact = 1.35D
                },
                new Dgm.Constraint {
                    Type = Dgm.ConstraintValues.ConnectionDistance
                }));
            connector.Append(new Dgm.RuleList());
            iterator.Append(connector);
            return iterator;
        }

        private static Dgm.Constraint CreateMarginConstraint(
            Dgm.ConstraintValues type) => new() {
                Type = type,
                ReferenceType = Dgm.ConstraintValues.PrimaryFontSize,
                Fact = 0.2D
            };

        private static Dgm.ForEach CreateNodeIterator(string name,
            Dgm.LayoutNode? child = null, uint? count = null) {
            Dgm.ForEach iterator = new() {
                Name = name,
                Axis = AxisList("ch"),
                PointType = PointTypeList("node")
            };
            if (count.HasValue) {
                iterator.Count = new ListValue<UInt32Value> {
                    InnerText = count.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                };
            }
            if (child != null) {
                iterator.Append(child);
            }
            return iterator;
        }

        private static Dgm.ForEach CreateSelfNodeIterator(string name,
            Dgm.LayoutNode child) {
            Dgm.ForEach iterator = new() {
                Name = name,
                Axis = AxisList("self"),
                PointType = PointTypeList("node")
            };
            iterator.Append(child);
            return iterator;
        }

        private static ListValue<EnumValue<Dgm.AxisValues>> AxisList(string value) =>
            new() { InnerText = value };

        private static ListValue<EnumValue<Dgm.ElementValues>> PointTypeList(string value) =>
            new() { InnerText = value };

        private static ListValue<UInt32Value> UIntList(params uint[] values) =>
            new() {
                InnerText = string.Join(" ", values.Select(value =>
                    value.ToString(System.Globalization.CultureInfo.InvariantCulture)))
            };

        private static ListValue<Int32Value> IntList(params int[] values) =>
            new() {
                InnerText = string.Join(" ", values.Select(value =>
                    value.ToString(System.Globalization.CultureInfo.InvariantCulture)))
            };
    }
}
