using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using Dsp = DocumentFormat.OpenXml.Office.Drawing;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Custom SmartArt template based on exported SmartArt2 (cycle2 layout).
    internal static class SmartArtCustom2 {
        internal static void PopulateColors(DiagramColorsPart part) {
            var colors = new Dgm.ColorsDefinition {
                UniqueId = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"
            };
            colors.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            colors.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            colors.Append(new Dgm.ColorDefinitionTitle { Val = "" });
            colors.Append(new Dgm.ColorTransformDescription { Val = "" });

            var catLst = new Dgm.ColorTransformCategories();
            catLst.Append(new Dgm.ColorTransformCategory { Type = "accent1", Priority = (UInt32Value)11200U });
            colors.Append(catLst);

            // Core labels used by nodes/connectors
            colors.Append(MakeStyleLabel("node0", accentFill: true, lineAccent: false));
            colors.Append(MakeStyleLabel("lnNode1", accentFill: false, lineAccent: true));
            colors.Append(MakeStyleLabel("alignNode1", accentFill: true, lineAccent: false));
            colors.Append(MakeStyleLabel("node1", accentFill: true, lineAccent: false));
            // sibTrans2D1: accent1 with tint for both fill and line
            colors.Append(MakeStyleLabelTint("sibTrans2D1", 60000));
            // Additional connector style used by exports
            colors.Append(MakeStyleLabel("sibTrans1D1", accentFill: true, lineAccent: true));

            part.ColorsDefinition = colors;
        }

        internal static void PopulateStyle(DiagramStylePart part) {
            var style = new Dgm.StyleDefinition { UniqueId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" };
            style.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            style.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            style.Append(new Dgm.StyleDefinitionTitle { Val = "" });
            style.Append(new Dgm.StyleLabelDescription { Val = "" });
            var displayCats = new Dgm.StyleDisplayCategories();
            displayCats.Append(new Dgm.StyleDisplayCategory { Type = "simple", Priority = (UInt32Value)10100U });
            style.Append(displayCats);

            // Primary style labels (parity with exports)
            style.Append(MakeStyleLabel("node0"));
            style.Append(MakeStyleLabel("lnNode1"));
            style.Append(MakeStyleLabel("alignNode1"));
            style.Append(MakeStyleLabel("node1"));
            style.Append(MakeStyleLabel("sibTrans2D1"));
            style.Append(MakeStyleLabel("sibTrans1D1"));

            part.StyleDefinition = style;
        }

        private static Dgm.ColorTransformStyleLabel MakeStyleLabel(string name, bool accentFill, bool lineAccent) {
            var lbl = new Dgm.ColorTransformStyleLabel { Name = name };
            var fill = new Dgm.FillColorList { Method = Dgm.ColorApplicationMethodValues.Repeat };
            fill.Append(new A.SchemeColor { Val = accentFill ? A.SchemeColorValues.Accent1 : A.SchemeColorValues.Light1 });
            var line = new Dgm.LineColorList { Method = Dgm.ColorApplicationMethodValues.Repeat };
            line.Append(new A.SchemeColor { Val = lineAccent ? A.SchemeColorValues.Accent1 : A.SchemeColorValues.Light1 });
            lbl.Append(fill);
            lbl.Append(line);
            lbl.Append(new Dgm.EffectColorList());
            lbl.Append(new Dgm.TextLineColorList());
            lbl.Append(new Dgm.TextFillColorList());
            lbl.Append(new Dgm.TextEffectColorList());
            return lbl;
        }

        private static Dgm.ColorTransformStyleLabel MakeStyleLabelTint(string name, int tint) {
            var lbl = new Dgm.ColorTransformStyleLabel { Name = name };
            var fill = new Dgm.FillColorList { Method = Dgm.ColorApplicationMethodValues.Repeat };
            var scFill = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
            scFill.Append(new A.Tint { Val = tint });
            fill.Append(scFill);
            var line = new Dgm.LineColorList { Method = Dgm.ColorApplicationMethodValues.Repeat };
            var scLine = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
            scLine.Append(new A.Tint { Val = tint });
            line.Append(scLine);
            lbl.Append(fill);
            lbl.Append(line);
            lbl.Append(new Dgm.EffectColorList());
            lbl.Append(new Dgm.TextLineColorList());
            lbl.Append(new Dgm.TextFillColorList());
            lbl.Append(new Dgm.TextEffectColorList());
            return lbl;
        }

        private static Dgm.StyleLabel MakeStyleLabel(string name) {
            var label = new Dgm.StyleLabel { Name = name };
            var dgmStyle = new Dgm.Style();
            var lnRef = new A.LineReference { Index = 2U };
            lnRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            var fillRef = new A.FillReference { Index = 1U };
            fillRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            var effRef = new A.EffectReference { Index = 0U };
            effRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
            var fontRef = new A.FontReference { Index = A.FontCollectionIndexValues.Minor };
            fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Light1 });
            dgmStyle.Append(lnRef);
            dgmStyle.Append(fillRef);
            dgmStyle.Append(effRef);
            dgmStyle.Append(fontRef);
            label.Append(dgmStyle);
            return label;
        }

        internal static void PopulateLayout(DiagramLayoutDefinitionPart part) {
            // Derived from Assets/WordTemplates/SmartArt2.cs (GenerateDiagramLayoutDefinitionPart1Content)
            // Simplified to essential structure for cycle diagrams
            var layout = new Dgm.LayoutDefinition { UniqueId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" };
            layout.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            layout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            layout.Append(new Dgm.Title { Val = "" });
            layout.Append(new Dgm.Description { Val = "" });

            var cats = new Dgm.CategoryList();
            cats.Append(new Dgm.Category { Type = "cycle", Priority = (UInt32Value)1000U });
            cats.Append(new Dgm.Category { Type = "convert", Priority = (UInt32Value)10000U });
            layout.Append(cats);

            var layoutNode = new Dgm.LayoutNode { Name = "cycle" };

            // Variables and direction
            var varList = new Dgm.VariableList();
            varList.Append(new Dgm.Direction());
            varList.Append(new Dgm.ResizeHandles { Val = Dgm.ResizeHandlesStringValues.Exact });
            layoutNode.Append(varList);

            // Choose based on dir and number of children
            var choose = new Dgm.Choose { Name = "Name0" };
            var ifDirNorm = new Dgm.DiagramChooseIf {
                Name = "Name1",
                Function = Dgm.FunctionValues.Variable,
                Argument = "dir",
                Operator = Dgm.FunctionOperatorValues.Equal,
                Val = "norm"
            };
            var choose2 = new Dgm.Choose { Name = "Name2" };
            var ifCountGt2 = new Dgm.DiagramChooseIf {
                Name = "Name3",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" },
                Function = Dgm.FunctionValues.Count,
                Operator = Dgm.FunctionOperatorValues.GreaterThan,
                Val = "2"
            };
            var algCycle1 = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Cycle };
            algCycle1.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.StartAngle, Val = "0" });
            algCycle1.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.SpanAngle, Val = "360" });
            ifCountGt2.Append(algCycle1);
            var elseLessEq2 = new Dgm.DiagramChooseElse { Name = "Name4" };
            var algCycle2 = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Cycle };
            algCycle2.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.StartAngle, Val = "-90" });
            algCycle2.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.SpanAngle, Val = "360" });
            elseLessEq2.Append(algCycle2);
            choose2.Append(ifCountGt2);
            choose2.Append(elseLessEq2);
            ifDirNorm.Append(choose2);

            var elseDir = new Dgm.DiagramChooseElse { Name = "Name5" };
            var choose3 = new Dgm.Choose { Name = "Name6" };
            var ifCountGt2b = new Dgm.DiagramChooseIf {
                Name = "Name7",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" },
                Function = Dgm.FunctionValues.Count,
                Operator = Dgm.FunctionOperatorValues.GreaterThan,
                Val = "2"
            };
            var algCycle3 = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Cycle };
            algCycle3.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.StartAngle, Val = "0" });
            algCycle3.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.SpanAngle, Val = "-360" });
            ifCountGt2b.Append(algCycle3);
            var elseLessEq2b = new Dgm.DiagramChooseElse { Name = "Name8" };
            var algCycle4 = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Cycle };
            algCycle4.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.StartAngle, Val = "90" });
            algCycle4.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.SpanAngle, Val = "-360" });
            elseLessEq2b.Append(algCycle4);
            choose3.Append(ifCountGt2b);
            choose3.Append(elseLessEq2b);
            elseDir.Append(choose3);
            choose.Append(ifDirNorm);
            choose.Append(elseDir);
            layoutNode.Append(choose);

            // Base diagram shape
            var shape = new Dgm.Shape { Blip = "" };
            shape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            shape.Append(new Dgm.AdjustList());
            layoutNode.Append(shape);

            // Layout algorithm: arrange children on a circle
            var cycleAlg = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Cycle };
            cycleAlg.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.StartAngle, Val = "-90" });
            cycleAlg.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.SpanAngle, Val = "360" });
            layoutNode.Append(cycleAlg);

            // Map presentation to descendants-or-self nodes (ensures algorithm applies to node points)
            var presentation = new Dgm.PresentationOf {
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "desOrSelf" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" }
            };
            layoutNode.Append(presentation);

            // Basic constraints to spread siblings and size nodes reasonably
            var constraints = new Dgm.Constraints();
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.Width,
                For = Dgm.ConstraintRelationshipValues.Child,
                PointType = Dgm.ElementValues.Node,
                ReferenceType = Dgm.ConstraintValues.Width
            });
            constraints.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.SiblingSpacing,
                ReferenceType = Dgm.ConstraintValues.Width,
                ReferenceFor = Dgm.ConstraintRelationshipValues.Child,
                ReferencePointType = Dgm.ElementValues.Node,
                Fact = 0.5D
            });
            layoutNode.Append(constraints);

            // For each child node render an ellipse (cycle step)
            var forEach = new Dgm.ForEach {
                Name = "nodesForEach",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" }
            };
            var node = new Dgm.LayoutNode { Name = "node" };
            var nodeVarList = new Dgm.VariableList();
            nodeVarList.Append(new Dgm.BulletEnabled { Val = true });
            node.Append(nodeVarList);
            var txtAlg = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Text };
            txtAlg.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.TextAnchorVerticalWithChildren, Val = "mid" });
            node.Append(txtAlg);
            var nodeShape = new Dgm.Shape { Type = "ellipse", Blip = "" };
            nodeShape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nodeShape.Append(new Dgm.AdjustList());
            node.Append(nodeShape);
            node.Append(new Dgm.PresentationOf { Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "desOrSelf" }, PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" } });
            var nodeConstr = new Dgm.Constraints();
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.Height, ReferenceType = Dgm.ConstraintValues.Width });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.LeftMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.1D });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.RightMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.1D });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.TopMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.1D });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.BottomMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.1D });
            node.Append(nodeConstr);
            var nodeRules = new Dgm.RuleList();
            nodeRules.Append(new Dgm.Rule { Type = Dgm.ConstraintValues.PrimaryFontSize, Val = 5D, Fact = new DoubleValue { InnerText = "NaN" }, Max = new DoubleValue { InnerText = "NaN" } });
            node.Append(nodeRules);
            forEach.Append(node);

            // Optional connectors when more than one node
            var chooseConn = new Dgm.Choose { Name = "Name9" };
            var ifHasSib = new Dgm.DiagramChooseIf {
                Name = "Name10",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "par ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "doc node" },
                Function = Dgm.FunctionValues.Count,
                Operator = Dgm.FunctionOperatorValues.GreaterThan,
                Val = "1"
            };
            var forEachSib = new Dgm.ForEach {
                Name = "sibTransForEach",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "followSib" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "sibTrans" },
                HideLastTrans = new ListValue<BooleanValue> { InnerText = "0" },
                Count = new ListValue<UInt32Value> { InnerText = "1" }
            };
            var sibNode = new Dgm.LayoutNode { Name = "sibTrans" };
            var chooseConnAlg = new Dgm.Choose { Name = "Name11" };
            var ifFew = new Dgm.DiagramChooseIf {
                Name = "Name12",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "par ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "doc node" },
                Function = Dgm.FunctionValues.Count,
                Operator = Dgm.FunctionOperatorValues.LessThan,
                Val = "3"
            };
            var connAlg1 = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Connector };
            connAlg1.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.BeginningPoints, Val = "radial" });
            connAlg1.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.EndPoints, Val = "radial" });
            ifFew.Append(connAlg1);
            var elseMany = new Dgm.DiagramChooseElse { Name = "Name13" };
            var connAlg2 = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Connector };
            connAlg2.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.BeginningPoints, Val = "auto" });
            connAlg2.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.EndPoints, Val = "auto" });
            elseMany.Append(connAlg2);
            chooseConnAlg.Append(ifFew);
            chooseConnAlg.Append(elseMany);
            sibNode.Append(chooseConnAlg);
            var sibShape = new Dgm.Shape { Type = "conn", Blip = "" };
            sibShape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            sibShape.Append(new Dgm.AdjustList());
            sibNode.Append(sibShape);
            sibNode.Append(new Dgm.PresentationOf { Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "self" } });
            var sibConstr = new Dgm.Constraints();
            sibConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.Height, ReferenceType = Dgm.ConstraintValues.Width, Fact = 1.35D });
            sibConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.ConnectionDistance });
            sibConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.Width, For = Dgm.ConstraintRelationshipValues.Child, ReferenceType = Dgm.ConstraintValues.ConnectionDistance, Fact = 0.45D });
            sibConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.Height, For = Dgm.ConstraintRelationshipValues.Child, ReferenceType = Dgm.ConstraintValues.Height });
            sibNode.Append(sibConstr);
            sibNode.Append(new Dgm.RuleList());

            // Connector text as hidden geometry
            var connText = new Dgm.LayoutNode { Name = "connectorText" };
            var connTextAlg = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Text };
            connTextAlg.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.AutoTextRotation, Val = "grav" });
            connText.Append(connTextAlg);
            var connTextShape = new Dgm.Shape { Type = "conn", Blip = "", HideGeometry = true };
            connTextShape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            connTextShape.Append(new Dgm.AdjustList());
            connText.Append(connTextShape);
            connText.Append(new Dgm.PresentationOf { Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "self" } });
            connText.Append(new Dgm.Constraints());
            var connTextRules = new Dgm.RuleList();
            connTextRules.Append(new Dgm.Rule { Type = Dgm.ConstraintValues.PrimaryFontSize, Val = 5D, Fact = new DoubleValue { InnerText = "NaN" }, Max = new DoubleValue { InnerText = "NaN" } });
            connText.Append(connTextRules);

            sibNode.Append(connText);
            forEachSib.Append(sibNode);
            ifHasSib.Append(forEachSib);
            chooseConn.Append(ifHasSib);
            chooseConn.Append(new Dgm.DiagramChooseElse { Name = "Name14" });
            forEach.Append(chooseConn);
            layoutNode.Append(forEach);

            layout.Append(layoutNode);
            part.LayoutDefinition = layout;
        }

        internal static void PopulateData(DiagramDataPart part, string? persistRelId = null) {
            // Ported subset (typed) from exported SmartArt2 to ensure mapping of dsp shapes
            var root = new Dgm.DataModelRoot();
            root.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            root.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var pts = new Dgm.PointList();

            // Base IDs from export (cycle2)
            const string docId = "{111C4888-DF27-4BEC-921D-264911430FFA}";
            const string child1Id = "{8321044F-28D8-416B-9551-074629DA1075}";
            const string parTrans1Id = "{283A91E5-E0E5-4FA4-9127-D0DCB6269F98}";
            const string sibTrans1Id = "{1ED01F69-F93F-4ED4-8596-657DE1C20CBC}";
            const string child2Id = "{D086731A-8F37-40E1-A750-469EB6E755AB}";
            const string cyclePresId = "{4C78BDB7-E2AB-4399-87CD-307C05A02F63}";
            const string node1PresId = "{EA6A36C8-8C9D-4EEB-8E22-E155EC714A8F}";
            const string arrow1PresId = "{0BA95F28-1B53-4967-BA10-28214B883828}";
            const string node2PresId = "{6E91FA55-2036-4528-BC0C-3D31BEA04727}";
            const string cxnId1 = "{6765E681-F519-4228-B58B-ADE08AF84F0D}";
            const string cxnId2 = "{889A3B69-20DF-479A-A1A1-7FBC7DE84A44}";
            const string cxnId3 = "{9CCB7C53-548A-45DF-AA34-4E5103151A5F}";
            const string cxnId4 = "{1B618C73-AA93-4069-9C37-DBF7E03AFEC6}";
            const string cxnId5 = "{F45CA66E-A1DE-4D62-BFDB-14CFF93A4661}";

            // Document
            var doc = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
            doc.Append(new Dgm.PropertySet {
                LayoutTypeId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2",
                LayoutCategoryId = "cycle",
                QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                QuickStyleCategoryId = "simple",
                ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
                ColorCategoryId = "accent1",
                Placeholder = false
            });
            doc.Append(new Dgm.ShapeProperties());
            var docText = new Dgm.TextBody(); docText.Append(new A.BodyProperties()); docText.Append(new A.ListStyle());
            var docPara = new A.Paragraph(); docPara.Append(new A.EndParagraphRunProperties { Language = "en-US" }); docText.Append(docPara);
            doc.Append(docText);
            pts.Append(doc);

            // Child 1
            var child1 = new Dgm.Point { ModelId = child1Id };
            child1.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" });
            child1.Append(new Dgm.ShapeProperties());
            var c1tb = new Dgm.TextBody(); c1tb.Append(new A.BodyProperties()); c1tb.Append(new A.ListStyle());
            var c1p = new A.Paragraph(); c1p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); c1tb.Append(c1p);
            child1.Append(c1tb);
            pts.Append(child1);

            // Parent/Sibling transition 1
            var parTrans1 = new Dgm.Point { ModelId = parTrans1Id, Type = Dgm.PointValues.ParentTransition, ConnectionId = cxnId1 };
            parTrans1.Append(new Dgm.PropertySet()); parTrans1.Append(new Dgm.ShapeProperties());
            var st1 = new Dgm.TextBody(); st1.Append(new A.BodyProperties()); st1.Append(new A.ListStyle()); var st1p = new A.Paragraph(); st1p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); st1.Append(st1p); parTrans1.Append(st1);
            pts.Append(parTrans1);

            var sibTrans1 = new Dgm.Point { ModelId = sibTrans1Id, Type = Dgm.PointValues.SiblingTransition, ConnectionId = cxnId1 };
            sibTrans1.Append(new Dgm.PropertySet()); sibTrans1.Append(new Dgm.ShapeProperties());
            var sb1 = new Dgm.TextBody(); sb1.Append(new A.BodyProperties()); sb1.Append(new A.ListStyle()); var sb1p = new A.Paragraph(); sb1p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sb1.Append(sb1p); sibTrans1.Append(sb1);
            pts.Append(sibTrans1);

            // Child 2
            var child2 = new Dgm.Point { ModelId = child2Id };
            child2.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" });
            child2.Append(new Dgm.ShapeProperties());
            var c2tb = new Dgm.TextBody(); c2tb.Append(new A.BodyProperties()); c2tb.Append(new A.ListStyle()); var c2p = new A.Paragraph(); c2p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); c2tb.Append(c2p); child2.Append(c2tb);
            pts.Append(child2);

            // Presentation mapping for cycle + nodes + arrows
            var presCycle = new Dgm.Point { ModelId = cyclePresId, Type = Dgm.PointValues.Presentation };
            var presCycleProps = new Dgm.PropertySet { PresentationElementId = docId, PresentationName = "cycle", PresentationStyleCount = 0 };
            var presVars = new Dgm.PresentationLayoutVariables(); presVars.Append(new Dgm.Direction()); presVars.Append(new Dgm.ResizeHandles { Val = Dgm.ResizeHandlesStringValues.Exact }); presCycleProps.Append(presVars);
            presCycle.Append(presCycleProps); presCycle.Append(new Dgm.ShapeProperties()); pts.Append(presCycle);

            var presNode1 = new Dgm.Point { ModelId = node1PresId, Type = Dgm.PointValues.Presentation };
            var presNode1Props = new Dgm.PropertySet { PresentationElementId = child1Id, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 0, PresentationStyleCount = 5 };
            var presNode1Vars = new Dgm.PresentationLayoutVariables(); presNode1Vars.Append(new Dgm.BulletEnabled { Val = true }); presNode1Props.Append(presNode1Vars);
            presNode1.Append(presNode1Props); presNode1.Append(new Dgm.ShapeProperties()); pts.Append(presNode1);

            var presArrow1 = new Dgm.Point { ModelId = arrow1PresId, Type = Dgm.PointValues.Presentation };
            var presArrow1Props = new Dgm.PropertySet { PresentationElementId = sibTrans1Id, PresentationName = "sibTrans", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 0, PresentationStyleCount = 5 };
            presArrow1.Append(presArrow1Props); presArrow1.Append(new Dgm.ShapeProperties()); pts.Append(presArrow1);
            // connectorText presentation for arrow 1 (hidden text geometry)
            const string connText1PresId = "{3D6D24F6-214A-4FEB-9D03-50D0FE4766EF}";
            var presConnText1 = new Dgm.Point { ModelId = connText1PresId, Type = Dgm.PointValues.Presentation };
            var presConnText1Props = new Dgm.PropertySet { PresentationElementId = sibTrans1Id, PresentationName = "connectorText", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 0, PresentationStyleCount = 5 };
            presConnText1.Append(presConnText1Props); presConnText1.Append(new Dgm.ShapeProperties()); pts.Append(presConnText1);

            var presNode2 = new Dgm.Point { ModelId = node2PresId, Type = Dgm.PointValues.Presentation };
            var presNode2Props = new Dgm.PropertySet { PresentationElementId = child2Id, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 1, PresentationStyleCount = 5 };
            var presNode2Vars = new Dgm.PresentationLayoutVariables(); presNode2Vars.Append(new Dgm.BulletEnabled { Val = true }); presNode2Props.Append(presNode2Vars);
            presNode2.Append(presNode2Props); presNode2.Append(new Dgm.ShapeProperties()); pts.Append(presNode2);

            // More children, transitions, and presentation mapping to match export
            const string child3Id = "{54E8997A-E156-4EFC-A886-153C90898B31}";
            const string parTrans2Id = "{DC6E593A-33F0-4BD6-A8FB-BAC9E25F868C}";
            const string sibTrans2Id = "{E7437055-F9C9-4D7E-B9F6-EA2ACB17BF27}";
            const string parTrans3Id = "{BF1C22EB-DC9B-4784-B281-99943CF26C7E}";
            const string sibTrans3Id = "{F269153A-E28E-4352-8D97-7B7399E00F46}";
            const string child4Id = "{8FA3C8D0-F752-4AA8-902C-C4215EBDFA48}";
            const string parTrans4Id = "{2050BFF9-4245-4CED-BB9D-DA36277734E0}";
            const string sibTrans4Id = "{D9E2FE70-2497-4DB9-A0B4-55726D9F891C}";
            const string child5Id = "{FF546BAC-A85C-4761-94A3-6D990039FE05}";
            const string parTrans5Id = "{FF546BAC-A85C-4761-94A3-6D990039FE05}"; // parTrans5 has different id in export; mapping uses cxn below
            const string sibTrans5Id = "{94FF4399-6442-4717-AF0C-2B3F625142C2}";

            const string node3PresId = "{AFFFFF7B-5085-434A-A3EC-39CA3570E329}";
            const string arrow3PresId = "{1C14F7F4-DFC4-4BCC-9211-35A893821A37}";
            const string node4PresId = "{AB9582F8-24C6-4DAA-8F90-5AD5EDEB4E2E}";
            const string arrow4PresId = "{7FA91D88-E855-4818-B461-6D97AC4D5A28}";
            const string node5PresId = "{CD10DB43-E63C-461D-BC2D-127B174639F0}";
            const string arrow5PresId = "{EA3375B3-251A-4C77-B464-BEF875A50217}";

            var child3 = new Dgm.Point { ModelId = child3Id }; child3.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" }); child3.Append(new Dgm.ShapeProperties()); var c3tb = new Dgm.TextBody(); c3tb.Append(new A.BodyProperties()); c3tb.Append(new A.ListStyle()); var c3p = new A.Paragraph(); c3p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); c3tb.Append(c3p); child3.Append(c3tb); pts.Append(child3);
            // Parent/Sibling transition 2
            var parTrans2 = new Dgm.Point { ModelId = parTrans2Id, Type = Dgm.PointValues.ParentTransition, ConnectionId = cxnId2 }; parTrans2.Append(new Dgm.PropertySet()); parTrans2.Append(new Dgm.ShapeProperties()); var pt2tb = new Dgm.TextBody(); pt2tb.Append(new A.BodyProperties()); pt2tb.Append(new A.ListStyle()); var pt2p = new A.Paragraph(); pt2p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); pt2tb.Append(pt2p); parTrans2.Append(pt2tb); pts.Append(parTrans2);
            var sib2 = new Dgm.Point { ModelId = sibTrans2Id, Type = Dgm.PointValues.SiblingTransition, ConnectionId = cxnId2 }; sib2.Append(new Dgm.PropertySet()); sib2.Append(new Dgm.ShapeProperties()); var sb2tb = new Dgm.TextBody(); sb2tb.Append(new A.BodyProperties()); sb2tb.Append(new A.ListStyle()); var sb2p = new A.Paragraph(); sb2p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sb2tb.Append(sb2p); sib2.Append(sb2tb); pts.Append(sib2);
            var parTrans3 = new Dgm.Point { ModelId = "{BF1C22EB-DC9B-4784-B281-99943CF26C7E}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{9CCB7C53-548A-45DF-AA34-4E5103151A5F}" }; parTrans3.Append(new Dgm.PropertySet()); parTrans3.Append(new Dgm.ShapeProperties()); var pt3tb = new Dgm.TextBody(); pt3tb.Append(new A.BodyProperties()); pt3tb.Append(new A.ListStyle()); var pt3p = new A.Paragraph(); pt3p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); pt3tb.Append(pt3p); parTrans3.Append(pt3tb); pts.Append(parTrans3);
            var sib3 = new Dgm.Point { ModelId = sibTrans3Id, Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{9CCB7C53-548A-45DF-AA34-4E5103151A5F}" }; sib3.Append(new Dgm.PropertySet()); sib3.Append(new Dgm.ShapeProperties()); var sb3tb = new Dgm.TextBody(); sb3tb.Append(new A.BodyProperties()); sb3tb.Append(new A.ListStyle()); var sb3p = new A.Paragraph(); sb3p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sb3tb.Append(sb3p); sib3.Append(sb3tb); pts.Append(sib3);

            var child4 = new Dgm.Point { ModelId = child4Id }; child4.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" }); child4.Append(new Dgm.ShapeProperties()); var c4tb = new Dgm.TextBody(); c4tb.Append(new A.BodyProperties()); c4tb.Append(new A.ListStyle()); var c4p = new A.Paragraph(); c4p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); c4tb.Append(c4p); child4.Append(c4tb); pts.Append(child4);
            var parTrans4 = new Dgm.Point { ModelId = parTrans4Id, Type = Dgm.PointValues.ParentTransition, ConnectionId = "{1B618C73-AA93-4069-9C37-DBF7E03AFEC6}" }; parTrans4.Append(new Dgm.PropertySet()); parTrans4.Append(new Dgm.ShapeProperties()); var pt4tb = new Dgm.TextBody(); pt4tb.Append(new A.BodyProperties()); pt4tb.Append(new A.ListStyle()); var pt4p = new A.Paragraph(); pt4p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); pt4tb.Append(pt4p); parTrans4.Append(pt4tb); pts.Append(parTrans4);
            var sib4 = new Dgm.Point { ModelId = sibTrans4Id, Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{1B618C73-AA93-4069-9C37-DBF7E03AFEC6}" }; sib4.Append(new Dgm.PropertySet()); sib4.Append(new Dgm.ShapeProperties()); var sb4tb = new Dgm.TextBody(); sb4tb.Append(new A.BodyProperties()); sb4tb.Append(new A.ListStyle()); var sb4p = new A.Paragraph(); sb4p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sb4tb.Append(sb4p); sib4.Append(sb4tb); pts.Append(sib4);

            var child5 = new Dgm.Point { ModelId = child5Id }; child5.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" }); child5.Append(new Dgm.ShapeProperties()); var c5tb = new Dgm.TextBody(); c5tb.Append(new A.BodyProperties()); c5tb.Append(new A.ListStyle()); var c5p = new A.Paragraph(); c5p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); c5tb.Append(c5p); child5.Append(c5tb); pts.Append(child5);
            var parTrans5 = new Dgm.Point { ModelId = "{FF546BAC-A85C-4761-94A3-6D990039FE05}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{F45CA66E-A1DE-4D62-BFDB-14CFF93A4661}" }; parTrans5.Append(new Dgm.PropertySet()); parTrans5.Append(new Dgm.ShapeProperties()); var pt5tb = new Dgm.TextBody(); pt5tb.Append(new A.BodyProperties()); pt5tb.Append(new A.ListStyle()); var pt5p = new A.Paragraph(); pt5p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); pt5tb.Append(pt5p); parTrans5.Append(pt5tb); pts.Append(parTrans5);
            var sib5 = new Dgm.Point { ModelId = sibTrans5Id, Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{F45CA66E-A1DE-4D62-BFDB-14CFF93A4661}" }; sib5.Append(new Dgm.PropertySet()); sib5.Append(new Dgm.ShapeProperties()); var sb5tb = new Dgm.TextBody(); sb5tb.Append(new A.BodyProperties()); sb5tb.Append(new A.ListStyle()); var sb5p = new A.Paragraph(); sb5p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sb5tb.Append(sb5p); sib5.Append(sb5tb); pts.Append(sib5);

            var presNode3 = new Dgm.Point { ModelId = node3PresId, Type = Dgm.PointValues.Presentation }; var presNode3Props = new Dgm.PropertySet { PresentationElementId = child3Id, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 2, PresentationStyleCount = 5 }; var presNode3Vars = new Dgm.PresentationLayoutVariables(); presNode3Vars.Append(new Dgm.BulletEnabled { Val = true }); presNode3Props.Append(presNode3Vars); presNode3.Append(presNode3Props); presNode3.Append(new Dgm.ShapeProperties()); pts.Append(presNode3);
            var presArrow3 = new Dgm.Point { ModelId = arrow3PresId, Type = Dgm.PointValues.Presentation }; var presArrow3Props = new Dgm.PropertySet { PresentationElementId = sibTrans3Id, PresentationName = "sibTrans", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 2, PresentationStyleCount = 5 }; presArrow3.Append(presArrow3Props); presArrow3.Append(new Dgm.ShapeProperties()); pts.Append(presArrow3);
            var presConnText3 = new Dgm.Point { ModelId = NewId(), Type = Dgm.PointValues.Presentation }; var presConnText3Props = new Dgm.PropertySet { PresentationElementId = sibTrans3Id, PresentationName = "connectorText", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 2, PresentationStyleCount = 5 }; presConnText3.Append(presConnText3Props); presConnText3.Append(new Dgm.ShapeProperties()); pts.Append(presConnText3);
            var presNode4 = new Dgm.Point { ModelId = node4PresId, Type = Dgm.PointValues.Presentation }; var presNode4Props = new Dgm.PropertySet { PresentationElementId = child4Id, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 3, PresentationStyleCount = 5 }; var presNode4Vars = new Dgm.PresentationLayoutVariables(); presNode4Vars.Append(new Dgm.BulletEnabled { Val = true }); presNode4Props.Append(presNode4Vars); presNode4.Append(presNode4Props); presNode4.Append(new Dgm.ShapeProperties()); pts.Append(presNode4);
            var presArrow4 = new Dgm.Point { ModelId = arrow4PresId, Type = Dgm.PointValues.Presentation }; var presArrow4Props = new Dgm.PropertySet { PresentationElementId = sibTrans4Id, PresentationName = "sibTrans", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 3, PresentationStyleCount = 5 }; presArrow4.Append(presArrow4Props); presArrow4.Append(new Dgm.ShapeProperties()); pts.Append(presArrow4);
            var presConnText4 = new Dgm.Point { ModelId = NewId(), Type = Dgm.PointValues.Presentation }; var presConnText4Props = new Dgm.PropertySet { PresentationElementId = sibTrans4Id, PresentationName = "connectorText", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 3, PresentationStyleCount = 5 }; presConnText4.Append(presConnText4Props); presConnText4.Append(new Dgm.ShapeProperties()); pts.Append(presConnText4);
            var presNode5 = new Dgm.Point { ModelId = node5PresId, Type = Dgm.PointValues.Presentation }; var presNode5Props = new Dgm.PropertySet { PresentationElementId = child5Id, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 4, PresentationStyleCount = 5 }; var presNode5Vars = new Dgm.PresentationLayoutVariables(); presNode5Vars.Append(new Dgm.BulletEnabled { Val = true }); presNode5Props.Append(presNode5Vars); presNode5.Append(presNode5Props); presNode5.Append(new Dgm.ShapeProperties()); pts.Append(presNode5);
            var presArrow5 = new Dgm.Point { ModelId = arrow5PresId, Type = Dgm.PointValues.Presentation }; var presArrow5Props = new Dgm.PropertySet { PresentationElementId = sibTrans5Id, PresentationName = "sibTrans", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 4, PresentationStyleCount = 5 }; presArrow5.Append(presArrow5Props); presArrow5.Append(new Dgm.ShapeProperties()); pts.Append(presArrow5);
            const string connText5PresId = "{B29E22C1-42AB-483A-910C-228869C199EB}";
            var presConnText5 = new Dgm.Point { ModelId = connText5PresId, Type = Dgm.PointValues.Presentation }; var presConnText5Props = new Dgm.PropertySet { PresentationElementId = sibTrans5Id, PresentationName = "connectorText", PresentationStyleLabel = "sibTrans2D1", PresentationStyleIndex = 4, PresentationStyleCount = 5 }; presConnText5.Append(presConnText5Props); presConnText5.Append(new Dgm.ShapeProperties()); pts.Append(presConnText5);

            // Connections doc->children and presentation mapping
            var cxns = new Dgm.ConnectionList();
            cxns.Append(new Dgm.Connection { ModelId = cxnId1, SourceId = docId, DestinationId = child1Id, SourcePosition = 0U, DestinationPosition = 0U, ParentTransitionId = parTrans1Id, SiblingTransitionId = sibTrans1Id });
            cxns.Append(new Dgm.Connection { ModelId = cxnId2, SourceId = docId, DestinationId = child2Id, SourcePosition = 1U, DestinationPosition = 0U, ParentTransitionId = parTrans2Id, SiblingTransitionId = sibTrans2Id });
            cxns.Append(new Dgm.Connection { ModelId = cxnId3, SourceId = docId, DestinationId = child3Id, SourcePosition = 2U, DestinationPosition = 0U, ParentTransitionId = parTrans3Id, SiblingTransitionId = sibTrans3Id });
            cxns.Append(new Dgm.Connection { ModelId = cxnId4, SourceId = docId, DestinationId = child4Id, SourcePosition = 3U, DestinationPosition = 0U, ParentTransitionId = parTrans4Id, SiblingTransitionId = sibTrans4Id });
            cxns.Append(new Dgm.Connection { ModelId = cxnId5, SourceId = docId, DestinationId = child5Id, SourcePosition = 4U, DestinationPosition = 0U, ParentTransitionId = parTrans5Id, SiblingTransitionId = sibTrans5Id });

            // PresentationOf (data -> persisted shapes) for sibTrans2 -> arrow(4) and connectorText(2)
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = sibTrans2Id, DestinationId = "{E49B6A3A-6C54-4136-886C-A34B2BF24378}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = sibTrans2Id, DestinationId = "{AC38F62A-4877-4993-9240-0EA2D93CF03C}", SourcePosition = 1U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });

            // PresentationOf mapping (data -> persisted shapes) for nodes and arrows
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = child1Id, DestinationId = node1PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = child2Id, DestinationId = node2PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = child3Id, DestinationId = node3PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = child4Id, DestinationId = node4PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = child5Id, DestinationId = node5PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });

            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = sibTrans1Id, DestinationId = arrow1PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = sibTrans3Id, DestinationId = arrow3PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = sibTrans4Id, DestinationId = arrow4PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = sibTrans5Id, DestinationId = arrow5PresId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" });

            root.Append(pts);
            root.Append(cxns);
            root.Append(new Dgm.Background());
            root.Append(new Dgm.Whole());

            if (!string.IsNullOrEmpty(persistRelId)) {
                var extList = new Dgm.DataModelExtensionList();
                var ext = new A.DataModelExtension { Uri = "http://schemas.microsoft.com/office/drawing/2008/diagram" };
                var dspBlock = new Dsp.DataModelExtensionBlock { RelId = persistRelId, MinVer = "http://schemas.openxmlformats.org/drawingml/2006/diagram" };
                dspBlock.AddNamespaceDeclaration("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram");
                ext.Append(dspBlock);
                extList.Append(ext);
                root.Append(extList);
            }

            part.DataModelRoot = root;
            // Ensure immediate readability of the part by writing the content to the stream
            var xmlCommit = root.OuterXml;
            using (var msCommit = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xmlCommit))) {
                part.FeedData(msCommit);
            }
        }

        private static string NewId() => "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";

        internal static void PopulatePersistLayout(DiagramPersistLayoutPart part) {
            // Ported (typed) from exported SmartArt2.cs -> GenerateDiagramPersistLayoutPart1Content
            Dsp.Drawing drawing2 = new Dsp.Drawing();
            drawing2.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            drawing2.AddNamespaceDeclaration("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram");
            drawing2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Dsp.ShapeTree shapeTree1 = new Dsp.ShapeTree();

            Dsp.GroupShapeNonVisualProperties groupShapeNonVisualProperties1 = new Dsp.GroupShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Dsp.NonVisualGroupDrawingShapeProperties();

            groupShapeNonVisualProperties1.Append(nonVisualDrawingProperties1);
            groupShapeNonVisualProperties1.Append(nonVisualGroupDrawingShapeProperties1);
            Dsp.GroupShapeProperties groupShapeProperties1 = new Dsp.GroupShapeProperties();

            // Node 1 ellipse
            Dsp.Shape shape1 = new Dsp.Shape() { ModelId = "{EA6A36C8-8C9D-4EEB-8E22-E155EC714A8F}" };
            Dsp.ShapeNonVisualProperties shapeNonVisualProperties1 = new Dsp.ShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Dsp.NonVisualDrawingShapeProperties();
            shapeNonVisualProperties1.Append(nonVisualDrawingProperties2);
            shapeNonVisualProperties1.Append(nonVisualDrawingShapeProperties1);
            Dsp.ShapeProperties shapeProperties1 = new Dsp.ShapeProperties();
            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 2259657L, Y = 390L };
            A.Extents extents1 = new A.Extents() { Cx = 967085L, Cy = 967085L };
            transform2D1.Append(offset1);
            transform2D1.Append(extents1);
            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();
            presetGeometry1.Append(adjustValueList1);
            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            schemeColor1.Append(new A.HueOffset() { Val = 0 });
            schemeColor1.Append(new A.SaturationOffset() { Val = 0 });
            schemeColor1.Append(new A.LuminanceOffset() { Val = 0 });
            schemeColor1.Append(new A.AlphaOffset() { Val = 0 });
            solidFill1.Append(schemeColor1);
            A.Outline outline1 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            schemeColor2.Append(new A.HueOffset() { Val = 0 });
            schemeColor2.Append(new A.SaturationOffset() { Val = 0 });
            schemeColor2.Append(new A.LuminanceOffset() { Val = 0 });
            schemeColor2.Append(new A.AlphaOffset() { Val = 0 });
            solidFill2.Append(schemeColor2);
            outline1.Append(solidFill2);
            outline1.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Solid });
            outline1.Append(new A.Miter() { Limit = 800000 });
            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);
            shapeProperties1.Append(new A.EffectList());
            Dsp.ShapeStyle shapeStyle1 = new Dsp.ShapeStyle();
            var lnRef1 = new A.LineReference() { Index = (UInt32Value)2U };
            lnRef1.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 });
            var fillRef1 = new A.FillReference() { Index = (UInt32Value)1U };
            fillRef1.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 });
            var effRef1 = new A.EffectReference() { Index = (UInt32Value)0U };
            effRef1.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 });
            var fontRef1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            fontRef1.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 });
            shapeStyle1.Append(lnRef1);
            shapeStyle1.Append(fillRef1);
            shapeStyle1.Append(effRef1);
            shapeStyle1.Append(fontRef1);
            Dsp.TextBody textBody1 = new Dsp.TextBody();
            var bodyProps1 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 26670, TopInset = 26670, RightInset = 26670, BottomInset = 26670, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            bodyProps1.Append(new A.NoAutoFit());
            textBody1.Append(bodyProps1);
            textBody1.Append(new A.ListStyle());
            var p1 = new A.Paragraph();
            var pPr1 = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 933450 };
            var lnSpc1 = new A.LineSpacing(); lnSpc1.Append(new A.SpacingPercent() { Val = 90000 });
            var spcBef1 = new A.SpaceBefore(); spcBef1.Append(new A.SpacingPercent() { Val = 0 });
            var spcAft1 = new A.SpaceAfter(); spcAft1.Append(new A.SpacingPercent() { Val = 35000 });
            pPr1.Append(lnSpc1); pPr1.Append(spcBef1); pPr1.Append(spcAft1); pPr1.Append(new A.NoBullet());
            p1.Append(pPr1);
            p1.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2100, Kerning = 1200 });
            textBody1.Append(p1);
            Dsp.Transform2D tx1 = new Dsp.Transform2D(); tx1.Append(new A.Offset() { X = 2401283, Y = 142016 }); tx1.Append(new A.Extents() { Cx = 683833, Cy = 683833 });
            shape1.Append(shapeNonVisualProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(shapeStyle1);
            shape1.Append(textBody1);
            shape1.Append(tx1);

            // Right arrow connector 1
            Dsp.Shape shape2 = new Dsp.Shape() { ModelId = "{0BA95F28-1B53-4967-BA10-28214B883828}" };
            var nv2 = new Dsp.ShapeNonVisualProperties();
            nv2.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" });
            nv2.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr2 = new Dsp.ShapeProperties();
            var xfrm2 = new A.Transform2D() { Rotation = 2160000 };
            xfrm2.Append(new A.Offset() { X = 3196004, Y = 742848 });
            xfrm2.Append(new A.Extents() { Cx = 256362, Cy = 326391 });
            spPr2.Append(xfrm2);
            var geom2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RightArrow };
            var av2 = new A.AdjustValueList();
            av2.Append(new A.ShapeGuide() { Name = "adj1", Formula = "val 60000" });
            av2.Append(new A.ShapeGuide() { Name = "adj2", Formula = "val 50000" });
            geom2.Append(av2);
            spPr2.Append(geom2);
            var fill2 = new A.SolidFill(); var sc9 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc9.Append(new A.Tint() { Val = 60000 }); fill2.Append(sc9);
            spPr2.Append(fill2);
            var ln2 = new A.Outline(); ln2.Append(new A.NoFill()); spPr2.Append(ln2);
            spPr2.Append(new A.EffectList());
            var style2 = new Dsp.ShapeStyle();
            style2.Append(new A.LineReference() { Index = (UInt32Value)0U, });
            style2.Append(new A.FillReference() { Index = (UInt32Value)1U });
            style2.Append(new A.EffectReference() { Index = (UInt32Value)0U });
            style2.Append(new A.FontReference() { Index = A.FontCollectionIndexValues.Minor });
            var txBody2 = new Dsp.TextBody();
            var bpr2 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            bpr2.Append(new A.NoAutoFit());
            txBody2.Append(bpr2);
            txBody2.Append(new A.ListStyle());
            var p2 = new A.Paragraph(); var p2Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 577850 };
            var lnSpc2 = new A.LineSpacing(); lnSpc2.Append(new A.SpacingPercent() { Val = 90000 });
            var spBef2 = new A.SpaceBefore(); spBef2.Append(new A.SpacingPercent() { Val = 0 });
            var spAft2 = new A.SpaceAfter(); spAft2.Append(new A.SpacingPercent() { Val = 35000 });
            p2Pr.Append(lnSpc2); p2Pr.Append(spBef2); p2Pr.Append(spAft2); p2Pr.Append(new A.NoBullet()); p2.Append(p2Pr);
            p2.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1300, Kerning = 1200 });
            txBody2.Append(p2);
            var txXfrm2 = new Dsp.Transform2D(); txXfrm2.Append(new A.Offset() { X = 3203348, Y = 785523 }); txXfrm2.Append(new A.Extents() { Cx = 179453, Cy = 195835 });
            shape2.Append(nv2); shape2.Append(spPr2); shape2.Append(style2); shape2.Append(txBody2); shape2.Append(txXfrm2);

            // Node 2 ellipse
            Dsp.Shape shape3 = new Dsp.Shape() { ModelId = "{6E91FA55-2036-4528-BC0C-3D31BEA04727}" };
            var nv3 = new Dsp.ShapeNonVisualProperties(); nv3.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv3.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr3 = new Dsp.ShapeProperties(); var xfrm3 = new A.Transform2D(); xfrm3.Append(new A.Offset() { X = 3433369, Y = 853142 }); xfrm3.Append(new A.Extents() { Cx = 967085, Cy = 967085 }); spPr3.Append(xfrm3);
            var geom3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse }; geom3.Append(new A.AdjustValueList()); spPr3.Append(geom3);
            var fill3 = new A.SolidFill(); var sc3 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc3.Append(new A.HueOffset() { Val = 0 }); sc3.Append(new A.SaturationOffset() { Val = 0 }); sc3.Append(new A.LuminanceOffset() { Val = 0 }); sc3.Append(new A.AlphaOffset() { Val = 0 }); fill3.Append(sc3); spPr3.Append(fill3);
            var ln3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            var ln3Fill = new A.SolidFill(); ln3Fill.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); ln3.Append(ln3Fill); ln3.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Solid }); ln3.Append(new A.Miter() { Limit = 800000 }); spPr3.Append(ln3);
            spPr3.Append(new A.EffectList());
            var style3 = new Dsp.ShapeStyle(); var lnRef3 = new A.LineReference() { Index = (UInt32Value)2U }; lnRef3.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fillRef3 = new A.FillReference() { Index = (UInt32Value)1U }; fillRef3.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var effRef3 = new A.EffectReference() { Index = (UInt32Value)0U }; effRef3.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fontRef3 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor }; fontRef3.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); style3.Append(lnRef3); style3.Append(fillRef3); style3.Append(effRef3); style3.Append(fontRef3);
            var txBody3 = new Dsp.TextBody(); var bpr3 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 26670, TopInset = 26670, RightInset = 26670, BottomInset = 26670, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr3.Append(new A.NoAutoFit()); txBody3.Append(bpr3); txBody3.Append(new A.ListStyle()); var p3 = new A.Paragraph(); var p3Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 933450 }; var lnSpc3 = new A.LineSpacing(); lnSpc3.Append(new A.SpacingPercent() { Val = 90000 }); var spBef3 = new A.SpaceBefore(); spBef3.Append(new A.SpacingPercent() { Val = 0 }); var spAft3 = new A.SpaceAfter(); spAft3.Append(new A.SpacingPercent() { Val = 35000 }); p3Pr.Append(lnSpc3); p3Pr.Append(spBef3); p3Pr.Append(spAft3); p3Pr.Append(new A.NoBullet()); p3.Append(p3Pr); p3.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2100, Kerning = 1200 }); txBody3.Append(p3);
            var txXfrm3 = new Dsp.Transform2D(); txXfrm3.Append(new A.Offset() { X = 3574995, Y = 994768 }); txXfrm3.Append(new A.Extents() { Cx = 683833, Cy = 683833 });
            shape3.Append(nv3); shape3.Append(spPr3); shape3.Append(style3); shape3.Append(txBody3); shape3.Append(txXfrm3);

            // Another right arrow shape (rotated)
            Dsp.Shape shape4 = new Dsp.Shape() { ModelId = "{E49B6A3A-6C54-4136-886C-A34B2BF24378}" };
            var nv4 = new Dsp.ShapeNonVisualProperties(); nv4.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv4.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr4 = new Dsp.ShapeProperties(); var xfrm4 = new A.Transform2D() { Rotation = 6480000 }; xfrm4.Append(new A.Offset() { X = 3566814, Y = 1856479 }); xfrm4.Append(new A.Extents() { Cx = 256362, Cy = 326391 }); spPr4.Append(xfrm4);
            var geom4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RightArrow }; var av4 = new A.AdjustValueList(); av4.Append(new A.ShapeGuide() { Name = "adj1", Formula = "val 60000" }); av4.Append(new A.ShapeGuide() { Name = "adj2", Formula = "val 50000" }); geom4.Append(av4); spPr4.Append(geom4);
            var fill4 = new A.SolidFill(); var sc4 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc4.Append(new A.Tint() { Val = 60000 }); fill4.Append(sc4); spPr4.Append(fill4); var ln4 = new A.Outline(); ln4.Append(new A.NoFill()); spPr4.Append(ln4); spPr4.Append(new A.EffectList());
            var style4 = new Dsp.ShapeStyle(); style4.Append(new A.LineReference() { Index = (UInt32Value)0U }); style4.Append(new A.FillReference() { Index = (UInt32Value)1U }); style4.Append(new A.EffectReference() { Index = (UInt32Value)0U }); style4.Append(new A.FontReference() { Index = A.FontCollectionIndexValues.Minor });
            var txBody4 = new Dsp.TextBody(); var bpr4 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr4.Append(new A.NoAutoFit()); txBody4.Append(bpr4); txBody4.Append(new A.ListStyle()); var p4 = new A.Paragraph(); var p4Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 577850 }; var lnSpc4 = new A.LineSpacing(); lnSpc4.Append(new A.SpacingPercent() { Val = 90000 }); var spBef4 = new A.SpaceBefore(); spBef4.Append(new A.SpacingPercent() { Val = 0 }); var spAft4 = new A.SpaceAfter(); spAft4.Append(new A.SpacingPercent() { Val = 35000 }); p4Pr.Append(lnSpc4); p4Pr.Append(spBef4); p4Pr.Append(spAft4); p4Pr.Append(new A.NoBullet()); p4.Append(p4Pr); p4.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2100, Kerning = 1200 }); txBody4.Append(p4);
            var txXfrm4 = new Dsp.Transform2D(); txXfrm4.Append(new A.Offset() { X = 3574995, Y = 994768 }); txXfrm4.Append(new A.Extents() { Cx = 683833, Cy = 683833 });
            shape4.Append(nv4); shape4.Append(spPr4); shape4.Append(style4); shape4.Append(txBody4); shape4.Append(txXfrm4);

            // Node 3 ellipse
            Dsp.Shape shape5 = new Dsp.Shape() { ModelId = "{AFFFFF7B-5085-434A-A3EC-39CA3570E329}" };
            var nv5 = new Dsp.ShapeNonVisualProperties(); nv5.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv5.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr5 = new Dsp.ShapeProperties(); var xfrm5 = new A.Transform2D(); xfrm5.Append(new A.Offset() { X = 2985051, Y = 2232924 }); xfrm5.Append(new A.Extents() { Cx = 967085, Cy = 967085 }); spPr5.Append(xfrm5);
            var geom5 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse }; geom5.Append(new A.AdjustValueList()); spPr5.Append(geom5);
            var fill5 = new A.SolidFill(); var sc11 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc11.Append(new A.HueOffset() { Val = 0 }); sc11.Append(new A.SaturationOffset() { Val = 0 }); sc11.Append(new A.LuminanceOffset() { Val = 0 }); sc11.Append(new A.AlphaOffset() { Val = 0 }); fill5.Append(sc11); spPr5.Append(fill5);
            var ln5 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            var ln5Fill = new A.SolidFill(); ln5Fill.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); ln5.Append(ln5Fill); ln5.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Solid }); ln5.Append(new A.Miter() { Limit = 800000 }); spPr5.Append(ln5);
            spPr5.Append(new A.EffectList());
            var style5 = new Dsp.ShapeStyle(); var lnRef5 = new A.LineReference() { Index = (UInt32Value)2U }; lnRef5.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fillRef5 = new A.FillReference() { Index = (UInt32Value)1U }; fillRef5.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var effRef5 = new A.EffectReference() { Index = (UInt32Value)0U }; effRef5.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fontRef5 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor }; fontRef5.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); style5.Append(lnRef5); style5.Append(fillRef5); style5.Append(effRef5); style5.Append(fontRef5);
            var txBody5 = new Dsp.TextBody(); var bpr5 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 26670, TopInset = 26670, RightInset = 26670, BottomInset = 26670, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr5.Append(new A.NoAutoFit()); txBody5.Append(bpr5); txBody5.Append(new A.ListStyle()); var p5 = new A.Paragraph(); var p5Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 933450 }; var lnSpc5 = new A.LineSpacing(); lnSpc5.Append(new A.SpacingPercent() { Val = 90000 }); var spBef5 = new A.SpaceBefore(); spBef5.Append(new A.SpacingPercent() { Val = 0 }); var spAft5 = new A.SpaceAfter(); spAft5.Append(new A.SpacingPercent() { Val = 35000 }); p5Pr.Append(lnSpc5); p5Pr.Append(spBef5); p5Pr.Append(spAft5); p5Pr.Append(new A.NoBullet()); p5.Append(p5Pr); p5.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2100, Kerning = 1200 }); txBody5.Append(p5);
            var txXfrm5 = new Dsp.Transform2D(); txXfrm5.Append(new A.Offset() { X = 3126677, Y = 2374550 }); txXfrm5.Append(new A.Extents() { Cx = 683833, Cy = 683833 }); shape5.Append(nv5); shape5.Append(spPr5); shape5.Append(style5); shape5.Append(txBody5); shape5.Append(txXfrm5);

            // Arrow 3
            Dsp.Shape shape6 = new Dsp.Shape() { ModelId = "{1C14F7F4-DFC4-4BCC-9211-35A893821A37}" };
            var nv6 = new Dsp.ShapeNonVisualProperties(); nv6.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv6.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr6 = new Dsp.ShapeProperties(); var xfrm6 = new A.Transform2D() { Rotation = 10800000 }; xfrm6.Append(new A.Offset() { X = 2622274, Y = 2553271 }); xfrm6.Append(new A.Extents() { Cx = 256362, Cy = 326391 }); spPr6.Append(xfrm6);
            var geom6 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RightArrow }; var av6 = new A.AdjustValueList(); av6.Append(new A.ShapeGuide() { Name = "adj1", Formula = "val 60000" }); av6.Append(new A.ShapeGuide() { Name = "adj2", Formula = "val 50000" }); geom6.Append(av6); spPr6.Append(geom6);
            var fill6 = new A.SolidFill(); var sc14 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc14.Append(new A.Tint() { Val = 60000 }); fill6.Append(sc14); spPr6.Append(fill6); var ln6 = new A.Outline(); ln6.Append(new A.NoFill()); spPr6.Append(ln6); spPr6.Append(new A.EffectList());
            var style6 = new Dsp.ShapeStyle(); style6.Append(new A.LineReference() { Index = (UInt32Value)0U }); style6.Append(new A.FillReference() { Index = (UInt32Value)1U }); style6.Append(new A.EffectReference() { Index = (UInt32Value)0U }); style6.Append(new A.FontReference() { Index = A.FontCollectionIndexValues.Minor });
            var txBody6 = new Dsp.TextBody(); var bpr6 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr6.Append(new A.NoAutoFit()); txBody6.Append(bpr6); txBody6.Append(new A.ListStyle()); var p6 = new A.Paragraph(); var p6Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 577850 }; var lnSpc6 = new A.LineSpacing(); lnSpc6.Append(new A.SpacingPercent() { Val = 90000 }); var spBef6 = new A.SpaceBefore(); spBef6.Append(new A.SpacingPercent() { Val = 0 }); var spAft6 = new A.SpaceAfter(); spAft6.Append(new A.SpacingPercent() { Val = 35000 }); p6Pr.Append(lnSpc6); p6Pr.Append(spBef6); p6Pr.Append(spAft6); p6Pr.Append(new A.NoBullet()); p6.Append(p6Pr); p6.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1300, Kerning = 1200 }); txBody6.Append(p6);
            var txXfrm6 = new Dsp.Transform2D(); txXfrm6.Append(new A.Offset() { X = 2622274, Y = 2553271 }); txXfrm6.Append(new A.Extents() { Cx = 256362, Cy = 326391 }); shape6.Append(nv6); shape6.Append(spPr6); shape6.Append(style6); shape6.Append(txBody6); shape6.Append(txXfrm6);

            // Node 4 ellipse
            Dsp.Shape shape7 = new Dsp.Shape() { ModelId = "{AB9582F8-24C6-4DAA-8F90-5AD5EDEB4E2E}" };
            var nv7 = new Dsp.ShapeNonVisualProperties(); nv7.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv7.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr7 = new Dsp.ShapeProperties(); var xfrm7 = new A.Transform2D(); xfrm7.Append(new A.Offset() { X = 1534263, Y = 2232924 }); xfrm7.Append(new A.Extents() { Cx = 967085, Cy = 967085 }); spPr7.Append(xfrm7);
            var geom7 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse }; geom7.Append(new A.AdjustValueList()); spPr7.Append(geom7);
            var fill7 = new A.SolidFill(); var sc10 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc10.Append(new A.HueOffset() { Val = 0 }); sc10.Append(new A.SaturationOffset() { Val = 0 }); sc10.Append(new A.LuminanceOffset() { Val = 0 }); sc10.Append(new A.AlphaOffset() { Val = 0 }); fill7.Append(sc10); spPr7.Append(fill7);
            var ln7 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            var ln7Fill = new A.SolidFill(); ln7Fill.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); ln7.Append(ln7Fill); ln7.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Solid }); ln7.Append(new A.Miter() { Limit = 800000 }); spPr7.Append(ln7); spPr7.Append(new A.EffectList());
            var style7 = new Dsp.ShapeStyle(); var lnRef7 = new A.LineReference() { Index = (UInt32Value)2U }; lnRef7.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fillRef7 = new A.FillReference() { Index = (UInt32Value)1U }; fillRef7.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var effRef7 = new A.EffectReference() { Index = (UInt32Value)0U }; effRef7.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fontRef7 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor }; fontRef7.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); style7.Append(lnRef7); style7.Append(fillRef7); style7.Append(effRef7); style7.Append(fontRef7);
            var txBody7 = new Dsp.TextBody(); var bpr7 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 26670, TopInset = 26670, RightInset = 26670, BottomInset = 26670, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr7.Append(new A.NoAutoFit()); txBody7.Append(bpr7); txBody7.Append(new A.ListStyle()); var p7 = new A.Paragraph(); var p7Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 933450 }; var lnSpc7 = new A.LineSpacing(); lnSpc7.Append(new A.SpacingPercent() { Val = 90000 }); var spBef7 = new A.SpaceBefore(); spBef7.Append(new A.SpacingPercent() { Val = 0 }); var spAft7 = new A.SpaceAfter(); spAft7.Append(new A.SpacingPercent() { Val = 35000 }); p7Pr.Append(lnSpc7); p7Pr.Append(spBef7); p7Pr.Append(spAft7); p7Pr.Append(new A.NoBullet()); p7.Append(p7Pr); p7.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2100, Kerning = 1200 }); txBody7.Append(p7);
            var txXfrm7 = new Dsp.Transform2D(); txXfrm7.Append(new A.Offset() { X = 1675889, Y = 2374550 }); txXfrm7.Append(new A.Extents() { Cx = 683833, Cy = 683833 }); shape7.Append(nv7); shape7.Append(spPr7); shape7.Append(style7); shape7.Append(txBody7); shape7.Append(txXfrm7);

            // Arrow 4
            Dsp.Shape shape8 = new Dsp.Shape() { ModelId = "{7FA91D88-E855-4818-B461-6D97AC4D5A28}" };
            var nv8 = new Dsp.ShapeNonVisualProperties(); nv8.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv8.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr8 = new Dsp.ShapeProperties(); var xfrm8 = new A.Transform2D() { Rotation = 15120000 }; xfrm8.Append(new A.Offset() { X = 1667707, Y = 1870280 }); xfrm8.Append(new A.Extents() { Cx = 256362, Cy = 326391 }); spPr8.Append(xfrm8);
            var geom8 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RightArrow }; var av8 = new A.AdjustValueList(); av8.Append(new A.ShapeGuide() { Name = "adj1", Formula = "val 60000" }); av8.Append(new A.ShapeGuide() { Name = "adj2", Formula = "val 50000" }); geom8.Append(av8); spPr8.Append(geom8);
            var fill8 = new A.SolidFill(); var sc19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc19.Append(new A.Tint() { Val = 60000 }); fill8.Append(sc19); spPr8.Append(fill8); var ln8 = new A.Outline(); ln8.Append(new A.NoFill()); spPr8.Append(ln8); spPr8.Append(new A.EffectList());
            var style8 = new Dsp.ShapeStyle(); style8.Append(new A.LineReference() { Index = (UInt32Value)0U }); style8.Append(new A.FillReference() { Index = (UInt32Value)1U }); style8.Append(new A.EffectReference() { Index = (UInt32Value)0U }); style8.Append(new A.FontReference() { Index = A.FontCollectionIndexValues.Minor });
            var txBody8 = new Dsp.TextBody(); var bpr8 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr8.Append(new A.NoAutoFit()); txBody8.Append(bpr8); txBody8.Append(new A.ListStyle()); var p8 = new A.Paragraph(); var p8Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 577850 }; var lnSpc8 = new A.LineSpacing(); lnSpc8.Append(new A.SpacingPercent() { Val = 90000 }); var spBef8 = new A.SpaceBefore(); spBef8.Append(new A.SpacingPercent() { Val = 0 }); var spAft8 = new A.SpaceAfter(); spAft8.Append(new A.SpacingPercent() { Val = 35000 }); p8Pr.Append(lnSpc8); p8Pr.Append(spBef8); p8Pr.Append(spAft8); p8Pr.Append(new A.NoBullet()); p8.Append(p8Pr); p8.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1300, Kerning = 1200 }); txBody8.Append(p8);
            var txXfrm8 = new Dsp.Transform2D(); txXfrm8.Append(new A.Offset() { X = 1667707, Y = 1870280 }); txXfrm8.Append(new A.Extents() { Cx = 256362, Cy = 326391 }); shape8.Append(nv8); shape8.Append(spPr8); shape8.Append(style8); shape8.Append(txBody8); shape8.Append(txXfrm8);

            // Node 5 ellipse
            Dsp.Shape shape9 = new Dsp.Shape() { ModelId = "{CD10DB43-E63C-461D-BC2D-127B174639F0}" };
            var nv9 = new Dsp.ShapeNonVisualProperties(); nv9.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv9.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr9 = new Dsp.ShapeProperties(); var xfrm9 = new A.Transform2D(); xfrm9.Append(new A.Offset() { X = 907806, Y = 1363107 }); xfrm9.Append(new A.Extents() { Cx = 967085, Cy = 967085 }); spPr9.Append(xfrm9);
            var geom9 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse }; geom9.Append(new A.AdjustValueList()); spPr9.Append(geom9);
            var fill9 = new A.SolidFill(); var sc21 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc21.Append(new A.HueOffset() { Val = 0 }); sc21.Append(new A.SaturationOffset() { Val = 0 }); sc21.Append(new A.LuminanceOffset() { Val = 0 }); sc21.Append(new A.AlphaOffset() { Val = 0 }); fill9.Append(sc21); spPr9.Append(fill9);
            var ln9 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            var ln9Fill = new A.SolidFill(); ln9Fill.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); ln9.Append(ln9Fill); ln9.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Solid }); ln9.Append(new A.Miter() { Limit = 800000 }); spPr9.Append(ln9); spPr9.Append(new A.EffectList());
            var style9 = new Dsp.ShapeStyle(); var lnRef9 = new A.LineReference() { Index = (UInt32Value)2U }; lnRef9.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fillRef9 = new A.FillReference() { Index = (UInt32Value)1U }; fillRef9.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var effRef9 = new A.EffectReference() { Index = (UInt32Value)0U }; effRef9.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fontRef9 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor }; fontRef9.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); style9.Append(lnRef9); style9.Append(fillRef9); style9.Append(effRef9); style9.Append(fontRef9);
            var txBody9 = new Dsp.TextBody(); var bpr9 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 26670, TopInset = 26670, RightInset = 26670, BottomInset = 26670, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr9.Append(new A.NoAutoFit()); txBody9.Append(bpr9); txBody9.Append(new A.ListStyle()); var p9 = new A.Paragraph(); var p9Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 933450 }; var lnSpc9 = new A.LineSpacing(); lnSpc9.Append(new A.SpacingPercent() { Val = 90000 }); var spBef9 = new A.SpaceBefore(); spBef9.Append(new A.SpacingPercent() { Val = 0 }); var spAft9 = new A.SpaceAfter(); spAft9.Append(new A.SpacingPercent() { Val = 35000 }); p9Pr.Append(lnSpc9); p9Pr.Append(spBef9); p9Pr.Append(spAft9); p9Pr.Append(new A.NoBullet()); p9.Append(p9Pr); p9.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2100, Kerning = 1200 }); txBody9.Append(p9);
            var txXfrm9 = new Dsp.Transform2D(); txXfrm9.Append(new A.Offset() { X = 1049432, Y = 1504733 }); txXfrm9.Append(new A.Extents() { Cx = 683833, Cy = 683833 }); shape9.Append(nv9); shape9.Append(spPr9); shape9.Append(style9); shape9.Append(txBody9); shape9.Append(txXfrm9);

            // Arrow 5
            Dsp.Shape shape10 = new Dsp.Shape() { ModelId = "{EA3375B3-251A-4C77-B464-BEF875A50217}" };
            var nv10 = new Dsp.ShapeNonVisualProperties(); nv10.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv10.Append(new Dsp.NonVisualDrawingShapeProperties());
            var spPr10 = new Dsp.ShapeProperties(); var xfrm10 = new A.Transform2D(); xfrm10.Append(new A.Offset() { X = 2029636, Y = 839259 }); xfrm10.Append(new A.Extents() { Cx = 179453, Cy = 195835 }); spPr10.Append(xfrm10);
            var geom10 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RightArrow }; var av10 = new A.AdjustValueList(); av10.Append(new A.ShapeGuide() { Name = "adj1", Formula = "val 60000" }); av10.Append(new A.ShapeGuide() { Name = "adj2", Formula = "val 50000" }); geom10.Append(av10); spPr10.Append(geom10);
            var fill10 = new A.SolidFill(); var sc25 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc25.Append(new A.Tint() { Val = 60000 }); fill10.Append(sc25); spPr10.Append(fill10); var ln10 = new A.Outline(); ln10.Append(new A.NoFill()); spPr10.Append(ln10); spPr10.Append(new A.EffectList());
            var style10 = new Dsp.ShapeStyle(); style10.Append(new A.LineReference() { Index = (UInt32Value)0U }); style10.Append(new A.FillReference() { Index = (UInt32Value)1U }); style10.Append(new A.EffectReference() { Index = (UInt32Value)0U }); style10.Append(new A.FontReference() { Index = A.FontCollectionIndexValues.Minor });
            var txBody10 = new Dsp.TextBody(); var bpr10 = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr10.Append(new A.NoAutoFit()); txBody10.Append(bpr10); txBody10.Append(new A.ListStyle()); var p10 = new A.Paragraph(); var p10Pr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 577850 }; var lnSpc10 = new A.LineSpacing(); lnSpc10.Append(new A.SpacingPercent() { Val = 90000 }); var spBef10 = new A.SpaceBefore(); spBef10.Append(new A.SpacingPercent() { Val = 0 }); var spAft10 = new A.SpaceAfter(); spAft10.Append(new A.SpacingPercent() { Val = 35000 }); p10Pr.Append(lnSpc10); p10Pr.Append(spBef10); p10Pr.Append(spAft10); p10Pr.Append(new A.NoBullet()); p10.Append(p10Pr); p10.Append(new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1300, Kerning = 1200 }); txBody10.Append(p10);
            var txXfrm10 = new Dsp.Transform2D(); txXfrm10.Append(new A.Offset() { X = 2029636, Y = 839259 }); txXfrm10.Append(new A.Extents() { Cx = 179453, Cy = 195835 }); shape10.Append(nv10); shape10.Append(spPr10); shape10.Append(style10); shape10.Append(txBody10); shape10.Append(txXfrm10);

            // Assemble tree
            shapeTree1.Append(groupShapeNonVisualProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(shape1);
            shapeTree1.Append(shape2);
            shapeTree1.Append(shape3);
            shapeTree1.Append(shape4);
            shapeTree1.Append(shape5);
            shapeTree1.Append(shape6);
            shapeTree1.Append(shape7);
            shapeTree1.Append(shape8);
            shapeTree1.Append(shape9);
            shapeTree1.Append(shape10);

            drawing2.Append(shapeTree1);
            part.Drawing = drawing2;
        }
    }
}
