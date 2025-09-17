using DocumentFormat.OpenXml.Packaging;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using Dsp = DocumentFormat.OpenXml.Office.Drawing;

namespace OfficeIMO.Word.SmartArt.Templates {
    /// Custom SmartArt template based on exported SmartArt1 (list/default layout).
    internal static class SmartArtCustom1 {
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

            // node1, node0 and sibTrans2D1 to align with exports
            colors.Append(MakeStyleLabel("node0", accentFill: true, lineAccent: false));
            colors.Append(MakeStyleLabel("node1", accentFill: true, lineAccent: false));
            colors.Append(MakeStyleLabelTint("sibTrans2D1", 60000));

            // Additional labels used by exported template for better visual parity
            colors.Append(MakeStyleLabel("lnNode1", accentFill: false, lineAccent: true));
            colors.Append(MakeStyleLabel("alignNode1", accentFill: true, lineAccent: false));
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
            // Core labels
            style.Append(MakeStyleLabelStyle("node0"));
            style.Append(MakeStyleLabelStyle("node1"));
            style.Append(MakeStyleLabelStyle("lnNode1"));
            style.Append(MakeStyleLabelStyle("alignNode1"));
            style.Append(MakeStyleLabelStyle("sibTrans2D1"));
            style.Append(MakeStyleLabelStyle("sibTrans1D1"));
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

        private static Dgm.StyleLabel MakeStyleLabelStyle(string name) {
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
            // Derived from Assets/WordTemplates/SmartArt1.cs (GenerateDiagramLayoutDefinitionPart1Content)
            // Simplified to the essential structure while preserving IDs and categories
            var layout = new Dgm.LayoutDefinition { UniqueId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            layout.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            layout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            layout.Append(new Dgm.Title { Val = "" });
            layout.Append(new Dgm.Description { Val = "" });

            var cats = new Dgm.CategoryList();
            cats.Append(new Dgm.Category { Type = "list", Priority = (UInt32Value)400U });
            layout.Append(cats);

            var layoutNode = new Dgm.LayoutNode { Name = "diagram" };

            // Variables drive direction and sizing semantics
            var varList = new Dgm.VariableList();
            varList.Append(new Dgm.Direction());
            varList.Append(new Dgm.ResizeHandles { Val = Dgm.ResizeHandlesStringValues.Exact });
            layoutNode.Append(varList);

            // Choose algorithm based on 'dir' variable
            var choose = new Dgm.Choose { Name = "Name0" };
            var ifDirNorm = new Dgm.DiagramChooseIf {
                Name = "Name1",
                Function = Dgm.FunctionValues.Variable,
                Argument = "dir",
                Operator = Dgm.FunctionOperatorValues.Equal,
                Val = "norm"
            };
            var algSnakeNorm = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Snake };
            algSnakeNorm.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.GrowDirection, Val = "tL" });
            algSnakeNorm.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.FlowDirection, Val = "row" });
            algSnakeNorm.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.ContinueDirection, Val = "sameDir" });
            algSnakeNorm.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.Offset, Val = "ctr" });
            ifDirNorm.Append(algSnakeNorm);

            var elseDir = new Dgm.DiagramChooseElse { Name = "Name2" };
            var algSnakeElse = new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Snake };
            algSnakeElse.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.GrowDirection, Val = "tR" });
            algSnakeElse.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.FlowDirection, Val = "row" });
            algSnakeElse.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.ContinueDirection, Val = "sameDir" });
            algSnakeElse.Append(new Dgm.Parameter { Type = Dgm.ParameterIdValues.Offset, Val = "ctr" });
            elseDir.Append(algSnakeElse);

            choose.Append(ifDirNorm);
            choose.Append(elseDir);
            layoutNode.Append(choose);

            // Base diagram shape
            var shape = new Dgm.Shape { Blip = "" };
            shape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            shape.Append(new Dgm.AdjustList());
            layoutNode.Append(shape);

            // Apply to descendants-or-self nodes
            layoutNode.Append(new Dgm.PresentationOf());

            // Constraints controlling size, spacing and font
            var constr = new Dgm.Constraints();
            constr.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.Width,
                For = Dgm.ConstraintRelationshipValues.Child,
                PointType = Dgm.ElementValues.Node,
                ReferenceType = Dgm.ConstraintValues.Width
            });
            constr.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.Height,
                For = Dgm.ConstraintRelationshipValues.Child,
                PointType = Dgm.ElementValues.Node,
                ReferenceType = Dgm.ConstraintValues.Width,
                ReferenceFor = Dgm.ConstraintRelationshipValues.Child,
                ReferencePointType = Dgm.ElementValues.Node,
                Fact = 0.6D
            });
            constr.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.Width,
                For = Dgm.ConstraintRelationshipValues.Child,
                PointType = Dgm.ElementValues.SiblingTransition,
                ReferenceType = Dgm.ConstraintValues.Width,
                ReferenceFor = Dgm.ConstraintRelationshipValues.Child,
                ReferencePointType = Dgm.ElementValues.Node,
                Fact = 0.1D
            });
            constr.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.SiblingSpacing,
                ReferenceType = Dgm.ConstraintValues.Width,
                ReferenceFor = Dgm.ConstraintRelationshipValues.Child,
                ReferencePointType = Dgm.ElementValues.Node
            });
            constr.Append(new Dgm.Constraint {
                Type = Dgm.ConstraintValues.PrimaryFontSize,
                For = Dgm.ConstraintRelationshipValues.Child,
                PointType = Dgm.ElementValues.Node,
                Operator = Dgm.BoolOperatorValues.Equal,
                Val = 65D
            });
            layoutNode.Append(constr);

            layoutNode.Append(new Dgm.RuleList());

            // For each child node render a rectangle (process step)
            var forEach = new Dgm.ForEach {
                Name = "Name3",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "ch" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" }
            };

            var node = new Dgm.LayoutNode { Name = "node" };
            var nodeVarList = new Dgm.VariableList();
            nodeVarList.Append(new Dgm.BulletEnabled { Val = true });
            node.Append(nodeVarList);
            node.Append(new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Text });
            var nodeShape = new Dgm.Shape { Type = "rect", Blip = "" };
            nodeShape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nodeShape.Append(new Dgm.AdjustList());
            node.Append(nodeShape);
            node.Append(new Dgm.PresentationOf {
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "desOrSelf" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "node" }
            });
            var nodeConstr = new Dgm.Constraints();
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.LeftMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.RightMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.TopMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D });
            nodeConstr.Append(new Dgm.Constraint { Type = Dgm.ConstraintValues.BottomMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D });
            node.Append(nodeConstr);
            var nodeRules = new Dgm.RuleList();
            nodeRules.Append(new Dgm.Rule { Type = Dgm.ConstraintValues.PrimaryFontSize, Val = 5D, Fact = new DoubleValue { InnerText = "NaN" }, Max = new DoubleValue { InnerText = "NaN" } });
            node.Append(nodeRules);
            forEach.Append(node);

            // Sibling transition spacer (once per sibling edge)
            var forEachSib = new Dgm.ForEach {
                Name = "Name4",
                Axis = new ListValue<EnumValue<Dgm.AxisValues>> { InnerText = "followSib" },
                PointType = new ListValue<EnumValue<Dgm.ElementValues>> { InnerText = "sibTrans" },
                Count = new ListValue<UInt32Value> { InnerText = "1" }
            };
            var sibNode = new Dgm.LayoutNode { Name = "sibTrans" };
            sibNode.Append(new Dgm.Algorithm { Type = Dgm.AlgorithmValues.Space });
            var sibShape = new Dgm.Shape { Blip = "" };
            sibShape.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            sibShape.Append(new Dgm.AdjustList());
            sibNode.Append(sibShape);
            sibNode.Append(new Dgm.PresentationOf());
            sibNode.Append(new Dgm.Constraints());
            sibNode.Append(new Dgm.RuleList());
            forEachSib.Append(sibNode);
            forEach.Append(forEachSib);

            layoutNode.Append(forEach);

            layout.Append(layoutNode);
            part.LayoutDefinition = layout;
        }

        internal static void PopulateData(DiagramDataPart part, string? persistRelId = null) {
            // Port mapping from exported SmartArt1 (list/default) for 5 nodes and presentation links
            var root = new Dgm.DataModelRoot();
            root.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            root.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var pts = new Dgm.PointList();
            const string docId = "{050810F5-6B93-4502-92CB-06E52385CAF3}";
            // Data nodes
            var doc = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
            doc.Append(new Dgm.PropertySet {
                LayoutTypeId = "urn:microsoft.com/office/officeart/2005/8/layout/default",
                LayoutCategoryId = "list",
                QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
                QuickStyleCategoryId = "simple",
                ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
                ColorCategoryId = "accent1",
                Placeholder = false
            });
            doc.Append(new Dgm.ShapeProperties()); var dTb = new Dgm.TextBody(); dTb.Append(new A.BodyProperties()); dTb.Append(new A.ListStyle()); var dP = new A.Paragraph(); dP.Append(new A.EndParagraphRunProperties { Language = "en-US" }); dTb.Append(dP); doc.Append(dTb); pts.Append(doc);

            const string n1 = "{2EEF2A58-A2D7-4991-8A78-E26574C46C74}";
            const string n2 = "{68641FAB-77F7-4312-BEB5-72B80B86845C}";
            const string n3 = "{89391C13-C504-4B29-8FBB-561C84CC10C1}";
            const string n4 = "{575655FE-BCCE-4492-8A81-F5E8A3A92F30}";
            const string n5 = "{17A72E10-5B50-4595-A1C0-8605F16EB379}";
            string[] children = { n1, n2, n3, n4, n5 };
            for (int i = 0; i < children.Length; i++) {
                var cid = children[i]; var pt = new Dgm.Point { ModelId = cid }; pt.Append(new Dgm.PropertySet { Placeholder = true, PlaceholderText = "[Text]" }); pt.Append(new Dgm.ShapeProperties()); var tb = new Dgm.TextBody(); tb.Append(new A.BodyProperties()); tb.Append(new A.ListStyle()); var p = new A.Paragraph(); p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); tb.Append(p); pt.Append(tb); pts.Append(pt);
            }

            // Presentation nodes mapping dsp rectangles
            const string presDiagramId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}";
            var presDiagram = new Dgm.Point { ModelId = presDiagramId, Type = Dgm.PointValues.Presentation };
            var presDProps = new Dgm.PropertySet { PresentationElementId = docId, PresentationName = "diagram", PresentationStyleCount = 0 };
            var presDVars = new Dgm.PresentationLayoutVariables(); presDVars.Append(new Dgm.Direction()); presDVars.Append(new Dgm.ResizeHandles { Val = Dgm.ResizeHandlesStringValues.Exact }); presDProps.Append(presDVars);
            presDiagram.Append(presDProps); presDiagram.Append(new Dgm.ShapeProperties()); pts.Append(presDiagram);

            var pres1 = new Dgm.Point { ModelId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", Type = Dgm.PointValues.Presentation };
            pres1.Append(new Dgm.PropertySet { PresentationElementId = n1, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 0, PresentationStyleCount = 5 }); pres1.Append(new Dgm.ShapeProperties()); pts.Append(pres1);
            var pres2 = new Dgm.Point { ModelId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}", Type = Dgm.PointValues.Presentation };
            pres2.Append(new Dgm.PropertySet { PresentationElementId = n2, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 1, PresentationStyleCount = 5 }); pres2.Append(new Dgm.ShapeProperties()); pts.Append(pres2);
            var pres3 = new Dgm.Point { ModelId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}", Type = Dgm.PointValues.Presentation };
            pres3.Append(new Dgm.PropertySet { PresentationElementId = n3, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 2, PresentationStyleCount = 5 }); pres3.Append(new Dgm.ShapeProperties()); pts.Append(pres3);
            var pres4 = new Dgm.Point { ModelId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}", Type = Dgm.PointValues.Presentation };
            pres4.Append(new Dgm.PropertySet { PresentationElementId = n4, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 3, PresentationStyleCount = 5 }); pres4.Append(new Dgm.ShapeProperties()); pts.Append(pres4);
            var pres5 = new Dgm.Point { ModelId = "{2B5DA877-174D-4060-8E30-014EB5090235}", Type = Dgm.PointValues.Presentation };
            pres5.Append(new Dgm.PropertySet { PresentationElementId = n5, PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 4, PresentationStyleCount = 5 }); pres5.Append(new Dgm.ShapeProperties()); pts.Append(pres5);

            // Parent/Sibling transition points (to match export)
            // child1
            var par1 = new Dgm.Point { ModelId = "{34BEA7A6-9CDE-4021-B5E2-BA989ECE9AE2}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{53DE9593-1885-4455-BA51-60A7F54355EE}" };
            par1.Append(new Dgm.PropertySet()); par1.Append(new Dgm.ShapeProperties()); var par1tb = new Dgm.TextBody(); par1tb.Append(new A.BodyProperties()); par1tb.Append(new A.ListStyle()); var par1p = new A.Paragraph(); par1p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); par1tb.Append(par1p); par1.Append(par1tb); pts.Append(par1);
            var sib1 = new Dgm.Point { ModelId = "{D1398D45-A4D5-4AEC-A3DC-A6C3D843196D}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{53DE9593-1885-4455-BA51-60A7F54355EE}" };
            sib1.Append(new Dgm.PropertySet()); sib1.Append(new Dgm.ShapeProperties()); var sib1tb = new Dgm.TextBody(); sib1tb.Append(new A.BodyProperties()); sib1tb.Append(new A.ListStyle()); var sib1p = new A.Paragraph(); sib1p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sib1tb.Append(sib1p); sib1.Append(sib1tb); pts.Append(sib1);
            // child2
            var par2 = new Dgm.Point { ModelId = "{3D52860C-F7E4-494B-8C26-803853184C4F}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{20801D82-1CB0-4B72-A2A5-2604C9CC9D7E}" };
            par2.Append(new Dgm.PropertySet()); par2.Append(new Dgm.ShapeProperties()); var par2tb = new Dgm.TextBody(); par2tb.Append(new A.BodyProperties()); par2tb.Append(new A.ListStyle()); var par2p = new A.Paragraph(); par2p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); par2tb.Append(par2p); par2.Append(par2tb); pts.Append(par2);
            var sib2 = new Dgm.Point { ModelId = "{93C814DC-C96A-464D-A307-B751C432E31D}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{20801D82-1CB0-4B72-A2A5-2604C9CC9D7E}" };
            sib2.Append(new Dgm.PropertySet()); sib2.Append(new Dgm.ShapeProperties()); var sib2tb = new Dgm.TextBody(); sib2tb.Append(new A.BodyProperties()); sib2tb.Append(new A.ListStyle()); var sib2p = new A.Paragraph(); sib2p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sib2tb.Append(sib2p); sib2.Append(sib2tb); pts.Append(sib2);
            // child3
            var par3 = new Dgm.Point { ModelId = "{EA6D4184-8096-4199-B209-0A82B4347DDE}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{192BD615-74D4-4535-B946-F0127C2DAB21}" };
            par3.Append(new Dgm.PropertySet()); par3.Append(new Dgm.ShapeProperties()); var par3tb = new Dgm.TextBody(); par3tb.Append(new A.BodyProperties()); par3tb.Append(new A.ListStyle()); var par3p = new A.Paragraph(); par3p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); par3tb.Append(par3p); par3.Append(par3tb); pts.Append(par3);
            var sib3 = new Dgm.Point { ModelId = "{839AC1D3-AFF4-44E4-AD2D-D35D5C61E303}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{192BD615-74D4-4535-B946-F0127C2DAB21}" };
            sib3.Append(new Dgm.PropertySet()); sib3.Append(new Dgm.ShapeProperties()); var sib3tb = new Dgm.TextBody(); sib3tb.Append(new A.BodyProperties()); sib3tb.Append(new A.ListStyle()); var sib3p = new A.Paragraph(); sib3p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sib3tb.Append(sib3p); sib3.Append(sib3tb); pts.Append(sib3);
            // child4
            var par4 = new Dgm.Point { ModelId = "{5DCAA827-5E19-4785-BA24-36EDD3800D76}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{5B0B9687-F94F-40E3-BC77-4E2284D67B82}" };
            par4.Append(new Dgm.PropertySet()); par4.Append(new Dgm.ShapeProperties()); var par4tb = new Dgm.TextBody(); par4tb.Append(new A.BodyProperties()); par4tb.Append(new A.ListStyle()); var par4p = new A.Paragraph(); par4p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); par4tb.Append(par4p); par4.Append(par4tb); pts.Append(par4);
            var sib4 = new Dgm.Point { ModelId = "{D9D42E39-AEF2-4B6C-BD25-361A6F0D8CB2}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{5B0B9687-F94F-40E3-BC77-4E2284D67B82}" };
            sib4.Append(new Dgm.PropertySet()); sib4.Append(new Dgm.ShapeProperties()); var sib4tb = new Dgm.TextBody(); sib4tb.Append(new A.BodyProperties()); sib4tb.Append(new A.ListStyle()); var sib4p = new A.Paragraph(); sib4p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sib4tb.Append(sib4p); sib4.Append(sib4tb); pts.Append(sib4);
            // child5
            var par5 = new Dgm.Point { ModelId = "{55D66658-965D-44CC-ABAE-6638BD3536C3}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{BD2FFBF0-81A7-4151-B2A4-72895B9DAC55}" };
            par5.Append(new Dgm.PropertySet()); par5.Append(new Dgm.ShapeProperties()); var par5tb = new Dgm.TextBody(); par5tb.Append(new A.BodyProperties()); par5tb.Append(new A.ListStyle()); var par5p = new A.Paragraph(); par5p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); par5tb.Append(par5p); par5.Append(par5tb); pts.Append(par5);
            var sib5 = new Dgm.Point { ModelId = "{0ADD6561-125E-4ED0-A0C2-6BC4566DCBA6}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{BD2FFBF0-81A7-4151-B2A4-72895B9DAC55}" };
            sib5.Append(new Dgm.PropertySet()); sib5.Append(new Dgm.ShapeProperties()); var sib5tb = new Dgm.TextBody(); sib5tb.Append(new A.BodyProperties()); sib5tb.Append(new A.ListStyle()); var sib5p = new A.Paragraph(); sib5p.Append(new A.EndParagraphRunProperties { Language = "en-US" }); sib5tb.Append(sib5p); sib5.Append(sib5tb); pts.Append(sib5);

            // Presentation points for sibTrans (subset present in export)
            var presSib1 = new Dgm.Point { ModelId = "{BB3DC0D3-305A-444F-A1F6-B180204939B5}", Type = Dgm.PointValues.Presentation };
            presSib1.Append(new Dgm.PropertySet { PresentationElementId = "{D1398D45-A4D5-4AEC-A3DC-A6C3D843196D}", PresentationName = "sibTrans", PresentationStyleCount = 0 }); presSib1.Append(new Dgm.ShapeProperties()); pts.Append(presSib1);
            var presSib2 = new Dgm.Point { ModelId = "{2AD38D92-7E4A-4864-A84B-BD1C732A96A3}", Type = Dgm.PointValues.Presentation };
            presSib2.Append(new Dgm.PropertySet { PresentationElementId = "{93C814DC-C96A-464D-A307-B751C432E31D}", PresentationName = "sibTrans", PresentationStyleCount = 0 }); presSib2.Append(new Dgm.ShapeProperties()); pts.Append(presSib2);
            var presSib3 = new Dgm.Point { ModelId = "{0052E64C-8063-4397-B08E-4468F28ADC78}", Type = Dgm.PointValues.Presentation };
            presSib3.Append(new Dgm.PropertySet { PresentationElementId = "{839AC1D3-AFF4-44E4-AD2D-D35D5C61E303}", PresentationName = "sibTrans", PresentationStyleCount = 0 }); presSib3.Append(new Dgm.ShapeProperties()); pts.Append(presSib3);
            var presSib4 = new Dgm.Point { ModelId = "{465E4E3A-A29E-4172-B02E-CDF8A89566C6}", Type = Dgm.PointValues.Presentation };
            presSib4.Append(new Dgm.PropertySet { PresentationElementId = "{D9D42E39-AEF2-4B6C-BD25-361A6F0D8CB2}", PresentationName = "sibTrans", PresentationStyleCount = 0 }); presSib4.Append(new Dgm.ShapeProperties()); pts.Append(presSib4);

            // Connections doc -> children (five)
            var cxns = new Dgm.ConnectionList();
            cxns.Append(new Dgm.Connection { ModelId = "{53DE9593-1885-4455-BA51-60A7F54355EE}", SourceId = docId, DestinationId = n1, SourcePosition = 0U, DestinationPosition = 0U, ParentTransitionId = "{34BEA7A6-9CDE-4021-B5E2-BA989ECE9AE2}", SiblingTransitionId = "{D1398D45-A4D5-4AEC-A3DC-A6C3D843196D}" });
            cxns.Append(new Dgm.Connection { ModelId = "{20801D82-1CB0-4B72-A2A5-2604C9CC9D7E}", SourceId = docId, DestinationId = n2, SourcePosition = 1U, DestinationPosition = 0U, ParentTransitionId = "{3D52860C-F7E4-494B-8C26-803853184C4F}", SiblingTransitionId = "{93C814DC-C96A-464D-A307-B751C432E31D}" });
            cxns.Append(new Dgm.Connection { ModelId = "{192BD615-74D4-4535-B946-F0127C2DAB21}", SourceId = docId, DestinationId = n3, SourcePosition = 2U, DestinationPosition = 0U, ParentTransitionId = "{EA6D4184-8096-4199-B209-0A82B4347DDE}", SiblingTransitionId = "{839AC1D3-AFF4-44E4-AD2D-D35D5C61E303}" });
            cxns.Append(new Dgm.Connection { ModelId = "{5B0B9687-F94F-40E3-BC77-4E2284D67B82}", SourceId = docId, DestinationId = n4, SourcePosition = 3U, DestinationPosition = 0U, ParentTransitionId = "{5DCAA827-5E19-4785-BA24-36EDD3800D76}", SiblingTransitionId = "{D9D42E39-AEF2-4B6C-BD25-361A6F0D8CB2}" });
            cxns.Append(new Dgm.Connection { ModelId = "{BD2FFBF0-81A7-4151-B2A4-72895B9DAC55}", SourceId = docId, DestinationId = n5, SourcePosition = 4U, DestinationPosition = 0U, ParentTransitionId = "{55D66658-965D-44CC-ABAE-6638BD3536C3}", SiblingTransitionId = "{0ADD6561-125E-4ED0-A0C2-6BC4566DCBA6}" });

            // Presentation mapping connections to persist shapes for visual parity
            // PresentationOf: map data nodes to persisted rectangle model ids
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = n1, DestinationId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = n2, DestinationId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = n3, DestinationId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = n4, DestinationId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = n5, DestinationId = "{2B5DA877-174D-4060-8E30-014EB5090235}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });

            // PresentationOf: document -> diagram presentation id
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationOf, SourceId = docId, DestinationId = presDiagramId, SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });

            // PresentationParentOf: presentation diagram -> each presentation node shape
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = presDiagramId, DestinationId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", SourcePosition = 0U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = presDiagramId, DestinationId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}", SourcePosition = 2U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = presDiagramId, DestinationId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}", SourcePosition = 4U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = presDiagramId, DestinationId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}", SourcePosition = 6U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });
            cxns.Append(new Dgm.Connection { ModelId = NewId(), Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = presDiagramId, DestinationId = "{2B5DA877-174D-4060-8E30-014EB5090235}", SourcePosition = 8U, DestinationPosition = 0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" });

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
            // Ensure the part has a physical stream so immediate consumers can read it
            var xmlCommit = root.OuterXml;
            using (var msCommit = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xmlCommit))) {
                part.FeedData(msCommit);
            }
        }

        private static string NewId() => "{" + System.Guid.NewGuid().ToString().ToUpper() + "}";

        internal static void PopulatePersistLayout(DiagramPersistLayoutPart part) {
            // Ported subset (typed) from exported SmartArt1.cs -> GenerateDiagramPersistLayoutPart1Content
            Dsp.Drawing drawing2 = new Dsp.Drawing();
            drawing2.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            drawing2.AddNamespaceDeclaration("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram");
            drawing2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Dsp.ShapeTree shapeTree1 = new Dsp.ShapeTree();
            var nvGrp = new Dsp.GroupShapeNonVisualProperties();
            nvGrp.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" });
            nvGrp.Append(new Dsp.NonVisualGroupDrawingShapeProperties());
            var grpProps = new Dsp.GroupShapeProperties();

            // Five rectangle nodes (3 on first row, 2 on second) per export
            Dsp.Shape MakeRect(string modelId, long x, long y, long cx, long cy) {
                var sp = new Dsp.Shape() { ModelId = modelId };
                var nv = new Dsp.ShapeNonVisualProperties(); nv.Append(new Dsp.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" }); nv.Append(new Dsp.NonVisualDrawingShapeProperties());
                var spPr = new Dsp.ShapeProperties(); var xfrm = new A.Transform2D(); xfrm.Append(new A.Offset() { X = x, Y = y }); xfrm.Append(new A.Extents() { Cx = cx, Cy = cy }); spPr.Append(xfrm);
                var geom = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle }; geom.Append(new A.AdjustValueList()); spPr.Append(geom);
                var fill = new A.SolidFill(); var sc = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }; sc.Append(new A.HueOffset() { Val = 0 }); sc.Append(new A.SaturationOffset() { Val = 0 }); sc.Append(new A.LuminanceOffset() { Val = 0 }); sc.Append(new A.AlphaOffset() { Val = 0 }); fill.Append(sc); spPr.Append(fill);
                var ln = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
                var lnFill = new A.SolidFill(); lnFill.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); ln.Append(lnFill); ln.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Solid }); ln.Append(new A.Miter() { Limit = 800000 }); spPr.Append(ln);
                spPr.Append(new A.EffectList());
                var style = new Dsp.ShapeStyle(); var lnRef = new A.LineReference() { Index = (UInt32Value)2U }; lnRef.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fillRef = new A.FillReference() { Index = (UInt32Value)1U }; fillRef.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var effRef = new A.EffectReference() { Index = (UInt32Value)0U }; effRef.Append(new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 }); var fontRef = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor }; fontRef.Append(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 }); style.Append(lnRef); style.Append(fillRef); style.Append(effRef); style.Append(fontRef);
            var tx = new Dsp.TextBody(); var bpr = new A.BodyProperties() { UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 179070, TopInset = 179070, RightInset = 179070, BottomInset = 179070, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false }; bpr.Append(new A.NoAutoFit()); tx.Append(bpr); tx.Append(new A.ListStyle()); var p = new A.Paragraph(); var pPr = new A.ParagraphProperties() { LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 2089150 }; var lns = new A.LineSpacing(); lns.Append(new A.SpacingPercent() { Val = 90000 }); var sb = new A.SpaceBefore(); sb.Append(new A.SpacingPercent() { Val = 0 }); var sa = new A.SpaceAfter(); sa.Append(new A.SpacingPercent() { Val = 35000 }); pPr.Append(lns); pPr.Append(sb); pPr.Append(sa); pPr.Append(new A.NoBullet()); p.Append(pPr); p.Append(new A.EndParagraphRunProperties() { Language = "pl-PL", FontSize = 4700, Kerning = 1200 }); tx.Append(p);
                var txXfrm = new Dsp.Transform2D(); txXfrm.Append(new A.Offset() { X = x, Y = y }); txXfrm.Append(new A.Extents() { Cx = cx, Cy = cy });
                sp.Append(nv); sp.Append(spPr); sp.Append(style); sp.Append(tx); sp.Append(txXfrm);
                return sp;
            }

            var rect1 = MakeRect("{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", 0L, 485774L, 1714499L, 1028700L);
            var rect2 = MakeRect("{DF06976E-7188-463E-AE39-BDA19617EFC4}", 1885950L, 485774L, 1714499L, 1028700L);
            var rect3 = MakeRect("{73C7BCEA-927D-4615-95E9-BB89F1A66540}", 3771900L, 485774L, 1714499L, 1028700L);
            var rect4 = MakeRect("{72A7A719-1E8E-46AB-8256-BB41F6817818}", 942975L, 1685925L, 1714499L, 1028700L);
            var rect5 = MakeRect("{2B5DA877-174D-4060-8E30-014EB5090235}", 2828925L, 1685925L, 1714499L, 1028700L);

            shapeTree1.Append(nvGrp); shapeTree1.Append(grpProps);
            shapeTree1.Append(rect1); shapeTree1.Append(rect2); shapeTree1.Append(rect3); shapeTree1.Append(rect4); shapeTree1.Append(rect5);

            drawing2.Append(shapeTree1);
            part.Drawing = drawing2;
        }
    }
}
