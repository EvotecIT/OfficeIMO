using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio;

internal static partial class VisioLegacyXmlCodec {
    // Named singleton rows and cells from the Microsoft Visio 2003 DatadiagramML schema.
    private static readonly Dictionary<string, string[]> SingletonRows = new(StringComparer.Ordinal) {
        ["XForm"] = "PinX PinY Width Height LocPinX LocPinY Angle FlipX FlipY ResizeMode".Split(' '),
        ["Line"] = "LineWeight LineColor LinePattern Rounding EndArrowSize BeginArrow EndArrow LineCap BeginArrowSize LineColorTrans".Split(' '),
        ["Fill"] = "FillForegnd FillBkgnd FillPattern ShdwForegnd ShdwBkgnd ShdwPattern FillForegndTrans FillBkgndTrans ShdwForegndTrans ShdwBkgndTrans ShapeShdwType ShapeShdwOffsetX ShapeShdwOffsetY ShapeShdwObliqueAngle ShapeShdwScaleFactor".Split(' '),
        ["XForm1D"] = "BeginX BeginY EndX EndY".Split(' '),
        ["Event"] = "TheData TheText EventDblClick EventXFMod EventDrop".Split(' '),
        ["LayerMem"] = "LayerMember".Split(' '),
        ["StyleProp"] = "EnableLineProps EnableFillProps EnableTextProps HideForApply".Split(' '),
        ["Foreign"] = "ImgOffsetX ImgOffsetY ImgWidth ImgHeight".Split(' '),
        ["PageProps"] = "PageWidth PageHeight ShdwOffsetX ShdwOffsetY PageScale DrawingScale DrawingSizeType DrawingScaleType InhibitSnap UIVisibility ShdwType ShdwObliqueAngle ShdwScaleFactor".Split(' '),
        ["TextBlock"] = "LeftMargin RightMargin TopMargin BottomMargin VerticalAlign TextBkgnd DefaultTabStop TextDirection TextBkgndTrans".Split(' '),
        ["TextXForm"] = "TxtPinX TxtPinY TxtWidth TxtHeight TxtLocPinX TxtLocPinY TxtAngle".Split(' '),
        ["Align"] = "AlignLeft AlignCenter AlignRight AlignTop AlignMiddle AlignBottom".Split(' '),
        ["Protection"] = "LockWidth LockHeight LockMoveX LockMoveY LockAspect LockDelete LockBegin LockEnd LockRotate LockCrop LockVtxEdit LockTextEdit LockFormat LockGroup LockCalcWH LockSelect LockCustProp".Split(' '),
        ["Help"] = "HelpTopic Copyright".Split(' '),
        ["Misc"] = "NoObjHandles NonPrinting NoCtlHandles NoAlignBox UpdateAlignBox HideText DynFeedback GlueType WalkPreference BegTrigger EndTrigger ObjType Comment IsDropSource NoLiveDynamics LocalizeMerge Calendar LangID ShapeKeywords DropOnPageScale".Split(' '),
        ["RulerGrid"] = "XRulerDensity YRulerDensity XRulerOrigin YRulerOrigin XGridDensity YGridDensity XGridSpacing YGridSpacing XGridOrigin YGridOrigin".Split(' '),
        ["DocProps"] = "OutputFormat LockPreview AddMarkup ViewMarkup PreviewQuality PreviewScope DocLangID".Split(' '),
        ["Image"] = "Gamma Contrast Brightness Sharpen Blur Denoise Transparency".Split(' '),
        ["Group"] = "SelectMode DisplayMode IsDropTarget IsSnapTarget IsTextEditTarget DontMoveChildren".Split(' '),
        ["Layout"] = "ShapePermeableX ShapePermeableY ShapePermeablePlace ShapeFixedCode ShapePlowCode ShapeRouteStyle ConFixedCode ConLineJumpCode ConLineJumpStyle ConLineJumpDirX ConLineJumpDirY ShapePlaceFlip ConLineRouteExt ShapeSplit ShapeSplittable".Split(' '),
        ["PageLayout"] = "ResizePage EnableGrid DynamicsOff CtrlAsInput PlaceStyle RouteStyle PlaceDepth PlowCode LineJumpCode LineJumpStyle PageLineJumpDirX PageLineJumpDirY LineToNodeX LineToNodeY BlockSizeX BlockSizeY AvenueSizeX AvenueSizeY LineToLineX LineToLineY LineJumpFactorX LineJumpFactorY LineAdjustFrom LineAdjustTo PlaceFlip LineRouteExt PageShapeSplit".Split(' '),
        ["PrintProps"] = "PageLeftMargin PageRightMargin PageTopMargin PageBottomMargin ScaleX ScaleY PagesX PagesY CenterX CenterY OnPage PrintGrid PrintPageOrientation PaperKind PaperSource".Split(' '),
    };
    private static readonly Dictionary<string, string> CellRows = SingletonRows
        .SelectMany(row => row.Value.Select(cell => new KeyValuePair<string, string>(cell, row.Key)))
        .GroupBy(item => item.Key, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.First().Value, StringComparer.Ordinal);
    private static readonly Dictionary<string, string> IndexedRows = new(StringComparer.Ordinal) {
        ["Char"] = "Character",
        ["Para"] = "Paragraph",
        ["Scratch"] = "Scratch",
        ["Connection"] = "Connection",
        ["Field"] = "Field",
        ["Control"] = "Control",
        ["Act"] = "Action",
        ["Layer"] = "Layer",
        ["User"] = "User",
        ["Prop"] = "Property",
        ["Hyperlink"] = "Hyperlink",
        ["Reviewer"] = "Reviewer",
        ["Annotation"] = "Annotation",
        ["SmartTagDef"] = "SmartTag",
    };
}
