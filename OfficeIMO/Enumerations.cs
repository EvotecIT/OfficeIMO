using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public enum PropertyTypes : int {
        Undefined,
        YesNo,
        Text,
        DateTime,
        NumberInteger,
        NumberDouble
    }
    public enum CapsStyle {
        /// <summary>
        /// No caps, characters as written.
        /// </summary>
        None,

        /// <summary>
        /// All caps, make every character uppercase.
        /// </summary>
        Caps,

        /// <summary>
        /// Small caps, make all characters capital but with a smaller font size.
        /// </summary>
        SmallCaps
    };

    public enum TableStyle {
        Custom,
        TableNormal,
        TableGrid,
        LightShading,
        LightShadingAccent1,
        LightShadingAccent2,
        LightShadingAccent3,
        LightShadingAccent4,
        LightShadingAccent5,
        LightShadingAccent6,
        LightList,
        LightListAccent1,
        LightListAccent2,
        LightListAccent3,
        LightListAccent4,
        LightListAccent5,
        LightListAccent6,
        LightGrid,
        LightGridAccent1,
        LightGridAccent2,
        LightGridAccent3,
        LightGridAccent4,
        LightGridAccent5,
        LightGridAccent6,
        MediumShading1,
        MediumShading1Accent1,
        MediumShading1Accent2,
        MediumShading1Accent3,
        MediumShading1Accent4,
        MediumShading1Accent5,
        MediumShading1Accent6,
        MediumShading2,
        MediumShading2Accent1,
        MediumShading2Accent2,
        MediumShading2Accent3,
        MediumShading2Accent4,
        MediumShading2Accent5,
        MediumShading2Accent6,
        MediumList1,
        MediumList1Accent1,
        MediumList1Accent2,
        MediumList1Accent3,
        MediumList1Accent4,
        MediumList1Accent5,
        MediumList1Accent6,
        MediumList2,
        MediumList2Accent1,
        MediumList2Accent2,
        MediumList2Accent3,
        MediumList2Accent4,
        MediumList2Accent5,
        MediumList2Accent6,
        MediumGrid1,
        MediumGrid1Accent1,
        MediumGrid1Accent2,
        MediumGrid1Accent3,
        MediumGrid1Accent4,
        MediumGrid1Accent5,
        MediumGrid1Accent6,
        MediumGrid2,
        MediumGrid2Accent1,
        MediumGrid2Accent2,
        MediumGrid2Accent3,
        MediumGrid2Accent4,
        MediumGrid2Accent5,
        MediumGrid2Accent6,
        MediumGrid3,
        MediumGrid3Accent1,
        MediumGrid3Accent2,
        MediumGrid3Accent3,
        MediumGrid3Accent4,
        MediumGrid3Accent5,
        MediumGrid3Accent6,
        DarkList,
        DarkListAccent1,
        DarkListAccent2,
        DarkListAccent3,
        DarkListAccent4,
        DarkListAccent5,
        DarkListAccent6,
        ColorfulShading,
        ColorfulShadingAccent1,
        ColorfulShadingAccent2,
        ColorfulShadingAccent3,
        ColorfulShadingAccent4,
        ColorfulShadingAccent5,
        ColorfulShadingAccent6,
        ColorfulList,
        ColorfulListAccent1,
        ColorfulListAccent2,
        ColorfulListAccent3,
        ColorfulListAccent4,
        ColorfulListAccent5,
        ColorfulListAccent6,
        ColorfulGrid,
        ColorfulGridAccent1,
        ColorfulGridAccent2,
        ColorfulGridAccent3,
        ColorfulGridAccent4,
        ColorfulGridAccent5,
        ColorfulGridAccent6,
        None
    };
}