using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        // Generates content of styleDefinitionsPart1.
        private static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument document) {
            var part = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        private static void CreateAndAddCharacterStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename, string aliases = "") {
            //Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            //Create a new character style and specify some of the attributes.
            Style style = new Style() {
                Type = StyleValues.Character,
                StyleId = styleid,
                CustomStyle = true
            };

            //Create and add the child elements (properties of the style).
            Aliases aliases1 = new Aliases() {Val = aliases};
            StyleName styleName1 = new StyleName() {Val = stylename};
            LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "OverdueAmountPara"};
            if (aliases != "")
                style.Append(aliases1);
            style.Append(styleName1);
            style.Append(linkedStyle1);

            //Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() {ThemeColor = ThemeColorValues.Accent2};
            RunFonts font1 = new RunFonts() {Ascii = "Tahoma"};
            Italic italic1 = new Italic();
            //Specify a 24 point size.
            FontSize fontSize1 = new FontSize() {Val = "48"};
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(italic1);

            //Add the run properties to the style.
            style.Append(styleRunProperties1);

            //Add the style to the styles part.
            styles.Append(style);
        }

        public static void AddDefaultStyleDefinitions(WordprocessingDocument document, StyleDefinitionsPart styleDefinitionsPart1) {
            if (styleDefinitionsPart1 == null) {
                styleDefinitionsPart1 = AddStylesPartToPackage(document);
            }

            DocumentFormat.OpenXml.Wordprocessing.Styles styles1 = new DocumentFormat.OpenXml.Wordprocessing.Styles() {MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "w14 w15 w16se"}};
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1039 = new RunFonts() {AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi};
            Languages languages1 = new Languages() {Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA"};

            runPropertiesBaseStyle1.Append(runFonts1039);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() {DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 372};
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() {Name = "Normal", Locked = true, UiPriority = 0, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() {Name = "heading 1", Locked = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() {Name = "heading 2", Locked = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() {Name = "heading 3", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() {Name = "heading 4", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() {Name = "heading 5", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() {Name = "heading 6", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() {Name = "heading 7", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() {Name = "heading 8", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() {Name = "heading 9", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() {Name = "index 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() {Name = "index 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() {Name = "index 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() {Name = "index 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() {Name = "index 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() {Name = "index 6", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() {Name = "index 7", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() {Name = "index 8", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() {Name = "index 9", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() {Name = "toc 1", Locked = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() {Name = "toc 2", Locked = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() {Name = "toc 3", Locked = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() {Name = "toc 4", Locked = true, UiPriority = 0};
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() {Name = "toc 5", Locked = true, UiPriority = 0};
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() {Name = "toc 6", Locked = true, UiPriority = 0};
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() {Name = "toc 7", Locked = true, UiPriority = 0};
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() {Name = "toc 8", Locked = true, UiPriority = 0};
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() {Name = "toc 9", Locked = true, UiPriority = 0};
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() {Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() {Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() {Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() {Name = "header", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() {Name = "footer", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() {Name = "index heading", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() {Name = "caption", Locked = true, UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() {Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() {Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() {Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() {Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() {Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() {Name = "line number", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() {Name = "page number", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() {Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() {Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() {Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() {Name = "macro", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() {Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() {Name = "List", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() {Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() {Name = "List Number", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() {Name = "List 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() {Name = "List 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() {Name = "List 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() {Name = "List 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() {Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() {Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() {Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() {Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() {Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() {Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() {Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() {Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() {Name = "Title", Locked = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() {Name = "Closing", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() {Name = "Signature", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() {Name = "Default Paragraph Font", Locked = true, UiPriority = 1};
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() {Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() {Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() {Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() {Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() {Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() {Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() {Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() {Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() {Name = "Subtitle", Locked = true, UiPriority = 0, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() {Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() {Name = "Date", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() {Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() {Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() {Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() {Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() {Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() {Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() {Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() {Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() {Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() {Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() {Name = "Strong", Locked = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() {Name = "Emphasis", Locked = true, UiPriority = 0, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() {Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() {Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() {Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() {Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() {Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() {Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() {Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() {Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() {Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() {Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() {Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() {Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() {Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() {Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() {Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() {Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() {Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() {Name = "No List", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() {Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() {Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() {Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() {Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() {Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() {Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() {Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() {Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() {Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() {Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() {Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() {Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() {Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() {Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() {Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() {Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() {Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() {Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() {Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() {Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() {Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() {Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() {Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() {Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() {Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() {Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() {Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() {Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() {Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() {Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() {Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() {Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() {Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() {Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() {Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() {Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() {Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() {Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() {Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() {Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() {Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() {Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() {Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() {Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() {Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() {Name = "Table Grid", Locked = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() {Name = "Placeholder Text", SemiHidden = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() {Name = "No Spacing", PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() {Name = "Light Shading"};
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() {Name = "Light List", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() {Name = "Light Grid", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() {Name = "Medium List 1", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() {Name = "Medium List 2", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() {Name = "Dark List", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() {Name = "Colorful Shading", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() {Name = "Colorful List", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() {Name = "Colorful Grid", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() {Name = "Light Shading Accent 1", UiPriority = 60};
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() {Name = "Light List Accent 1", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() {Name = "Light Grid Accent 1", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1 Accent 1", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2 Accent 1", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() {Name = "Medium List 1 Accent 1", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() {Name = "Revision", SemiHidden = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() {Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() {Name = "Quote", UiPriority = 29, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() {Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() {Name = "Medium List 2 Accent 1", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1 Accent 1", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2 Accent 1", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3 Accent 1", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() {Name = "Dark List Accent 1", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() {Name = "Colorful Shading Accent 1", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() {Name = "Colorful List Accent 1", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() {Name = "Colorful Grid Accent 1", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() {Name = "Light Shading Accent 2", UiPriority = 60};
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() {Name = "Light List Accent 2", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() {Name = "Light Grid Accent 2", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1 Accent 2", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2 Accent 2", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() {Name = "Medium List 1 Accent 2", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() {Name = "Medium List 2 Accent 2", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1 Accent 2", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2 Accent 2", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3 Accent 2", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() {Name = "Dark List Accent 2", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() {Name = "Colorful Shading Accent 2", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() {Name = "Colorful List Accent 2", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() {Name = "Colorful Grid Accent 2", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() {Name = "Light Shading Accent 3", UiPriority = 60};
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() {Name = "Light List Accent 3", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() {Name = "Light Grid Accent 3", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1 Accent 3", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2 Accent 3", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() {Name = "Medium List 1 Accent 3", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() {Name = "Medium List 2 Accent 3", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1 Accent 3", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2 Accent 3", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3 Accent 3", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() {Name = "Dark List Accent 3", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() {Name = "Colorful Shading Accent 3", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() {Name = "Colorful List Accent 3", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() {Name = "Colorful Grid Accent 3", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() {Name = "Light Shading Accent 4", UiPriority = 60};
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() {Name = "Light List Accent 4", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() {Name = "Light Grid Accent 4", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1 Accent 4", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2 Accent 4", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() {Name = "Medium List 1 Accent 4", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() {Name = "Medium List 2 Accent 4", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1 Accent 4", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2 Accent 4", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3 Accent 4", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() {Name = "Dark List Accent 4", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() {Name = "Colorful Shading Accent 4", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() {Name = "Colorful List Accent 4", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() {Name = "Colorful Grid Accent 4", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() {Name = "Light Shading Accent 5", UiPriority = 60};
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() {Name = "Light List Accent 5", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() {Name = "Light Grid Accent 5", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1 Accent 5", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2 Accent 5", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() {Name = "Medium List 1 Accent 5", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() {Name = "Medium List 2 Accent 5", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1 Accent 5", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2 Accent 5", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3 Accent 5", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() {Name = "Dark List Accent 5", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() {Name = "Colorful Shading Accent 5", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() {Name = "Colorful List Accent 5", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() {Name = "Colorful Grid Accent 5", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() {Name = "Light Shading Accent 6", UiPriority = 60};
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() {Name = "Light List Accent 6", UiPriority = 61};
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() {Name = "Light Grid Accent 6", UiPriority = 62};
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() {Name = "Medium Shading 1 Accent 6", UiPriority = 63};
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() {Name = "Medium Shading 2 Accent 6", UiPriority = 64};
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() {Name = "Medium List 1 Accent 6", UiPriority = 65};
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() {Name = "Medium List 2 Accent 6", UiPriority = 66};
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() {Name = "Medium Grid 1 Accent 6", UiPriority = 67};
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() {Name = "Medium Grid 2 Accent 6", UiPriority = 68};
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() {Name = "Medium Grid 3 Accent 6", UiPriority = 69};
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() {Name = "Dark List Accent 6", UiPriority = 70};
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() {Name = "Colorful Shading Accent 6", UiPriority = 71};
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() {Name = "Colorful List Accent 6", UiPriority = 72};
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() {Name = "Colorful Grid Accent 6", UiPriority = 73};
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() {Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() {Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() {Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() {Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() {Name = "Book Title", UiPriority = 33, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() {Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() {Name = "TOC Heading", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() {Name = "Plain Table 1", UiPriority = 41};
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() {Name = "Plain Table 2", UiPriority = 42};
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() {Name = "Plain Table 3", UiPriority = 43};
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() {Name = "Plain Table 4", UiPriority = 44};
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() {Name = "Plain Table 5", UiPriority = 45};
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() {Name = "Grid Table Light", UiPriority = 40};
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() {Name = "Grid Table 2", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() {Name = "Grid Table 3", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() {Name = "Grid Table 4", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light Accent 1", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() {Name = "Grid Table 2 Accent 1", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() {Name = "Grid Table 3 Accent 1", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() {Name = "Grid Table 4 Accent 1", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark Accent 1", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light Accent 2", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() {Name = "Grid Table 2 Accent 2", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() {Name = "Grid Table 3 Accent 2", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() {Name = "Grid Table 4 Accent 2", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark Accent 2", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light Accent 3", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() {Name = "Grid Table 2 Accent 3", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() {Name = "Grid Table 3 Accent 3", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() {Name = "Grid Table 4 Accent 3", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark Accent 3", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light Accent 4", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() {Name = "Grid Table 2 Accent 4", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() {Name = "Grid Table 3 Accent 4", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() {Name = "Grid Table 4 Accent 4", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark Accent 4", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light Accent 5", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() {Name = "Grid Table 2 Accent 5", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() {Name = "Grid Table 3 Accent 5", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() {Name = "Grid Table 4 Accent 5", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark Accent 5", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() {Name = "Grid Table 1 Light Accent 6", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() {Name = "Grid Table 2 Accent 6", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() {Name = "Grid Table 3 Accent 6", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() {Name = "Grid Table 4 Accent 6", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() {Name = "Grid Table 5 Dark Accent 6", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() {Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() {Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() {Name = "List Table 2", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() {Name = "List Table 3", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() {Name = "List Table 4", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light Accent 1", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() {Name = "List Table 2 Accent 1", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() {Name = "List Table 3 Accent 1", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() {Name = "List Table 4 Accent 1", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark Accent 1", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful Accent 1", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful Accent 1", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light Accent 2", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() {Name = "List Table 2 Accent 2", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() {Name = "List Table 3 Accent 2", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() {Name = "List Table 4 Accent 2", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark Accent 2", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful Accent 2", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful Accent 2", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light Accent 3", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() {Name = "List Table 2 Accent 3", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() {Name = "List Table 3 Accent 3", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() {Name = "List Table 4 Accent 3", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark Accent 3", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful Accent 3", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful Accent 3", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light Accent 4", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() {Name = "List Table 2 Accent 4", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() {Name = "List Table 3 Accent 4", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() {Name = "List Table 4 Accent 4", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark Accent 4", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful Accent 4", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful Accent 4", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light Accent 5", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() {Name = "List Table 2 Accent 5", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() {Name = "List Table 3 Accent 5", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() {Name = "List Table 4 Accent 5", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark Accent 5", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful Accent 5", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful Accent 5", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() {Name = "List Table 1 Light Accent 6", UiPriority = 46};
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() {Name = "List Table 2 Accent 6", UiPriority = 47};
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() {Name = "List Table 3 Accent 6", UiPriority = 48};
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() {Name = "List Table 4 Accent 6", UiPriority = 49};
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() {Name = "List Table 5 Dark Accent 6", UiPriority = 50};
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() {Name = "List Table 6 Colorful Accent 6", UiPriority = 51};
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() {Name = "List Table 7 Colorful Accent 6", UiPriority = 52};
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() {Name = "Mention", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() {Name = "Smart Hyperlink", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() {Name = "Hashtag", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo372 = new LatentStyleExceptionInfo() {Name = "Unresolved Mention", SemiHidden = true, UnhideWhenUsed = true};
            LatentStyleExceptionInfo latentStyleExceptionInfo373 = new LatentStyleExceptionInfo() {Name = "Smart Link", SemiHidden = true, UnhideWhenUsed = true};

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);
            latentStyles1.Append(latentStyleExceptionInfo372);
            latentStyles1.Append(latentStyleExceptionInfo373);

            Style style1 = new Style() {Type = StyleValues.Paragraph, StyleId = "Normal", Default = true};
            StyleName styleName1 = new StyleName() {Val = "Normal"};
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid384 = new Rsid() {Val = "007756BD"};

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1056 = new FontSize() {Val = "24"};
            FontSizeComplexScript fontSizeComplexScript1042 = new FontSizeComplexScript() {Val = "24"};

            styleRunProperties1.Append(fontSize1056);
            styleRunProperties1.Append(fontSizeComplexScript1042);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid384);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading1"};
            StyleName styleName2 = new StyleName() {Val = "heading 1"};
            Aliases aliases1 = new Aliases() {Val = "H1,DO NOT USE_h1,H11,H12,H111,H13,H112,H121,H1111,H14"};
            BasedOn basedOn1 = new BasedOn() {Val = "Normal"};
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() {Val = "Normal"};
            LinkedStyle linkedStyle1 = new LinkedStyle() {Val = "Heading1Char"};
            UIPriority uIPriority1 = new UIPriority() {Val = 99};
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid385 = new Rsid() {Val = "005244EE"};

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() {Before = "240", After = "60"};
            OutlineLevel outlineLevel1 = new OutlineLevel() {Val = 0};

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(spacingBetweenLines2);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            Bold bold119 = new Bold();
            Kern kern1 = new Kern() {Val = (UInt32Value) 28U};
            FontSize fontSize1057 = new FontSize() {Val = "32"};
            FontSizeComplexScript fontSizeComplexScript1043 = new FontSizeComplexScript() {Val = "20"};

            styleRunProperties2.Append(bold119);
            styleRunProperties2.Append(kern1);
            styleRunProperties2.Append(fontSize1057);
            styleRunProperties2.Append(fontSizeComplexScript1043);

            style2.Append(styleName2);
            style2.Append(aliases1);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(uIPriority1);
            style2.Append(primaryStyle2);
            style2.Append(rsid385);
            style2.Append(styleParagraphProperties1);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() {Type = StyleValues.Paragraph, StyleId = "Heading2"};
            StyleName styleName3 = new StyleName() {Val = "heading 2"};
            BasedOn basedOn2 = new BasedOn() {Val = "Normal"};
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() {Val = "Normal"};
            LinkedStyle linkedStyle2 = new LinkedStyle() {Val = "Heading2Char"};
            UIPriority uIPriority2 = new UIPriority() {Val = 99};
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Locked locked1 = new Locked();
            Rsid rsid386 = new Rsid() {Val = "0077670E"};

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() {Before = "240", After = "60"};
            OutlineLevel outlineLevel2 = new OutlineLevel() {Val = 1};

            styleParagraphProperties2.Append(keepNext2);
            styleParagraphProperties2.Append(spacingBetweenLines3);
            styleParagraphProperties2.Append(outlineLevel2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts1040 = new RunFonts() {Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial"};
            Bold bold120 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();
            Italic italic8 = new Italic();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            FontSize fontSize1058 = new FontSize() {Val = "28"};
            FontSizeComplexScript fontSizeComplexScript1044 = new FontSizeComplexScript() {Val = "28"};

            styleRunProperties3.Append(runFonts1040);
            styleRunProperties3.Append(bold120);
            styleRunProperties3.Append(boldComplexScript14);
            styleRunProperties3.Append(italic8);
            styleRunProperties3.Append(italicComplexScript4);
            styleRunProperties3.Append(fontSize1058);
            styleRunProperties3.Append(fontSizeComplexScript1044);

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(nextParagraphStyle2);
            style3.Append(linkedStyle2);
            style3.Append(uIPriority2);
            style3.Append(primaryStyle3);
            style3.Append(locked1);
            style3.Append(rsid386);
            style3.Append(styleParagraphProperties2);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() {Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true};
            StyleName styleName4 = new StyleName() {Val = "Default Paragraph Font"};
            UIPriority uIPriority3 = new UIPriority() {Val = 1};
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden1);
            style4.Append(unhideWhenUsed1);

            Style style5 = new Style() {Type = StyleValues.Table, StyleId = "TableNormal", Default = true};
            StyleName styleName5 = new StyleName() {Val = "Normal Table"};
            UIPriority uIPriority4 = new UIPriority() {Val = 99};
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation5 = new TableIndentation() {Width = 0, Type = TableWidthUnitValues.Dxa};

            TableCellMarginDefault tableCellMarginDefault4 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() {Width = "0", Type = TableWidthUnitValues.Dxa};
            TableCellLeftMargin tableCellLeftMargin4 = new TableCellLeftMargin() {Width = 108, Type = TableWidthValues.Dxa};
            BottomMargin bottomMargin1 = new BottomMargin() {Width = "0", Type = TableWidthUnitValues.Dxa};
            TableCellRightMargin tableCellRightMargin4 = new TableCellRightMargin() {Width = 108, Type = TableWidthValues.Dxa};

            tableCellMarginDefault4.Append(topMargin1);
            tableCellMarginDefault4.Append(tableCellLeftMargin4);
            tableCellMarginDefault4.Append(bottomMargin1);
            tableCellMarginDefault4.Append(tableCellRightMargin4);

            styleTableProperties1.Append(tableIndentation5);
            styleTableProperties1.Append(tableCellMarginDefault4);

            style5.Append(styleName5);
            style5.Append(uIPriority4);
            style5.Append(semiHidden2);
            style5.Append(unhideWhenUsed2);
            style5.Append(styleTableProperties1);

            Style style6 = new Style() {Type = StyleValues.Numbering, StyleId = "NoList", Default = true};
            StyleName styleName6 = new StyleName() {Val = "No List"};
            UIPriority uIPriority5 = new UIPriority() {Val = 99};
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style6.Append(styleName6);
            style6.Append(uIPriority5);
            style6.Append(semiHidden3);
            style6.Append(unhideWhenUsed3);

            Style style7 = new Style() {Type = StyleValues.Character, StyleId = "Heading1Char", CustomStyle = true};
            StyleName styleName7 = new StyleName() {Val = "Heading 1 Char"};
            Aliases aliases2 = new Aliases() {Val = "H1 Char,DO NOT USE_h1 Char,H11 Char,H12 Char,H111 Char,H13 Char,H112 Char,H121 Char,H1111 Char,H14 Char"};
            LinkedStyle linkedStyle3 = new LinkedStyle() {Val = "Heading1"};
            UIPriority uIPriority6 = new UIPriority() {Val = 99};
            Locked locked2 = new Locked();
            Rsid rsid387 = new Rsid() {Val = "00972E14"};

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts1041 = new RunFonts() {Ascii = "Cambria", HighAnsi = "Cambria", ComplexScript = "Times New Roman"};
            Bold bold121 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            Kern kern2 = new Kern() {Val = (UInt32Value) 32U};
            FontSize fontSize1059 = new FontSize() {Val = "32"};
            FontSizeComplexScript fontSizeComplexScript1045 = new FontSizeComplexScript() {Val = "32"};

            styleRunProperties4.Append(runFonts1041);
            styleRunProperties4.Append(bold121);
            styleRunProperties4.Append(boldComplexScript15);
            styleRunProperties4.Append(kern2);
            styleRunProperties4.Append(fontSize1059);
            styleRunProperties4.Append(fontSizeComplexScript1045);

            style7.Append(styleName7);
            style7.Append(aliases2);
            style7.Append(linkedStyle3);
            style7.Append(uIPriority6);
            style7.Append(locked2);
            style7.Append(rsid387);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() {Type = StyleValues.Character, StyleId = "Heading2Char", CustomStyle = true};
            StyleName styleName8 = new StyleName() {Val = "Heading 2 Char"};
            LinkedStyle linkedStyle4 = new LinkedStyle() {Val = "Heading2"};
            UIPriority uIPriority7 = new UIPriority() {Val = 99};
            SemiHidden semiHidden4 = new SemiHidden();
            Locked locked3 = new Locked();
            Rsid rsid388 = new Rsid() {Val = "00B43BFA"};

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts1042 = new RunFonts() {Ascii = "Cambria", HighAnsi = "Cambria", ComplexScript = "Times New Roman"};
            Bold bold122 = new Bold();
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();
            Italic italic9 = new Italic();
            ItalicComplexScript italicComplexScript5 = new ItalicComplexScript();
            FontSize fontSize1060 = new FontSize() {Val = "28"};
            FontSizeComplexScript fontSizeComplexScript1046 = new FontSizeComplexScript() {Val = "28"};

            styleRunProperties5.Append(runFonts1042);
            styleRunProperties5.Append(bold122);
            styleRunProperties5.Append(boldComplexScript16);
            styleRunProperties5.Append(italic9);
            styleRunProperties5.Append(italicComplexScript5);
            styleRunProperties5.Append(fontSize1060);
            styleRunProperties5.Append(fontSizeComplexScript1046);

            style8.Append(styleName8);
            style8.Append(linkedStyle4);
            style8.Append(uIPriority7);
            style8.Append(semiHidden4);
            style8.Append(locked3);
            style8.Append(rsid388);
            style8.Append(styleRunProperties5);

            Style style9 = new Style() {Type = StyleValues.Paragraph, StyleId = "Header"};
            StyleName styleName9 = new StyleName() {Val = "header"};
            BasedOn basedOn3 = new BasedOn() {Val = "Normal"};
            LinkedStyle linkedStyle5 = new LinkedStyle() {Val = "HeaderChar"};
            UIPriority uIPriority8 = new UIPriority() {Val = 99};
            Rsid rsid389 = new Rsid() {Val = "000F55C6"};

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs9 = new Tabs();
            TabStop tabStop9 = new TabStop() {Val = TabStopValues.Center, Position = 4320};
            TabStop tabStop10 = new TabStop() {Val = TabStopValues.Right, Position = 8640};

            tabs9.Append(tabStop9);
            tabs9.Append(tabStop10);

            styleParagraphProperties3.Append(tabs9);

            style9.Append(styleName9);
            style9.Append(basedOn3);
            style9.Append(linkedStyle5);
            style9.Append(uIPriority8);
            style9.Append(rsid389);
            style9.Append(styleParagraphProperties3);

            Style style10 = new Style() {Type = StyleValues.Character, StyleId = "HeaderChar", CustomStyle = true};
            StyleName styleName10 = new StyleName() {Val = "Header Char"};
            LinkedStyle linkedStyle6 = new LinkedStyle() {Val = "Header"};
            UIPriority uIPriority9 = new UIPriority() {Val = 99};
            SemiHidden semiHidden5 = new SemiHidden();
            Locked locked4 = new Locked();
            Rsid rsid390 = new Rsid() {Val = "00972E14"};

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts1043 = new RunFonts() {ComplexScript = "Times New Roman"};
            FontSize fontSize1061 = new FontSize() {Val = "24"};
            FontSizeComplexScript fontSizeComplexScript1047 = new FontSizeComplexScript() {Val = "24"};

            styleRunProperties6.Append(runFonts1043);
            styleRunProperties6.Append(fontSize1061);
            styleRunProperties6.Append(fontSizeComplexScript1047);

            style10.Append(styleName10);
            style10.Append(linkedStyle6);
            style10.Append(uIPriority9);
            style10.Append(semiHidden5);
            style10.Append(locked4);
            style10.Append(rsid390);
            style10.Append(styleRunProperties6);

            Style style11 = new Style() {Type = StyleValues.Paragraph, StyleId = "Footer"};
            StyleName styleName11 = new StyleName() {Val = "footer"};
            BasedOn basedOn4 = new BasedOn() {Val = "Normal"};
            LinkedStyle linkedStyle7 = new LinkedStyle() {Val = "FooterChar"};
            UIPriority uIPriority10 = new UIPriority() {Val = 99};
            Rsid rsid391 = new Rsid() {Val = "000F55C6"};

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();

            Tabs tabs10 = new Tabs();
            TabStop tabStop11 = new TabStop() {Val = TabStopValues.Center, Position = 4320};
            TabStop tabStop12 = new TabStop() {Val = TabStopValues.Right, Position = 8640};

            tabs10.Append(tabStop11);
            tabs10.Append(tabStop12);

            styleParagraphProperties4.Append(tabs10);

            style11.Append(styleName11);
            style11.Append(basedOn4);
            style11.Append(linkedStyle7);
            style11.Append(uIPriority10);
            style11.Append(rsid391);
            style11.Append(styleParagraphProperties4);

            Style style12 = new Style() {Type = StyleValues.Character, StyleId = "FooterChar", CustomStyle = true};
            StyleName styleName12 = new StyleName() {Val = "Footer Char"};
            LinkedStyle linkedStyle8 = new LinkedStyle() {Val = "Footer"};
            UIPriority uIPriority11 = new UIPriority() {Val = 99};
            Locked locked5 = new Locked();
            Rsid rsid392 = new Rsid() {Val = "00972E14"};

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts1044 = new RunFonts() {ComplexScript = "Times New Roman"};
            FontSize fontSize1062 = new FontSize() {Val = "24"};
            FontSizeComplexScript fontSizeComplexScript1048 = new FontSizeComplexScript() {Val = "24"};

            styleRunProperties7.Append(runFonts1044);
            styleRunProperties7.Append(fontSize1062);
            styleRunProperties7.Append(fontSizeComplexScript1048);

            style12.Append(styleName12);
            style12.Append(linkedStyle8);
            style12.Append(uIPriority11);
            style12.Append(locked5);
            style12.Append(rsid392);
            style12.Append(styleRunProperties7);

            Style style13 = new Style() {Type = StyleValues.Character, StyleId = "PageNumber"};
            StyleName styleName13 = new StyleName() {Val = "page number"};
            UIPriority uIPriority12 = new UIPriority() {Val = 99};
            Rsid rsid393 = new Rsid() {Val = "000F55C6"};

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts1045 = new RunFonts() {ComplexScript = "Times New Roman"};

            styleRunProperties8.Append(runFonts1045);

            style13.Append(styleName13);
            style13.Append(uIPriority12);
            style13.Append(rsid393);
            style13.Append(styleRunProperties8);

            Style style14 = new Style() {Type = StyleValues.Paragraph, StyleId = "BodyText"};
            StyleName styleName14 = new StyleName() {Val = "Body Text"};
            BasedOn basedOn5 = new BasedOn() {Val = "Normal"};
            LinkedStyle linkedStyle9 = new LinkedStyle() {Val = "BodyTextChar"};
            UIPriority uIPriority13 = new UIPriority() {Val = 99};
            Rsid rsid394 = new Rsid() {Val = "00076FEE"};

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() {Before = "120", After = "120"};

            styleParagraphProperties5.Append(spacingBetweenLines4);

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            FontSize fontSize1063 = new FontSize() {Val = "22"};
            FontSizeComplexScript fontSizeComplexScript1049 = new FontSizeComplexScript() {Val = "20"};

            styleRunProperties9.Append(fontSize1063);
            styleRunProperties9.Append(fontSizeComplexScript1049);

            style14.Append(styleName14);
            style14.Append(basedOn5);
            style14.Append(linkedStyle9);
            style14.Append(uIPriority13);
            style14.Append(rsid394);
            style14.Append(styleParagraphProperties5);
            style14.Append(styleRunProperties9);

            Style style15 = new Style() {Type = StyleValues.Character, StyleId = "BodyTextChar", CustomStyle = true};
            StyleName styleName15 = new StyleName() {Val = "Body Text Char"};
            LinkedStyle linkedStyle10 = new LinkedStyle() {Val = "BodyText"};
            UIPriority uIPriority14 = new UIPriority() {Val = 99};
            Locked locked6 = new Locked();
            Rsid rsid395 = new Rsid() {Val = "00076FEE"};

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts1046 = new RunFonts() {ComplexScript = "Times New Roman"};
            FontSize fontSize1064 = new FontSize() {Val = "22"};

            styleRunProperties10.Append(runFonts1046);
            styleRunProperties10.Append(fontSize1064);

            style15.Append(styleName15);
            style15.Append(linkedStyle10);
            style15.Append(uIPriority14);
            style15.Append(locked6);
            style15.Append(rsid395);
            style15.Append(styleRunProperties10);

            Style style16 = new Style() {Type = StyleValues.Paragraph, StyleId = "Title2", CustomStyle = true};
            StyleName styleName16 = new StyleName() {Val = "Title 2"};
            BasedOn basedOn6 = new BasedOn() {Val = "Title"};
            UIPriority uIPriority15 = new UIPriority() {Val = 99};
            Rsid rsid396 = new Rsid() {Val = "00076FEE"};

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() {Val = false};
            AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() {Val = false};
            AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() {Val = false};
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() {Before = "120", After = "120"};
            OutlineLevel outlineLevel3 = new OutlineLevel() {Val = 9};

            styleParagraphProperties6.Append(autoSpaceDE2);
            styleParagraphProperties6.Append(autoSpaceDN2);
            styleParagraphProperties6.Append(adjustRightIndent2);
            styleParagraphProperties6.Append(spacingBetweenLines5);
            styleParagraphProperties6.Append(outlineLevel3);

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts1047 = new RunFonts() {Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial"};
            Kern kern3 = new Kern() {Val = (UInt32Value) 0U};
            FontSize fontSize1065 = new FontSize() {Val = "28"};

            styleRunProperties11.Append(runFonts1047);
            styleRunProperties11.Append(kern3);
            styleRunProperties11.Append(fontSize1065);

            style16.Append(styleName16);
            style16.Append(basedOn6);
            style16.Append(uIPriority15);
            style16.Append(rsid396);
            style16.Append(styleParagraphProperties6);
            style16.Append(styleRunProperties11);

            Style style17 = new Style() {Type = StyleValues.Paragraph, StyleId = "TableHeading", CustomStyle = true};
            StyleName styleName17 = new StyleName() {Val = "Table Heading"};
            BasedOn basedOn7 = new BasedOn() {Val = "BodyText"};
            UIPriority uIPriority16 = new UIPriority() {Val = 99};
            Rsid rsid397 = new Rsid() {Val = "00076FEE"};

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() {Before = "60", After = "60"};

            styleParagraphProperties7.Append(spacingBetweenLines6);

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts1048 = new RunFonts() {Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial"};
            Bold bold123 = new Bold();
            FontSizeComplexScript fontSizeComplexScript1050 = new FontSizeComplexScript() {Val = "22"};

            styleRunProperties12.Append(runFonts1048);
            styleRunProperties12.Append(bold123);
            styleRunProperties12.Append(fontSizeComplexScript1050);

            style17.Append(styleName17);
            style17.Append(basedOn7);
            style17.Append(uIPriority16);
            style17.Append(rsid397);
            style17.Append(styleParagraphProperties7);
            style17.Append(styleRunProperties12);

            Style style18 = new Style() {Type = StyleValues.Paragraph, StyleId = "TableText", CustomStyle = true};
            StyleName styleName18 = new StyleName() {Val = "Table Text"};
            BasedOn basedOn8 = new BasedOn() {Val = "BodyText"};
            UIPriority uIPriority17 = new UIPriority() {Val = 99};
            Rsid rsid398 = new Rsid() {Val = "00076FEE"};

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() {Before = "60", After = "60"};

            styleParagraphProperties8.Append(spacingBetweenLines7);

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts1049 = new RunFonts() {Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial"};

            styleRunProperties13.Append(runFonts1049);

            style18.Append(styleName18);
            style18.Append(basedOn8);
            style18.Append(uIPriority17);
            style18.Append(rsid398);
            style18.Append(styleParagraphProperties8);
            style18.Append(styleRunProperties13);

            Style style19 = new Style() {Type = StyleValues.Paragraph, StyleId = "Title"};
            StyleName styleName19 = new StyleName() {Val = "Title"};
            BasedOn basedOn9 = new BasedOn() {Val = "Normal"};
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() {Val = "Normal"};
            LinkedStyle linkedStyle11 = new LinkedStyle() {Val = "TitleChar"};
            UIPriority uIPriority18 = new UIPriority() {Val = 99};
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid399 = new Rsid() {Val = "00076FEE"};

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() {Before = "240", After = "60"};
            Justification justification78 = new Justification() {Val = JustificationValues.Center};
            OutlineLevel outlineLevel4 = new OutlineLevel() {Val = 0};

            styleParagraphProperties9.Append(spacingBetweenLines8);
            styleParagraphProperties9.Append(justification78);
            styleParagraphProperties9.Append(outlineLevel4);

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts1050 = new RunFonts() {Ascii = "Cambria", HighAnsi = "Cambria"};
            Bold bold124 = new Bold();
            BoldComplexScript boldComplexScript17 = new BoldComplexScript();
            Kern kern4 = new Kern() {Val = (UInt32Value) 28U};
            FontSize fontSize1066 = new FontSize() {Val = "32"};
            FontSizeComplexScript fontSizeComplexScript1051 = new FontSizeComplexScript() {Val = "32"};

            styleRunProperties14.Append(runFonts1050);
            styleRunProperties14.Append(bold124);
            styleRunProperties14.Append(boldComplexScript17);
            styleRunProperties14.Append(kern4);
            styleRunProperties14.Append(fontSize1066);
            styleRunProperties14.Append(fontSizeComplexScript1051);

            style19.Append(styleName19);
            style19.Append(basedOn9);
            style19.Append(nextParagraphStyle3);
            style19.Append(linkedStyle11);
            style19.Append(uIPriority18);
            style19.Append(primaryStyle4);
            style19.Append(rsid399);
            style19.Append(styleParagraphProperties9);
            style19.Append(styleRunProperties14);

            Style style20 = new Style() {Type = StyleValues.Character, StyleId = "TitleChar", CustomStyle = true};
            StyleName styleName20 = new StyleName() {Val = "Title Char"};
            LinkedStyle linkedStyle12 = new LinkedStyle() {Val = "Title"};
            UIPriority uIPriority19 = new UIPriority() {Val = 99};
            Locked locked7 = new Locked();
            Rsid rsid400 = new Rsid() {Val = "00076FEE"};

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            RunFonts runFonts1051 = new RunFonts() {Ascii = "Cambria", HighAnsi = "Cambria", ComplexScript = "Times New Roman"};
            Bold bold125 = new Bold();
            BoldComplexScript boldComplexScript18 = new BoldComplexScript();
            Kern kern5 = new Kern() {Val = (UInt32Value) 28U};
            FontSize fontSize1067 = new FontSize() {Val = "32"};
            FontSizeComplexScript fontSizeComplexScript1052 = new FontSizeComplexScript() {Val = "32"};

            styleRunProperties15.Append(runFonts1051);
            styleRunProperties15.Append(bold125);
            styleRunProperties15.Append(boldComplexScript18);
            styleRunProperties15.Append(kern5);
            styleRunProperties15.Append(fontSize1067);
            styleRunProperties15.Append(fontSizeComplexScript1052);

            style20.Append(styleName20);
            style20.Append(linkedStyle12);
            style20.Append(uIPriority19);
            style20.Append(locked7);
            style20.Append(rsid400);
            style20.Append(styleRunProperties15);

            Style style21 = new Style() {Type = StyleValues.Paragraph, StyleId = "TOCHeading"};
            StyleName styleName21 = new StyleName() {Val = "TOC Heading"};
            BasedOn basedOn10 = new BasedOn() {Val = "Heading1"};
            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle() {Val = "Normal"};
            UIPriority uIPriority20 = new UIPriority() {Val = 99};
            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            Rsid rsid401 = new Rsid() {Val = "00513B3F"};

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            KeepLines keepLines2 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() {Before = "480", After = "0", Line = "276", LineRule = LineSpacingRuleValues.Auto};
            OutlineLevel outlineLevel5 = new OutlineLevel() {Val = 9};

            styleParagraphProperties10.Append(keepLines2);
            styleParagraphProperties10.Append(spacingBetweenLines9);
            styleParagraphProperties10.Append(outlineLevel5);

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts1052 = new RunFonts() {Ascii = "Cambria", HighAnsi = "Cambria"};
            BoldComplexScript boldComplexScript19 = new BoldComplexScript();
            Color color375 = new Color() {Val = "365F91"};
            Kern kern6 = new Kern() {Val = (UInt32Value) 0U};
            FontSize fontSize1068 = new FontSize() {Val = "28"};
            FontSizeComplexScript fontSizeComplexScript1053 = new FontSizeComplexScript() {Val = "28"};

            styleRunProperties16.Append(runFonts1052);
            styleRunProperties16.Append(boldComplexScript19);
            styleRunProperties16.Append(color375);
            styleRunProperties16.Append(kern6);
            styleRunProperties16.Append(fontSize1068);
            styleRunProperties16.Append(fontSizeComplexScript1053);

            style21.Append(styleName21);
            style21.Append(basedOn10);
            style21.Append(nextParagraphStyle4);
            style21.Append(uIPriority20);
            style21.Append(primaryStyle5);
            style21.Append(rsid401);
            style21.Append(styleParagraphProperties10);
            style21.Append(styleRunProperties16);

            Style style22 = new Style() {Type = StyleValues.Paragraph, StyleId = "TOC2"};
            StyleName styleName22 = new StyleName() {Val = "toc 2"};
            BasedOn basedOn11 = new BasedOn() {Val = "Normal"};
            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle() {Val = "Normal"};
            AutoRedefine autoRedefine1 = new AutoRedefine();
            UIPriority uIPriority21 = new UIPriority() {Val = 99};
            SemiHidden semiHidden6 = new SemiHidden();
            Rsid rsid402 = new Rsid() {Val = "00513B3F"};

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() {After = "100", Line = "276", LineRule = LineSpacingRuleValues.Auto};
            Indentation indentation26 = new Indentation() {Start = "220"};

            styleParagraphProperties11.Append(spacingBetweenLines10);
            styleParagraphProperties11.Append(indentation26);

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts1053 = new RunFonts() {Ascii = "Calibri", HighAnsi = "Calibri"};
            FontSize fontSize1069 = new FontSize() {Val = "22"};
            FontSizeComplexScript fontSizeComplexScript1054 = new FontSizeComplexScript() {Val = "22"};

            styleRunProperties17.Append(runFonts1053);
            styleRunProperties17.Append(fontSize1069);
            styleRunProperties17.Append(fontSizeComplexScript1054);

            style22.Append(styleName22);
            style22.Append(basedOn11);
            style22.Append(nextParagraphStyle5);
            style22.Append(autoRedefine1);
            style22.Append(uIPriority21);
            style22.Append(semiHidden6);
            style22.Append(rsid402);
            style22.Append(styleParagraphProperties11);
            style22.Append(styleRunProperties17);

            Style style23 = new Style() {Type = StyleValues.Paragraph, StyleId = "TOC1"};
            StyleName styleName23 = new StyleName() {Val = "toc 1"};
            BasedOn basedOn12 = new BasedOn() {Val = "Normal"};
            NextParagraphStyle nextParagraphStyle6 = new NextParagraphStyle() {Val = "Normal"};
            AutoRedefine autoRedefine2 = new AutoRedefine();
            UIPriority uIPriority22 = new UIPriority() {Val = 99};
            SemiHidden semiHidden7 = new SemiHidden();
            Rsid rsid403 = new Rsid() {Val = "009E06D4"};

            StyleParagraphProperties styleParagraphProperties12 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders2 = new ParagraphBorders();
            BottomBorder bottomBorder99 = new BottomBorder() {Val = BorderValues.Single, Color = "auto", Size = (UInt32Value) 12U, Space = (UInt32Value) 1U};

            paragraphBorders2.Append(bottomBorder99);

            Tabs tabs11 = new Tabs();
            TabStop tabStop13 = new TabStop() {Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9360};

            tabs11.Append(tabStop13);
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() {After = "100", Line = "276", LineRule = LineSpacingRuleValues.Auto};
            OutlineLevel outlineLevel6 = new OutlineLevel() {Val = 0};

            styleParagraphProperties12.Append(paragraphBorders2);
            styleParagraphProperties12.Append(tabs11);
            styleParagraphProperties12.Append(spacingBetweenLines11);
            styleParagraphProperties12.Append(outlineLevel6);

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            RunFonts runFonts1054 = new RunFonts() {Ascii = "Verdana", HighAnsi = "Verdana"};
            Bold bold126 = new Bold();
            FontSize fontSize1070 = new FontSize() {Val = "21"};
            FontSizeComplexScript fontSizeComplexScript1055 = new FontSizeComplexScript() {Val = "21"};

            styleRunProperties18.Append(runFonts1054);
            styleRunProperties18.Append(bold126);
            styleRunProperties18.Append(fontSize1070);
            styleRunProperties18.Append(fontSizeComplexScript1055);

            style23.Append(styleName23);
            style23.Append(basedOn12);
            style23.Append(nextParagraphStyle6);
            style23.Append(autoRedefine2);
            style23.Append(uIPriority22);
            style23.Append(semiHidden7);
            style23.Append(rsid403);
            style23.Append(styleParagraphProperties12);
            style23.Append(styleRunProperties18);

            Style style24 = new Style() {Type = StyleValues.Paragraph, StyleId = "TOC3"};
            StyleName styleName24 = new StyleName() {Val = "toc 3"};
            BasedOn basedOn13 = new BasedOn() {Val = "Normal"};
            NextParagraphStyle nextParagraphStyle7 = new NextParagraphStyle() {Val = "Normal"};
            AutoRedefine autoRedefine3 = new AutoRedefine();
            UIPriority uIPriority23 = new UIPriority() {Val = 99};
            SemiHidden semiHidden8 = new SemiHidden();
            Rsid rsid404 = new Rsid() {Val = "00513B3F"};

            StyleParagraphProperties styleParagraphProperties13 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() {After = "100", Line = "276", LineRule = LineSpacingRuleValues.Auto};
            Indentation indentation27 = new Indentation() {Start = "440"};

            styleParagraphProperties13.Append(spacingBetweenLines12);
            styleParagraphProperties13.Append(indentation27);

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            RunFonts runFonts1055 = new RunFonts() {Ascii = "Calibri", HighAnsi = "Calibri"};
            FontSize fontSize1071 = new FontSize() {Val = "22"};
            FontSizeComplexScript fontSizeComplexScript1056 = new FontSizeComplexScript() {Val = "22"};

            styleRunProperties19.Append(runFonts1055);
            styleRunProperties19.Append(fontSize1071);
            styleRunProperties19.Append(fontSizeComplexScript1056);

            style24.Append(styleName24);
            style24.Append(basedOn13);
            style24.Append(nextParagraphStyle7);
            style24.Append(autoRedefine3);
            style24.Append(uIPriority23);
            style24.Append(semiHidden8);
            style24.Append(rsid404);
            style24.Append(styleParagraphProperties13);
            style24.Append(styleRunProperties19);

            Style style25 = new Style() {Type = StyleValues.Paragraph, StyleId = "BalloonText"};
            StyleName styleName25 = new StyleName() {Val = "Balloon Text"};
            BasedOn basedOn14 = new BasedOn() {Val = "Normal"};
            LinkedStyle linkedStyle13 = new LinkedStyle() {Val = "BalloonTextChar"};
            UIPriority uIPriority24 = new UIPriority() {Val = 99};
            SemiHidden semiHidden9 = new SemiHidden();
            Rsid rsid405 = new Rsid() {Val = "00513B3F"};

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            RunFonts runFonts1056 = new RunFonts() {Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma"};
            FontSize fontSize1072 = new FontSize() {Val = "16"};
            FontSizeComplexScript fontSizeComplexScript1057 = new FontSizeComplexScript() {Val = "16"};

            styleRunProperties20.Append(runFonts1056);
            styleRunProperties20.Append(fontSize1072);
            styleRunProperties20.Append(fontSizeComplexScript1057);

            style25.Append(styleName25);
            style25.Append(basedOn14);
            style25.Append(linkedStyle13);
            style25.Append(uIPriority24);
            style25.Append(semiHidden9);
            style25.Append(rsid405);
            style25.Append(styleRunProperties20);

            Style style26 = new Style() {Type = StyleValues.Character, StyleId = "BalloonTextChar", CustomStyle = true};
            StyleName styleName26 = new StyleName() {Val = "Balloon Text Char"};
            LinkedStyle linkedStyle14 = new LinkedStyle() {Val = "BalloonText"};
            UIPriority uIPriority25 = new UIPriority() {Val = 99};
            Locked locked8 = new Locked();
            Rsid rsid406 = new Rsid() {Val = "00513B3F"};

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts1057 = new RunFonts() {Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma"};
            FontSize fontSize1073 = new FontSize() {Val = "16"};
            FontSizeComplexScript fontSizeComplexScript1058 = new FontSizeComplexScript() {Val = "16"};

            styleRunProperties21.Append(runFonts1057);
            styleRunProperties21.Append(fontSize1073);
            styleRunProperties21.Append(fontSizeComplexScript1058);

            style26.Append(styleName26);
            style26.Append(linkedStyle14);
            style26.Append(uIPriority25);
            style26.Append(locked8);
            style26.Append(rsid406);
            style26.Append(styleRunProperties21);

            Style style27 = new Style() {Type = StyleValues.Character, StyleId = "Strong"};
            StyleName styleName27 = new StyleName() {Val = "Strong"};
            UIPriority uIPriority26 = new UIPriority() {Val = 99};
            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            Rsid rsid407 = new Rsid() {Val = "00E07DEF"};

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            RunFonts runFonts1058 = new RunFonts() {ComplexScript = "Times New Roman"};
            Bold bold127 = new Bold();
            BoldComplexScript boldComplexScript20 = new BoldComplexScript();

            styleRunProperties22.Append(runFonts1058);
            styleRunProperties22.Append(bold127);
            styleRunProperties22.Append(boldComplexScript20);

            style27.Append(styleName27);
            style27.Append(uIPriority26);
            style27.Append(primaryStyle6);
            style27.Append(rsid407);
            style27.Append(styleRunProperties22);

            Style style28 = new Style() {Type = StyleValues.Character, StyleId = "CommentReference"};
            StyleName styleName28 = new StyleName() {Val = "annotation reference"};
            UIPriority uIPriority27 = new UIPriority() {Val = 99};
            SemiHidden semiHidden10 = new SemiHidden();
            Rsid rsid408 = new Rsid() {Val = "00E04E9E"};

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            RunFonts runFonts1059 = new RunFonts() {ComplexScript = "Times New Roman"};
            FontSize fontSize1074 = new FontSize() {Val = "16"};
            FontSizeComplexScript fontSizeComplexScript1059 = new FontSizeComplexScript() {Val = "16"};

            styleRunProperties23.Append(runFonts1059);
            styleRunProperties23.Append(fontSize1074);
            styleRunProperties23.Append(fontSizeComplexScript1059);

            style28.Append(styleName28);
            style28.Append(uIPriority27);
            style28.Append(semiHidden10);
            style28.Append(rsid408);
            style28.Append(styleRunProperties23);

            Style style29 = new Style() {Type = StyleValues.Paragraph, StyleId = "CommentText"};
            StyleName styleName29 = new StyleName() {Val = "annotation text"};
            BasedOn basedOn15 = new BasedOn() {Val = "Normal"};
            LinkedStyle linkedStyle15 = new LinkedStyle() {Val = "CommentTextChar"};
            UIPriority uIPriority28 = new UIPriority() {Val = 99};
            SemiHidden semiHidden11 = new SemiHidden();
            Rsid rsid409 = new Rsid() {Val = "00E04E9E"};

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            FontSize fontSize1075 = new FontSize() {Val = "20"};
            FontSizeComplexScript fontSizeComplexScript1060 = new FontSizeComplexScript() {Val = "20"};

            styleRunProperties24.Append(fontSize1075);
            styleRunProperties24.Append(fontSizeComplexScript1060);

            style29.Append(styleName29);
            style29.Append(basedOn15);
            style29.Append(linkedStyle15);
            style29.Append(uIPriority28);
            style29.Append(semiHidden11);
            style29.Append(rsid409);
            style29.Append(styleRunProperties24);

            Style style30 = new Style() {Type = StyleValues.Character, StyleId = "CommentTextChar", CustomStyle = true};
            StyleName styleName30 = new StyleName() {Val = "Comment Text Char"};
            LinkedStyle linkedStyle16 = new LinkedStyle() {Val = "CommentText"};
            UIPriority uIPriority29 = new UIPriority() {Val = 99};
            Locked locked9 = new Locked();
            Rsid rsid410 = new Rsid() {Val = "00E04E9E"};

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            RunFonts runFonts1060 = new RunFonts() {ComplexScript = "Times New Roman"};

            styleRunProperties25.Append(runFonts1060);

            style30.Append(styleName30);
            style30.Append(linkedStyle16);
            style30.Append(uIPriority29);
            style30.Append(locked9);
            style30.Append(rsid410);
            style30.Append(styleRunProperties25);

            Style style31 = new Style() {Type = StyleValues.Paragraph, StyleId = "CommentSubject"};
            StyleName styleName31 = new StyleName() {Val = "annotation subject"};
            BasedOn basedOn16 = new BasedOn() {Val = "CommentText"};
            NextParagraphStyle nextParagraphStyle8 = new NextParagraphStyle() {Val = "CommentText"};
            LinkedStyle linkedStyle17 = new LinkedStyle() {Val = "CommentSubjectChar"};
            UIPriority uIPriority30 = new UIPriority() {Val = 99};
            SemiHidden semiHidden12 = new SemiHidden();
            Rsid rsid411 = new Rsid() {Val = "00E04E9E"};

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            Bold bold128 = new Bold();
            BoldComplexScript boldComplexScript21 = new BoldComplexScript();

            styleRunProperties26.Append(bold128);
            styleRunProperties26.Append(boldComplexScript21);

            style31.Append(styleName31);
            style31.Append(basedOn16);
            style31.Append(nextParagraphStyle8);
            style31.Append(linkedStyle17);
            style31.Append(uIPriority30);
            style31.Append(semiHidden12);
            style31.Append(rsid411);
            style31.Append(styleRunProperties26);

            Style style32 = new Style() {Type = StyleValues.Character, StyleId = "CommentSubjectChar", CustomStyle = true};
            StyleName styleName32 = new StyleName() {Val = "Comment Subject Char"};
            LinkedStyle linkedStyle18 = new LinkedStyle() {Val = "CommentSubject"};
            UIPriority uIPriority31 = new UIPriority() {Val = 99};
            Locked locked10 = new Locked();
            Rsid rsid412 = new Rsid() {Val = "00E04E9E"};

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            RunFonts runFonts1061 = new RunFonts() {ComplexScript = "Times New Roman"};
            Bold bold129 = new Bold();
            BoldComplexScript boldComplexScript22 = new BoldComplexScript();

            styleRunProperties27.Append(runFonts1061);
            styleRunProperties27.Append(bold129);
            styleRunProperties27.Append(boldComplexScript22);

            style32.Append(styleName32);
            style32.Append(linkedStyle18);
            style32.Append(uIPriority31);
            style32.Append(locked10);
            style32.Append(rsid412);
            style32.Append(styleRunProperties27);

            Style style33 = new Style() {Type = StyleValues.Character, StyleId = "Hyperlink"};
            StyleName styleName33 = new StyleName() {Val = "Hyperlink"};
            UIPriority uIPriority32 = new UIPriority() {Val = 99};
            Rsid rsid413 = new Rsid() {Val = "0077670E"};

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            RunFonts runFonts1062 = new RunFonts() {ComplexScript = "Times New Roman"};
            Color color376 = new Color() {Val = "0000FF"};
            Underline underline96 = new Underline() {Val = UnderlineValues.Single};

            styleRunProperties28.Append(runFonts1062);
            styleRunProperties28.Append(color376);
            styleRunProperties28.Append(underline96);

            style33.Append(styleName33);
            style33.Append(uIPriority32);
            style33.Append(rsid413);
            style33.Append(styleRunProperties28);

            Style style34 = new Style() {Type = StyleValues.Paragraph, StyleId = "NoSpacing"};
            StyleName styleName34 = new StyleName() {Val = "No Spacing"};
            UIPriority uIPriority33 = new UIPriority() {Val = 99};
            PrimaryStyle primaryStyle7 = new PrimaryStyle();
            Rsid rsid414 = new Rsid() {Val = "00CF3870"};

            StyleRunProperties styleRunProperties29 = new StyleRunProperties();
            RunFonts runFonts1063 = new RunFonts() {Ascii = "Calibri", HighAnsi = "Calibri"};
            FontSize fontSize1076 = new FontSize() {Val = "22"};
            FontSizeComplexScript fontSizeComplexScript1061 = new FontSizeComplexScript() {Val = "22"};

            styleRunProperties29.Append(runFonts1063);
            styleRunProperties29.Append(fontSize1076);
            styleRunProperties29.Append(fontSizeComplexScript1061);

            style34.Append(styleName34);
            style34.Append(uIPriority33);
            style34.Append(primaryStyle7);
            style34.Append(rsid414);
            style34.Append(styleRunProperties29);

            Style style35 = new Style() {Type = StyleValues.Table, StyleId = "TableGrid"};
            StyleName styleName35 = new StyleName() {Val = "Table Grid"};
            BasedOn basedOn17 = new BasedOn() {Val = "TableNormal"};
            UIPriority uIPriority34 = new UIPriority() {Val = 99};
            Locked locked11 = new Locked();
            Rsid rsid415 = new Rsid() {Val = "00614236"};

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();

            TableBorders tableBorders5 = new TableBorders();
            TopBorder topBorder94 = new TopBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 4U, Space = (UInt32Value) 0U};
            LeftBorder leftBorder93 = new LeftBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 4U, Space = (UInt32Value) 0U};
            BottomBorder bottomBorder100 = new BottomBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 4U, Space = (UInt32Value) 0U};
            RightBorder rightBorder93 = new RightBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 4U, Space = (UInt32Value) 0U};
            InsideHorizontalBorder insideHorizontalBorder5 = new InsideHorizontalBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 4U, Space = (UInt32Value) 0U};
            InsideVerticalBorder insideVerticalBorder5 = new InsideVerticalBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 4U, Space = (UInt32Value) 0U};

            tableBorders5.Append(topBorder94);
            tableBorders5.Append(leftBorder93);
            tableBorders5.Append(bottomBorder100);
            tableBorders5.Append(rightBorder93);
            tableBorders5.Append(insideHorizontalBorder5);
            tableBorders5.Append(insideVerticalBorder5);

            styleTableProperties2.Append(tableBorders5);

            style35.Append(styleName35);
            style35.Append(basedOn17);
            style35.Append(uIPriority34);
            style35.Append(locked11);
            style35.Append(rsid415);
            style35.Append(styleTableProperties2);

            Style style36 = new Style() {Type = StyleValues.Table, StyleId = "LightShading"};
            StyleName styleName36 = new StyleName() {Val = "Light Shading"};
            BasedOn basedOn18 = new BasedOn() {Val = "TableNormal"};
            UIPriority uIPriority35 = new UIPriority() {Val = 99};
            Rsid rsid416 = new Rsid() {Val = "00614236"};

            StyleRunProperties styleRunProperties30 = new StyleRunProperties();
            Color color377 = new Color() {Val = "000000"};

            styleRunProperties30.Append(color377);

            StyleTableProperties styleTableProperties3 = new StyleTableProperties();
            TableStyleRowBandSize tableStyleRowBandSize1 = new TableStyleRowBandSize() {Val = 1};
            TableStyleColumnBandSize tableStyleColumnBandSize1 = new TableStyleColumnBandSize() {Val = 1};

            TableBorders tableBorders6 = new TableBorders();
            TopBorder topBorder95 = new TopBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 8U, Space = (UInt32Value) 0U};
            BottomBorder bottomBorder101 = new BottomBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 8U, Space = (UInt32Value) 0U};

            tableBorders6.Append(topBorder95);
            tableBorders6.Append(bottomBorder101);

            styleTableProperties3.Append(tableStyleRowBandSize1);
            styleTableProperties3.Append(tableStyleColumnBandSize1);
            styleTableProperties3.Append(tableBorders6);

            TableStyleProperties tableStyleProperties1 = new TableStyleProperties() {Type = TableStyleOverrideValues.FirstRow};

            StyleParagraphProperties styleParagraphProperties14 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() {Before = "0", After = "0"};

            styleParagraphProperties14.Append(spacingBetweenLines13);

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            RunFonts runFonts1064 = new RunFonts() {ComplexScript = "Times New Roman"};
            Bold bold130 = new Bold();
            BoldComplexScript boldComplexScript23 = new BoldComplexScript();

            runPropertiesBaseStyle2.Append(runFonts1064);
            runPropertiesBaseStyle2.Append(bold130);
            runPropertiesBaseStyle2.Append(boldComplexScript23);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties1 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties1 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders91 = new TableCellBorders();
            TopBorder topBorder96 = new TopBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 8U, Space = (UInt32Value) 0U};
            LeftBorder leftBorder94 = new LeftBorder() {Val = BorderValues.Nil};
            BottomBorder bottomBorder102 = new BottomBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 8U, Space = (UInt32Value) 0U};
            RightBorder rightBorder94 = new RightBorder() {Val = BorderValues.Nil};
            InsideHorizontalBorder insideHorizontalBorder6 = new InsideHorizontalBorder() {Val = BorderValues.Nil};
            InsideVerticalBorder insideVerticalBorder6 = new InsideVerticalBorder() {Val = BorderValues.Nil};

            tableCellBorders91.Append(topBorder96);
            tableCellBorders91.Append(leftBorder94);
            tableCellBorders91.Append(bottomBorder102);
            tableCellBorders91.Append(rightBorder94);
            tableCellBorders91.Append(insideHorizontalBorder6);
            tableCellBorders91.Append(insideVerticalBorder6);

            tableStyleConditionalFormattingTableCellProperties1.Append(tableCellBorders91);

            tableStyleProperties1.Append(styleParagraphProperties14);
            tableStyleProperties1.Append(runPropertiesBaseStyle2);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableProperties1);
            tableStyleProperties1.Append(tableStyleConditionalFormattingTableCellProperties1);

            TableStyleProperties tableStyleProperties2 = new TableStyleProperties() {Type = TableStyleOverrideValues.LastRow};

            StyleParagraphProperties styleParagraphProperties15 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() {Before = "0", After = "0"};

            styleParagraphProperties15.Append(spacingBetweenLines14);

            RunPropertiesBaseStyle runPropertiesBaseStyle3 = new RunPropertiesBaseStyle();
            RunFonts runFonts1065 = new RunFonts() {ComplexScript = "Times New Roman"};
            Bold bold131 = new Bold();
            BoldComplexScript boldComplexScript24 = new BoldComplexScript();

            runPropertiesBaseStyle3.Append(runFonts1065);
            runPropertiesBaseStyle3.Append(bold131);
            runPropertiesBaseStyle3.Append(boldComplexScript24);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties2 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties2 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders92 = new TableCellBorders();
            TopBorder topBorder97 = new TopBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 8U, Space = (UInt32Value) 0U};
            LeftBorder leftBorder95 = new LeftBorder() {Val = BorderValues.Nil};
            BottomBorder bottomBorder103 = new BottomBorder() {Val = BorderValues.Single, Color = "000000", Size = (UInt32Value) 8U, Space = (UInt32Value) 0U};
            RightBorder rightBorder95 = new RightBorder() {Val = BorderValues.Nil};
            InsideHorizontalBorder insideHorizontalBorder7 = new InsideHorizontalBorder() {Val = BorderValues.Nil};
            InsideVerticalBorder insideVerticalBorder7 = new InsideVerticalBorder() {Val = BorderValues.Nil};

            tableCellBorders92.Append(topBorder97);
            tableCellBorders92.Append(leftBorder95);
            tableCellBorders92.Append(bottomBorder103);
            tableCellBorders92.Append(rightBorder95);
            tableCellBorders92.Append(insideHorizontalBorder7);
            tableCellBorders92.Append(insideVerticalBorder7);

            tableStyleConditionalFormattingTableCellProperties2.Append(tableCellBorders92);

            tableStyleProperties2.Append(styleParagraphProperties15);
            tableStyleProperties2.Append(runPropertiesBaseStyle3);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableProperties2);
            tableStyleProperties2.Append(tableStyleConditionalFormattingTableCellProperties2);

            TableStyleProperties tableStyleProperties3 = new TableStyleProperties() {Type = TableStyleOverrideValues.FirstColumn};

            RunPropertiesBaseStyle runPropertiesBaseStyle4 = new RunPropertiesBaseStyle();
            RunFonts runFonts1066 = new RunFonts() {ComplexScript = "Times New Roman"};
            Bold bold132 = new Bold();
            BoldComplexScript boldComplexScript25 = new BoldComplexScript();

            runPropertiesBaseStyle4.Append(runFonts1066);
            runPropertiesBaseStyle4.Append(bold132);
            runPropertiesBaseStyle4.Append(boldComplexScript25);

            tableStyleProperties3.Append(runPropertiesBaseStyle4);

            TableStyleProperties tableStyleProperties4 = new TableStyleProperties() {Type = TableStyleOverrideValues.LastColumn};

            RunPropertiesBaseStyle runPropertiesBaseStyle5 = new RunPropertiesBaseStyle();
            RunFonts runFonts1067 = new RunFonts() {ComplexScript = "Times New Roman"};
            Bold bold133 = new Bold();
            BoldComplexScript boldComplexScript26 = new BoldComplexScript();

            runPropertiesBaseStyle5.Append(runFonts1067);
            runPropertiesBaseStyle5.Append(bold133);
            runPropertiesBaseStyle5.Append(boldComplexScript26);

            tableStyleProperties4.Append(runPropertiesBaseStyle5);

            TableStyleProperties tableStyleProperties5 = new TableStyleProperties() {Type = TableStyleOverrideValues.Band1Vertical};

            RunPropertiesBaseStyle runPropertiesBaseStyle6 = new RunPropertiesBaseStyle();
            RunFonts runFonts1068 = new RunFonts() {ComplexScript = "Times New Roman"};

            runPropertiesBaseStyle6.Append(runFonts1068);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties3 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties3 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders93 = new TableCellBorders();
            LeftBorder leftBorder96 = new LeftBorder() {Val = BorderValues.Nil};
            RightBorder rightBorder96 = new RightBorder() {Val = BorderValues.Nil};
            InsideHorizontalBorder insideHorizontalBorder8 = new InsideHorizontalBorder() {Val = BorderValues.Nil};
            InsideVerticalBorder insideVerticalBorder8 = new InsideVerticalBorder() {Val = BorderValues.Nil};

            tableCellBorders93.Append(leftBorder96);
            tableCellBorders93.Append(rightBorder96);
            tableCellBorders93.Append(insideHorizontalBorder8);
            tableCellBorders93.Append(insideVerticalBorder8);
            Shading shading55 = new Shading() {Val = ShadingPatternValues.Clear, Color = "auto", Fill = "C0C0C0"};

            tableStyleConditionalFormattingTableCellProperties3.Append(tableCellBorders93);
            tableStyleConditionalFormattingTableCellProperties3.Append(shading55);

            tableStyleProperties5.Append(runPropertiesBaseStyle6);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableProperties3);
            tableStyleProperties5.Append(tableStyleConditionalFormattingTableCellProperties3);

            TableStyleProperties tableStyleProperties6 = new TableStyleProperties() {Type = TableStyleOverrideValues.Band1Horizontal};

            RunPropertiesBaseStyle runPropertiesBaseStyle7 = new RunPropertiesBaseStyle();
            RunFonts runFonts1069 = new RunFonts() {ComplexScript = "Times New Roman"};

            runPropertiesBaseStyle7.Append(runFonts1069);
            TableStyleConditionalFormattingTableProperties tableStyleConditionalFormattingTableProperties4 = new TableStyleConditionalFormattingTableProperties();

            TableStyleConditionalFormattingTableCellProperties tableStyleConditionalFormattingTableCellProperties4 = new TableStyleConditionalFormattingTableCellProperties();

            TableCellBorders tableCellBorders94 = new TableCellBorders();
            LeftBorder leftBorder97 = new LeftBorder() {Val = BorderValues.Nil};
            RightBorder rightBorder97 = new RightBorder() {Val = BorderValues.Nil};
            InsideHorizontalBorder insideHorizontalBorder9 = new InsideHorizontalBorder() {Val = BorderValues.Nil};
            InsideVerticalBorder insideVerticalBorder9 = new InsideVerticalBorder() {Val = BorderValues.Nil};

            tableCellBorders94.Append(leftBorder97);
            tableCellBorders94.Append(rightBorder97);
            tableCellBorders94.Append(insideHorizontalBorder9);
            tableCellBorders94.Append(insideVerticalBorder9);
            Shading shading56 = new Shading() {Val = ShadingPatternValues.Clear, Color = "auto", Fill = "C0C0C0"};

            tableStyleConditionalFormattingTableCellProperties4.Append(tableCellBorders94);
            tableStyleConditionalFormattingTableCellProperties4.Append(shading56);

            tableStyleProperties6.Append(runPropertiesBaseStyle7);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableProperties4);
            tableStyleProperties6.Append(tableStyleConditionalFormattingTableCellProperties4);

            style36.Append(styleName36);
            style36.Append(basedOn18);
            style36.Append(uIPriority35);
            style36.Append(rsid416);
            style36.Append(styleRunProperties30);
            style36.Append(styleTableProperties3);
            style36.Append(tableStyleProperties1);
            style36.Append(tableStyleProperties2);
            style36.Append(tableStyleProperties3);
            style36.Append(tableStyleProperties4);
            style36.Append(tableStyleProperties5);
            style36.Append(tableStyleProperties6);

            Style style37 = new Style() {Type = StyleValues.Paragraph, StyleId = "ListParagraph"};
            StyleName styleName37 = new StyleName() {Val = "List Paragraph"};
            BasedOn basedOn19 = new BasedOn() {Val = "Normal"};
            UIPriority uIPriority36 = new UIPriority() {Val = 34};
            PrimaryStyle primaryStyle8 = new PrimaryStyle();
            Rsid rsid417 = new Rsid() {Val = "00016236"};

            StyleParagraphProperties styleParagraphProperties16 = new StyleParagraphProperties();
            Indentation indentation28 = new Indentation() {Start = "720"};

            styleParagraphProperties16.Append(indentation28);

            style37.Append(styleName37);
            style37.Append(basedOn19);
            style37.Append(uIPriority36);
            style37.Append(primaryStyle8);
            style37.Append(rsid417);
            style37.Append(styleParagraphProperties16);

            Style style38 = new Style() {Type = StyleValues.Character, StyleId = "PlaceholderText"};
            StyleName styleName38 = new StyleName() {Val = "Placeholder Text"};
            BasedOn basedOn20 = new BasedOn() {Val = "DefaultParagraphFont"};
            UIPriority uIPriority37 = new UIPriority() {Val = 99};
            SemiHidden semiHidden13 = new SemiHidden();
            Rsid rsid418 = new Rsid() {Val = "005E53DA"};

            StyleRunProperties styleRunProperties31 = new StyleRunProperties();
            Color color378 = new Color() {Val = "808080"};

            styleRunProperties31.Append(color378);

            style38.Append(styleName38);
            style38.Append(basedOn20);
            style38.Append(uIPriority37);
            style38.Append(semiHidden13);
            style38.Append(rsid418);
            style38.Append(styleRunProperties31);

            Style style39 = new Style() {Type = StyleValues.Paragraph, StyleId = "Revision"};
            StyleName styleName39 = new StyleName() {Val = "Revision"};
            StyleHidden styleHidden1 = new StyleHidden();
            UIPriority uIPriority38 = new UIPriority() {Val = 99};
            SemiHidden semiHidden14 = new SemiHidden();
            Rsid rsid419 = new Rsid() {Val = "000E2AF5"};

            StyleRunProperties styleRunProperties32 = new StyleRunProperties();
            FontSize fontSize1077 = new FontSize() {Val = "24"};
            FontSizeComplexScript fontSizeComplexScript1062 = new FontSizeComplexScript() {Val = "24"};

            styleRunProperties32.Append(fontSize1077);
            styleRunProperties32.Append(fontSizeComplexScript1062);

            style39.Append(styleName39);
            style39.Append(styleHidden1);
            style39.Append(uIPriority38);
            style39.Append(semiHidden14);
            style39.Append(rsid419);
            style39.Append(styleRunProperties32);

            Style style40 = new Style() {Type = StyleValues.Character, StyleId = "FollowedHyperlink"};
            StyleName styleName40 = new StyleName() {Val = "FollowedHyperlink"};
            BasedOn basedOn21 = new BasedOn() {Val = "DefaultParagraphFont"};
            UIPriority uIPriority39 = new UIPriority() {Val = 99};
            SemiHidden semiHidden15 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            Rsid rsid420 = new Rsid() {Val = "00BC1A91"};

            StyleRunProperties styleRunProperties33 = new StyleRunProperties();
            Color color379 = new Color() {Val = "800080", ThemeColor = ThemeColorValues.FollowedHyperlink};
            Underline underline97 = new Underline() {Val = UnderlineValues.Single};

            styleRunProperties33.Append(color379);
            styleRunProperties33.Append(underline97);

            style40.Append(styleName40);
            style40.Append(basedOn21);
            style40.Append(uIPriority39);
            style40.Append(semiHidden15);
            style40.Append(unhideWhenUsed4);
            style40.Append(rsid420);
            style40.Append(styleRunProperties33);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);
            styles1.Append(style25);
            styles1.Append(style26);
            styles1.Append(style27);
            styles1.Append(style28);
            styles1.Append(style29);
            styles1.Append(style30);
            styles1.Append(style31);
            styles1.Append(style32);
            styles1.Append(style33);
            styles1.Append(style34);
            styles1.Append(style35);
            styles1.Append(style36);
            styles1.Append(style37);
            styles1.Append(style38);
            styles1.Append(style39);
            styles1.Append(style40);

            styleDefinitionsPart1.Styles = styles1;
        }
    }
}