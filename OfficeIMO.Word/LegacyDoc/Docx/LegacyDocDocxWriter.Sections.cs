using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using static OfficeIMO.Word.LegacyDoc.Docx.LegacyDocXml;

namespace OfficeIMO.Word.LegacyDoc.Docx {
    internal sealed partial class LegacyDocDocxWriter {
        private const string HeaderContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
        private const string FooterContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
        private OfficeXmlTree? _settings;
        private readonly Dictionary<(int SectionProperties, bool IsHeader, LegacyDocHeaderFooterType Type), LegacyDocDocxPart> _headerFooterParts = new();

        private OfficeXmlTree Settings => _settings ??= _package.GetXml(SettingsPartName);

        /// <summary>
        /// Ends the current section at <paramref name="boundaryParagraph"/> (or a new empty paragraph) and starts a new
        /// one. The ending section keeps a copy of the current settings in the boundary paragraph; the body-level
        /// properties go on describing the new last section, which starts as <paramref name="format"/> says.
        /// </summary>
        private void AddSection(List<(int SectionProperties, LegacyDocSectionFormat Format)> sections, LegacyDocSectionFormat format, int boundaryParagraph) {
            OfficeXmlTree tree = _main.Tree;
            if (boundaryParagraph == OfficeXmlTree.NoNode) {
                boundaryParagraph = Element(tree, "p");
                AppendBodyBlock(boundaryParagraph);
            }

            int last = sections[sections.Count - 1].SectionProperties;
            int ended = Element(tree, "sectPr");
            foreach (int element in Elements(tree, last)) {
                string name = tree.LocalNameOf(element);
                if (tree.NamespaceOf(element) == W && (name == "headerReference" || name == "footerReference" || name == "titlePg")) {
                    tree.Remove(element);
                    Append(tree, ended, element);
                } else if (tree.NamespaceOf(element) == W && (name == "pgSz" || name == "pgMar")) {
                    if (Child(tree, ended, name) == OfficeXmlTree.NoNode) Append(tree, ended, tree.CloneElement(element));
                } else if (tree.NamespaceOf(element) == W && (name == "footnotePr" || name == "endnotePr")) {
                    // The copy belongs to the ending section; the new last section continues numbering by default.
                    Insert(tree, ended, tree.CloneElement(element), SectionPropertiesOrder);
                    RemoveAll(tree, element, "numRestart");
                } else {
                    Append(tree, ended, tree.CloneElement(element));
                }
            }

            Insert(tree, ParagraphProperties(tree, boundaryParagraph), ended, ParagraphPropertiesOrder);
            sections[sections.Count - 1] = (ended, sections[sections.Count - 1].Format);
            sections.Add((last, format));
            // Section properties describe how their own section starts.
            Replace(tree, last, Element(tree, "type", (format.SectionBreakType ?? LegacyDocSectionMark.NextPage).ToXml()), SectionPropertiesOrder);
            int pageNumbering = Child(tree, last, "pgNumType");
            if (pageNumbering != OfficeXmlTree.NoNode) tree.RemoveAttribute(pageNumbering, W, "start");
        }

        private void ApplySectionFormatting(int sectionProperties, LegacyDocSectionFormat sectionFormat) {
            OfficeXmlTree tree = _main.Tree;
            // Native section dimensions do not carry a DOCX printer-form code. Do not keep the creation package's
            // paper hint for another paper size.
            int pageSize = Child(tree, sectionProperties, "pgSz");
            if (pageSize != OfficeXmlTree.NoNode) tree.RemoveAttribute(pageSize, W, "code");
            if (!sectionFormat.HasFormatting) return;

            if (sectionFormat.SectionBreakType != null) {
                RemoveAll(tree, sectionProperties, "type");
                Insert(tree, sectionProperties, Element(tree, "type", sectionFormat.SectionBreakType.Value.ToXml()), SectionPropertiesOrder);
            }

            if (sectionFormat.DifferentFirstPage) {
                int titlePage = Child(tree, sectionProperties, "titlePg");
                if (titlePage == OfficeXmlTree.NoNode) Append(tree, sectionProperties, Element(tree, "titlePg"));
                else Set(tree, titlePage, "val", "true");
                AddHeaderFooterReference(sectionProperties, isHeader: true, LegacyDocHeaderFooterType.First);
                AddHeaderFooterReference(sectionProperties, isHeader: false, LegacyDocHeaderFooterType.First);
            }

            if (sectionFormat.Orientation != null) {
                // DOC stores physical dimensions independently of its orientation flag.
                pageSize = Child(tree, sectionProperties, "pgSz");
                if (pageSize == OfficeXmlTree.NoNode) pageSize = Insert(tree, sectionProperties, Element(tree, "pgSz"), SectionPropertiesOrder);
                Set(tree, pageSize, "orient", sectionFormat.Orientation.Value.ToXml());
            }

            if (sectionFormat.PageWidthTwips != null) Set(tree, PageSize(tree, sectionProperties), "w", Text(sectionFormat.PageWidthTwips.Value));
            if (sectionFormat.PageHeightTwips != null) Set(tree, PageSize(tree, sectionProperties), "h", Text(sectionFormat.PageHeightTwips.Value));
            if (sectionFormat.MarginTopTwips != null) Set(tree, PageMargin(tree, sectionProperties), "top", Text(sectionFormat.MarginTopTwips.Value));
            if (sectionFormat.MarginRightTwips != null) Set(tree, PageMargin(tree, sectionProperties), "right", Text(sectionFormat.MarginRightTwips.Value));
            if (sectionFormat.MarginBottomTwips != null) Set(tree, PageMargin(tree, sectionProperties), "bottom", Text(sectionFormat.MarginBottomTwips.Value));
            if (sectionFormat.MarginLeftTwips != null) Set(tree, PageMargin(tree, sectionProperties), "left", Text(sectionFormat.MarginLeftTwips.Value));

            int pageMargin = Child(tree, sectionProperties, "pgMar");
            if (pageMargin == OfficeXmlTree.NoNode && (sectionFormat.HeaderDistanceTwips != null || sectionFormat.FooterDistanceTwips != null || sectionFormat.GutterTwips != null)) {
                pageMargin = Append(tree, sectionProperties, Element(tree, "pgMar"));
                Set(tree, pageMargin, "top", "1440");
                Set(tree, pageMargin, "right", "1440");
                Set(tree, pageMargin, "bottom", "1440");
                Set(tree, pageMargin, "left", "1440");
                Set(tree, pageMargin, "header", "720");
                Set(tree, pageMargin, "footer", "720");
                Set(tree, pageMargin, "gutter", "0");
            }

            if (pageMargin != OfficeXmlTree.NoNode) {
                if (sectionFormat.HeaderDistanceTwips != null) Set(tree, pageMargin, "header", Text(sectionFormat.HeaderDistanceTwips.Value));
                if (sectionFormat.FooterDistanceTwips != null) Set(tree, pageMargin, "footer", Text(sectionFormat.FooterDistanceTwips.Value));
                if (sectionFormat.GutterTwips != null) Set(tree, pageMargin, "gutter", Text(sectionFormat.GutterTwips.Value));
            }

            if (sectionFormat.ColumnCount != null) Set(tree, Columns(tree, sectionProperties), "num", Text(sectionFormat.ColumnCount.Value));
            if (sectionFormat.ColumnSpacingTwips != null) Set(tree, Columns(tree, sectionProperties), "space", Text(sectionFormat.ColumnSpacingTwips.Value));
            if (sectionFormat.ColumnDefinitions != null) SetColumnDefinitions(tree, sectionProperties, sectionFormat.ColumnDefinitions);
            if (sectionFormat.HasColumnSeparator) Set(tree, Columns(tree, sectionProperties), "sep", "true");
            if (sectionFormat.RtlGutter) {
                Replace(tree, sectionProperties, Element(tree, "rtlGutter", "true"), SectionPropertiesOrder);
            }

            if (sectionFormat.VerticalAlignment != null) {
                RemoveAll(tree, sectionProperties, "vAlign");
                Append(tree, sectionProperties, Element(tree, "vAlign", sectionFormat.VerticalAlignment.Value.ToXml()));
            }

            if (sectionFormat.PageBorders != null && sectionFormat.PageBorders.Value.HasAny) ApplyPageBorders(tree, sectionProperties, sectionFormat.PageBorders.Value);

            if (sectionFormat.LineNumberCountBy != null || sectionFormat.LineNumberDistanceTwips != null
                || sectionFormat.LineNumberStart != null || sectionFormat.LineNumberRestart != null) {
                RemoveAll(tree, sectionProperties, "lnNumType");
                int lineNumbering = Element(tree, "lnNumType");
                if (sectionFormat.LineNumberCountBy != null) Set(tree, lineNumbering, "countBy", Text(sectionFormat.LineNumberCountBy.Value));
                if (sectionFormat.LineNumberDistanceTwips != null) Set(tree, lineNumbering, "distance", Text(sectionFormat.LineNumberDistanceTwips.Value));
                if (sectionFormat.LineNumberStart != null) Set(tree, lineNumbering, "start", Text(sectionFormat.LineNumberStart.Value));
                if (sectionFormat.LineNumberRestart != null) Set(tree, lineNumbering, "restart", sectionFormat.LineNumberRestart.Value.ToXml());
                Append(tree, sectionProperties, lineNumbering);
            }

            if (sectionFormat.PageNumberStart != null || sectionFormat.PageNumberFormat != null) {
                RemoveAll(tree, sectionProperties, "pgNumType");
                int pageNumbering = Element(tree, "pgNumType");
                if (sectionFormat.PageNumberFormat != null) Set(tree, pageNumbering, "fmt", sectionFormat.PageNumberFormat.Value.ToXml());
                if (sectionFormat.PageNumberStart != null) Set(tree, pageNumbering, "start", Text(sectionFormat.PageNumberStart.Value));
                Insert(tree, sectionProperties, pageNumbering, SectionPropertiesOrder);
            }

            if (sectionFormat.FootnotePosition != null || sectionFormat.FootnoteRestart != null
                || sectionFormat.FootnoteStart != null || sectionFormat.FootnoteNumberFormat != null) {
                int properties = GetOrAdd(tree, sectionProperties, "footnotePr", SectionPropertiesOrder);
                RemoveAll(tree, properties, "numFmt");
                RemoveAll(tree, properties, "pos");
                RemoveAll(tree, properties, "numRestart");
                RemoveAll(tree, properties, "numStart");
                AddNoteProperties(tree, properties, sectionFormat.FootnotePosition?.ToXml(), sectionFormat.FootnoteNumberFormat, sectionFormat.FootnoteStart, sectionFormat.FootnoteRestart);
            }

            if (sectionFormat.EndnotePosition != null || sectionFormat.EndnoteRestart != null
                || sectionFormat.EndnoteStart != null || sectionFormat.EndnoteNumberFormat != null) {
                int properties = GetOrAdd(tree, sectionProperties, "endnotePr", SectionPropertiesOrder);
                RemoveAll(tree, properties, "pos");
                bool positionOnly = sectionFormat.EndnotePosition != null && sectionFormat.EndnoteNumberFormat == null
                    && sectionFormat.EndnoteRestart == null && sectionFormat.EndnoteStart == null;
                if (!positionOnly) {
                    RemoveAll(tree, properties, "numFmt");
                    RemoveAll(tree, properties, "numRestart");
                    RemoveAll(tree, properties, "numStart");
                }

                AddNoteProperties(tree, properties, sectionFormat.EndnotePosition?.ToXml(), sectionFormat.EndnoteNumberFormat, sectionFormat.EndnoteStart, sectionFormat.EndnoteRestart);
                if (sectionFormat.EndnotePosition != null) SetDocumentEndnotePosition(sectionFormat.EndnotePosition.Value.ToXml());
            }

            // Setters may append properties after page geometry. Restore schema order.
            List<int> ordered = Elements(tree, sectionProperties);
            foreach (int property in ordered) tree.Remove(property);
            foreach (int property in ordered) Insert(tree, sectionProperties, property, SectionPropertiesOrder);
        }

        private static readonly OfficeXmlChildOrder NotePropertiesOrder = Order("pos", "numFmt", "numStart", "numRestart");

        private static void AddNoteProperties(OfficeXmlTree tree, int properties, string? position, LegacyDocNumberFormat? numberFormat, int? start, LegacyDocRestartNumber? restart) {
            if (numberFormat != null) Insert(tree, properties, Element(tree, "numFmt", numberFormat.Value.ToXml()), NotePropertiesOrder);
            if (position != null) Insert(tree, properties, Element(tree, "pos", position), NotePropertiesOrder);
            if (restart != null) Insert(tree, properties, Element(tree, "numRestart", restart.Value.ToXml()), NotePropertiesOrder);
            if (start != null) Insert(tree, properties, Element(tree, "numStart", Text(start.Value)), NotePropertiesOrder);
        }

        private void SetDocumentEndnotePosition(string position) {
            OfficeXmlTree settings = Settings;
            int properties = GetOrAdd(settings, settings.Root, "endnotePr", SettingsOrder);
            RemoveAll(settings, properties, "pos");
            Insert(settings, properties, Element(settings, "pos", position), NotePropertiesOrder);
        }

        private static int PageSize(OfficeXmlTree tree, int sectionProperties) =>
            GetOrAdd(tree, sectionProperties, "pgSz", SectionPropertiesOrder);

        private static int PageMargin(OfficeXmlTree tree, int sectionProperties) =>
            GetOrAdd(tree, sectionProperties, "pgMar", SectionPropertiesOrder);

        private static int Columns(OfficeXmlTree tree, int sectionProperties) {
            int columns = Child(tree, sectionProperties, "cols");
            return columns != OfficeXmlTree.NoNode ? columns : Append(tree, sectionProperties, Element(tree, "cols"));
        }

        private static void SetColumnDefinitions(OfficeXmlTree tree, int sectionProperties, IReadOnlyList<WordSectionColumn> definitions) {
            int columns = Child(tree, sectionProperties, "cols");
            if (definitions.Count == 0 && columns == OfficeXmlTree.NoNode) return;
            string? currentCount = columns == OfficeXmlTree.NoNode ? null : Get(tree, columns, "num");
            if (columns == OfficeXmlTree.NoNode) columns = Append(tree, sectionProperties, Element(tree, "cols"));
            RemoveAll(tree, columns, "col");
            Set(tree, columns, "equalWidth", Text(definitions.Count == 0));
            Set(tree, columns, "num", definitions.Count == 0 ? currentCount ?? "1" : Text(definitions.Count));
            foreach (WordSectionColumn definition in definitions) {
                int column = Element(tree, "col");
                Set(tree, column, "w", Text(checked((int)definition.Width.ToTwips())));
                if (definition.SpaceAfter.HasValue) Set(tree, column, "space", Text(checked((int)definition.SpaceAfter.Value.ToTwips())));
                Append(tree, columns, column);
            }
        }

        private static void ApplyPageBorders(OfficeXmlTree tree, int sectionProperties, LegacyDocParagraphBorders borders) {
            RemoveAll(tree, sectionProperties, "pgBorders");
            int pageBorders = Element(tree, "pgBorders");
            if (borders.PageOptions.HasNonDefault) {
                Set(tree, pageBorders, "display", borders.PageOptions.Display switch {
                    LegacyDocPageBorderDisplay.FirstPage => "firstPage",
                    LegacyDocPageBorderDisplay.NotFirstPage => "notFirstPage",
                    _ => "allPages"
                });
                Set(tree, pageBorders, "offsetFrom", borders.PageOptions.OffsetFrom == LegacyDocPageBorderOffsetFrom.Page ? "page" : "text");
                Set(tree, pageBorders, "zOrder", borders.PageOptions.ZOrder == LegacyDocPageBorderZOrder.Back ? "back" : "front");
            }

            if (borders.Top.HasAny) AddParagraphBorder(tree, pageBorders, "top", borders.Top, explicitSizeAndSpace: true);
            if (borders.Left.HasAny) AddParagraphBorder(tree, pageBorders, "left", borders.Left, explicitSizeAndSpace: true);
            if (borders.Bottom.HasAny) AddParagraphBorder(tree, pageBorders, "bottom", borders.Bottom, explicitSizeAndSpace: true);
            if (borders.Right.HasAny) AddParagraphBorder(tree, pageBorders, "right", borders.Right, explicitSizeAndSpace: true);
            if (HasElements(tree, pageBorders)) Append(tree, sectionProperties, pageBorders);
        }

        // ---- Document options -----------------------------------------------------------------------------------

        private void ApplyDocumentOptions(IReadOnlyList<int> sections) {
            OfficeXmlTree settings = Settings;
            if (_source.MirrorMargins) Replace(settings, settings.Root, OnOffValue(settings, "mirrorMargins", true), SettingsOrder);
            else RemoveAll(settings, settings.Root, "mirrorMargins");
            Replace(settings, settings.Root, OnOffValue(settings, "gutterAtTop", _source.GutterAtTop), SettingsOrder);
            int compatibility = GetOrAdd(settings, settings.Root, "compat", SettingsOrder);
            SetCompatibilityToggle(settings, compatibility, "doNotBalanceTextColumns", _source.NoColumnBalance);
            SetCompatibilitySetting(settings, compatibility, "compatibilityMode", "11");
            // Word's native DOC layout moves trailing paragraph marks even when later compatibility storage contains
            // a disabled option. Write that effective layout into the editable DOCX.
            SetCompatibilityToggle(settings, compatibility, "splitPgBreakAndParaMark", true);
            // Native DOC numbering advances to an authored or default tab rather than treating the hanging indent as
            // an implicit numbering stop.
            SetCompatibilityToggle(settings, compatibility, "doNotUseIndentAsNumberingTabStop", true);
            if (_source.DefaultTabStop != null) {
                Replace(settings, settings.Root, Element(settings, "defaultTabStop", Text(_source.DefaultTabStop.Value)), SettingsOrder);
            }

            if (_source.RevisionMarkingEnabled || _source.LockedRevisionTrackingEnabled) SetSettingsToggle(settings, "trackRevisions", true);
            if (_source.LockedRevisionTrackingEnabled) {
                int protection = GetOrAdd(settings, settings.Root, "documentProtection", SettingsOrder);
                Set(settings, protection, "edit", "trackedChanges");
                Set(settings, protection, "enforcement", "true");
            }

            if (!_source.DifferentOddAndEvenPages) return;
            foreach (int section in sections) {
                AddHeaderFooterReference(section, isHeader: true, LegacyDocHeaderFooterType.Even);
                AddHeaderFooterReference(section, isHeader: false, LegacyDocHeaderFooterType.Even);
                int setting = Child(settings, settings.Root, "evenAndOddHeaders");
                if (setting == OfficeXmlTree.NoNode) Insert(settings, settings.Root, Element(settings, "evenAndOddHeaders"), SettingsOrder);
                else Set(settings, setting, "val", "true");
            }
        }

        private static void SetSettingsToggle(OfficeXmlTree settings, string name, bool value) {
            int existing = Child(settings, settings.Root, name);
            if (!value) {
                if (existing != OfficeXmlTree.NoNode) settings.Remove(existing);
                return;
            }

            if (existing == OfficeXmlTree.NoNode) Insert(settings, settings.Root, Element(settings, name), SettingsOrder);
            else settings.RemoveAttribute(existing, W, "val");
        }

        private static readonly OfficeXmlChildOrder CompatibilityOrder = Order(
            "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd", "noLeading", "spaceForUL", "noColumnBalance",
            "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing", "doNotLeaveBackslashAlone", "ulTrailSpace", "doNotExpandShiftReturn",
            "spacingInWholePoints", "lineWrapLikeWord6", "printBodyTextBeforeHeader", "printColBlack", "wpSpaceWidth", "showBreaksInFrames",
            "subFontBySize", "suppressBottomSpacing", "suppressTopSpacing", "suppressSpacingAtTopOfPage", "suppressTopSpacingWP",
            "suppressSpBfAfterPgBrk", "swapBordersFacingPages", "convMailMergeEsc", "truncateFontHeightsLikeWP6", "mwSmallCaps",
            "usePrinterMetrics", "doNotSuppressParagraphBorders", "wrapTrailSpaces", "footnoteLayoutLikeWW8", "shapeLayoutLikeWW8",
            "alignTablesRowByRow", "forgetLastTabAlignment", "adjustLineHeightInTable", "autoSpaceLikeWord95", "noSpaceRaiseLower",
            "doNotUseHTMLParagraphAutoSpacing", "layoutRawTableWidth", "layoutTableRowsApart", "useWord97LineBreakRules",
            "doNotBreakWrappedTables", "doNotSnapToGridInCell", "selectFldWithFirstOrLastChar", "applyBreakingRules",
            "doNotWrapTextWithPunct", "doNotUseEastAsianBreakRules", "useWord2002TableStyleRules", "growAutofit", "useFELayout",
            "useNormalStyleForList", "doNotUseIndentAsNumberingTabStop", "useAltKinsokuLineBreakRules", "allowSpaceOfSameStyleInTable",
            "doNotSuppressIndentation", "doNotAutofitConstrainedTables", "autofitToFirstFixedWidthCell", "underlineTabInNumList",
            "displayHangulFixedWidth", "splitPgBreakAndParaMark", "doNotVertAlignCellWithSp", "doNotBreakConstrainedForcedTable",
            "doNotVertAlignInTxbx", "useAnsiKerningPairs", "cachedColBalance", "compatSetting");

        private static void SetCompatibilityToggle(OfficeXmlTree settings, int compatibility, string name, bool value) {
            string localName = name == "doNotBalanceTextColumns" ? "noColumnBalance" : name;
            int existing = Child(settings, compatibility, localName);
            if (!value) {
                if (existing != OfficeXmlTree.NoNode) settings.Remove(existing);
                return;
            }

            if (existing == OfficeXmlTree.NoNode) Insert(settings, compatibility, OnOffValue(settings, localName, true), CompatibilityOrder);
            else Set(settings, existing, "val", "true");
        }

        private static void SetCompatibilitySetting(OfficeXmlTree settings, int compatibility, string name, string value) {
            foreach (int setting in Elements(settings, compatibility, "compatSetting")) {
                if (Get(settings, setting, "name") == name) {
                    Set(settings, setting, "val", value);
                    return;
                }
            }

            int created = Element(settings, "compatSetting");
            Set(settings, created, "name", name);
            Set(settings, created, "uri", "http://schemas.microsoft.com/office/word");
            Set(settings, created, "val", value);
            Insert(settings, compatibility, created, CompatibilityOrder);
        }

        // ---- Header and footer parts ----------------------------------------------------------------------------

        /// <summary>
        /// Adds an empty header or footer of <paramref name="type"/> to the section unless it already has one, and
        /// returns the part.
        /// </summary>
        private LegacyDocDocxPart AddHeaderFooterReference(int sectionProperties, bool isHeader, LegacyDocHeaderFooterType type) {
            OfficeXmlTree tree = _main.Tree;
            string referenceName = isHeader ? "headerReference" : "footerReference";
            string typeName = type.ToXml();
            if (_headerFooterParts.TryGetValue((sectionProperties, isHeader, type), out LegacyDocDocxPart? known)) return known;
            foreach (int reference in Elements(tree, sectionProperties, referenceName)) {
                if ((Get(tree, reference, "type") ?? "default") != typeName) continue;
                string? id = tree.GetAttribute(reference, OfficeXmlNamespaces.Relationships, "id");
                string? target = id == null ? null : _package.GetRelationshipTarget(MainPartName, id);
                string? partName = target == null ? null : OfficeOpcPartNames.ResolveRelationshipTarget(OfficeOpcPartNames.GetRelationshipsPartName(MainPartName), target);
                if (partName != null) return new LegacyDocDocxPart(partName, _package.GetXml(partName));
            }

            string name = NextPartName(isHeader ? "/word/header" : "/word/footer", ".xml");
            OfficeXmlTree part = _package.CreateXmlPart(name, isHeader ? HeaderContentType : FooterContentType, W, isHeader ? "hdr" : "ftr");
            DeclareWordNamespaces(part);
            string relationshipId = _package.AddRelationship(MainPartName, RelationshipTypePrefix + (isHeader ? "header" : "footer"), name, id: NextRelationshipId());
            int created = Element(tree, referenceName);
            Set(tree, created, "type", typeName);
            tree.SetAttribute(created, OfficeXmlNamespaces.Relationships, "id", relationshipId);
            int lastReference = OfficeXmlTree.NoNode;
            foreach (int element in Elements(tree, sectionProperties)) {
                if (Is(tree, element, "headerReference") || Is(tree, element, "footerReference")) lastReference = element;
            }

            if (lastReference == OfficeXmlTree.NoNode) tree.InsertBefore(sectionProperties, created, tree.FirstChildOf(sectionProperties));
            else tree.InsertBefore(sectionProperties, created, tree.NextSiblingOf(lastReference));
            var createdPart = new LegacyDocDocxPart(name, part);
            _headerFooterParts[(sectionProperties, isHeader, type)] = createdPart;
            return createdPart;
        }

        private void AddHeaderFooterStories(IReadOnlyList<int> sections) {
            foreach (LegacyDocHeaderFooterStory story in _source.HeaderFooterStories) {
                if (story.SectionIndex < 0 || story.SectionIndex >= sections.Count || story.Blocks.Count == 0) continue;
                // A section gets its header and footer of one type together.
                LegacyDocDocxPart header = AddHeaderFooterReference(sections[story.SectionIndex], isHeader: true, story.Type);
                LegacyDocDocxPart footer = AddHeaderFooterReference(sections[story.SectionIndex], isHeader: false, story.Type);
                LegacyDocDocxPart part = story.IsHeader ? header : footer;
                OfficeXmlTree tree = part.Tree;
                foreach (int paragraph in Elements(tree, tree.Root, "p")) tree.Remove(paragraph);
                foreach (LegacyDocBodyBlock block in story.Blocks) {
                    if (block is LegacyDocTableBlock tableBlock) {
                        int table = AddTable(part, tableBlock, LegacyDocNoteProjection.Empty, projectNestedTables: true);
                        if (table != OfficeXmlTree.NoNode) {
                            Append(tree, tree.Root, table);
                            AddTableBlockBookmarks(tree, table, tableBlock);
                        }
                        continue;
                    }

                    if (block is not LegacyDocParagraphBlock sourceParagraph) continue;
                    int paragraph = Append(tree, tree.Root, Element(tree, "p"));
                    Append(tree, paragraph, Element(tree, "pPr"));
                    ApplyParagraphFormatting(part, paragraph, sourceParagraph.Format);
                    var bookmarks = LegacyDocBookmarkProjection.Create(sourceParagraph.Bookmarks, sourceParagraph.StartCharacter, sourceParagraph.EndCharacter);
                    AddRuns(new RunTarget(part, paragraph, paragraph), sourceParagraph.Runs, 0, LegacyDocNoteProjection.Empty, bookmarks);
                    bookmarks.EmitRemaining(tree, paragraph);
                }
            }
        }
    }
}
