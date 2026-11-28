using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>Sequence metadata for document settings in ISO/IEC 29500-4.</summary>
internal static class WordSettingsXml {
    private const int W = OfficeXmlNamespaces.WordprocessingML;

    internal static readonly OfficeXmlChildOrder SettingsOrder = WordOrder(
        "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
        "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
        "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
        "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges",
        "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
        "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate",
        "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge",
        "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection",
        "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
        "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope", "summaryLength",
        "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting",
        "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
        "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
        "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
        "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne",
        "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture",
        "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
        "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags",
        "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
        "docVars", "rsids", "m:mathPr", "uiCompat97To2003", "attachedSchema", "themeFontLang", "clrSchemeMapping",
        "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions",
        "readModeInkLockDown", "smartTagType", "sl:schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
        "decimalSymbol", "listSeparator");

    internal static readonly OfficeXmlChildOrder CompatibilityOrder = WordOrder(
        "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd", "noLeading", "spaceForUL",
        "noColumnBalance", "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing", "doNotLeaveBackslashAlone",
        "ulTrailSpace", "doNotExpandShiftReturn", "spacingInWholePoints", "lineWrapLikeWord6",
        "printBodyTextBeforeHeader", "printColBlack", "wpSpaceWidth", "showBreaksInFrames", "subFontBySize",
        "suppressBottomSpacing", "suppressTopSpacing", "suppressSpacingAtTopOfPage", "suppressTopSpacingWP",
        "suppressSpBfAfterPgBrk", "swapBordersFacingPages", "convMailMergeEsc", "truncateFontHeightsLikeWP6",
        "mwSmallCaps", "usePrinterMetrics", "doNotSuppressParagraphBorders", "wrapTrailSpaces",
        "footnoteLayoutLikeWW8", "shapeLayoutLikeWW8", "alignTablesRowByRow", "forgetLastTabAlignment",
        "adjustLineHeightInTable", "autoSpaceLikeWord95", "noSpaceRaiseLower",
        "doNotUseHTMLParagraphAutoSpacing", "layoutRawTableWidth", "layoutTableRowsApart",
        "useWord97LineBreakRules", "doNotBreakWrappedTables", "doNotSnapToGridInCell",
        "selectFldWithFirstOrLastChar", "applyBreakingRules", "doNotWrapTextWithPunct",
        "doNotUseEastAsianBreakRules", "useWord2002TableStyleRules", "growAutofit", "useFELayout",
        "useNormalStyleForList", "doNotUseIndentAsNumberingTabStop", "useAltKinsokuLineBreakRules",
        "allowSpaceOfSameStyleInTable", "doNotSuppressIndentation", "doNotAutofitConstrainedTables",
        "autofitToFirstFixedWidthCell", "underlineTabInNumList", "displayHangulFixedWidth",
        "splitPgBreakAndParaMark", "doNotVertAlignCellWithSp", "doNotBreakConstrainedForcedTable",
        "doNotVertAlignInTxbx", "useAnsiKerningPairs", "cachedColBalance", "compatSetting");

    internal static OfficeXmlNode EnsureChild(OfficeXmlNode parent, string name, OfficeXmlChildOrder order) {
        OfficeXmlNode existing = parent.WChild(name);
        return existing.IsNone ? AddChild(parent, name, order) : existing;
    }

    internal static OfficeXmlNode AddChild(OfficeXmlNode parent, string name, OfficeXmlChildOrder order) =>
        InsertChild(parent, parent.CreateElement(W, name), order);

    internal static void SetChild(OfficeXmlNode parent, OfficeXmlNode child, OfficeXmlChildOrder order) {
        OfficeXmlNode previous = parent.WChild(child.LocalName);
        if (previous.IsNone) {
            InsertChild(parent, child, order);
        } else {
            parent.InsertBefore(child, previous);
            previous.Remove();
        }
    }

    private static OfficeXmlNode InsertChild(OfficeXmlNode parent, OfficeXmlNode child, OfficeXmlChildOrder order) {
        int rank = order.RankOf(child.NamespaceId, child.LocalName);
        OfficeXmlNode terminalExtension = default;
        for (OfficeXmlNode sibling = parent.FirstChild; sibling.Exists; sibling = sibling.NextSibling) {
            int siblingRank = order.RankOf(sibling.NamespaceId, sibling.LocalName);
            if (siblingRank > rank) return parent.InsertBefore(child, sibling);
            if (siblingRank >= 0) {
                terminalExtension = default;
            } else if (terminalExtension.IsNone &&
                (sibling.LocalName == "extLst" || sibling.NamespaceUri is WordXml.W14 or WordXml.W15)) {
                terminalExtension = sibling;
            }
        }
        return parent.InsertBefore(child, terminalExtension);
    }

    private static OfficeXmlChildOrder WordOrder(params string[] names) {
        var namespaces = new int[names.Length];
        for (int index = 0; index < names.Length; index++) {
            string name = names[index];
            namespaces[index] = name switch {
                "m:mathPr" => OfficeXmlNamespaces.Math,
                "sl:schemaLibrary" => OfficeXmlNamespaces.SchemaLibrary,
                _ => W
            };
            int colon = name.IndexOf(':');
            if (colon >= 0) names[index] = name.Substring(colon + 1);
        }
        return new OfficeXmlChildOrder(namespaces, names);
    }
}
