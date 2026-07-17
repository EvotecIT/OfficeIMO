namespace OfficeIMO.OneNote;

/// <summary>
/// Checks the generated section's page identity/order/relationships, structural content hierarchy,
/// rich-text runs, supported layout and media metadata, and binary payload resolution state.
/// </summary>
internal static class OneNoteWriteRoundTripValidator {
    internal static void ValidateSection(OneNoteSection expected, OneNoteSection actual) {
        if (expected == null) throw new ArgumentNullException(nameof(expected));
        if (actual == null) throw new ArgumentNullException(nameof(actual));
        if (!string.Equals(expected.Name, actual.Name, StringComparison.Ordinal)) Fail("section name");
        ValidatePageList(expected.Pages, actual.Pages, "section", PageRole.Current);
    }

    private static void ValidatePageList(
        IList<OneNotePage> expected,
        IList<OneNotePage> actual,
        string path,
        PageRole role) {
        if (expected.Count != actual.Count) Fail(path + " page count");
        for (int index = 0; index < expected.Count; index++) {
            ValidatePage(expected[index], actual[index], path + "/page[" + index + "]", role);
        }
    }

    private static void ValidatePage(OneNotePage expected, OneNotePage actual, string path, PageRole role) {
        if (!Equals(expected.Id, actual.Id)) Fail(path + " identity");
        if (!Equals(expected.RevisionContextId, actual.RevisionContextId)) Fail(path + " revision context");
        if (!string.Equals(expected.Title, actual.Title, StringComparison.Ordinal)) Fail(path + " title");
        bool expectedConflict = role == PageRole.Conflict || expected.IsConflictPage;
        bool expectedVersion = role == PageRole.Version || expected.IsVersionHistoryPage;
        if (expected.Level != actual.Level ||
            expected.IsDeleted != actual.IsDeleted ||
            expectedConflict != actual.IsConflictPage ||
            expectedVersion != actual.IsVersionHistoryPage ||
            !NormalizedStringEquals(expected.OriginalAuthor, actual.OriginalAuthor) ||
            !NormalizedStringEquals(expected.MostRecentAuthor, actual.MostRecentAuthor) ||
            !FloatEquals(expected.Width, actual.Width) ||
            !FloatEquals(expected.Height, actual.Height)) {
            Fail(path + " metadata");
        }

        if (expected.Outlines.Count != actual.Outlines.Count) Fail(path + " outline count");
        for (int index = 0; index < expected.Outlines.Count; index++) {
            ValidateElement(expected.Outlines[index], actual.Outlines[index], path + "/outline[" + index + "]");
        }
        ValidateElementList(expected.DirectContent, actual.DirectContent, path + "/direct");
        ValidatePageList(expected.ConflictPages, actual.ConflictPages, path + "/conflicts", PageRole.Conflict);
        ValidatePageList(expected.VersionHistory, actual.VersionHistory, path + "/versions", PageRole.Version);
    }

    private static void ValidateElementList(
        IList<OneNoteElement> expected,
        IList<OneNoteElement> actual,
        string path) {
        if (expected.Count != actual.Count) Fail(path + " content count");
        for (int index = 0; index < expected.Count; index++) {
            ValidateElement(expected[index], actual[index], path + "[" + index + "]");
        }
    }

    private static void ValidateElement(OneNoteElement expected, OneNoteElement actual, string path) {
        if (expected.Kind != actual.Kind) Fail(path + " kind");
        ValidateLayout(expected.Layout, actual.Layout, path + " layout");
        if (!NormalizedStringEquals(expected.Author?.Name, actual.Author?.Name)) Fail(path + " author");

        if (expected is OneNoteOutline expectedOutline && actual is OneNoteOutline actualOutline) {
            ValidateElementList(expectedOutline.Children, actualOutline.Children, path + "/children");
            return;
        }
        if (expected is OneNoteParagraph expectedParagraph && actual is OneNoteParagraph actualParagraph) {
            ValidateParagraph(expectedParagraph, actualParagraph, path);
            return;
        }
        if (expected is OneNoteTable expectedTable && actual is OneNoteTable actualTable) {
            ValidateTable(expectedTable, actualTable, path);
            return;
        }
        if (expected is OneNoteImage expectedImage && actual is OneNoteImage actualImage) {
            ValidateBinary(expectedImage, actualImage, path);
            if (!NormalizedStringEquals(expectedImage.AltText, actualImage.AltText) ||
                !NormalizedStringEquals(expectedImage.SourcePath, actualImage.SourcePath) ||
                !NormalizedStringEquals(expectedImage.Hyperlink, actualImage.Hyperlink) ||
                !FloatEquals(expectedImage.WidthHalfInches, actualImage.WidthHalfInches) ||
                !FloatEquals(expectedImage.HeightHalfInches, actualImage.HeightHalfInches)) {
                Fail(path + " image metadata");
            }
            return;
        }
        if (expected is OneNoteEmbeddedFile expectedFile && actual is OneNoteEmbeddedFile actualFile) {
            ValidateBinary(expectedFile, actualFile, path);
            if (!NormalizedStringEquals(expectedFile.SourcePath, actualFile.SourcePath)) Fail(path + " source path");
            return;
        }
        if (expected is OneNoteMedia expectedMedia && actual is OneNoteMedia actualMedia) {
            ValidateBinary(expectedMedia, actualMedia, path);
            if (expectedMedia.RecordingKind != actualMedia.RecordingKind ||
                !NormalizedStringEquals(expectedMedia.SourcePath, actualMedia.SourcePath)) {
                Fail(path + " media metadata");
            }
            return;
        }
        if (expected is OneNoteMath expectedMath && actual is OneNoteMath actualMath) {
            if (!string.Equals(expectedMath.Text, actualMath.Text, StringComparison.Ordinal) ||
                !string.Equals(expectedMath.MathMl, actualMath.MathMl, StringComparison.Ordinal) ||
                !string.Equals(expectedMath.Latex, actualMath.Latex, StringComparison.Ordinal)) {
                Fail(path + " math projection");
            }
            ValidatePayload(expectedMath.RawPayload, actualMath.RawPayload, path + " math payload");
            return;
        }

        if (expected is OneNoteBinaryElement expectedBinary && actual is OneNoteBinaryElement actualBinary) {
            ValidateBinary(expectedBinary, actualBinary, path);
        }
    }

    private static void ValidateParagraph(OneNoteParagraph expected, OneNoteParagraph actual, string path) {
        int expectedRunCount = Math.Max(1, expected.Runs.Count);
        if (expectedRunCount != actual.Runs.Count) Fail(path + " run count");
        for (int index = 0; index < expectedRunCount; index++) {
            OneNoteTextRun? expectedRun = expected.Runs.Count == 0 ? null : expected.Runs[index];
            OneNoteTextRun actualRun = actual.Runs[index];
            if (!string.Equals(expectedRun?.Text ?? string.Empty, actualRun.Text, StringComparison.Ordinal) ||
                !NormalizedStringEquals(expectedRun?.Hyperlink, actualRun.Hyperlink) ||
                (expectedRun?.HyperlinkProtected ?? false) != actualRun.HyperlinkProtected) {
                Fail(path + "/run[" + index + "]");
            }
            if (expectedRun != null) ValidateTextStyle(expectedRun.Style, actualRun.Style, path + "/run[" + index + "] style");
        }
        ValidateList(expected.List, actual.List, path + " list");
        ValidateParagraphStyle(expected.Style, actual.Style, path + " paragraph style");
        ValidateElementList(expected.Children, actual.Children, path + "/children");
    }

    private static void ValidateList(OneNoteListInfo? expected, OneNoteListInfo? actual, string path) {
        if (expected == null || actual == null) {
            if (expected != null || actual != null) Fail(path);
            return;
        }

        uint? expectedFormat = expected.Ordered ? expected.Format ?? 0U : null;
        bool expectedRestart = expected.Restart || expected.DisplayIndex.HasValue;
        int? expectedDisplayIndex = expectedRestart ? Math.Max(1, expected.DisplayIndex ?? 1) : (int?)null;
        if (expected.Ordered != actual.Ordered ||
            expectedFormat != actual.Format ||
            expected.Level != actual.Level ||
            expectedRestart != actual.Restart ||
            expectedDisplayIndex != actual.DisplayIndex ||
            !NormalizedStringEquals(expected.FontFamily, actual.FontFamily)) {
            Fail(path);
        }
    }

    private static void ValidateTable(OneNoteTable expected, OneNoteTable actual, string path) {
        if (expected.BordersVisible != actual.BordersVisible || expected.ColumnWidths.Count != actual.ColumnWidths.Count) {
            Fail(path + " table metadata");
        }
        for (int index = 0; index < expected.ColumnWidths.Count; index++) {
            if (!FloatEquals(expected.ColumnWidths[index], actual.ColumnWidths[index])) Fail(path + "/column[" + index + "]");
        }
        if (expected.Rows.Count != actual.Rows.Count) Fail(path + " row count");
        for (int rowIndex = 0; rowIndex < expected.Rows.Count; rowIndex++) {
            OneNoteTableRow expectedRow = expected.Rows[rowIndex];
            OneNoteTableRow actualRow = actual.Rows[rowIndex];
            if (expectedRow.Cells.Count != actualRow.Cells.Count) Fail(path + "/row[" + rowIndex + "] cell count");
            for (int cellIndex = 0; cellIndex < expectedRow.Cells.Count; cellIndex++) {
                OneNoteTableCell expectedCell = expectedRow.Cells[cellIndex];
                OneNoteTableCell actualCell = actualRow.Cells[cellIndex];
                if (expectedCell.ShadingColorArgb != actualCell.ShadingColorArgb) {
                    Fail(path + "/row[" + rowIndex + "]/cell[" + cellIndex + "] shading");
                }
                ValidateElementList(
                    expectedCell.Content,
                    actualCell.Content,
                    path + "/row[" + rowIndex + "]/cell[" + cellIndex + "]");
            }
        }
    }

    private static void ValidateBinary(OneNoteBinaryElement expected, OneNoteBinaryElement actual, string path) {
        if (!NormalizedStringEquals(expected.FileName, actual.FileName)) Fail(path + " file name");
        ValidatePayload(expected.Payload, actual.Payload, path + " payload");
    }

    private static void ValidatePayload(OneNoteBinaryPayload? expected, OneNoteBinaryPayload? actual, string path) {
        if ((expected == null) != (actual == null)) Fail(path + " resolution");
        if (expected?.Length.HasValue == true && expected.Length != actual?.Length) Fail(path + " length");
    }

    private static void ValidateLayout(OneNoteLayout? expected, OneNoteLayout? actual, string path) {
        if (!FloatEquals(expected?.X, actual?.X) ||
            !FloatEquals(expected?.Y, actual?.Y) ||
            !FloatEquals(expected?.Width, actual?.Width) ||
            !FloatEquals(expected?.Height, actual?.Height) ||
            !ExpectedNullableEquals(expected?.Tight, actual?.Tight) ||
            !ExpectedNullableEquals(expected?.RightToLeft, actual?.RightToLeft)) {
            Fail(path);
        }
    }

    private static void ValidateTextStyle(OneNoteTextStyle expected, OneNoteTextStyle actual, string path) {
        double? expectedFontSize = expected.FontSize.HasValue
            ? Math.Round(expected.FontSize.Value * 2, MidpointRounding.AwayFromZero) / 2.0
            : (double?)null;
        if (!NormalizedStringEquals(expected.FontFamily, actual.FontFamily) ||
            !ExpectedNullableEquals(expectedFontSize, actual.FontSize) ||
            !ExpectedNullableEquals(expected.ColorArgb, actual.ColorArgb) ||
            !ExpectedNullableEquals(expected.HighlightColorArgb, actual.HighlightColorArgb) ||
            !ExpectedNullableEquals(expected.Bold, actual.Bold) ||
            !ExpectedNullableEquals(expected.Italic, actual.Italic) ||
            !ExpectedNullableEquals(expected.Underline, actual.Underline) ||
            !ExpectedNullableEquals(expected.Strikethrough, actual.Strikethrough) ||
            !ExpectedNullableEquals(expected.Superscript, actual.Superscript) ||
            !ExpectedNullableEquals(expected.Subscript, actual.Subscript) ||
            !ExpectedNullableEquals(expected.LanguageId, actual.LanguageId) ||
            !ExpectedNullableEquals(expected.IsMath, actual.IsMath)) {
            Fail(path);
        }
    }

    private static void ValidateParagraphStyle(OneNoteParagraphStyle expected, OneNoteParagraphStyle actual, string path) {
        if (!NormalizedStringEquals(expected.StyleId, actual.StyleId) ||
            !ExpectedNullableEquals(expected.Alignment, actual.Alignment) ||
            !FloatEquals(expected.SpaceBefore, actual.SpaceBefore) ||
            !FloatEquals(expected.SpaceAfter, actual.SpaceAfter) ||
            !FloatEquals(expected.ExactLineSpacing, actual.ExactLineSpacing)) {
            Fail(path);
        }
    }

    private static bool FloatEquals(double expected, double actual) => (double)(float)expected == actual;

    private static bool FloatEquals(double? expected, double? actual) =>
        !expected.HasValue ? !actual.HasValue : actual.HasValue && FloatEquals(expected.Value, actual.Value);

    private static bool ExpectedNullableEquals<T>(T? expected, T? actual) where T : struct =>
        !expected.HasValue || expected.Equals(actual);

    private static bool NormalizedStringEquals(string? expected, string? actual) =>
        string.Equals(NormalizeString(expected), NormalizeString(actual), StringComparison.Ordinal);

    private static string? NormalizeString(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    private static void Fail(string boundary) => throw new OneNoteFormatException(
        "ONENOTE_WRITE_ROUNDTRIP_SEMANTICS",
        "The generated OneNote section changed validated public-model semantics at " + boundary + ".");

    private enum PageRole {
        Current,
        Conflict,
        Version
    }
}
