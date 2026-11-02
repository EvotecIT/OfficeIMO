using System;
using System.Management.Automation;
using System.Net;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

public sealed partial class ReadOfficeDocumentCommand {
    private OfficeReadConfiguration BuildConfiguration() {
        var configuration = OfficeReadConfiguration.Bind(Options, Reader?.DefaultOptions);
        bool Bound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
        bool excel = Bound(nameof(Sheet)) || Bound(nameof(Range)) || Bound(nameof(ChunkRows)) || Bound(nameof(NoHeaderRow));
        bool word = Bound(nameof(NoFootnotes)) || Bound(nameof(IncludePageLocations));
        bool powerPoint = Bound(nameof(NoNotes));
        bool markdown = Bound(nameof(NoHeadingChunks));
        bool mail = Bound(nameof(MaximumStoreItems));
        bool formatOverrides = excel || word || powerPoint || markdown || mail || Bound(nameof(IncludePageLocations));
        if (Reader != null && (configuration.HasFormatOptions || formatOverrides))
            throw new PSArgumentException("Configure format options when creating a caller-provided immutable Reader; common ReaderOptions can still be supplied per read.");
        var common = configuration.Common;
        if (MaximumInputBytes.HasValue) common.MaxInputBytes = MaximumInputBytes.Value;
        if (MaximumXmlCharactersInPart.HasValue) common.OpenXmlMaxCharactersInPart = MaximumXmlCharactersInPart.Value;
        if (MaximumChunkCharacters.HasValue) common.MaxChars = MaximumChunkCharacters.Value;
        if (MaximumTableRows.HasValue) common.MaxTableRows = MaximumTableRows.Value;
        if (Bound(nameof(NoHashes))) common.ComputeHashes = !NoHashes;
        if (Bound(nameof(Password))) common.OpenPassword = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        if (excel) {
            var options = configuration.Formats.Excel = OfficeReadConfiguration.Copy(configuration.Formats.Excel);
            if (Sheet != null) {
                options.SheetNumber = Sheet.Position;
                options.SheetName = Sheet.Name ?? (Sheet.Sheet as OfficeIMO.Excel.ExcelSheet)?.Name;
            }
            if (Bound(nameof(Range))) options.A1Range = Range;
            if (ChunkRows.HasValue) options.ChunkRows = ChunkRows.Value;
            if (Bound(nameof(NoHeaderRow))) options.HeadersInFirstRow = !NoHeaderRow;
        }
        if (word) {
            var options = configuration.Formats.Word = OfficeReadConfiguration.Copy(configuration.Formats.Word);
            if (Bound(nameof(NoFootnotes))) options.IncludeFootnotes = !NoFootnotes;
            if (Bound(nameof(IncludePageLocations))) options.IncludePageLocations = IncludePageLocations;
        }
        if (powerPoint) {
            var options = configuration.Formats.PowerPoint = OfficeReadConfiguration.Copy(configuration.Formats.PowerPoint);
            options.IncludeNotes = !NoNotes;
        }
        if (markdown) {
            var options = configuration.Formats.Markdown = OfficeReadConfiguration.Copy(configuration.Formats.Markdown);
            options.ChunkByHeadings = !NoHeadingChunks;
        }
        if (Bound(nameof(IncludePageLocations))) {
            var options = configuration.Formats.Rtf = OfficeReadConfiguration.Copy(configuration.Formats.Rtf);
            options.IncludePageLocations = IncludePageLocations;
        }
        if (Reader == null && (mail || configuration.Formats.Email == null)) {
            var options = configuration.Formats.Email = OfficeReadConfiguration.Copy(configuration.Formats.Email);
            options.Stores = OfficeReadConfiguration.Copy(options.Stores);
            options.Stores.MaxItems = MaximumStoreItems == 0 ? int.MaxValue : MaximumStoreItems;
        }
        return configuration;
    }
}
