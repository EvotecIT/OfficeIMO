using OfficeIMO.OneNote;

namespace PSWriteOffice.Cmdlets.OneNote;

internal static class OneNoteCommandUtilities
{
    internal static object Read(string path, OneNoteLoadOptions? options, OneNoteNotebookReaderOptions? notebookOptions)
    {
        var extension = Path.GetExtension(path);
        if (extension.Equals(".one", StringComparison.OrdinalIgnoreCase))
        {
            return OneNoteSection.Load(path, options);
        }

        var effectiveNotebookOptions = CloneNotebookOptions(notebookOptions, options);
        if (extension.Equals(".onetoc2", StringComparison.OrdinalIgnoreCase))
        {
            return OneNoteNotebookReader.Read(path, effectiveNotebookOptions);
        }

        if (extension.Equals(".onepkg", StringComparison.OrdinalIgnoreCase))
        {
            return OneNotePackageReader.Read(path, effectiveNotebookOptions);
        }

        throw new InvalidDataException("OneNote input must use the .one, .onetoc2, or .onepkg extension.");
    }

    private static OneNoteNotebookReaderOptions CloneNotebookOptions(
        OneNoteNotebookReaderOptions? source,
        OneNoteLoadOptions? sectionOptions)
    {
        var effective = source ?? new OneNoteNotebookReaderOptions();
        return new OneNoteNotebookReaderOptions
        {
            OneNoteOptions = sectionOptions ?? effective.OneNoteOptions,
            LoadSectionContent = effective.LoadSectionContent,
            ContinueOnSectionError = effective.ContinueOnSectionError,
            RecurseSectionGroups = effective.RecurseSectionGroups,
            IncludeRecycleBin = effective.IncludeRecycleBin,
            MaxSectionGroupDepth = effective.MaxSectionGroupDepth,
            MaxNotebookEntries = effective.MaxNotebookEntries,
            MaxPackageEntries = effective.MaxPackageEntries,
            MaxPackageExpandedBytes = effective.MaxPackageExpandedBytes,
            MaxPackageEntryBytes = effective.MaxPackageEntryBytes
        };
    }
}
