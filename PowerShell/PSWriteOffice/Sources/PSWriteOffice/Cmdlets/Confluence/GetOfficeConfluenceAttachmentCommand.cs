using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Confluence;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Lists or downloads Confluence page attachments.</summary>
/// <example>
/// <summary>List all attachments for a page.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Get-OfficeConfluenceAttachment -Session $session -PageId 12345</code>
/// <para>Follows attachment cursors and streams attachment metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeConfluenceAttachment", DefaultParameterSetName = ParameterSetList, SupportsShouldProcess = true)]
[OutputType(typeof(ConfluenceAttachment), typeof(ConfluenceAttachmentBatch), typeof(byte[]), typeof(FileInfo))]
public sealed class GetOfficeConfluenceAttachmentCommand : AsyncPSCmdlet
{
    private const string ParameterSetList = "List";
    private const string ParameterSetDownload = "Download";

    /// <summary>Configured Confluence session.</summary>
    [Parameter(Mandatory = true)]
    public ConfluenceSession Session { get; set; } = null!;

    /// <summary>Page identifier.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string PageId { get; set; } = string.Empty;

    /// <summary>Attachment identifier to download.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDownload)]
    [ValidateNotNullOrEmpty]
    public string AttachmentId { get; set; } = string.Empty;

    /// <summary>Optional destination path. Without this parameter, the download is returned as one byte array.</summary>
    [Parameter(ParameterSetName = ParameterSetDownload)]
    public string? OutFile { get; set; }

    /// <summary>Overwrite an existing destination file.</summary>
    [Parameter(ParameterSetName = ParameterSetDownload)]
    public SwitchParameter Force { get; set; }

    /// <summary>Optional cursor at which to resume attachment listing.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    public string? Cursor { get; set; }

    /// <summary>Maximum attachments requested per listing batch.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    [ValidateRange(1, 250)]
    public int Limit { get; set; } = 50;

    /// <summary>Return attachment batches rather than individual metadata objects.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    public SwitchParameter AsPage { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        using var client = Session.CreateClient();
        if (ParameterSetName == ParameterSetDownload)
        {
            if (string.IsNullOrWhiteSpace(OutFile))
            {
                var downloadedBytes = await client.DownloadAttachmentAsync(PageId, AttachmentId, CancelToken).ConfigureAwait(false);
                WriteObject(downloadedBytes, enumerateCollection: false);
                return;
            }

            var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutFile);
            if (File.Exists(path) && !Force.IsPresent)
            {
                throw new IOException($"File '{path}' already exists. Use -Force to overwrite it.");
            }

            if (!ShouldProcess(path, "Write Confluence attachment"))
            {
                return;
            }

            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }
            else
            {
                directory = Directory.GetCurrentDirectory();
            }

            string temporaryPath = AtomicFileWriter.CreateTemporaryPath(directory, Path.GetFileName(path));
            try
            {
                using (var destination = new FileStream(
                           temporaryPath,
                           FileMode.CreateNew,
                           FileAccess.Write,
                           FileShare.None,
                           81920,
                           useAsync: true))
                {
                    await client.DownloadAttachmentAsync(PageId, AttachmentId, destination, CancelToken).ConfigureAwait(false);
                }

                if (Force.IsPresent)
                {
                    ReplaceFile(temporaryPath, path);
                }
                else
                {
                    File.Move(temporaryPath, path);
                }
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
            WriteObject(new FileInfo(path));
            return;
        }

        var cursor = Cursor;
        var observed = new HashSet<string>(StringComparer.Ordinal);
        if (cursor != null)
        {
            observed.Add(cursor);
        }

        do
        {
            var requestedCursor = cursor;
            var batch = await client.GetAttachmentsAsync(PageId, requestedCursor, Limit, CancelToken).ConfigureAwait(false);
            if (AsPage.IsPresent)
            {
                WriteObject(batch);
            }
            else
            {
                foreach (var attachment in batch.Attachments)
                {
                    WriteObject(attachment);
                }
            }

            cursor = batch.NextCursor;
            if (cursor != null &&
                (string.Equals(cursor, requestedCursor, StringComparison.Ordinal) || !observed.Add(cursor)))
            {
                throw new InvalidOperationException("Confluence returned a repeated attachment cursor and cannot make progress.");
            }
        }
        while (cursor != null);
    }

    private static void ReplaceFile(string temporaryPath, string destinationPath)
    {
#if FRAMEWORK
        for (var attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                File.Move(temporaryPath, destinationPath);
                return;
            }
            catch (IOException) when (File.Exists(destinationPath))
            {
                try
                {
                    File.Replace(temporaryPath, destinationPath, null);
                    return;
                }
                catch (FileNotFoundException) when (attempt < 2)
                {
                }
            }
        }
        throw new IOException($"Could not atomically replace file '{destinationPath}'.");
#else
        File.Move(temporaryPath, destinationPath, overwrite: true);
#endif
    }
}
