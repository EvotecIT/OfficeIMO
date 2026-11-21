namespace OfficeIMO.Workflows;

/// <summary>The local file a file printer, such as Microsoft Print to PDF, writes a job to.</summary>
internal static class PrintOutputFile {
    /// <summary>
    /// Removes the output of a job that failed or was canceled when the job created the file and nothing was written to it.
    /// A file that existed before the job, or that holds any data, is kept. Cleanup problems are ignored; the job's own error is the result.
    /// </summary>
    /// <param name="path">Output path handed to the printer.</param>
    /// <param name="existedBefore">Whether the path existed when the job started.</param>
    /// <returns>True when an empty output file was removed.</returns>
    internal static bool DiscardEmpty(string path, bool existedBefore) {
        if (existedBefore) return false;
        for (int attempt = 0; ; attempt++) {
            try {
                // Sharing only deletion keeps writers out between the length check and the delete.
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Delete);
                if (stream.Length != 0) return false;
                File.Delete(path);
                return true;
            } catch (FileNotFoundException) {
                return false;
            } catch (DirectoryNotFoundException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (IOException) when (attempt < 3) {
                // The spooler can hold the file briefly after the job is aborted.
                Thread.Sleep(100);
            } catch (IOException) {
                return false;
            }
        }
    }
}
