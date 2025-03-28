namespace OfficeIMO.Word;
/// <summary>
/// Provides helper methods for stream and file operations.
/// </summary>
public static partial class Helpers {
    /// <summary>
    /// Reads all bytes from a file and writes them to a MemoryStream.
    /// </summary>
    /// <param name="path">The path of the file to read.</param>
    /// <returns>A MemoryStream containing the file's bytes.</returns>
    public static MemoryStream ReadAllBytesToMemoryStream(string path) {
        byte[] buffer = File.ReadAllBytes(path);
        var destStream = new MemoryStream(buffer.Length);
        destStream.Write(buffer, 0, buffer.Length);
        destStream.Seek(0, SeekOrigin.Begin);
        return destStream;
    }

    /// <summary>
    /// Copies the contents of a file stream to a MemoryStream.
    /// </summary>
    /// <param name="path">The path of the file to read.</param>
    /// <returns>A MemoryStream containing the file's contents.</returns>
    public static MemoryStream CopyFileStreamToMemoryStream(string path) {
        FileStream sourceStream = File.OpenRead(path);
        var destStream = new MemoryStream((int)sourceStream.Length);
        sourceStream.CopyTo(destStream);
        destStream.Seek(0, SeekOrigin.Begin);
        return destStream;
    }

    /// <summary>
    /// Copies the contents of one file stream to another file stream.
    /// </summary>
    /// <param name="sourcePath">The path of the source file.</param>
    /// <param name="destPath">The path of the destination file.</param>
    /// <returns>A FileStream for the destination file.</returns>
    public static FileStream CopyFileStreamToFileStream(string sourcePath, string destPath) {
        FileStream sourceStream = File.OpenRead(sourcePath);
        FileStream destStream = File.Create(destPath);
        sourceStream.CopyTo(destStream);
        destStream.Seek(0, SeekOrigin.Begin);
        return destStream;
    }

    /// <summary>
    /// Copies a file and opens a file stream for the copied file.
    /// </summary>
    /// <param name="sourcePath">The path of the source file.</param>
    /// <param name="destPath">The path of the destination file.</param>
    /// <returns>A FileStream for the copied file.</returns>
    public static FileStream CopyFileAndOpenFileStream(string sourcePath, string destPath) {
        File.Copy(sourcePath, destPath, true);
        return new FileStream(destPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }

    /// <summary>
    /// Reads the contents of a file into a MemoryStream.
    /// </summary>
    /// <param name="path">The path of the file to read.</param>
    /// <returns>A MemoryStream containing the file's contents.</returns>
    public static MemoryStream ReadFileToMemoryStream(string path) {
        using (Stream sourceStream = File.OpenRead(path)) {
            var destStream = new MemoryStream((int)sourceStream.Length);
            CopyStream(sourceStream, destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }
    }

    /// <summary>
    /// Copies the contents of one stream to another stream.
    /// </summary>
    /// <param name="source">The source stream.</param>
    /// <param name="target">The target stream.</param>
    public static void CopyStream(Stream source, Stream target) {
        if (source != null) {
            MemoryStream mstream = source as MemoryStream;
            if (mstream != null) {
                mstream.WriteTo(target);
            } else {
                byte[] buffer = new byte[2048];
                int length = buffer.Length, size;

                while ((size = source.Read(buffer, 0, length)) != 0) {
                    target.Write(buffer, 0, size);
                }
            }
        }
    }
}
