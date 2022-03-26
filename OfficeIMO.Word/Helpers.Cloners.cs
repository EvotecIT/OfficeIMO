using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Word {
    public static partial class Helpers {
        public static MemoryStream ReadAllBytesToMemoryStream(string path) {
            byte[] buffer = File.ReadAllBytes(path);
            var destStream = new MemoryStream(buffer.Length);
            destStream.Write(buffer, 0, buffer.Length);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static MemoryStream CopyFileStreamToMemoryStream(string path) {
            FileStream sourceStream = File.OpenRead(path);
            var destStream = new MemoryStream((int)sourceStream.Length);
            sourceStream.CopyTo(destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static FileStream CopyFileStreamToFileStream(string sourcePath, string destPath) {
            FileStream sourceStream = File.OpenRead(sourcePath);
            FileStream destStream = File.Create(destPath);
            sourceStream.CopyTo(destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static FileStream CopyFileAndOpenFileStream(string sourcePath, string destPath) {
            File.Copy(sourcePath, destPath, true);
            return new FileStream(destPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
    }
}
