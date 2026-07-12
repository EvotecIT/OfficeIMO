namespace OfficeIMO.Tests;

internal static class TestFileExtensions {
    internal static bool IsFileLocked(this string fileName) {
        if (string.IsNullOrEmpty(fileName) || !File.Exists(fileName)) {
            return false;
        }

        try {
            using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.None);
            return false;
        } catch (IOException) {
            return true;
        }
    }
}
