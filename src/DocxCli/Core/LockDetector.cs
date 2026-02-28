namespace DocxCli.Core;

public static class LockDetector
{
    public static string? GetLockFile(string docxPath)
    {
        var dir = Path.GetDirectoryName(docxPath) ?? ".";
        var name = Path.GetFileName(docxPath);

        // Word lock file: ~$filename.docx
        var wordLock = Path.Combine(dir, "~$" + name);
        if (File.Exists(wordLock))
            return wordLock;

        // LibreOffice lock file: .~lock.filename.docx#
        var libreOfficeLock = Path.Combine(dir, ".~lock." + name + "#");
        if (File.Exists(libreOfficeLock))
            return libreOfficeLock;

        return null;
    }

    public static void CheckLock(string docxPath, bool force)
    {
        var lockFile = GetLockFile(docxPath);
        if (lockFile != null && !force)
        {
            Console.Error.WriteLine($"error: file appears to be open in another application ({Path.GetFileName(lockFile)} exists)");
            Console.Error.WriteLine("hint: close the file first, or use --force to edit anyway");
            throw new InvalidOperationException("File is locked");
        }
    }
}
