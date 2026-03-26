using System.Text;
using MsAccessExtract;
using MsAccessExtract.Helpers;

Console.OutputEncoding = Encoding.UTF8;

// Detect if running interactively (double-click vs terminal)
bool isInteractive = !Console.IsInputRedirected && Environment.UserInteractive;

// Banner
Console.ForegroundColor = ConsoleColor.Cyan;
Console.WriteLine();
Console.WriteLine("  +============================================+");
Console.WriteLine("  |          MsAccessExtract v1.0.0             |");
Console.WriteLine("  |   MS Access -> Git Source Code Extractor    |");
Console.WriteLine("  +============================================+");
Console.ResetColor();
Console.WriteLine();

var logger = new ConsoleLogger();

try
{
    // Determine working directory (where the exe is located)
    string exeDir = AppContext.BaseDirectory;

    // Allow override via command-line argument
    string? dbPath = null;

    if (args.Length > 0)
    {
        // User specified a path
        string arg = args[0];
        if (File.Exists(arg))
        {
            dbPath = Path.GetFullPath(arg);
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"  File not found: {arg}");
            Console.ResetColor();
            WaitForExit(isInteractive);
            return 1;
        }
    }
    else
    {
        // Auto-detect: find .accdb or .mdb in the exe directory
        dbPath = FindAccessDatabase(exeDir);

        if (dbPath == null)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("  No Access database found in the current directory.");
            Console.WriteLine();
            Console.WriteLine("  Usage:");
            Console.WriteLine("    MsAccessExtract.exe                       (auto-detect .accdb/.mdb)");
            Console.WriteLine("    MsAccessExtract.exe <path-to-database>    (specify file)");
            Console.WriteLine();
            Console.WriteLine("  Place this executable in the same folder as your .accdb or .mdb file.");
            Console.ResetColor();
            WaitForExit(isInteractive);
            return 1;
        }
    }

    logger.Info($"Database: {dbPath}");

    using var extractor = new AccessExtractor(logger);
    extractor.Extract(dbPath);

    WaitForExit(isInteractive);
    return logger.HasErrors ? 2 : 0;
}
catch (InvalidOperationException ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine();
    Console.WriteLine($"  ERROR: {ex.Message}");
    Console.ResetColor();
    WaitForExit(isInteractive);
    return 1;
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine();
    Console.WriteLine("  UNEXPECTED ERROR:");
    Console.WriteLine($"  {ex.Message}");
    Console.ForegroundColor = ConsoleColor.DarkGray;
    Console.WriteLine($"\n  {ex.StackTrace}");
    Console.ResetColor();
    WaitForExit(isInteractive);
    return 1;
}

static string? FindAccessDatabase(string directory)
{
    // Prefer .accdb over .mdb
    var accdbFiles = Directory.GetFiles(directory, "*.accdb");
    if (accdbFiles.Length == 1)
        return accdbFiles[0];

    var mdbFiles = Directory.GetFiles(directory, "*.mdb");
    if (mdbFiles.Length == 1)
        return mdbFiles[0];

    // Multiple files found — pick the first .accdb, then .mdb
    if (accdbFiles.Length > 1)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"  Multiple .accdb files found ({accdbFiles.Length}). Using: {Path.GetFileName(accdbFiles[0])}");
        Console.ResetColor();
        return accdbFiles[0];
    }

    if (mdbFiles.Length > 1)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"  Multiple .mdb files found ({mdbFiles.Length}). Using: {Path.GetFileName(mdbFiles[0])}");
        Console.ResetColor();
        return mdbFiles[0];
    }

    // Check all access files together
    if (accdbFiles.Length == 0 && mdbFiles.Length == 0)
        return null;

    return accdbFiles.Length > 0 ? accdbFiles[0] : mdbFiles[0];
}

static void WaitForExit(bool isInteractive)
{
    if (!isInteractive) return;
    Console.WriteLine();
    Console.ForegroundColor = ConsoleColor.DarkGray;
    Console.Write("  Press any key to exit...");
    Console.ResetColor();
    Console.ReadKey(intercept: true);
    Console.WriteLine();
}
