using System.Diagnostics;

namespace MsAccessExtract.Helpers;

/// <summary>
/// Verbose console logger with colors and counters.
/// </summary>
internal sealed class ConsoleLogger
{
    private readonly Dictionary<string, int> _counts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Stopwatch _stopwatch = new();
    private int _errors;
    private int _warnings;

    public void StartTimer() => _stopwatch.Start();
    public void StopTimer() => _stopwatch.Stop();

    public void Header(string message)
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine();
        Console.WriteLine($"═══ {message} ═══");
        Console.ResetColor();
    }

    public void Section(string category)
    {
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.WriteLine();
        Console.WriteLine($"── {category} ──");
        Console.ResetColor();
    }

    public void Success(string category, string name)
    {
        IncrementCount(category);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  ✓ ");
        Console.ResetColor();
        Console.WriteLine(name);
    }

    public void Warn(string message)
    {
        _warnings++;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.Write("  ⚠ ");
        Console.ResetColor();
        Console.WriteLine(message);
    }

    public void Error(string message)
    {
        _errors++;
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Write("  ✗ ");
        Console.ResetColor();
        Console.WriteLine(message);
    }

    public void Info(string message)
    {
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.Write("  ℹ ");
        Console.ResetColor();
        Console.WriteLine(message);
    }

    public void Skip(string name, string reason)
    {
        Console.ForegroundColor = ConsoleColor.DarkYellow;
        Console.Write("  ○ ");
        Console.ResetColor();
        Console.WriteLine($"{name} (skipped: {reason})");
    }

    public void PrintSummary()
    {
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("═══ Summary ═══");
        Console.ResetColor();

        foreach (var kvp in _counts.OrderBy(x => x.Key))
        {
            Console.Write("  ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write($"{kvp.Value,4}");
            Console.ResetColor();
            Console.WriteLine($"  {kvp.Key}");
        }

        int total = _counts.Values.Sum();
        Console.WriteLine($"  ──────────────");
        Console.ForegroundColor = ConsoleColor.White;
        Console.Write($"  {total,4}");
        Console.ResetColor();
        Console.WriteLine("  Total objects exported");

        if (_warnings > 0)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"  {_warnings,4}  Warnings");
            Console.ResetColor();
        }

        if (_errors > 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"  {_errors,4}  Errors");
            Console.ResetColor();
        }

        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine($"\n  Elapsed: {_stopwatch.Elapsed:hh\\:mm\\:ss\\.fff}");
        Console.ResetColor();
    }

    public bool HasErrors => _errors > 0;

    private void IncrementCount(string category)
    {
        if (!_counts.TryGetValue(category, out int count))
            count = 0;
        _counts[category] = count + 1;
    }
}
