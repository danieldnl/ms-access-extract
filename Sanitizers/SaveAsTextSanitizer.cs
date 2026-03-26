using System.Text;
using System.Text.RegularExpressions;

namespace MsAccessExtract.Sanitizers;

/// <summary>
/// Removes noise from SaveAsText exports to produce clean, diff-friendly output.
/// Inspired by MSAccessVCS sanitization patterns.
/// </summary>
internal static partial class SaveAsTextSanitizer
{
    // Properties that produce single-line noise
    private static readonly HashSet<string> SingleLineRemovals = new(StringComparer.OrdinalIgnoreCase)
    {
        "Checksum",
        "NoSaveCTID",
        "GUID",
    };

    // Properties that produce multi-line binary blocks (base64/hex encoded)
    private static readonly HashSet<string> MultiLineBlockRemovals = new(StringComparer.OrdinalIgnoreCase)
    {
        "PrtMip",
        "PrtDevMode",
        "PrtDevNames",
        "PrtDevModeW",
        "PrtDevNamesW",
        "NameMap",
    };

    [GeneratedRegex(@"^\s+(\w+)\s*=\s*", RegexOptions.Compiled)]
    private static partial Regex PropertyLineRegex();

    [GeneratedRegex(@"^\s+dbLongBinary\s+""OLE""", RegexOptions.Compiled)]
    private static partial Regex OleBlobRegex();

    /// <summary>
    /// Sanitize the content of a SaveAsText export file.
    /// </summary>
    public static string Sanitize(string content)
    {
        var lines = content.Split('\n');
        var result = new StringBuilder(content.Length);
        bool skippingBlock = false;
        bool skippingOleBlob = false;

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            string trimmed = line.TrimEnd('\r');

            // Handle OLE blob removal (dbLongBinary "OLE" blocks)
            if (OleBlobRegex().IsMatch(trimmed))
            {
                skippingOleBlob = true;
                continue;
            }

            if (skippingOleBlob)
            {
                // OLE blobs end when we hit a line that starts a new property at the same or lesser indent
                if (!trimmed.StartsWith("    ") && trimmed.Length > 0
                    || (trimmed.Length > 0 && !char.IsWhiteSpace(trimmed[0])))
                {
                    skippingOleBlob = false;
                    // Fall through to process this line normally
                }
                else
                {
                    continue;
                }
            }

            // Handle multi-line block removal
            if (skippingBlock)
            {
                // Block continues while lines are indented continuation (hex/base64 data)
                // Blocks end when we hit a non-continuation line
                if (IsContinuationLine(trimmed))
                {
                    continue;
                }
                skippingBlock = false;
                // Fall through to process this line normally
            }

            // Check for property assignments
            var match = PropertyLineRegex().Match(trimmed);
            if (match.Success)
            {
                string propName = match.Groups[1].Value;

                // Single-line removals
                if (SingleLineRemovals.Contains(propName))
                {
                    continue;
                }

                // Multi-line block removals
                if (MultiLineBlockRemovals.Contains(propName))
                {
                    skippingBlock = true;
                    continue;
                }
            }

            result.Append(trimmed);
            result.Append('\n');
        }

        // Remove trailing newlines but keep one
        string output = result.ToString().TrimEnd('\n') + "\n";
        return output;
    }

    /// <summary>
    /// Sanitize a file in-place.
    /// </summary>
    public static void SanitizeFile(string filePath)
    {
        if (!File.Exists(filePath)) return;

        string content = File.ReadAllText(filePath, Encoding.UTF8);
        string sanitized = Sanitize(content);
        File.WriteAllText(filePath, sanitized, new UTF8Encoding(false));
    }

    /// <summary>
    /// Checks if a line is a continuation of a multi-line property value.
    /// These are typically indented hex/base64 data blocks.
    /// Pattern: lines that start with spaces and contain hex-like data or "Begin"/"End" blocks.
    /// </summary>
    private static bool IsContinuationLine(string trimmedLine)
    {
        if (string.IsNullOrWhiteSpace(trimmedLine))
            return false;

        // Continuation lines in SaveAsText are indented with spaces and contain
        // base64 or hex data like: "    0x0000..." or "    AAAA..."
        if (trimmedLine.Length > 4 && trimmedLine.StartsWith("    "))
        {
            string data = trimmedLine.Trim();
            // Hex data
            if (data.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                return true;
            // Base64-like data (pure alphanumeric + padding)
            if (IsBase64Like(data))
                return true;
        }

        return false;
    }

    private static bool IsBase64Like(string s)
    {
        if (s.Length < 4) return false;
        foreach (char c in s)
        {
            if (!char.IsLetterOrDigit(c) && c != '+' && c != '/' && c != '=')
                return false;
        }
        return true;
    }
}
