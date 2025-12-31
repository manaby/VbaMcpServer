using System.Text.RegularExpressions;

namespace VbaMcpServer.Helpers;

/// <summary>
/// Utility class for normalizing and preprocessing VBA code
/// </summary>
public static class CodeNormalizer
{
    /// <summary>
    /// Normalize line endings to CR+LF (required by VBA engine)
    /// </summary>
    /// <param name="code">VBA code with potentially mixed line endings</param>
    /// <returns>Code with CR+LF line endings</returns>
    public static string NormalizeLineEndings(string code)
    {
        if (string.IsNullOrEmpty(code))
            return code;

        // Convert all line endings to LF first, then to CR+LF
        return code
            .Replace("\r\n", "\n")  // CR+LF → LF
            .Replace("\r", "\n")    // CR → LF
            .Replace("\n", "\r\n"); // LF → CR+LF
    }

    /// <summary>
    /// Unescape XML entities that may have been incorrectly applied by AI
    /// Note: MCP communication uses JSON, not XML, so XML escaping is incorrect
    /// </summary>
    /// <param name="code">VBA code potentially containing XML entities</param>
    /// <returns>Code with XML entities unescaped</returns>
    public static string UnescapeXmlEntities(string code)
    {
        if (string.IsNullOrEmpty(code))
            return code;

        return code
            .Replace("&amp;", "&")      // String concatenation operator
            .Replace("&lt;", "<")       // Less than operator
            .Replace("&gt;", ">")       // Greater than operator
            .Replace("&quot;", "\"")    // String literal
            .Replace("&apos;", "'");    // String literal
    }

    /// <summary>
    /// Extract procedure name from VBA code
    /// </summary>
    /// <param name="code">VBA procedure code starting with Sub/Function/Property declaration</param>
    /// <returns>Procedure name</returns>
    /// <exception cref="ArgumentException">Thrown when procedure name cannot be extracted</exception>
    public static string ExtractProcedureName(string code)
    {
        if (string.IsNullOrEmpty(code))
            throw new ArgumentException("Code cannot be empty", nameof(code));

        // Patterns to match Sub, Function, and Property declarations
        var patterns = new[]
        {
            @"^\s*(?:Public\s+|Private\s+|Friend\s+)?Sub\s+(\w+)",
            @"^\s*(?:Public\s+|Private\s+|Friend\s+)?Function\s+(\w+)",
            @"^\s*(?:Public\s+|Private\s+|Friend\s+)?Property\s+(?:Get|Let|Set)\s+(\w+)"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(code, pattern,
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
        }

        throw new ArgumentException(
            "Could not extract procedure name from code. " +
            "Code must start with Sub, Function, or Property declaration.",
            nameof(code));
    }

    /// <summary>
    /// Preprocess VBA code before writing to module
    /// Applies XML entity unescaping and line ending normalization
    /// </summary>
    /// <param name="code">Raw VBA code</param>
    /// <returns>Preprocessed VBA code ready for insertion</returns>
    public static string PreprocessCode(string code)
    {
        if (string.IsNullOrEmpty(code))
            return code;

        // 1. Unescape XML entities (defensive against AI mistakes)
        code = UnescapeXmlEntities(code);

        // 2. Normalize line endings to CR+LF
        code = NormalizeLineEndings(code);

        return code;
    }
}
