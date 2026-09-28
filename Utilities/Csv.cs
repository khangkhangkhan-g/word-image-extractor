namespace WordImageExtractor.Utilities;

public static class Csv
{
    public static string Escape(string? value)
    {
        value ??= string.Empty;
        if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0) return value;
        return $"\"{value.Replace("\"", "\"\"")}\"";
    }
}
