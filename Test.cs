using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

class JsonToCsvStreaming
{
    static void Main(string[] args)
    {
        string inputPath  = args.Length > 0 ? args[0] : "input.json";
        string outputPath = args.Length > 1 ? args[1] : "output.csv";
        char   delimiter  = args.Length > 2 ? args[2][0] : ',';

        Console.WriteLine($"▶ Lecture : {inputPath}");
        Console.WriteLine($"▶ Sortie  : {outputPath}");
        Console.WriteLine($"▶ Délimiteur : '{delimiter}'\n");

        Convert(inputPath, outputPath, delimiter);
    }

    static void Convert(string inputPath, string outputPath, char delimiter)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        // ── Passe 1 : collecter les headers ──
        Console.WriteLine("Passe 1/2 — Collecte des colonnes…");
        var headers     = new List<string>();
        var headerIndex = new Dictionary<string, int>();

        using (var fs = OpenRead(inputPath))
        {
            var reader = new Utf8JsonReader(ReadAllBytes(fs), new JsonReaderOptions
            {
                AllowTrailingCommas = true,
                CommentHandling     = JsonCommentHandling.Skip
            });

            SkipToArrayStart(ref reader);

            while (reader.Read())
            {
                if (reader.TokenType == JsonTokenType.EndArray) break;
                if (reader.TokenType != JsonTokenType.StartObject) continue;

                while (reader.Read() && reader.TokenType != JsonTokenType.EndObject)
                {
                    if (reader.TokenType == JsonTokenType.PropertyName)
                    {
                        string key = reader.GetString()!;
                        if (!headerIndex.ContainsKey(key))
                        {
                            headerIndex[key] = headers.Count;
                            headers.Add(key);
                        }
                        reader.Read(); // skip value
                    }
                }
            }
        }

        Console.WriteLine($"  → {headers.Count} colonnes : {string.Join(", ", headers)}\n");

        // ── Passe 2 : écrire le CSV ──
        Console.WriteLine("Passe 2/2 — Conversion en cours…");

        using var outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write,
                                             FileShare.None, bufferSize: 1 << 20);
        using var writer    = new StreamWriter(outStream, Encoding.UTF8, bufferSize: 1 << 20);

        writer.WriteLine(string.Join(delimiter, headers.ConvertAll(h => Escape(h, delimiter))));

        long rowCount = 0;
        var  buffer   = new string[headers.Count];

        using (var fs = OpenRead(inputPath))
        {
            var reader = new Utf8JsonReader(ReadAllBytes(fs), new JsonReaderOptions
            {
                AllowTrailingCommas = true,
                CommentHandling     = JsonCommentHandling.Skip
            });

            SkipToArrayStart(ref reader);

            while (reader.Read())
            {
                if (reader.TokenType == JsonTokenType.EndArray) break;
                if (reader.TokenType != JsonTokenType.StartObject) continue;

                Array.Clear(buffer, 0, buffer.Length);

                while (reader.Read() && reader.TokenType != JsonTokenType.EndObject)
                {
                    if (reader.TokenType == JsonTokenType.PropertyName)
                    {
                        string key = reader.GetString()!;
                        reader.Read();
                        if (headerIndex.TryGetValue(key, out int idx))
                            buffer[idx] = GetValue(ref reader);
                        else
                            SkipValue(ref reader);
                    }
                }

                writer.WriteLine(string.Join(delimiter,
                    Array.ConvertAll(buffer, v => Escape(v ?? "", delimiter))));

                rowCount++;
                if (rowCount % 500_000 == 0)
                {
                    double mb = GC.GetTotalMemory(false) / 1_048_576.0;
                    Console.WriteLine($"  {rowCount:N0} lignes — RAM : {mb:F0} Mo — {sw.Elapsed:mm\\:ss}");
                }
            }
        }

        writer.Flush();
        sw.Stop();
        Console.WriteLine($"\n✓ {rowCount:N0} lignes en {sw.ElapsedMilliseconds:N0} ms → {outputPath}");
    }

    static FileStream OpenRead(string path) =>
        new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 1 << 20);

    static byte[] ReadAllBytes(FileStream fs)
    {
        var bytes = new byte[fs.Length];
        int total = 0;
        while (total < bytes.Length)
            total += fs.Read(bytes, total, bytes.Length - total);
        return bytes;
    }

    static void SkipToArrayStart(ref Utf8JsonReader reader)
    {
        while (reader.Read())
            if (reader.TokenType == JsonTokenType.StartArray) return;
        throw new Exception("Tableau JSON '[' introuvable.");
    }

    static string GetValue(ref Utf8JsonReader reader) => reader.TokenType switch
    {
        JsonTokenType.String      => reader.GetString() ?? "",
        JsonTokenType.Number      => reader.GetDecimal().ToString(),
        JsonTokenType.True        => "true",
        JsonTokenType.False       => "false",
        JsonTokenType.Null        => "",
        JsonTokenType.StartObject => ReadRaw(ref reader),
        JsonTokenType.StartArray  => ReadRaw(ref reader),
        _                         => ""
    };

    static string ReadRaw(ref Utf8JsonReader reader)
    {
        int depth = 1;
        var sb = new StringBuilder();
        sb.Append(reader.TokenType == JsonTokenType.StartObject ? '{' : '[');

        while (reader.Read() && depth > 0)
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.StartObject:  depth++; sb.Append('{'); break;
                case JsonTokenType.EndObject:    depth--; sb.Append('}'); break;
                case JsonTokenType.StartArray:   depth++; sb.Append('['); break;
                case JsonTokenType.EndArray:     depth--; sb.Append(']'); break;
                case JsonTokenType.PropertyName: sb.Append($"\"{reader.GetString()}\":"); break;
                case JsonTokenType.String:       sb.Append($"\"{reader.GetString()}\""); break;
                case JsonTokenType.Number:       sb.Append(reader.GetDecimal()); break;
                case JsonTokenType.True:         sb.Append("true"); break;
                case JsonTokenType.False:        sb.Append("false"); break;
                case JsonTokenType.Null:         sb.Append("null"); break;
            }
        }
        return sb.ToString();
    }

    static void SkipValue(ref Utf8JsonReader reader)
    {
        if (reader.TokenType is JsonTokenType.StartObject or JsonTokenType.StartArray)
        {
            int depth = 1;
            while (reader.Read() && depth > 0)
            {
                if (reader.TokenType is JsonTokenType.StartObject or JsonTokenType.StartArray) depth++;
                else if (reader.TokenType is JsonTokenType.EndObject or JsonTokenType.EndArray) depth--;
            }
        }
    }

    static string Escape(string value, char delimiter)
    {
        if (value.Contains(delimiter) || value.Contains('"') ||
            value.Contains('\n')      || value.Contains('\r'))
            return '"' + value.Replace("\"", "\"\"") + '"';
        return value;
    }
}
