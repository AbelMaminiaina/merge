using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

class JsonToCsv
{
    static void Main(string[] args)
    {
        string inputPath  = args.Length > 0 ? args[0] : "input.json";
        string outputPath = args.Length > 1 ? args[1] : "output.csv";
        char   delimiter  = args.Length > 2 ? args[2][0] : ',';

        Convert(inputPath, outputPath, delimiter);
    }

    static void Convert(string inputPath, string outputPath, char delimiter = ',')
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        using var inputStream  = new FileStream(inputPath,  FileMode.Open,   FileAccess.Read,  FileShare.Read,  bufferSize: 1 << 20);
        using var outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None,  bufferSize: 1 << 20);
        using var writer       = new StreamWriter(outputStream, Encoding.UTF8, bufferSize: 1 << 20);

        var jsonDoc = JsonDocument.Parse(inputStream, new JsonDocumentOptions
        {
            AllowTrailingCommas = true,
            CommentHandling     = JsonCommentHandling.Skip
        });

        var root = jsonDoc.RootElement;
        if (root.ValueKind != JsonValueKind.Array)
            throw new Exception("Le JSON doit être un tableau d'objets.");

        // ── 1. Collecter tous les headers (union des clés) ──
        var headers = new List<string>();
        var headerIndex = new Dictionary<string, int>();

        foreach (var element in root.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object) continue;
            foreach (var prop in element.EnumerateObject())
            {
                if (!headerIndex.ContainsKey(prop.Name))
                {
                    headerIndex[prop.Name] = headers.Count;
                    headers.Add(prop.Name);
                }
            }
        }

        // ── 2. Écrire les headers ──
        writer.WriteLine(string.Join(delimiter, headers.ConvertAll(h => Escape(h, delimiter))));

        // ── 3. Écrire les lignes ──
        long rowCount = 0;
        var  buffer   = new string[headers.Count];

        foreach (var element in root.EnumerateArray())
        {
            Array.Clear(buffer, 0, buffer.Length);

            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in element.EnumerateObject())
                {
                    if (headerIndex.TryGetValue(prop.Name, out int idx))
                        buffer[idx] = GetValue(prop.Value);
                }
            }

            writer.WriteLine(string.Join(delimiter, Array.ConvertAll(buffer, v => Escape(v ?? "", delimiter))));
            rowCount++;

            if (rowCount % 500_000 == 0)
                Console.WriteLine($"  {rowCount:N0} lignes traitées…");
        }

        writer.Flush();
        sw.Stop();

        Console.WriteLine($"\n✓ {rowCount:N0} lignes en {sw.ElapsedMilliseconds} ms → {outputPath}");
    }

    static string GetValue(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.String  => el.GetString() ?? "",
        JsonValueKind.Number  => el.GetRawText(),
        JsonValueKind.True    => "true",
        JsonValueKind.False   => "false",
        JsonValueKind.Null    => "",
        JsonValueKind.Object  => el.GetRawText(),
        JsonValueKind.Array   => el.GetRawText(),
        _                     => ""
    };

    static string Escape(string value, char delimiter)
    {
        if (value.Contains(delimiter) || value.Contains('"') ||
            value.Contains('\n')      || value.Contains('\r'))
            return '"' + value.Replace("\"", "\"\"") + '"';
        return value;
    }
}
