using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;  // NuGet: System.Management
using System.Text;
using System.Threading;
using System.Threading.Tasks;

class Program
{
    static readonly HashSet<string> MdExts = new(StringComparer.OrdinalIgnoreCase)
    {
        ".md", ".markdown", ".mdown", ".mkd", ".mkdn", ".mdtxt", ".mdtext"
    };

    // スピナー表示用（メインスレッドから更新）
    static volatile string CurrentItem = "";
    static volatile int ProcessedCount = 0;
    static volatile int TotalCount = 0;

    static int Main(string[] args)
    {
        CancellationTokenSource? spinnerCts = null;
        Task? spinnerTask = null;

        try
        {
            var (inputs, outputPathOpt, addSourceComments, sortZip) = ParseArgs(args);

            if (inputs.Count == 0)
            {
                PrintUsage();
                Console.WriteLine();
                WriteInfo("このexeに、zip / Markdown / フォルダ をドラッグ＆ドロップすると実行できます。");
                WriteInfo("Enter を押すと終了します。");
                Console.ReadLine();
                return 2;
            }

            // 合計件数を先に数える（表示用）
            TotalCount = CountTotalMarkdownItems(inputs, sortZip);

            // 自動出力先（-oが無い場合）
            var outputPath = string.IsNullOrWhiteSpace(outputPathOpt)
                ? GetAutoOutputPath(inputs)
                : EnsureMdExtension(outputPathOpt!);

            // スピナー開始
            spinnerCts = new CancellationTokenSource();
            spinnerTask = RunSpinner(
                spinnerCts.Token,
                () => CurrentItem,
                () => ProcessedCount,
                () => TotalCount
            );

            var parts = new List<(string Name, string Text)>();
            var skipped = new List<string>();

            foreach (var input in inputs)
            {
                if (Directory.Exists(input))
                {
                    var mdFiles = Directory.EnumerateFiles(input, "*.*", SearchOption.AllDirectories)
                        .Where(f => MdExts.Contains(Path.GetExtension(f)))
                        .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    foreach (var f in mdFiles)
                    {
                        CurrentItem = f;
                        parts.Add((Path.GetFileName(f), NormalizeNewlines(ReadTextSmart(f))));
                        ProcessedCount++;
                    }
                    continue;
                }

                if (!File.Exists(input))
                {
                    StopSpinner(spinnerCts, spinnerTask);
                    WriteError($"見つかりません: {input}");
                    PauseIfExplorerLaunched();
                    return 2;
                }

                var ext = Path.GetExtension(input);

                if (ext.Equals(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var p in ExtractMarkdownFromZip(input, sortZip))
                    {
                        CurrentItem = p.Name; // "zip名:entry名"
                        parts.Add(p);
                        ProcessedCount++;
                    }
                }
                else if (MdExts.Contains(ext))
                {
                    CurrentItem = input;
                    parts.Add((Path.GetFileName(input), NormalizeNewlines(ReadTextSmart(input))));
                    ProcessedCount++;
                }
                else
                {
                    skipped.Add(input);
                }
            }

            if (parts.Count == 0)
            {
                StopSpinner(spinnerCts, spinnerTask);
                WriteError("結合対象のMarkdownが見つかりませんでした。");
                if (skipped.Count > 0)
                    WriteWarn($"対象外: {skipped.Count}件");
                PauseIfExplorerLaunched();
                return 2;
            }

            var merged = Merge(parts, addSourceComments);

            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
            File.WriteAllText(outputPath, merged, new UTF8Encoding(false));

            StopSpinner(spinnerCts, spinnerTask);

            WriteSuccess($"完了: {outputPath}");
            WriteInfo($"結合: {parts.Count}件");
            if (skipped.Count > 0)
                WriteWarn($"対象外: {skipped.Count}件");

            PauseIfExplorerLaunched();
            return 0;
        }
        catch (Exception ex)
        {
            StopSpinner(spinnerCts, spinnerTask);
            WriteError("エラー: " + ex.Message);
            PauseIfExplorerLaunched();
            return 1;
        }
    }

    static void PauseIfExplorerLaunched()
    {
        // Explorer（ダブルクリック / ドラッグ＆ドロップ）のときだけ待つ
        try
        {
            if (Console.IsInputRedirected)
                return;

            var parent = GetParentProcessName();
            if (parent.Equals("explorer.exe", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Enter を押すと終了します。");
                Console.ResetColor();
                Console.ReadLine();
            }
        }
        catch
        {
            // 判定に失敗しても、ここで落とさない
        }
    }

    static string GetParentProcessName()
    {
        using var current = Process.GetCurrentProcess();
        var query = $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {current.Id}";
        using var searcher = new ManagementObjectSearcher(query);
        using var results = searcher.Get();

        foreach (ManagementObject mo in results)
        {
            var ppid = (uint)mo["ParentProcessId"];
            using var parent = Process.GetProcessById((int)ppid);
            return parent.ProcessName + ".exe";
        }
        return "";
    }

    static void PrintUsage()
    {
        Console.WriteLine(
@"使い方:
  MarkdownMerger.exe <入力...> [-o <出力.md>] [--no-source] [--no-sort-zip]
  MarkdownMerger.exe -i <入力...> [-o <出力.md>] [--no-source] [--no-sort-zip]

入力:
  zip / Markdown / フォルダ を混在して複数指定できます
  ドラッグ＆ドロップでもOKです（ドロップしたパスがそのまま入力になります）

出力:
  -o を省略すると、最初の入力と同じフォルダに自動出力します（上書き回避で連番）

オプション:
  -i            入力を複数指定（省略して直接パス指定でも可）
  -o            出力ファイル（省略すると自動）
  --no-source   元ファイル名コメントを入れません
  --no-sort-zip zip内の順序を並べ替えません
");
    }

    static (List<string> Inputs, string? OutputPath, bool AddSourceComments, bool SortZip) ParseArgs(string[] args)
    {
        var inputs = new List<string>();
        string? output = null;
        bool addSourceComments = true;
        bool sortZip = true;

        for (int i = 0; i < args.Length; i++)
        {
            var a = args[i];

            if (a == "-i")
            {
                i++;
                while (i < args.Length && !args[i].StartsWith("-"))
                {
                    inputs.Add(TrimQuotes(args[i]));
                    i++;
                }
                i--;
            }
            else if (a == "-o")
            {
                if (i + 1 >= args.Length) throw new ArgumentException("-o の後に出力パスが必要です。");
                output = TrimQuotes(args[++i]);
            }
            else if (a == "--no-source")
            {
                addSourceComments = false;
            }
            else if (a == "--no-sort-zip")
            {
                sortZip = false;
            }
            else
            {
                if (!a.StartsWith("-"))
                    inputs.Add(TrimQuotes(a));
                else
                    throw new ArgumentException($"不明なオプション: {a}");
            }
        }

        return (inputs, output, addSourceComments, sortZip);
    }

    static string TrimQuotes(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        s = s.Trim();
        if (s.Length >= 2 && s.StartsWith("\"") && s.EndsWith("\""))
            return s.Substring(1, s.Length - 2);
        return s;
    }

    static int CountTotalMarkdownItems(List<string> inputs, bool sortZip)
    {
        int total = 0;

        foreach (var input in inputs)
        {
            if (Directory.Exists(input))
            {
                total += Directory.EnumerateFiles(input, "*.*", SearchOption.AllDirectories)
                    .Count(f => MdExts.Contains(Path.GetExtension(f)));
                continue;
            }

            if (!File.Exists(input))
                continue;

            var ext = Path.GetExtension(input);
            if (ext.Equals(".zip", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    using var fs = File.OpenRead(input);
                    using var zf = new ZipArchive(fs, ZipArchiveMode.Read);
                    total += zf.Entries
                        .Where(e => !string.IsNullOrEmpty(e.FullName) && !e.FullName.EndsWith("/"))
                        .Where(e => !e.FullName.StartsWith("__MACOSX/", StringComparison.OrdinalIgnoreCase))
                        .Count(e => MdExts.Contains(Path.GetExtension(e.FullName)));
                }
                catch
                {
                    // 壊れているzipなど。ここでは数えられないので、後段でエラーにする
                }
            }
            else if (MdExts.Contains(ext))
            {
                total += 1;
            }
        }

        return Math.Max(total, 1);
    }

    static string GetAutoOutputPath(List<string> inputs)
    {
        var first = inputs[0];
        string dir;

        if (Directory.Exists(first))
            dir = Path.GetFullPath(first);
        else if (File.Exists(first))
            dir = Path.GetDirectoryName(Path.GetFullPath(first))!;
        else
            dir = Directory.GetCurrentDirectory();

        string baseName;
        if (inputs.Count == 1 && File.Exists(first))
            baseName = Path.GetFileNameWithoutExtension(first) + "_merged";
        else
            baseName = "merged";

        var candidate = Path.Combine(dir, baseName + ".md");
        return MakeUniquePath(candidate);
    }

    static string EnsureMdExtension(string path)
    {
        if (path.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
            return path;
        return path + ".md";
    }

    static string MakeUniquePath(string path)
    {
        if (!File.Exists(path))
            return path;

        var dir = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);

        for (int i = 1; i < 10000; i++)
        {
            var p = Path.Combine(dir, $"{name} ({i}){ext}");
            if (!File.Exists(p))
                return p;
        }

        throw new IOException("出力先の重複回避に失敗しました。");
    }

    static IEnumerable<(string Name, string Text)> ExtractMarkdownFromZip(string zipPath, bool sortZip)
    {
        using var fs = File.OpenRead(zipPath);
        using var zf = new ZipArchive(fs, ZipArchiveMode.Read);

        var entries = zf.Entries
            .Where(e => !string.IsNullOrEmpty(e.FullName) && !e.FullName.EndsWith("/"))
            .Where(e => !e.FullName.StartsWith("__MACOSX/", StringComparison.OrdinalIgnoreCase))
            .Where(e => MdExts.Contains(Path.GetExtension(e.FullName)));

        entries = sortZip
            ? entries.OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
            : entries;

        foreach (var e in entries)
        {
            var label = $"{Path.GetFileName(zipPath)}:{e.FullName}";
            using var s = e.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);

            var text = DecodeSmart(ms.ToArray());
            text = NormalizeNewlines(text);

            yield return (label, text);
        }
    }

    static string Merge(List<(string Name, string Text)> parts, bool addSourceComments)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < parts.Count; i++)
        {
            var (name, text) = parts[i];

            if (i > 0)
            {
                sb.AppendLine().AppendLine("---").AppendLine();
            }

            if (addSourceComments)
            {
                sb.Append("<!-- source: ").Append(name).AppendLine(" -->");
                sb.AppendLine();
            }

            if (!text.EndsWith("\n"))
                text += "\n";

            sb.Append(text);
        }
        return sb.ToString();
    }

    static string ReadTextSmart(string path)
    {
        var bytes = File.ReadAllBytes(path);
        return DecodeSmart(bytes);
    }

    static string DecodeSmart(byte[] bytes)
    {
        var utf8Strict = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
        try
        {
            return utf8Strict.GetString(bytes);
        }
        catch
        {
            var sjis = Encoding.GetEncoding(932);
            return sjis.GetString(bytes);
        }
    }

    static string NormalizeNewlines(string text)
    {
        return text.Replace("\r\n", "\n").Replace("\r", "\n");
    }

    static Task RunSpinner(CancellationToken token, Func<string> current, Func<int> processed, Func<int> total)
    {
        return Task.Run(async () =>
        {
            var frames = new[] { '|', '/', '-', '\\' };
            int idx = 0;

            while (!token.IsCancellationRequested)
            {
                var c = current() ?? "";
                var p = processed();
                var t = total();

                var msg = $"{frames[idx++ % frames.Length]} {p}/{t} 処理中: {c}";
                if (msg.Length > 160)
                    msg = msg.Substring(0, 157) + "...";

                int width;
                try { width = Math.Max(20, Console.WindowWidth - 1); }
                catch { width = 120; }

                if (msg.Length < width) msg = msg.PadRight(width);

                Console.Write("\r" + msg);
                await Task.Delay(100, token).ContinueWith(_ => { });
            }
        }, token);
    }

    static void StopSpinner(CancellationTokenSource? cts, Task? spinnerTask)
    {
        if (cts == null) return;

        try
        {
            cts.Cancel();
            spinnerTask?.Wait(300);
        }
        catch
        {
        }
        finally
        {
            int width;
            try { width = Math.Max(20, Console.WindowWidth - 1); }
            catch { width = 120; }

            Console.Write("\r" + new string(' ', width) + "\r");
        }
    }

    // 色付き出力（見た目だけ。ログが必要なら後でファイル出力も足せます）
    static void WriteSuccess(string msg)
    {
        var old = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(msg);
        Console.ForegroundColor = old;
    }

    static void WriteError(string msg)
    {
        var old = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(msg);
        Console.ForegroundColor = old;
    }

    static void WriteWarn(string msg)
    {
        var old = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(msg);
        Console.ForegroundColor = old;
    }

    static void WriteInfo(string msg)
    {
        Console.WriteLine(msg);
    }
}
