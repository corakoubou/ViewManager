using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Threading;

namespace PureRefBoardWpf;

public sealed class Storage
{
    private readonly JsonSerializerOptions _opt = new() { WriteIndented = true };

    public string? SaveFolder { get; private set; }
    public string FileName { get; private set; } = "board.json";
    public bool AutoSaveOn { get; set; } = false;
    public DateTime LastSavedAt { get; private set; } = DateTime.MinValue;

    private CancellationTokenSource? _debounceCts;

    public string AppDataRoot =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PureRefBoardWpf");

    public string AppImageCacheDir => Path.Combine(AppDataRoot, "images");

    public void SetFolder(string folder, string fileName)
    {
        SaveFolder = folder;
        FileName = string.IsNullOrWhiteSpace(fileName) ? "board.json" : fileName.Trim();
        if (!FileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase)) FileName += ".json";
        Directory.CreateDirectory(SaveFolder);
        Directory.CreateDirectory(AppDataRoot);
        Directory.CreateDirectory(AppImageCacheDir);
    }

    public string? CurrentSavePath =>
        string.IsNullOrWhiteSpace(SaveFolder) ? null : Path.Combine(SaveFolder!, FileName);

    public void ScheduleAutoSave(Func<Task> saveFunc, int ms = 500)
    {
        if (!AutoSaveOn) return;

        _debounceCts?.Cancel();
        _debounceCts = new CancellationTokenSource();
        var token = _debounceCts.Token;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(ms, token);
                if (token.IsCancellationRequested) return;
                await saveFunc();
            }
            catch { }
        }, token);
    }

    public async Task SaveAsync(BoardDoc doc, string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var json = JsonSerializer.Serialize(doc, _opt);
        await File.WriteAllTextAsync(path, json);
        LastSavedAt = DateTime.Now;
    }

    public async Task<BoardDoc> LoadAsync(string path)
    {
        var json = await File.ReadAllTextAsync(path);
        var doc = JsonSerializer.Deserialize<BoardDoc>(json) ?? throw new Exception("JSONが不正っす");
        if (doc.Tabs.Count == 0) throw new Exception("タブが空っす");
        return doc;
    }

    /// <summary>
    /// クリップボード画像をPNGで保存して、ファイルパスを返すっす。
    /// まず SaveFolder\images を使い、未設定なら LocalAppData を使うっす。
    /// </summary>
    public string SavePastedPng(byte[] pngBytes)
    {
        var baseDir = !string.IsNullOrWhiteSpace(SaveFolder)
            ? Path.Combine(SaveFolder!, "images")
            : AppImageCacheDir;

        Directory.CreateDirectory(baseDir);
        var name = "paste_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".png";
        var path = Path.Combine(baseDir, name);
        File.WriteAllBytes(path, pngBytes);
        return path;
    }
}
