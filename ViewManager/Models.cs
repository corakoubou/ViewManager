using System.Collections.Generic;
using System;
using System.Text.Json.Serialization;

namespace PureRefBoardWpf;

public sealed class BoardDoc
{
    public string V { get; set; } = "wpf-14-like";
    public UiSettings Ui { get; set; } = new();
    public string ActiveTabId { get; set; } = "";
    public List<BoardTab> Tabs { get; set; } = new();
}

public sealed class UiSettings
{
    public double Scale { get; set; } = 1.0; // UIスケール
}

public sealed class BoardTab
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = "タブ";
    public CameraState Cam { get; set; } = new() { Tx = 200, Ty = 120, Scale = 1 };
    public List<BoardItem> Items { get; set; } = new();
}

public sealed class CameraState
{
    public double Tx { get; set; }
    public double Ty { get; set; }
    public double Scale { get; set; }
}

public sealed class BoardItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Path { get; set; } = ""; // 画像ファイル参照
    public double X { get; set; } = 0;     // world座標（中心）
    public double Y { get; set; } = 0;
    public double W { get; set; } = 420;
    public double H { get; set; } = 300;
    public int Z { get; set; } = 1;

    public bool Lock { get; set; } = false;
    public bool FlipH { get; set; } = false;
    public bool FlipV { get; set; } = false;

    [JsonIgnore] public long CachedBytes { get; set; } = 0;
}
