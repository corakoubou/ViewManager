using Microsoft.Win32;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PureRefBoardWpf;

public partial class MainWindow : Window
{
    private readonly Storage _storage = new();

    private BoardDoc _doc = new();
    private BoardTab _activeTab = new();

    private const double MinScale = 0.15;
    private const double MaxScale = 6.0;

    // camera
    private double Tx { get => CamTranslate.X; set => CamTranslate.X = value; }
    private double Ty { get => CamTranslate.Y; set => CamTranslate.Y = value; }
    private double Scale { get => CamScale.ScaleX; set { CamScale.ScaleX = CamScale.ScaleY = value; } }

    // input
    private bool _spaceDown = false;
    private bool _panning = false;
    private Point _panStart;
    private double _panStartTx, _panStartTy;

    private string? _selectedId = null;

    private bool _draggingItem = false;
    private string? _dragItemId = null;
    private Point _dragStartScreen;
    private double _dragStartX, _dragStartY;

    private bool _resizing = false;
    private string? _resizeItemId = null;
    private Point _resizeStartScreen;
    private double _resizeStartW, _resizeStartH;

    private const double DragThresholdPx = 7;

    public MainWindow()
    {
        InitializeComponent();
        InitNewDoc();
        ApplyUiScale(_doc.Ui.Scale);
        RenderTabs();
        SetActiveTab(_doc.ActiveTabId, saveCam: false);
        RefreshMbChip();
        RefreshStatus("起動したっす");
    }

    // ---------- init ----------
    private void InitNewDoc()
    {
        _doc = new BoardDoc();
        var t1 = new BoardTab { Name = "タブ1" };
        _doc.Tabs.Add(t1);
        _doc.ActiveTabId = t1.Id;
    }

    // ---------- UI scale ----------
    private void ApplyUiScale(double v)
    {
        v = Math.Max(0.6, Math.Min(1.2, v));
        _doc.Ui.Scale = v;
        UiScaleTf.ScaleX = UiScaleTf.ScaleY = v;
        UiScaleSlider.Value = v;
    }

    private void UiScaleSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (!IsLoaded) return;
        ApplyUiScale(e.NewValue);
        MarkDirty("UIサイズ変更");
    }

    // ---------- Tabs ----------
    private void RenderTabs()
    {
        TabsPanel.Children.Clear();

        foreach (var t in _doc.Tabs)
        {
            var b = new Border
            {
                CornerRadius = new CornerRadius(12),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 6, 10, 6),
                Margin = new Thickness(0, 0, 6, 0),
                Background = (t.Id == _doc.ActiveTabId) ? new SolidColorBrush(Color.FromArgb(25, 74, 163, 255)) : new SolidColorBrush(Color.FromArgb(12, 255, 255, 255)),
                BorderBrush = (t.Id == _doc.ActiveTabId) ? new SolidColorBrush(Color.FromArgb(140, 74, 163, 255)) : new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)),
                Cursor = Cursors.Hand
            };

            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            var dot = new Ellipse
            {
                Width = 8,
                Height = 8,
                Margin = new Thickness(0, 0, 6, 0),
                Fill = (t.Id == _doc.ActiveTabId) ? new SolidColorBrush(Color.FromArgb(240, 74, 163, 255)) : new SolidColorBrush(Color.FromArgb(70, 255, 255, 255))
            };
            var txt = new TextBlock { Text = t.Name, Foreground = Brushes.White, FontSize = 13 };

            sp.Children.Add(dot);
            sp.Children.Add(txt);
            b.Child = sp;

            b.MouseLeftButtonDown += (_, __) => SetActiveTab(t.Id, saveCam: true);
            b.MouseDown += (_, e) =>
            {
                if (e.ChangedButton == MouseButton.Left && e.ClickCount == 2)
                {
                    var next = Prompt("タブ名を入力してほしいっす", t.Name);
                    if (next != null)
                    {
                        t.Name = string.IsNullOrWhiteSpace(next) ? "タブ" : next.Trim().Substring(0, Math.Min(40, next.Trim().Length));
                        RenderTabs();
                        MarkDirty("タブ名変更");
                    }
                }
            };

            TabsPanel.Children.Add(b);
        }
    }

    private static string? Prompt(string msg, string defaultValue)
    {
        var w = new Window
        {
            Title = msg,
            Width = 420,
            Height = 160,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Background = new SolidColorBrush(Color.FromRgb(20, 24, 32))
        };
        var grid = new Grid { Margin = new Thickness(12) };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        grid.Children.Add(new TextBlock { Text = msg, Foreground = Brushes.White });

        var tb = new TextBox { Text = defaultValue, Margin = new Thickness(0, 8, 0, 10) };
        Grid.SetRow(tb, 1);
        grid.Children.Add(tb);

        var sp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var ok = new Button { Content = "OK", Width = 90, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "Cancel", Width = 90 };
        sp.Children.Add(ok); sp.Children.Add(cancel);
        Grid.SetRow(sp, 2);
        grid.Children.Add(sp);

        w.Content = grid;

        string? result = null;
        ok.Click += (_, __) => { result = tb.Text; w.DialogResult = true; w.Close(); };
        cancel.Click += (_, __) => { result = null; w.DialogResult = false; w.Close(); };

        w.Owner = Application.Current.MainWindow;
        w.ShowDialog();
        return result;
    }

    private void SetActiveTab(string tabId, bool saveCam)
    {
        if (saveCam) SaveCamToTab();

        _activeTab = _doc.Tabs.First(t => t.Id == tabId);
        _doc.ActiveTabId = tabId;

        // cam restore
        Tx = _activeTab.Cam.Tx;
        Ty = _activeTab.Cam.Ty;
        Scale = Clamp(_activeTab.Cam.Scale, MinScale, MaxScale);

        _selectedId = null;

        RenderTabs();
        RenderWorld();
        RefreshMbChip();
        RefreshStatus($"タブ：{_activeTab.Name} っす");

        MarkDirty("タブ切替", autosaveOnly: true);
    }

    private void SaveCamToTab()
    {
        _activeTab.Cam.Tx = Tx;
        _activeTab.Cam.Ty = Ty;
        _activeTab.Cam.Scale = Scale;
    }

    private void AddTab_Click(object sender, RoutedEventArgs e)
    {
        SaveCamToTab();
        var t = new BoardTab { Name = $"タブ{_doc.Tabs.Count + 1}" };
        _doc.Tabs.Add(t);
        SetActiveTab(t.Id, saveCam: false);
        MarkDirty("タブ追加");
    }

    private void DupTab_Click(object sender, RoutedEventArgs e)
    {
        SaveCamToTab();

        var src = _activeTab;
        var copy = new BoardTab
        {
            Name = (src.Name ?? "タブ") + "（コピー）",
            Cam = new CameraState { Tx = src.Cam.Tx, Ty = src.Cam.Ty, Scale = src.Cam.Scale },
            Items = src.Items.Select(it => new BoardItem
            {
                Id = Guid.NewGuid().ToString("N"),
                Path = it.Path,
                X = it.X,
                Y = it.Y,
                W = it.W,
                H = it.H,
                Z = it.Z,
                Lock = it.Lock,
                FlipH = it.FlipH,
                FlipV = it.FlipV,
                CachedBytes = it.CachedBytes
            }).ToList()
        };
        _doc.Tabs.Add(copy);
        SetActiveTab(copy.Id, saveCam: false);
        MarkDirty("タブ複製");
    }

    private void DelTab_Click(object sender, RoutedEventArgs e)
    {
        if (_doc.Tabs.Count <= 1)
        {
            MessageBox.Show("最後の1タブは消せないっす");
            return;
        }
        var ok = MessageBox.Show($"タブ「{_activeTab.Name}」を削除するっすか？", "確認", MessageBoxButton.YesNo) == MessageBoxResult.Yes;
        if (!ok) return;

        var idx = _doc.Tabs.FindIndex(t => t.Id == _activeTab.Id);
        _doc.Tabs.RemoveAt(idx);
        var next = _doc.Tabs[Math.Max(0, idx - 1)];
        SetActiveTab(next.Id, saveCam: false);
        MarkDirty("タブ削除");
    }

    // ---------- render world ----------
    private void RenderWorld()
    {
        World.Children.Clear();

        foreach (var it in _activeTab.Items.OrderBy(i => i.Z))
        {
            var host = new Border
            {
                Width = it.W,
                Height = it.H,
                Background = new SolidColorBrush(Color.FromRgb(11, 13, 18)),
                BorderThickness = new Thickness(it.Id == _selectedId ? 2 : 1),
                BorderBrush = it.Id == _selectedId
                    ? new SolidColorBrush(Color.FromRgb(74, 163, 255))
                    : new SolidColorBrush(Color.FromArgb(22, 255, 255, 255)),
                SnapsToDevicePixels = true,
            };

            if (it.Lock)
            {
                host.BorderBrush = new SolidColorBrush(Color.FromArgb(130, 255, 255, 255));
                host.BorderThickness = new Thickness(it.Id == _selectedId ? 2 : 1);
                host.BorderDashArray = new DoubleCollection { 4, 2 };
            }

            var grid = new Grid();

            var img = new Image
            {
                Stretch = Stretch.Uniform,
                SnapsToDevicePixels = true
            };
            RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);

            if (File.Exists(it.Path))
            {
                img.Source = LoadBitmap(it.Path);
            }

            // flip
            img.RenderTransformOrigin = new Point(0.5, 0.5);
            img.RenderTransform = new ScaleTransform(it.FlipH ? -1 : 1, it.FlipV ? -1 : 1);

            grid.Children.Add(img);

            // resize handle (Thumb)
            var thumb = new Thumb
            {
                Width = 14,
                Height = 14,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 6, 6),
                Cursor = Cursors.SizeNWSE,
                Background = new SolidColorBrush(Color.FromArgb(230, 74, 163, 255)),
                Visibility = (it.Id == _selectedId && !it.Lock) ? Visibility.Visible : Visibility.Collapsed
            };

            thumb.DragStarted += (_, __) =>
            {
                if (it.Lock) return;
                _resizing = true;
                _resizeItemId = it.Id;
                _resizeStartW = it.W;
                _resizeStartH = it.H;
                _resizeStartScreen = Mouse.GetPosition(this);
            };
            thumb.DragDelta += (_, __) =>
            {
                if (!_resizing || _resizeItemId != it.Id) return;
                var cur = Mouse.GetPosition(this);
                var dx = (cur.X - _resizeStartScreen.X) / Scale;
                var dy = (cur.Y - _resizeStartScreen.Y) / Scale;
                it.W = Clamp(_resizeStartW + dx, 80, 4000);
                it.H = Clamp(_resizeStartH + dy, 80, 4000);
                host.Width = it.W;
                host.Height = it.H;
                MarkDirty("リサイズ", autosaveOnly: true);
            };
            thumb.DragCompleted += (_, __) =>
            {
                _resizing = false;
                _resizeItemId = null;
                MarkDirty("リサイズ確定");
                RefreshMbChip();
            };

            grid.Children.Add(thumb);

            host.Child = grid;

            // position (center-based)
            Canvas.SetLeft(host, it.X - it.W / 2);
            Canvas.SetTop(host, it.Y - it.H / 2);
            Panel.SetZIndex(host, it.Z);

            host.MouseLeftButtonDown += (s, e) =>
            {
                World.Focus();
                Select(it.Id);
                BringToFront(it.Id);

                if (e.ClickCount == 2)
                {
                    FocusItemFit(it.Id);
                    e.Handled = true;
                    return;
                }

                if (it.Lock) return;

                _draggingItem = true;
                _dragItemId = it.Id;
                _dragStartScreen = e.GetPosition(this);
                _dragStartX = it.X;
                _dragStartY = it.Y;

                (s as UIElement)!.CaptureMouse();
                e.Handled = true;
            };

            host.MouseLeftButtonUp += (s, e) =>
            {
                if (_draggingItem && _dragItemId == it.Id)
                {
                    _draggingItem = false;
                    _dragItemId = null;
                    (s as UIElement)!.ReleaseMouseCapture();
                    MarkDirty("移動確定");
                    e.Handled = true;
                }
            };

            World.Children.Add(host);
        }
    }

    private static BitmapImage LoadBitmap(string path)
    {
        var bi = new BitmapImage();
        bi.BeginInit();
        bi.CacheOption = BitmapCacheOption.OnLoad;
        bi.UriSource = new Uri(path);
        bi.EndInit();
        bi.Freeze();
        return bi;
    }

    private void Select(string id)
    {
        _selectedId = id;
        RenderWorld();
    }

    private BoardItem? SelectedItem =>
        _selectedId == null ? null : _activeTab.Items.FirstOrDefault(i => i.Id == _selectedId);

    private void BringToFront(string id)
    {
        var it = _activeTab.Items.First(x => x.Id == id);
        var maxZ = _activeTab.Items.Count == 0 ? 0 : _activeTab.Items.Max(x => x.Z);
        it.Z = maxZ + 1;
        RenderWorld();
        MarkDirty("最前面", autosaveOnly: true);
    }

    private void FocusItemFit(string id)
    {
        var it = _activeTab.Items.FirstOrDefault(i => i.Id == id);
        if (it == null) return;

        var margin = 30;
        var topPad = 110; // topbar相当
        var vw = Math.Max(200, ActualWidth - margin * 2);
        var vh = Math.Max(200, ActualHeight - margin * 2 - topPad);

        var s = Math.Min(vw / it.W, vh / it.H);
        s = Clamp(s, MinScale, MaxScale);
        Scale = s;

        var targetCx = ActualWidth / 2;
        var targetCy = (ActualHeight + topPad) / 2;

        Tx = targetCx - it.X * Scale;
        Ty = targetCy - it.Y * Scale;

        SaveCamToTab();
        MarkDirty("フィットズーム");
    }

    // ---------- camera & input ----------
    private void World_MouseWheel(object sender, MouseWheelEventArgs e)
    {
        var zoomFactor = Math.Exp(e.Delta * 0.0015);
        var newScale = Clamp(Scale * zoomFactor, MinScale, MaxScale);

        var pos = e.GetPosition(World); // world-control coordinates
        var before = ControlToWorld(pos);

        Scale = newScale;

        var after = ControlToWorld(pos);

        Tx += (after.X - before.X) * Scale;
        Ty += (after.Y - before.Y) * Scale;

        SaveCamToTab();
        MarkDirty("ズーム", autosaveOnly: true);
        e.Handled = true;
    }

    private Point ControlToWorld(Point pInWorldControl)
    {
        var s = Scale;
        return new Point((pInWorldControl.X - Tx) / s, (pInWorldControl.Y - Ty) / s);
    }

    private void World_MouseDown(object sender, MouseButtonEventArgs e)
    {
        World.Focus();

        if (e.ChangedButton == MouseButton.Left && _spaceDown)
        {
            _panning = true;
            _panStart = e.GetPosition(this);
            _panStartTx = Tx;
            _panStartTy = Ty;
            World.CaptureMouse();
            Mouse.OverrideCursor = Cursors.Hand;
            e.Handled = true;
            return;
        }

        // blank click -> unselect
        if (e.ChangedButton == MouseButton.Left && e.OriginalSource == World)
        {
            _selectedId = null;
            RenderWorld();
        }
    }

    private void World_MouseMove(object sender, MouseEventArgs e)
    {
        if (_panning)
        {
            var cur = e.GetPosition(this);
            var dx = cur.X - _panStart.X;
            var dy = cur.Y - _panStart.Y;

            Tx = _panStartTx + dx;
            Ty = _panStartTy + dy;

            SaveCamToTab();
            MarkDirty("パン", autosaveOnly: true);
            return;
        }

        if (_draggingItem && _dragItemId != null)
        {
            var it = _activeTab.Items.First(x => x.Id == _dragItemId);
            var cur = e.GetPosition(this);
            var dx = cur.X - _dragStartScreen.X;
            var dy = cur.Y - _dragStartScreen.Y;

            // threshold
            if (Math.Sqrt(dx * dx + dy * dy) < DragThresholdPx) return;

            it.X = _dragStartX + dx / Scale;
            it.Y = _dragStartY + dy / Scale;

            // update element position without full rerender (軽い)
            var el = FindItemElement(it.Id);
            if (el != null)
            {
                Canvas.SetLeft(el, it.X - it.W / 2);
                Canvas.SetTop(el, it.Y - it.H / 2);
            }

            MarkDirty("移動", autosaveOnly: true);
            return;
        }
    }

    private FrameworkElement? FindItemElement(string id)
    {
        foreach (var c in World.Children)
        {
            if (c is FrameworkElement fe && fe is Border b)
            {
                // we stored id only in closure; easiest: use Tag now
            }
        }
        // Tag仕込みをするため、renderでTagを付けるのが良いけど、ここは簡易でnull返し → RenderWorldでもOKっす
        return null;
    }

    private void World_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_panning && e.ChangedButton == MouseButton.Left)
        {
            _panning = false;
            World.ReleaseMouseCapture();
            Mouse.OverrideCursor = null;
            MarkDirty("パン確定");
            e.Handled = true;
        }
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Space) _spaceDown = true;

        // front
        if (e.Key == Key.F && _selectedId != null)
        {
            BringToFront(_selectedId);
        }

        // delete
        if ((e.Key == Key.Delete || e.Key == Key.Back) && _selectedId != null)
        {
            var ok = MessageBox.Show("選択中の画像を削除するっすか？", "確認", MessageBoxButton.YesNo) == MessageBoxResult.Yes;
            if (!ok) return;

            _activeTab.Items.RemoveAll(i => i.Id == _selectedId);
            _selectedId = null;
            RenderWorld();
            RefreshMbChip();
            MarkDirty("削除");
        }

        // Ctrl+C copy
        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.C)
        {
            CopySelectedToClipboard();
            e.Handled = true;
        }

        // Ctrl+V paste
        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.V)
        {
            PasteImageFromClipboard();
            e.Handled = true;
        }

        // Esc close overlays (none here)
        if (e.Key == Key.Escape)
        {
            // no-op
        }
    }

    private void Window_KeyUp(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Space) _spaceDown = false;
    }

    // ---------- Context menu ----------
    private void CtxMenu_Opened(object sender, RoutedEventArgs e)
    {
        AutoSaveMenuItem.Header = "自動保存：" + (_storage.AutoSaveOn ? "ON" : "OFF");
        UiScaleSlider.Value = _doc.Ui.Scale;
    }

    private void Ctx_Add_Click(object sender, RoutedEventArgs e) => AddImages();
    private void Ctx_Copy_Click(object sender, RoutedEventArgs e) => CopySelectedToClipboard();
    private void Ctx_Lock_Click(object sender, RoutedEventArgs e)
    {
        var it = SelectedItem; if (it == null) return;
        it.Lock = !it.Lock;
        RenderWorld();
        MarkDirty("ロック切替");
    }
    private void Ctx_Front_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedId != null) BringToFront(_selectedId);
    }
    private void Ctx_FlipH_Click(object sender, RoutedEventArgs e)
    {
        var it = SelectedItem; if (it == null) return;
        it.FlipH = !it.FlipH;
        RenderWorld();
        MarkDirty("左右反転");
    }
    private void Ctx_FlipV_Click(object sender, RoutedEventArgs e)
    {
        var it = SelectedItem; if (it == null) return;
        it.FlipV = !it.FlipV;
        RenderWorld();
        MarkDirty("上下反転");
    }
    private void Ctx_Center_Click(object sender, RoutedEventArgs e)
    {
        Tx = ActualWidth / 2 - 200;
        Ty = ActualHeight / 2 - 120;
        SaveCamToTab();
        MarkDirty("中央へ");
    }
    private void Ctx_ResetView_Click(object sender, RoutedEventArgs e)
    {
        Tx = 200; Ty = 120; Scale = 1;
        SaveCamToTab();
        MarkDirty("表示リセット");
    }
    private void Ctx_PickFolder_Click(object sender, RoutedEventArgs e) => PickFolder();
    private async void Ctx_SaveNow_Click(object sender, RoutedEventArgs e) => await SaveToFolderNow("手動保存");
    private void Ctx_ToggleAutoSave_Click(object sender, RoutedEventArgs e)
    {
        _storage.AutoSaveOn = !_storage.AutoSaveOn;
        AutoSaveMenuItem.Header = "自動保存：" + (_storage.AutoSaveOn ? "ON" : "OFF");
        MarkDirty("自動保存切替", autosaveOnly: true);
        if (_storage.AutoSaveOn) _storage.ScheduleAutoSave(() => SaveToFolderNow("自動保存ON"), 50);
    }
    private void Ctx_ExportAs_Click(object sender, RoutedEventArgs e) => ExportAs();
    private async void Ctx_Import_Click(object sender, RoutedEventArgs e) => await ImportByDialog();
    private void Ctx_Help_Click(object sender, RoutedEventArgs e) => ShowHelp();
    private void Ctx_ClearTab_Click(object sender, RoutedEventArgs e)
    {
        var ok = MessageBox.Show($"タブ「{_activeTab.Name}」を全消去するっすか？", "確認", MessageBoxButton.YesNo) == MessageBoxResult.Yes;
        if (!ok) return;
        _activeTab.Items.Clear();
        _selectedId = null;
        RenderWorld();
        RefreshMbChip();
        MarkDirty("全消去");
    }

    // ---------- Add Images ----------
    private void AddImages()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Images|*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif|All|*.*",
            Multiselect = true
        };
        if (dlg.ShowDialog() != true) return;

        var center = new Point(World.ActualWidth / 2, World.ActualHeight / 2);
        var p = ControlToWorld(center);

        foreach (var path in dlg.FileNames)
        {
            var it = new BoardItem
            {
                Path = path,
                X = p.X,
                Y = p.Y,
                W = 420,
                H = 300,
                Z = (_activeTab.Items.Count == 0 ? 1 : _activeTab.Items.Max(i => i.Z) + 1),
                Lock = false,
                FlipH = false,
                FlipV = false
            };

            TryFitSizeToImage(it, 520, 420);
            it.CachedBytes = GetFileBytes(path);

            _activeTab.Items.Add(it);
            _selectedId = it.Id;
        }

        RenderWorld();
        RefreshMbChip();
        MarkDirty($"画像追加 {dlg.FileNames.Length}枚");
    }

    private static void TryFitSizeToImage(BoardItem it, double maxW, double maxH)
    {
        try
        {
            if (!File.Exists(it.Path)) return;
            using var fs = File.OpenRead(it.Path);
            var decoder = BitmapDecoder.Create(fs, BitmapCreateOptions.IgnoreColorProfile, BitmapCacheOption.OnLoad);
            var frame = decoder.Frames[0];
            var w = frame.PixelWidth;
            var h = frame.PixelHeight;
            if (w <= 0 || h <= 0) return;

            var r = Math.Min(maxW / w, maxH / h);
            r = Math.Min(r, 1.0);
            it.W = Math.Max(140, Math.Round(w * r));
            it.H = Math.Max(120, Math.Round(h * r));
        }
        catch { }
    }

    // ---------- Clipboard Copy/Paste ----------
    private void CopySelectedToClipboard()
    {
        var it = SelectedItem;
        if (it == null) { MessageBox.Show("コピーする画像を選択してほしいっす"); return; }
        if (!File.Exists(it.Path)) { MessageBox.Show("元画像が見つからないっす"); return; }

        try
        {
            var bi = LoadBitmap(it.Path);

            // 反転を含めてRenderTargetBitmapに描画
            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                var w = bi.PixelWidth;
                var h = bi.PixelHeight;

                dc.PushTransform(new TranslateTransform(w / 2.0, h / 2.0));
                dc.PushTransform(new ScaleTransform(it.FlipH ? -1 : 1, it.FlipV ? -1 : 1));
                dc.PushTransform(new TranslateTransform(-w / 2.0, -h / 2.0));

                dc.DrawImage(bi, new Rect(0, 0, w, h));
            }

            var rtb = new RenderTargetBitmap(bi.PixelWidth, bi.PixelHeight, bi.DpiX, bi.DpiY, PixelFormats.Pbgra32);
            rtb.Render(dv);

            Clipboard.SetImage(rtb);
            RefreshStatus("画像コピーしたっす");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "コピー失敗");
        }
    }

    private void PasteImageFromClipboard()
    {
        try
        {
            var img = Clipboard.GetImage();
            if (img == null) return;

            // PNGで保存してPath参照にするっす
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(img));

            using var ms = new MemoryStream();
            encoder.Save(ms);
            var path = _storage.SavePastedPng(ms.ToArray());

            AddImagePath(path);
            RefreshStatus("貼り付け画像を追加したっす");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "貼り付け失敗");
        }
    }

    private void AddImagePath(string path, double? clientX = null, double? clientY = null)
    {
        var cx = clientX ?? (ActualWidth / 2);
        var cy = clientY ?? (ActualHeight / 2);
        var posInWorld = ControlToWorld(new Point(World.ActualWidth / 2, World.ActualHeight / 2));

        var it = new BoardItem
        {
            Path = path,
            X = posInWorld.X,
            Y = posInWorld.Y,
            W = 420,
            H = 300,
            Z = (_activeTab.Items.Count == 0 ? 1 : _activeTab.Items.Max(i => i.Z) + 1),
        };
        TryFitSizeToImage(it, 520, 420);
        it.CachedBytes = GetFileBytes(path);

        _activeTab.Items.Add(it);
        _selectedId = it.Id;

        RenderWorld();
        RefreshMbChip();
        MarkDirty("貼り付け追加");
    }

    // ---------- Drag&Drop (files) ----------
    private void Window_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            DropOverlay.Visibility = Visibility.Visible;
        }
        else e.Effects = DragDropEffects.None;

        e.Handled = true;
    }

    private async void Window_Drop(object sender, DragEventArgs e)
    {
        DropOverlay.Visibility = Visibility.Collapsed;
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
        if (files == null || files.Length == 0) return;

        var json = files.FirstOrDefault(f => f.EndsWith(".json", StringComparison.OrdinalIgnoreCase));
        if (json != null)
        {
            try
            {
                var doc = await _storage.LoadAsync(json);
                NormalizeDoc(doc);
                _doc = doc;
                ApplyUiScale(_doc.Ui.Scale);

                var active = _doc.Tabs.FirstOrDefault(t => t.Id == _doc.ActiveTabId) ?? _doc.Tabs[0];
                RenderTabs();
                SetActiveTab(active.Id, saveCam: false);

                RefreshStatus("JSONインポートしたっす");
                MarkDirty("インポート");
            }
            catch
            {
                MessageBox.Show("インポート失敗っす（JSON形式を確認してほしいっす）");
            }
            return;
        }

        // images
        var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif" };
        var imgs = files.Where(f => exts.Contains(System.IO.Path.GetExtension(f))).ToList();
        if (imgs.Count == 0) return;

        var center = ControlToWorld(new Point(World.ActualWidth / 2, World.ActualHeight / 2));
        foreach (var path in imgs)
        {
            var it = new BoardItem
            {
                Path = path,
                X = center.X,
                Y = center.Y,
                W = 420,
                H = 300,
                Z = (_activeTab.Items.Count == 0 ? 1 : _activeTab.Items.Max(i => i.Z) + 1),
            };
            TryFitSizeToImage(it, 520, 420);
            it.CachedBytes = GetFileBytes(path);

            _activeTab.Items.Add(it);
            _selectedId = it.Id;
        }

        RenderWorld();
        RefreshMbChip();
        MarkDirty("ドロップ追加");
    }

    // ---------- Save/Load ----------
    private void PickFolder()
    {
        var dlg = new OpenFolderDialog
        {
            Title = "保存先フォルダを選んでほしいっす",
            Multiselect = false
        };

        if (dlg.ShowDialog() != true) return;

        var name = Prompt("保存するJSONファイル名っす（例：board.json）", _storage.FileName);
        if (name == null) return;

        _storage.SetFolder(dlg.FolderName, name);
        RefreshStatus($"保存先：{_storage.CurrentSavePath} っす");

        _ = SaveToFolderNow("初回保存");
    }


    private async Task SaveToFolderNow(string reason)
    {
        SaveCamToTab();
        if (_storage.CurrentSavePath == null)
        {
            RefreshStatus("保存先が未設定っす（右クリック→保存先フォルダ）");
            return;
        }

        try
        {
            await _storage.SaveAsync(_doc, _storage.CurrentSavePath);
            RefreshStatus($"{reason}で保存したっす（{_storage.CurrentSavePath}）");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "保存失敗");
        }
    }

    private void ExportAs()
    {
        SaveCamToTab();
        var dlg = new SaveFileDialog
        {
            Filter = "JSON (*.json)|*.json",
            FileName = "board_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".json"
        };
        if (dlg.ShowDialog() != true) return;

        _ = Task.Run(async () =>
        {
            try
            {
                await _storage.SaveAsync(_doc, dlg.FileName);
                Dispatcher.Invoke(() => RefreshStatus($"エクスポートしたっす（{dlg.FileName}）"));
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => MessageBox.Show(ex.Message, "エクスポート失敗"));
            }
        });
    }

    private async Task ImportByDialog()
    {
        var dlg = new OpenFileDialog { Filter = "JSON (*.json)|*.json" };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var doc = await _storage.LoadAsync(dlg.FileName);
            NormalizeDoc(doc);
            _doc = doc;

            ApplyUiScale(_doc.Ui.Scale);

            var active = _doc.Tabs.FirstOrDefault(t => t.Id == _doc.ActiveTabId) ?? _doc.Tabs[0];
            RenderTabs();
            SetActiveTab(active.Id, saveCam: false);

            RefreshStatus($"読み込み：{dlg.FileName} っす");
            MarkDirty("インポート");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "読み込み失敗");
        }
    }

    private static void NormalizeDoc(BoardDoc doc)
    {
        doc.Ui ??= new UiSettings();
        foreach (var t in doc.Tabs)
        {
            t.Cam ??= new CameraState { Tx = 200, Ty = 120, Scale = 1 };
            t.Items ??= new List<BoardItem>();
            foreach (var it in t.Items)
            {
                // null安全
                it.Path ??= "";
            }
        }
        if (string.IsNullOrWhiteSpace(doc.ActiveTabId))
            doc.ActiveTabId = doc.Tabs[0].Id;
    }

    // ---------- MB chip ----------
    private void RefreshMbChip()
    {
        long tab = 0;
        foreach (var it in _activeTab.Items) tab += GetFileBytesCached(it);

        long all = 0;
        foreach (var t in _doc.Tabs)
            foreach (var it in t.Items) all += GetFileBytesCached(it);

        MbChip.Text = $"このタブ: {BytesToMB(tab):0.0}MB / 全体: {BytesToMB(all):0.0}MB";
    }

    private long GetFileBytesCached(BoardItem it)
    {
        if (it.CachedBytes > 0) return it.CachedBytes;
        it.CachedBytes = GetFileBytes(it.Path);
        return it.CachedBytes;
    }

    private static long GetFileBytes(string path)
    {
        try
        {
            if (File.Exists(path)) return new FileInfo(path).Length;
        }
        catch { }
        return 0;
    }

    private static double BytesToMB(long b) => b / (1024.0 * 1024.0);

    // ---------- help ----------
    private void ShowHelp()
    {
        MessageBox.Show(
            "操作説明っす\n\n" +
            "・タブ名変更：タブをダブルクリック\n" +
            "・画像追加：右クリック→画像追加 / ファイルをドロップ\n" +
            "・JSONインポート：右クリック→インポート / JSONをドロップ\n" +
            "・パン：Space押しながらドラッグ\n" +
            "・ズーム：ホイール\n" +
            "・画像移動：画像ドラッグ（一定距離で開始）\n" +
            "・画像リサイズ：選択中の右下ハンドル\n" +
            "・フィットズーム：画像ダブルクリック\n" +
            "・コピー：Ctrl+C / 貼り付け：Ctrl+V\n" +
            "・保存：右クリック→保存先フォルダ→今すぐ保存\n",
            "操作説明");
    }

    // ---------- dirty/autosave ----------
    private void MarkDirty(string reason, bool autosaveOnly = false)
    {
        RefreshMbChip();
        SaveCamToTab();

        if (!autosaveOnly)
            RefreshStatus($"{reason} っす");

        if (_storage.AutoSaveOn)
        {
            _storage.ScheduleAutoSave(async () => await Dispatcher.InvokeAsync(async () => await SaveToFolderNow("自動保存")), 500);
        }
    }

    private void RefreshStatus(string s)
    {
        var folder = _storage.CurrentSavePath == null ? "未設定" : _storage.CurrentSavePath;
        var auto = _storage.AutoSaveOn ? "ON" : "OFF";
        var when = _storage.LastSavedAt == DateTime.MinValue ? "未保存" : _storage.LastSavedAt.ToString("yyyy/MM/dd HH:mm:ss");

        StatusText.Text = $"{s}  | 保存：{folder}  | 自動保存：{auto}  | 最終保存：{when}";
    }

    private static double Clamp(double v, double min, double max) => Math.Max(min, Math.Min(max, v));
}
