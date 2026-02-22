using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Platform.Storage;

namespace AudioSR.NET;

/// <summary>
/// MainWindow のコードビハインド (IWindowService 実装 + ドラッグ&ドロップ)
/// </summary>
public partial class MainWindow : Window, IWindowService
{
    private readonly MainWindowViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        _viewModel = new MainWindowViewModel(this);
        DataContext = _viewModel;

        // ドラッグ&ドロップの設定
        AddHandler(DragDrop.DragOverEvent, Window_DragOver);
        AddHandler(DragDrop.DropEvent, Window_Drop);

        // ライフサイクルイベント
        Opened += (_, _) => _viewModel.OnLoaded();
        Closing += (_, _) => _viewModel.OnClosing();
    }

    #region IWindowService 実装

    public async Task ShowMessageAsync(string title, string message)
    {
        var dialog = new Window
        {
            Title = title,
            Width = 450,
            Height = 220,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false,
            Content = new StackPanel
            {
                Margin = new Thickness(20),
                Spacing = 15,
                Children =
                {
                    new TextBlock
                    {
                        Text = message,
                        TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                        MaxWidth = 400
                    },
                    new Button
                    {
                        Content = "OK",
                        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                        Width = 80
                    }
                }
            }
        };

        var button = ((StackPanel)dialog.Content).Children.OfType<Button>().First();
        button.Click += (_, _) => dialog.Close();

        await dialog.ShowDialog(this);
    }

    public async Task<string[]> OpenFilePickerAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            AllowMultiple = true,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("音声/動画ファイル")
                {
                    Patterns = new[] { "*.wav", "*.mp3", "*.ogg", "*.flac", "*.aac", "*.mp4" }
                },
                new FilePickerFileType("すべてのファイル")
                {
                    Patterns = new[] { "*.*" }
                }
            }
        });

        return files
            .Select(f => f.TryGetLocalPath())
            .Where(p => p != null)
            .Cast<string>()
            .ToArray();
    }

    public async Task<string?> OpenFolderPickerAsync()
    {
        var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "出力フォルダを選択してください",
            AllowMultiple = false
        });

        if (folders.Count > 0)
        {
            return folders[0].TryGetLocalPath();
        }

        return null;
    }

    #endregion

    #region ドラッグ&ドロップ

    private void Window_DragOver(object? sender, DragEventArgs e)
    {
#pragma warning disable CS0618 // Avalonia 11.3: DataTransfer API は GetFiles 未対応のため旧 API を使用
        e.DragEffects = e.Data.Contains(DataFormats.Files)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
#pragma warning restore CS0618
    }

    private void Window_Drop(object? sender, DragEventArgs e)
    {
#pragma warning disable CS0618
        if (e.Data.Contains(DataFormats.Files))
        {
            var files = e.Data.GetFiles();
#pragma warning restore CS0618
            if (files != null)
            {
                var paths = files
                    .Select(f => f.TryGetLocalPath())
                    .Where(p => p != null)
                    .Cast<string>()
                    .ToArray();
                _viewModel.AddFilesToList(paths);
            }
        }
    }

    #endregion
}
