using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace AudioSR.NET;

/// <summary>
/// MainWindow.xaml の相互作用ロジック
/// </summary>
public partial class MainWindow : INotifyPropertyChanged
{
    private AppSettings _settings;
    private ObservableCollection<FileItem> _fileList = new();
    private CancellationTokenSource? _cancellationTokenSource;
    private double _progress;
    private readonly SynchronizationContext _syncContext;

    // プロパティ
    public ObservableCollection<FileItem> FileList => _fileList;
    private string _statusText = "準備完了";

    public string StatusText
    {
        get => _statusText;
        set
        {
            if (_statusText != value)
            {
                _statusText = value;
                OnPropertyChanged();
            }
        }
    }

    public string PythonHome
    {
        get => _settings.PythonHome;
        set
        {
            if (_settings.PythonHome != value)
            {
                _settings.PythonHome = value;
                OnPropertyChanged();
            }
        }
    }

    public string ModelName
    {
        get => _settings.ModelName;
        set
        {
            if (_settings.ModelName != value)
            {
                _settings.ModelName = value;
                OnPropertyChanged();
            }
        }
    }

    public int DdimSteps
    {
        get => _settings.DdimSteps;
        set
        {
            if (_settings.DdimSteps != value)
            {
                _settings.DdimSteps = value;
                OnPropertyChanged();
            }
        }
    }

    public float GuidanceScale
    {
        get => _settings.GuidanceScale;
        set
        {
            if (_settings.GuidanceScale != value)
            {
                _settings.GuidanceScale = value;
                OnPropertyChanged();
            }
        }
    }

    public int? Seed
    {
        get => _settings.Seed;
        set
        {
            if (_settings.Seed != value)
            {
                _settings.Seed = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsSeedEnabled => !UseRandomSeed;

    public bool UseRandomSeed
    {
        get => _settings.UseRandomSeed;
        set
        {
            if (_settings.UseRandomSeed != value)
            {
                _settings.UseRandomSeed = value;
                OnPropertyChanged();

                // ランダム設定が変更されたらSeedのUIを更新
                OnPropertyChanged(nameof(Seed));
                OnPropertyChanged(nameof(IsSeedEnabled));
            }
        }
    }

    public string OutputFolder
    {
        get => _settings.OutputFolder;
        set
        {
            if (_settings.OutputFolder != value)
            {
                _settings.OutputFolder = value;
                OnPropertyChanged();
            }
        }
    }

    public double Progress
    {
        get => _progress;
        set
        {
            if (_progress != value)
            {
                _progress = value;
                OnPropertyChanged();
            }
        }
    }

    public MainWindow()
    {
        // 設定の読み込み
        _settings = AppSettings.Load();

        // 同期コンテキストを保存（UIスレッドからの操作用）
        _syncContext = SynchronizationContext.Current ?? throw new InvalidOperationException("同期コンテキストが見つかりません。");

        InitializeComponent();
        DataContext = this;

        // モデル選択の初期設定
        if (_settings.ModelName == "basic")
            cmbModel.SelectedIndex = 0;
        else if (_settings.ModelName == "speech")
            cmbModel.SelectedIndex = 1;
        else
            cmbModel.SelectedIndex = 0;

        // 出力フォルダが存在しない場合は作成
        if (!Directory.Exists(OutputFolder))
        {
            try
            {
                Directory.CreateDirectory(OutputFolder);
            }
            catch (Exception ex)
            {
                LogMessage($"出力フォルダの作成に失敗しました: {ex.Message}");
            }
        }

        // 組み込みPythonを検出して設定する
        FindEmbeddedPython();

        if (string.IsNullOrEmpty(PythonHome))
        {
            LogMessage("警告: 組み込みPythonが見つかりませんでした。設定メニューから手動で指定するか、アプリケーションと同じ場所に組み込みPythonフォルダがあるか確認してください。");
        }
    }

    /// <summary>
    /// 組み込みPythonを検出するメソッド
    /// </summary>
    private void FindEmbeddedPython()
    {
        try
        {
            // アプリケーションと同じディレクトリの組み込みPythonディレクトリを確認
            var appDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // 新しいディレクトリ構造: /lib/python
            var libPythonDir = Path.Combine(appDirectory, "lib", "python");
            LogMessage($"新しい組み込みPythonを探します: {libPythonDir}");

            if (Directory.Exists(libPythonDir))
            {
                // lib/pythonディレクトリに埋め込みPythonフォルダを探す
                var pythonDirs = Directory.GetDirectories(libPythonDir, "python*embed*");
                if (pythonDirs.Length > 0)
                {
                    // 最初に見つかった埋め込みPythonを使用
                    var embeddedPythonPath = pythonDirs[0];
                    if (Directory.Exists(embeddedPythonPath) &&
                        File.Exists(Path.Combine(embeddedPythonPath, "python.exe")))
                    {
                        // 埋め込みPythonを設定
                        PythonHome = embeddedPythonPath;
                        LogMessage($"lib/python内の埋め込みPythonを検出しました: {PythonHome}");
                        return;
                    }
                }
            }

            // 古い方式: 直接pythonディレクトリを確認（後方互換性のため）
            var oldPythonDir = Path.Combine(appDirectory, "python");
            LogMessage($"旧形式の組み込みPythonを探します: {oldPythonDir}");

            if (Directory.Exists(oldPythonDir))
            {
                var pythonDirs = Directory.GetDirectories(oldPythonDir, "python*embed*");
                if (pythonDirs.Length > 0)
                {
                    var embeddedPythonPath = pythonDirs[0];
                    if (Directory.Exists(embeddedPythonPath) &&
                        File.Exists(Path.Combine(embeddedPythonPath, "python.exe")))
                    {
                        PythonHome = embeddedPythonPath;
                        LogMessage($"旧形式の埋め込みPythonを検出しました: {PythonHome}");
                        return;
                    }
                }
            }

            // 直接埋め込みPythonフォルダが存在するか確認
            var directEmbeddedPath = Path.Combine(appDirectory, "python-embed");
            if (Directory.Exists(directEmbeddedPath) && File.Exists(Path.Combine(directEmbeddedPath, "python.exe")))
            {
                PythonHome = directEmbeddedPath;
                LogMessage($"直接的な組み込みPythonを検出しました: {PythonHome}");
                return;
            }

            // 組み込みPythonが見つからない場合はエラーを表示
            LogMessage("エラー: 組み込みPythonが見つかりません。アプリケーションと同じディレクトリに lib/python フォルダが存在するか確認してください。");
        }
        catch (Exception ex)
        {
            LogMessage($"エラー: Pythonパスの設定中に例外が発生しました: {ex.Message}");
        }
    }

    #region イベントハンドラ

    /// <summary>
    /// ドラッグオーバー時のイベントハンドラ
    /// </summary>
    private void Window_DragOver(object sender, DragEventArgs e)
    {
        // ファイルのドラッグを許可
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    /// <summary>
    /// ドロップ時のイベントハンドラ
    /// </summary>
    private void Window_Drop(object sender, DragEventArgs e)
    {
        // ドロップされたファイルを処理
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddFilesToList(files);
        }
    }

    /// <summary>
    /// ファイル追加ボタンのクリックイベント
    /// </summary>
    private void AddFiles_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "音声/動画ファイル|*.wav;*.mp3;*.ogg;*.flac;*.aac;*.mp4|すべてのファイル|*.*",
            Multiselect = true
        };

        if (openFileDialog.ShowDialog() == true)
        {
            AddFilesToList(openFileDialog.FileNames);
        }
    }

    /// <summary>
    /// ファイル削除ボタンのクリックイベント
    /// </summary>
    private void RemoveFiles_Click(object sender, RoutedEventArgs e)
    {
        var selectedItems = lvFiles.SelectedItems.Cast<FileItem>().ToList();
        foreach (var item in selectedItems)
        {
            _fileList.Remove(item);
        }
    }

    /// <summary>
    /// ファイルリストクリアボタンのクリックイベント
    /// </summary>
    private void ClearFiles_Click(object sender, RoutedEventArgs e)
    {
        _fileList.Clear();
    }

    /// <summary>
    /// ログクリアボタンのクリックイベント
    /// </summary>
    private void ClearLog_Click(object sender, RoutedEventArgs e)
    {
        txtLog.Clear();
    }

    /// <summary>
    /// 出力フォルダ参照ボタンのクリックイベント
    /// </summary>
    private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folderDialog = new FolderBrowserDialog
        {
            Description = "出力フォルダを選択してください",
            UseDescriptionForTitle = true
        };

        // 現在の出力フォルダが存在する場合は初期フォルダとして設定
        if (!string.IsNullOrEmpty(OutputFolder) && Directory.Exists(OutputFolder))
        {
            folderDialog.SelectedPath = OutputFolder;
        }

        if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            OutputFolder = folderDialog.SelectedPath;
        }
    }

    /// <summary>
    /// 設定保存ボタンのクリックイベント
    /// </summary>
    private void SaveSettings_Click(object sender, RoutedEventArgs e)
    {
        // コンボボックスから選択値を取得
        if (cmbModel.SelectedIndex == 0)
            ModelName = "basic";
        else if (cmbModel.SelectedIndex == 1)
            ModelName = "speech";

        // 設定を保存
        if (_settings.Save())
        {
            LogMessage("設定を保存しました。");
        }
        else
        {
            LogMessage("設定の保存に失敗しました。");
        }
    }

    /// <summary>
    /// 処理開始ボタンのクリックイベント
    /// </summary>
    private void StartProcessing_Click(object sender, RoutedEventArgs e)
    {
        if (_fileList.Count == 0)
        {
            LogMessage("処理するファイルがありません。ファイルを追加してください。");
            return;
        }

        if (string.IsNullOrEmpty(PythonHome))
        {
            // 埋め込み Python が設定されていない場合は再検出を試行
            if (string.IsNullOrEmpty(PythonHome))
            {
                FindEmbeddedPython();

                if (string.IsNullOrEmpty(PythonHome))
                {
                    LogMessage("Pythonホームディレクトリが設定されていません。設定メニューから指定してください。");
                    return;
                }
            }
        }

        // 既に実行中の場合は何もしない
        if (_cancellationTokenSource != null)
        {
            LogMessage("既に処理を実行中です。");
            return;
        }

        // 処理の開始
        _cancellationTokenSource = new CancellationTokenSource();
        StartAudioProcessing(_cancellationTokenSource.Token);
    }

    /// <summary>
    /// 処理停止ボタンのクリックイベント
    /// </summary>
    private void StopProcessing_Click(object sender, RoutedEventArgs e)
    {
        if (_cancellationTokenSource != null)
        {
            _cancellationTokenSource.Cancel();
            LogMessage("処理を中断しています...");
        }
    }

    #endregion

    #region ヘルパーメソッド

    /// <summary>
    /// ファイルリストにファイルを追加
    /// </summary>
    private void AddFilesToList(string[] filePaths)
    {
        var addedCount = 0;
        var skippedCount = 0;

        foreach (var filePath in filePaths)
        {
            // 既に同じパスのファイルがリストにあれば追加しない
            if (_fileList.Any(f => f.Path.Equals(filePath, StringComparison.OrdinalIgnoreCase)))
            {
                skippedCount++;
                continue;
            }

            // ファイルアイテムを作成して追加
            var fileItem = new FileItem { Path = filePath };

            // 音声ファイルでない場合は警告をログに出す
            if (!fileItem.IsAudioFile())
            {
                LogMessage($"警告: {Path.GetFileName(filePath)} は音声ファイルではない可能性があります。");
            }

            // ファイルが存在しない場合もスキップ
            if (!fileItem.Exists())
            {
                LogMessage($"エラー: ファイルが存在しません: {filePath}");
                skippedCount++;
                continue;
            }

            _fileList.Add(fileItem);
            addedCount++;
        }

        LogMessage($"{addedCount}個のファイルを追加しました。{skippedCount}個のファイルはスキップされました。");
    }

    /// <summary>
    /// ログメッセージを出力
    /// </summary>
    private void LogMessage(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        _syncContext.Post(_ =>
        {
            txtLog.AppendText($"[{timestamp}] {message}{Environment.NewLine}");
            txtLog.ScrollToEnd();
        }, null);
    }

    /// <summary>
    /// 処理状況の更新
    /// </summary>
    private void UpdateStatus(string status)
    {
        _syncContext.Post(_ => { StatusText = status; }, null);
    }

    /// <summary>
    /// 音声処理を開始
    /// </summary>
    private async void StartAudioProcessing(CancellationToken cancellationToken)
    {
        LogMessage("処理を開始します...");
        Progress = 0;

        // 処理中の状態表示用
        UpdateStatus("初期化中...");

        try
        {
            // 詳細なデバッグログを追加
            LogMessage($"Pythonホームディレクトリ: {PythonHome}");
            LogMessage($"初期化開始: {DateTime.Now:HH:mm:ss.fff}");

            // カレントディレクトリとPythonホームディレクトリの内容確認
            var currentDir = Environment.CurrentDirectory;
            LogMessage($"カレントディレクトリ: {currentDir}");

            if (Directory.Exists(PythonHome))
            {
                // 特にPython DLLファイルの確認
                var pythonDlls = Directory.GetFiles(PythonHome, "python*.dll");
                LogMessage($"Python DLLファイル一覧:");
                foreach (var dll in pythonDlls)
                {
                    LogMessage($"  - {Path.GetFileName(dll)}");
                }
            }
            else
            {
                LogMessage($"警告: Pythonホームディレクトリ {PythonHome} が存在しません");
            }

            UpdateStatus("Pythonランタイム初期化中...");
            LogMessage($"Python.Runtime初期化開始: {DateTime.Now:HH:mm:ss.fff}");

            // AudioSRラッパーを初期化（これは時間がかかるので別スレッドで実行）
            AudioSrWrapper? audioSr = null;
            try
            {
                LogMessage("Task.Run開始...");
                await Task.Run(() =>
                {
                    try
                    {
                        LogMessage("バックグラウンドタスク開始: Pythonランタイム初期化");
                        LogMessage("Pythonランタイム初期化を開始します...");

                        // 各ステップでログを記録
                        LogMessage($"AudioSrWrapper コンストラクタ呼び出し開始: {DateTime.Now:HH:mm:ss.fff}");
                        audioSr = new AudioSrWrapper(PythonHome);
                        LogMessage($"AudioSrWrapper コンストラクタ呼び出し完了: {DateTime.Now:HH:mm:ss.fff}");
                        
                        // 明示的にInitializeメソッドを呼び出す
                        LogMessage($"AudioSrWrapper.Initialize 呼び出し開始: {DateTime.Now:HH:mm:ss.fff}");
                        audioSr.Initialize();
                        LogMessage($"AudioSrWrapper.Initialize 呼び出し完了: {DateTime.Now:HH:mm:ss.fff}");
                        LogMessage("Pythonランタイム初期化完了");
                    }
                    catch (Exception innerEx)
                    {
                        // Task内の例外をキャプチャして詳細にログ出力
                        LogMessage($"初期化中の例外（詳細）: {innerEx.GetType().Name}: {innerEx.Message}");
                        if (innerEx.InnerException != null)
                        {
                            LogMessage(
                                $"  内部例外: {innerEx.InnerException.GetType().Name}: {innerEx.InnerException.Message}");
                        }

                        throw; // 上位の例外ハンドラに再スロー
                    }
                }, cancellationToken);
                LogMessage($"初期化完了: {DateTime.Now:HH:mm:ss.fff}");
            }
            catch (Exception ex)
            {
                LogMessage($"初期化中の例外: {ex.GetType().Name}: {ex.Message}");
                throw; // 最上位の例外ハンドラに再スロー
            }

            LogMessage("AudioSRを初期化しました。");

            // 処理するファイル数
            var totalFiles = _fileList.Count;
            var processedFiles = 0;

            // コンボボックスから選択値を取得
            if (cmbModel.SelectedIndex == 0)
                ModelName = "basic";
            else if (cmbModel.SelectedIndex == 1)
                ModelName = "speech";

            // 各ファイルを処理
            foreach (var fileItem in _fileList)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }

                try
                {
                    // 処理中のステータスを更新（UIスレッドで実行）
                    _syncContext.Post(_ => { fileItem.Status = "処理中..."; }, null);

                    // ファイル名を取得
                    var fileName = Path.GetFileName(fileItem.Path);
                    var statusMsg = $"{fileName} を処理中...";
                    UpdateStatus(statusMsg);
                    LogMessage(statusMsg);

                    // 出力ファイルパスを生成
                    var outputFile = Path.Combine(OutputFolder, fileName);

                    // AudioSRで処理（非同期で実行）
                    await Task.Run(() =>
                    {
                        audioSr!.ProcessFile(
                            fileItem.Path,
                            outputFile,
                            ModelName,
                            DdimSteps,
                            GuidanceScale,
                            UseRandomSeed ? null : Seed
                        );
                    }, cancellationToken);

                    // 処理完了のステータスを更新（UIスレッドで実行）
                    _syncContext.Post(_ => { fileItem.Status = "完了"; }, null);
                    LogMessage($"{fileName} の処理が完了しました。");
                }
                catch (Exception ex)
                {
                    // エラー時のステータスを更新（UIスレッドで実行）
                    _syncContext.Post(_ => { fileItem.Status = "エラー"; }, null);
                    LogMessage($"{Path.GetFileName(fileItem.Path)} の処理中にエラーが発生しました: {ex.Message}");
                }

                // 進捗を更新
                processedFiles++;
                var progressPercent = (double)processedFiles / totalFiles * 100;
                Progress = progressPercent;
                UpdateStatus($"処理進捗: {processedFiles}/{totalFiles} ({progressPercent:F1}%)");
            }

            UpdateStatus("処理完了");
            LogMessage("すべての処理が完了しました。");
        }
        catch (OperationCanceledException)
        {
            UpdateStatus("処理が中断されました");
            LogMessage("処理が中断されました。");
        }
        catch (Exception ex)
        {
            UpdateStatus("エラーが発生しました");
            LogMessage($"エラーが発生しました: {ex.Message}");
        }
        finally
        {
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    #endregion

    #region INotifyPropertyChangedの実装

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    #endregion
}

public sealed class SeedValidationRule : ValidationRule
{
    public override ValidationResult Validate(object value, CultureInfo cultureInfo)
    {
        if (value is null)
        {
            return ValidationResult.ValidResult;
        }

        var text = value.ToString();
        if (string.IsNullOrWhiteSpace(text))
        {
            return ValidationResult.ValidResult;
        }

        return int.TryParse(text, NumberStyles.Integer, cultureInfo, out var seed) && seed >= 0
            ? ValidationResult.ValidResult
            : new ValidationResult(false, "0以上の整数を入力してください。");
    }
}
