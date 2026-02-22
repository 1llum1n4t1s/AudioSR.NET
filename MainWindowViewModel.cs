using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

using static AudioSR.NET.Logger;

namespace AudioSR.NET;

/// <summary>
/// MainWindow の ViewModel
/// </summary>
public partial class MainWindowViewModel : ObservableObject, IDisposable
{
    private readonly IWindowService _windowService;
    private readonly AppSettings _settings;
    private readonly HashSet<string> _filePaths = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _cancellationTokenSource;
    private AudioSrWrapper? _audioSrInstance;
    private bool _initializationInProgress;
    private static readonly object _initializationLock = new();
    private static readonly Version MaxEmbeddedPythonVersion = new(3, 11, 99);
    private readonly SemaphoreSlim _pythonDetectionSemaphore = new(1, 1);
    private static readonly HttpClient _httpClientForVersion = new(new SocketsHttpHandler { AutomaticDecompression = DecompressionMethods.All }) { Timeout = TimeSpan.FromSeconds(10) };

    public ObservableCollection<FileItem> FileList { get; } = new();

    #region Observable Properties - 設定

    [ObservableProperty]
    private int _modelSelectedIndex;

    [ObservableProperty]
    private int _ddimSteps;

    [ObservableProperty]
    private float _guidanceScale;

    [ObservableProperty]
    private long? _seed;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSeedEnabled))]
    [NotifyPropertyChangedFor(nameof(SeedText))]
    private bool _useRandomSeed;

    [ObservableProperty]
    private string _outputFolder = "";

    [ObservableProperty]
    private bool _overwriteOutputFiles;

    [ObservableProperty]
    private string _pythonHome = "";

    #endregion

    #region Observable Properties - UI状態

    [ObservableProperty]
    private double _progress;

    [ObservableProperty]
    private string _statusText = "準備完了";

    #endregion

    #region Observable Properties - 初期化オーバーレイ

    [ObservableProperty]
    private bool _isInitializing;

    [ObservableProperty]
    private double _initializeProgress;

    [ObservableProperty]
    private string _initializeStatusText = "依存パッケージをインストール中...";

    [ObservableProperty]
    private string _initializeProgressText = "0/10";

    #endregion

    #region Observable Properties - 処理中オーバーレイ

    [ObservableProperty]
    private bool _isProcessing;

    [ObservableProperty]
    private string _processingFileText = "ファイル名...";

    [ObservableProperty]
    private string _processingStatusText = "準備中...";

    [ObservableProperty]
    private double _processingProgress;

    [ObservableProperty]
    private string _processingDetailText = "ステップ: 0/100";

    [ObservableProperty]
    private string _processingOverallText = "全体進捗: 0/0";

    #endregion

    #region Computed Properties

    public bool IsSeedEnabled => !UseRandomSeed;

    public string SeedText
    {
        get => Seed?.ToString() ?? "";
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                Seed = null;
            else if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var seed) && seed >= 0)
                Seed = seed;
        }
    }

    #endregion

    #region Property Change Handlers

    partial void OnModelSelectedIndexChanged(int value)
    {
        _settings.ModelName = value switch
        {
            0 => "basic",
            1 => "speech",
            _ => "basic"
        };
    }

    partial void OnDdimStepsChanged(int value) => _settings.DdimSteps = value;
    partial void OnGuidanceScaleChanged(float value) => _settings.GuidanceScale = value;

    partial void OnSeedChanged(long? value)
    {
        _settings.Seed = value;
        OnPropertyChanged(nameof(SeedText));
    }

    partial void OnUseRandomSeedChanged(bool value) => _settings.UseRandomSeed = value;
    partial void OnOutputFolderChanged(string value) => _settings.OutputFolder = value;
    partial void OnOverwriteOutputFilesChanged(bool value) => _settings.OverwriteOutputFiles = value;
    partial void OnPythonHomeChanged(string value) => _settings.PythonHome = value;

    #endregion

    public MainWindowViewModel(IWindowService windowService)
    {
        _windowService = windowService;
        _settings = AppSettings.Load();

        // 設定からプロパティを初期化
        _modelSelectedIndex = _settings.ModelName switch
        {
            "speech" => 1,
            _ => 0
        };
        _ddimSteps = _settings.DdimSteps;
        _guidanceScale = _settings.GuidanceScale;
        _seed = _settings.Seed;
        _useRandomSeed = _settings.UseRandomSeed;
        _outputFolder = _settings.OutputFolder;
        _overwriteOutputFiles = _settings.OverwriteOutputFiles;
        _pythonHome = _settings.PythonHome;

        // 出力フォルダが存在しない場合は作成
        if (!Directory.Exists(OutputFolder))
        {
            try { Directory.CreateDirectory(OutputFolder); }
            catch (Exception ex) { Logger.Log($"出力フォルダの作成に失敗しました: {ex.Message}", LogLevel.Info); }
        }
    }

    #region ライフサイクル

    /// <summary>
    /// ウィンドウロード時に呼ばれる
    /// </summary>
    public void OnLoaded()
    {
        Logger.Log("アプリケーションが起動しました。", LogLevel.Info);
        Logger.Log("MainWindowがロードされました", LogLevel.Info);
        InitializeAudioSRAsync();
    }

    /// <summary>
    /// ウィンドウ閉じる時に呼ばれる
    /// </summary>
    public void OnClosing()
    {
        SaveSettingsInternal("アプリ終了時の自動保存", logSuccess: false);
        _cancellationTokenSource?.Cancel();
        _audioSrInstance?.Dispose();
    }

    public void Dispose()
    {
        _pythonDetectionSemaphore.Dispose();
        _cancellationTokenSource?.Dispose();
        _audioSrInstance?.Dispose();
    }

    #endregion

    #region Commands

    [RelayCommand]
    private async Task AddFilesAsync()
    {
        var paths = await _windowService.OpenFilePickerAsync();
        if (paths.Length > 0)
        {
            AddFilesToList(paths);
        }
    }

    [RelayCommand]
    private void RemoveFiles(IList? selectedItems)
    {
        if (selectedItems == null) return;
        var items = selectedItems.Cast<FileItem>().ToList();
        foreach (var item in items)
        {
            _filePaths.Remove(item.Path);
            FileList.Remove(item);
        }
    }

    [RelayCommand]
    private void ClearFiles()
    {
        _filePaths.Clear();
        FileList.Clear();
    }

    [RelayCommand]
    private async Task BrowseOutputFolderAsync()
    {
        var path = await _windowService.OpenFolderPickerAsync();
        if (path != null)
        {
            OutputFolder = path;
        }
    }

    [RelayCommand]
    private void SaveSettings()
    {
        SaveSettingsInternal("手動保存", logSuccess: true);
    }

    [RelayCommand]
    private async Task StartProcessingAsync()
    {
        SaveSettingsInternal("処理開始時の自動保存", logSuccess: false);

        if (FileList.Count == 0)
        {
            Logger.Log("処理するファイルがありません。ファイルを追加してください。", LogLevel.Info);
            return;
        }

        if (_cancellationTokenSource != null)
        {
            Logger.Log("既に処理を実行中です。", LogLevel.Info);
            return;
        }

        if (_audioSrInstance == null)
        {
            if (!await EnsureAudioSrInitializedAsync())
            {
                Logger.Log("AudioSRの初期化に失敗しました。詳細はログを確認してください。", LogLevel.Info);
                return;
            }
        }

        if (_audioSrInstance == null)
        {
            Logger.Log("AudioSRの初期化に失敗しました。", LogLevel.Info);
            return;
        }

        _cancellationTokenSource = new CancellationTokenSource();
        StartAudioProcessing(_audioSrInstance, _cancellationTokenSource.Token);
    }

    [RelayCommand]
    private void StopProcessing()
    {
        if (_cancellationTokenSource != null)
        {
            _cancellationTokenSource.Cancel();
            Logger.Log("処理を中断しています...", LogLevel.Info);
        }
    }

    [RelayCommand]
    private async Task RetryEmbeddedPythonAsync()
    {
        Logger.Log("埋め込みPythonの再試行を開始します。", LogLevel.Info);
        var pythonReady = await FindEmbeddedPythonAsync();
        if (pythonReady && !string.IsNullOrEmpty(PythonHome))
        {
            await _windowService.ShowMessageAsync("再試行完了",
                $"埋め込みPythonの準備が完了しました: {PythonHome}");
        }
    }

    [RelayCommand]
    private async Task ShowEmbeddedPythonHelpAsync()
    {
        await _windowService.ShowMessageAsync("Pythonランタイムのヘルプ", BuildEmbeddedPythonGuidance());
    }

    #endregion

    #region ファイルリスト操作

    /// <summary>
    /// ドラッグ&ドロップ等からファイルを追加
    /// </summary>
    public void AddFilesToList(string[] filePaths)
    {
        var addedCount = 0;
        var skippedCount = 0;

        foreach (var filePath in filePaths)
        {
            if (!_filePaths.Add(filePath))
            {
                skippedCount++;
                continue;
            }

            var fileItem = new FileItem { Path = filePath };

            if (!fileItem.Exists())
            {
                Logger.Log($"ファイルが存在しません: {filePath}", LogLevel.Error);
                _filePaths.Remove(filePath);
                skippedCount++;
                continue;
            }

            if (!fileItem.IsAudioFile())
            {
                Logger.Log($"{Path.GetFileName(filePath)} は音声ファイルではない可能性があります。", LogLevel.Warning);
            }

            FileList.Add(fileItem);
            addedCount++;
        }

        Logger.Log($"{addedCount}個のファイルを追加しました。{skippedCount}個のファイルはスキップされました。", LogLevel.Info);
    }

    private string GetAvailableOutputFilePath(string outputFilePath)
    {
        if (OverwriteOutputFiles || !File.Exists(outputFilePath))
        {
            return outputFilePath;
        }

        var directory = Path.GetDirectoryName(outputFilePath) ?? "";
        var fileName = Path.GetFileNameWithoutExtension(outputFilePath);
        var extension = Path.GetExtension(outputFilePath);

        for (var index = 1; index <= 9999; index++)
        {
            var candidate = Path.Combine(directory, $"{fileName} ({index}){extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
        return Path.Combine(directory, $"{fileName}_{timestamp}{extension}");
    }

    #endregion

    #region 設定保存

    private bool SaveSettingsInternal(string reason, bool logSuccess)
    {
        if (_settings.Save())
        {
            if (logSuccess)
            {
                Logger.Log("設定を保存しました。", LogLevel.Info);
            }
            return true;
        }

        Logger.Log($"警告: 設定の保存に失敗しました（{reason}）。", LogLevel.Warning);
        Dispatcher.UIThread.InvokeAsync(async () =>
        {
            await _windowService.ShowMessageAsync("設定保存エラー",
                "設定の保存に失敗しました。ディスクの空き容量や書き込み権限を確認してください。");
        });
        return false;
    }

    #endregion

    #region AudioSR初期化

    private async void InitializeAudioSRAsync()
    {
        try
        {
            if (!await EnsureAudioSrInitializedAsync())
            {
                await _windowService.ShowMessageAsync("エラー", "初期化に失敗しました。詳細はログを確認してください。");
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"初期化中に致命的なエラーが発生しました: {ex.Message}", LogLevel.Error);
            await _windowService.ShowMessageAsync("エラー", "初期化に失敗しました。詳細はログを確認してください。");
        }
    }

    private async Task<bool> EnsureAudioSrInitializedAsync()
    {
        lock (_initializationLock)
        {
            if (_audioSrInstance != null)
            {
                return true;
            }

            if (_initializationInProgress)
            {
                Logger.Log("AudioSRを初期化中です。少しお待ちください...", LogLevel.Info);
                return false;
            }

            _initializationInProgress = true;
        }

        try
        {
            var pythonReady = await FindEmbeddedPythonAsync();
            if (!pythonReady || string.IsNullOrEmpty(PythonHome))
            {
                Logger.Log("Pythonホームディレクトリの準備に失敗しました。", LogLevel.Info);
                Logger.Log("Pythonホームディレクトリが設定されていません", LogLevel.Warning);
                return false;
            }

            IsInitializing = true;

            try
            {
                _audioSrInstance = new AudioSrWrapper(PythonHome);
                await _audioSrInstance.InitializeAsync((step, totalSteps, message) =>
                {
                    Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        InitializeProgress = (double)step / totalSteps * 100;
                        InitializeStatusText = message;
                        InitializeProgressText = $"{step}/{totalSteps}";
                    });
                });
                Logger.Log("AudioSRの初期化が完了しました", LogLevel.Info);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"初期化エラー: {ex.Message}", LogLevel.Info);
                Logger.LogException("AudioSR初期化中にエラーが発生しました", ex);
                _audioSrInstance = null;
                return false;
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"初期化中に致命的なエラーが発生しました: {ex.Message}", LogLevel.Error);
            return false;
        }
        finally
        {
            IsInitializing = false;

            lock (_initializationLock)
            {
                _initializationInProgress = false;
            }
        }
    }

    #endregion

    #region 音声処理

    private async void StartAudioProcessing(AudioSrWrapper audioSr, CancellationToken cancellationToken)
    {
        Logger.Log("処理を開始します...", LogLevel.Info);
        Progress = 0;
        StatusText = "処理中...";
        IsProcessing = true;

        try
        {
            Logger.Log("AudioSRを初期化しました。", LogLevel.Info);

            var filesToProcess = FileList.ToList();
            var totalFiles = filesToProcess.Count;
            var processedFiles = 0;
            var modelName = _settings.ModelName;

            foreach (var fileItem in filesToProcess)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    _ = Dispatcher.UIThread.InvokeAsync(() => { fileItem.Status = "キャンセル"; });
                    break;
                }

                try
                {
                    _ = Dispatcher.UIThread.InvokeAsync(() => { fileItem.Status = "処理中..."; });

                    var fileName = Path.GetFileName(fileItem.Path);
                    var statusMsg = $"{fileName} を処理中...";
                    StatusText = statusMsg;
                    Logger.Log(statusMsg, LogLevel.Info);

                    ProcessingFileText = fileName;
                    ProcessingStatusText = "処理を実行中...";
                    ProcessingProgress = 0;
                    ProcessingDetailText = $"ステップ: 0/{DdimSteps}";
                    ProcessingOverallText = $"全体進捗: {processedFiles + 1}/{totalFiles}";

                    var outputFile = GetAvailableOutputFilePath(Path.Combine(OutputFolder, fileName));

                    await audioSr.ProcessFileAsync(
                        fileItem.Path,
                        outputFile,
                        modelName,
                        DdimSteps,
                        GuidanceScale,
                        UseRandomSeed ? null : Seed,
                        (currentStep, totalSteps) =>
                        {
                            var effectiveTotal = totalSteps > 0 ? totalSteps : DdimSteps;
                            var percent = effectiveTotal > 0 ? (double)currentStep / effectiveTotal * 100 : 0;
                            if (percent > 100) percent = 100;

                            Progress = percent;
                            ProcessingStatusText = "サンプリング実行中...";
                            ProcessingProgress = percent;
                            ProcessingDetailText = $"ステップ: {currentStep}/{effectiveTotal}";
                            ProcessingOverallText = $"全体進捗: {processedFiles + 1}/{totalFiles}";
                        }
                    );

                    _ = Dispatcher.UIThread.InvokeAsync(() => { fileItem.Status = "完了"; });
                    Logger.Log($"{fileName} の処理が完了しました。", LogLevel.Info);
                }
                catch (OperationCanceledException)
                {
                    _ = Dispatcher.UIThread.InvokeAsync(() => { fileItem.Status = "キャンセル"; });
                    throw;
                }
                catch (Exception ex)
                {
                    _ = Dispatcher.UIThread.InvokeAsync(() => { fileItem.Status = "エラー"; });
                    Logger.Log($"{Path.GetFileName(fileItem.Path)} の処理中にエラーが発生しました: {ex.Message}", LogLevel.Error);
                }

                processedFiles++;
                var progressPercent = (double)processedFiles / totalFiles * 100;
                Progress = progressPercent;
                StatusText = $"処理進捗: {processedFiles}/{totalFiles} ({progressPercent:F1}%)";
            }

            StatusText = "処理完了";
            Logger.Log("すべての処理が完了しました。", LogLevel.Info);
        }
        catch (OperationCanceledException)
        {
            StatusText = "処理が中断されました";
            Logger.Log("処理が中断されました。", LogLevel.Warning);
        }
        catch (Exception ex)
        {
            StatusText = "エラーが発生しました";
            Logger.Log($"エラーが発生しました: {ex.Message}", LogLevel.Error);
            var userMessage = "処理中にエラーが発生しました。Python設定や入出力フォルダを確認のうえ、再試行してください。";
            var invalidOperation = ex as InvalidOperationException ?? ex.InnerException as InvalidOperationException;
            if (invalidOperation != null && !string.IsNullOrWhiteSpace(invalidOperation.Message))
            {
                userMessage = invalidOperation.Message;
            }
            await _windowService.ShowMessageAsync("処理エラー", userMessage);
        }
        finally
        {
            IsProcessing = false;
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    #endregion

    #region Python検出・ダウンロード

    private void NotifyEmbeddedPythonFailure(string reason)
    {
        Logger.Log(reason, LogLevel.Info);
        StatusText = "ランタイムの準備に失敗しました。";
        _ = _windowService.ShowMessageAsync("Pythonランタイムの準備に失敗しました",
            $"{reason}{Environment.NewLine}{Environment.NewLine}{BuildEmbeddedPythonGuidance()}");
    }

    private string BuildEmbeddedPythonGuidance()
    {
        var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var managedEmbeddedPath = Path.Combine(appDirectory, "lib", "python", "python-embed");
        var legacyEmbeddedPath = Path.Combine(appDirectory, "python-embed");
        var legacyParentPath = Path.Combine(appDirectory, "python");
        var versionLabel = string.IsNullOrEmpty(_settings.PythonVersion) ? "未取得" : _settings.PythonVersion;

        return string.Join(
            Environment.NewLine,
            "次に試せる対応:",
            "・「ランタイム再試行」ボタンで再ダウンロードを試す",
            "・ネットワーク接続やプロキシ設定を確認する",
            $"・手動配置する場合は次のいずれかに python-embed フォルダを配置: {managedEmbeddedPath}",
            $"・旧形式フォルダでも可: {legacyEmbeddedPath} または {legacyParentPath}",
            "・自動ダウンロードは Python 3.12 系までが対象",
            $"・現在の目標バージョン: {versionLabel}",
            "・ログを確認して失敗理由を確認する");
    }

    private async Task<bool> FindEmbeddedPythonAsync()
    {
        await _pythonDetectionSemaphore.WaitAsync();
        try
        {
            Logger.Log("埋め込みPythonの検出を開始します...", LogLevel.Info);
            var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            Logger.Log($"アプリケーションディレクトリ: {appDirectory}", LogLevel.Info);
            var managedEmbeddedPath = Path.Combine(appDirectory, "lib", "python", "python-embed");
            Logger.Log($"管理対象Pythonパス: {managedEmbeddedPath}", LogLevel.Info);
            Logger.Log("対象バージョンを取得中...", LogLevel.Info);
            var targetVersion = await GetTargetEmbeddedPythonVersionAsync();
            Logger.Log($"取得バージョン: {(string.IsNullOrEmpty(targetVersion) ? "なし" : targetVersion)}", LogLevel.Info);
            if (!string.IsNullOrEmpty(targetVersion))
            {
                Logger.Log($"EnsureEmbeddedPythonAsync を実行中: パス={managedEmbeddedPath}, バージョン={targetVersion}", LogLevel.Info);
                var managedReady = await EnsureEmbeddedPythonAsync(managedEmbeddedPath, targetVersion);
                Logger.Log($"EnsureEmbeddedPythonAsync 完了: {managedReady}", LogLevel.Info);
                if (managedReady)
                {
                    PythonHome = managedEmbeddedPath;
                    Logger.Log($"管理対象の埋め込みPythonを使用します: {PythonHome}", LogLevel.Info);
                    return true;
                }
            }
            Logger.Log("管理対象Pythonは利用できません。別の場所を検索します...", LogLevel.Info);

            var libPythonDir = Path.Combine(appDirectory, "lib", "python");
            Logger.Log($"新しい組み込みPythonを探します: {libPythonDir}", LogLevel.Info);

            if (Directory.Exists(libPythonDir))
            {
                var pythonDirs = Directory.GetDirectories(libPythonDir, "python*embed*");
                if (pythonDirs.Length > 0)
                {
                    var embeddedPythonPath = pythonDirs[0];
                    if (Directory.Exists(embeddedPythonPath) &&
                        File.Exists(Path.Combine(embeddedPythonPath, "python.exe")))
                    {
                        PythonHome = embeddedPythonPath;
                        Logger.Log($"lib/python内の埋め込みPythonを検出しました: {PythonHome}", LogLevel.Info);
                        return true;
                    }
                }
            }

            var oldPythonDir = Path.Combine(appDirectory, "python");
            Logger.Log($"旧形式の組み込みPythonを探します: {oldPythonDir}", LogLevel.Info);

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
                        Logger.Log($"旧形式の埋め込みPythonを検出しました: {PythonHome}", LogLevel.Info);
                        return true;
                    }
                }
            }

            var directEmbeddedPath = Path.Combine(appDirectory, "python-embed");
            if (Directory.Exists(directEmbeddedPath) && File.Exists(Path.Combine(directEmbeddedPath, "python.exe")))
            {
                PythonHome = directEmbeddedPath;
                Logger.Log($"直接的な組み込みPythonを検出しました: {PythonHome}", LogLevel.Info);
                return true;
            }

            Logger.Log("組み込みPythonが見つからないため、ダウンロードを試行します。", LogLevel.Info);
            if (await DownloadEmbeddedPythonAsync())
            {
                return true;
            }

            NotifyEmbeddedPythonFailure("必要なランタイムが見つかりません。");
            return false;
        }
        catch (Exception ex)
        {
            NotifyEmbeddedPythonFailure($"Pythonパスの設定中に例外が発生しました: {ex.Message}");
            return false;
        }
        finally
        {
            _pythonDetectionSemaphore.Release();
        }
    }

    private async Task<bool> DownloadEmbeddedPythonAsync()
    {
        var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var basePythonDir = Path.Combine(appDirectory, "lib", "python");
        var embeddedPythonPath = Path.Combine(basePythonDir, "python-embed");

        try
        {
            Logger.Log("DownloadEmbeddedPythonAsync 開始", LogLevel.Info);

            Logger.Log("ResolveDownloadablePythonVersionAsync を実行中...", LogLevel.Info);
            var targetVersion = await ResolveDownloadablePythonVersionAsync(_httpClientForVersion);

            Logger.Log($"ResolveDownloadablePythonVersionAsync 完了: {(targetVersion ?? "失敗")}", LogLevel.Info);
            if (string.IsNullOrEmpty(targetVersion))
            {
                throw new InvalidOperationException("最新のPythonバージョン情報を取得できませんでした。");
            }

            var versionFilePath = GetEmbeddedPythonVersionFilePath(embeddedPythonPath);
            var currentVersion = ReadEmbeddedPythonVersion(versionFilePath);

            if (IsEmbeddedPythonReady(embeddedPythonPath) && string.Equals(currentVersion, targetVersion, StringComparison.OrdinalIgnoreCase))
            {
                PythonHome = embeddedPythonPath;
                Logger.Log($"既に埋め込みPythonが存在します: {PythonHome}", LogLevel.Info);
                StatusText = "準備完了";
                return true;
            }

            var zipFileName = $"python-{targetVersion}-embed-amd64.zip";
            var zipPath = Path.Combine(basePythonDir, zipFileName);
            var downloadUrl = new Uri($"https://www.python.org/ftp/python/{targetVersion}/{zipFileName}");

            if (!string.IsNullOrEmpty(currentVersion))
            {
                Logger.Log($"埋め込みPythonのバージョンが古いため更新します。現在: {currentVersion}, 予定: {targetVersion}", LogLevel.Info);
            }
            else
            {
                Logger.Log("埋め込みPythonが未インストールのためダウンロードを開始します。", LogLevel.Info);
            }

            if (Directory.Exists(embeddedPythonPath))
            {
                Directory.Delete(embeddedPythonPath, true);
            }
            Directory.CreateDirectory(embeddedPythonPath);

            if (!File.Exists(zipPath))
            {
                StatusText = "埋め込みPythonをダウンロード中...";
                Logger.Log($"埋め込みPythonをダウンロードします: {downloadUrl}", LogLevel.Info);

                using var response = await _httpClientForVersion.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                await using var contentStream = await response.Content.ReadAsStreamAsync();
                await using var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await contentStream.CopyToAsync(fileStream);
            }
            else
            {
                Logger.Log($"ダウンロード済みのZIPを使用します: {zipPath}", LogLevel.Info);
            }

            StatusText = "埋め込みPythonを展開中...";
            Logger.Log("埋め込みPythonの展開を開始します。", LogLevel.Info);
            ZipFile.ExtractToDirectory(zipPath, embeddedPythonPath, true);

            StatusText = "埋め込みPythonを確認中...";
            if (!IsEmbeddedPythonReady(embeddedPythonPath))
            {
                throw new InvalidOperationException("埋め込みPythonの展開に失敗しました。python.exeまたはpython DLLが見つかりません。");
            }

            WriteEmbeddedPythonVersion(versionFilePath, targetVersion);
            if (!string.Equals(_settings.PythonVersion, targetVersion, StringComparison.OrdinalIgnoreCase))
            {
                _settings.PythonVersion = targetVersion;
                SaveSettingsInternal("埋め込みPythonのバージョン更新", logSuccess: false);
            }

            PythonHome = embeddedPythonPath;
            Logger.Log($"埋め込みPythonの準備が完了しました: {PythonHome}", LogLevel.Info);
            StatusText = "準備完了";
            return true;
        }
        catch (Exception ex)
        {
            NotifyEmbeddedPythonFailure($"埋め込みPythonのダウンロード/展開に失敗しました: {ex.Message}");
            return false;
        }
    }

    private static bool IsEmbeddedPythonReady(string embeddedPythonPath)
    {
        if (!Directory.Exists(embeddedPythonPath))
            return false;

        var pythonExePath = Path.Combine(embeddedPythonPath, "python.exe");
        if (!File.Exists(pythonExePath))
            return false;

        return Directory.GetFiles(embeddedPythonPath, "python*.dll").Length > 0;
    }

    private static string GetEmbeddedPythonVersionFilePath(string embeddedPythonPath)
    {
        return Path.Combine(embeddedPythonPath, "python-embed-version.txt");
    }

    private static string? ReadEmbeddedPythonVersion(string versionFilePath)
    {
        if (!File.Exists(versionFilePath))
            return null;

        var content = File.ReadAllText(versionFilePath).Trim();
        return string.IsNullOrEmpty(content) ? null : content;
    }

    private static void WriteEmbeddedPythonVersion(string versionFilePath, string version)
    {
        File.WriteAllText(versionFilePath, version);
    }

    private async Task<bool> EnsureEmbeddedPythonAsync(string embeddedPythonPath, string targetVersion)
    {
        Logger.Log($"EnsureEmbeddedPythonAsync 開始: パス={embeddedPythonPath}", LogLevel.Info);
        var versionFilePath = GetEmbeddedPythonVersionFilePath(embeddedPythonPath);
        Logger.Log($"バージョンファイルパス: {versionFilePath}", LogLevel.Info);
        var currentVersion = ReadEmbeddedPythonVersion(versionFilePath);
        Logger.Log($"現在のバージョン: {(currentVersion ?? "なし")}, 対象バージョン: {targetVersion}", LogLevel.Info);

        Logger.Log("IsEmbeddedPythonReady をチェック中...", LogLevel.Info);
        if (!IsEmbeddedPythonReady(embeddedPythonPath))
        {
            Logger.Log("Pythonが準備できていません。ダウンロードを開始します...", LogLevel.Info);
            var result = await DownloadEmbeddedPythonAsync();
            Logger.Log($"ダウンロード結果: {result}", LogLevel.Info);
            return result;
        }

        Logger.Log("Pythonは準備完了です。バージョンを確認します...", LogLevel.Info);
        if (!string.Equals(currentVersion, targetVersion, StringComparison.OrdinalIgnoreCase))
        {
            Logger.Log($"バージョン不一致: {currentVersion} != {targetVersion}。アップデートを試みます...", LogLevel.Info);
            try
            {
                var result = await DownloadEmbeddedPythonAsync();
                Logger.Log($"アップデート結果: {result}", LogLevel.Info);
                if (result)
                {
                    return true;
                }
                Logger.Log("警告: アップデートに失敗しました。既存のPythonを使用します。", LogLevel.Info);
            }
            catch (Exception ex)
            {
                Logger.Log($"警告: アップデート中にエラー: {ex.Message}。既存のPythonを使用します。", LogLevel.Info);
            }
            return true;
        }

        Logger.Log("Pythonは最新バージョンです。処理を続行します。", LogLevel.Info);
        return true;
    }

    private async Task<string?> GetTargetEmbeddedPythonVersionAsync()
    {
        Logger.Log($"GetTargetEmbeddedPythonVersionAsync 開始: 保存済みバージョン={(_settings.PythonVersion ?? "なし")}", LogLevel.Info);

        if (!string.IsNullOrEmpty(_settings.PythonVersion))
        {
            Logger.Log($"保存済みバージョンの検証中: {_settings.PythonVersion}", LogLevel.Info);

            if (Version.TryParse(_settings.PythonVersion, out var configuredVersion) &&
                !IsSupportedEmbeddedPythonVersion(configuredVersion))
            {
                Logger.Log($"保存済みのPythonバージョン {_settings.PythonVersion} は上限を超えるため再取得します。", LogLevel.Info);
                _settings.PythonVersion = "";
            }
            else
            {
                try
                {
                    Logger.Log("IsEmbeddedPythonDownloadAvailableAsync をチェック中...", LogLevel.Info);
                    if (await IsEmbeddedPythonDownloadAvailableAsync(_httpClientForVersion, _settings.PythonVersion))
                    {
                        Logger.Log($"保存済みバージョン {_settings.PythonVersion} はダウンロード可能です。", LogLevel.Info);
                        return _settings.PythonVersion;
                    }

                    Logger.Log($"保存済みのPythonバージョン {_settings.PythonVersion} はダウンロード不可のため再取得します。", LogLevel.Info);
                }
                catch (Exception ex)
                {
                    Logger.Log($"保存済みのPythonバージョン {_settings.PythonVersion} の確認に失敗しました: {ex.Message}", LogLevel.Info);
                }

                _settings.PythonVersion = "";
            }
        }

        Logger.Log("最新の安定したPythonバージョンを取得中...", LogLevel.Info);
        var latestVersion = await GetLatestStablePythonVersionAsync();
        Logger.Log($"最新バージョン取得結果: {(latestVersion ?? "失敗")}", LogLevel.Info);
        if (string.IsNullOrEmpty(latestVersion))
        {
            return null;
        }

        _settings.PythonVersion = latestVersion;
        Logger.Log($"Pythonバージョンを設定に保存中: {latestVersion}", LogLevel.Info);
        SaveSettingsInternal("Pythonバージョン情報の保存", logSuccess: false);
        return latestVersion;
    }

    [System.Text.RegularExpressions.GeneratedRegex(@"href=""(?<version>\d+\.\d+\.\d+)/""")]
    private static partial System.Text.RegularExpressions.Regex VersionRegex();

    private static async Task<string?> GetLatestStablePythonVersionAsync()
    {
        try
        {
            Logger.Log("GetLatestStablePythonVersionAsync 開始: Python.orgからバージョン一覧を取得中...", LogLevel.Info);
            Logger.Log("python.org/ftp/python/ へのアクセス中...", LogLevel.Info);
            var indexContent = await _httpClientForVersion.GetStringAsync("https://www.python.org/ftp/python/");

            Logger.Log($"python.orgから応答を取得しました（{indexContent.Length}バイト）", LogLevel.Info);

            var versions = new System.Collections.Generic.List<Version>();
            var matches = VersionRegex().Matches(indexContent);

            Logger.Log($"見つかったバージョン数: {matches.Count}", LogLevel.Info);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                var versionText = match.Groups["version"].Value;
                if (Version.TryParse(versionText, out var version))
                {
                    versions.Add(version);
                }
            }

            Logger.Log($"パース可能なバージョン数: {versions.Count}", LogLevel.Info);

            if (versions.Count == 0)
            {
                Logger.Log("エラー: バージョン一覧から有効なバージョンが見つかりません。", LogLevel.Info);
                return null;
            }

            versions.Sort((a, b) => b.CompareTo(a));
            Logger.Log($"最新バージョン候補: {versions[0]}", LogLevel.Info);

            Logger.Log("ダウンロード可能なバージョンを検索中...", LogLevel.Info);
            foreach (var version in versions)
            {
                try
                {
                    if (!IsSupportedEmbeddedPythonVersion(version))
                    {
                        Logger.Log($"バージョン {version} はサポート対象外です。スキップします。", LogLevel.Info);
                        continue;
                    }

                    Logger.Log($"バージョン {version} のダウンロード可能性をチェック中...", LogLevel.Info);
                    if (await IsEmbeddedPythonDownloadAvailableAsync(_httpClientForVersion, version.ToString()))
                    {
                        Logger.Log($"ダウンロード可能なPythonバージョンを検出しました: {version}", LogLevel.Info);
                        return version.ToString();
                    }

                    Logger.Log($"バージョン {version} はダウンロード不可です。", LogLevel.Info);
                }
                catch (Exception ex)
                {
                    Logger.Log($"Pythonバージョン {version} の確認中にエラー: {ex.Message}", LogLevel.Info);
                    continue;
                }
            }

            Logger.Log("エラー: ダウンロード可能なPythonバージョンが見つかりません。", LogLevel.Info);
            return null;
        }
        catch (Exception ex)
        {
            Logger.Log($"エラー: 最新のPythonバージョン取得に失敗しました: {ex.GetType().Name}: {ex.Message}", LogLevel.Info);
            return null;
        }
    }

    private static async Task<bool> IsEmbeddedPythonDownloadAvailableAsync(HttpClient httpClient, string version)
    {
        var zipFileName = $"python-{version}-embed-amd64.zip";
        var downloadUrl = new Uri($"https://www.python.org/ftp/python/{version}/{zipFileName}");
        Logger.Log($"ダウンロードURL確認中: {downloadUrl}", LogLevel.Info);

        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));

            using var request = new HttpRequestMessage(HttpMethod.Head, downloadUrl);
            Logger.Log("HEADリクエスト送信中（タイムアウト3秒）...", LogLevel.Info);
            using var response = await httpClient.SendAsync(request, cts.Token);

            Logger.Log($"HEADリクエスト応答: {response.StatusCode}", LogLevel.Info);
            if (response.IsSuccessStatusCode)
            {
                Logger.Log("ファイルは利用可能です（HEAD成功）", LogLevel.Info);
                return true;
            }

            if (response.StatusCode != HttpStatusCode.MethodNotAllowed)
            {
                Logger.Log($"ファイルは利用不可です（状態: {response.StatusCode}）", LogLevel.Info);
                return false;
            }

            Logger.Log("HEADがサポートされていません。GETで確認中...", LogLevel.Info);
            using var getRequest = new HttpRequestMessage(HttpMethod.Get, downloadUrl);
            using var getResponse = await httpClient.SendAsync(getRequest, cts.Token);
            Logger.Log($"GETリクエスト応答: {getResponse.StatusCode}", LogLevel.Info);
            return getResponse.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            Logger.Log($"ダウンロード確認エラー: {ex.GetType().Name}: {ex.Message}", LogLevel.Info);
            return false;
        }
    }

    private async Task<string?> ResolveDownloadablePythonVersionAsync(HttpClient httpClient)
    {
        Logger.Log("ResolveDownloadablePythonVersionAsync 開始", LogLevel.Info);

        var targetVersion = await GetTargetEmbeddedPythonVersionAsync();
        Logger.Log($"取得した対象バージョン: {(targetVersion ?? "なし")}", LogLevel.Info);

        if (!string.IsNullOrEmpty(targetVersion) &&
            Version.TryParse(targetVersion, out var parsedTargetVersion) &&
            IsSupportedEmbeddedPythonVersion(parsedTargetVersion) &&
            await IsEmbeddedPythonDownloadAvailableAsync(httpClient, targetVersion))
        {
            Logger.Log($"対象バージョンは利用可能: {targetVersion}", LogLevel.Info);
            return targetVersion;
        }

        if (!string.IsNullOrEmpty(targetVersion))
        {
            Logger.Log($"埋め込みPython {targetVersion} はダウンロード不可のため再取得します。", LogLevel.Info);
            _settings.PythonVersion = "";
        }

        Logger.Log("最新バージョンの再取得...", LogLevel.Info);
        var latestVersion = await GetLatestStablePythonVersionAsync();
        if (string.IsNullOrEmpty(latestVersion))
        {
            Logger.Log("エラー: ダウンロード可能なバージョンが見つかりません。", LogLevel.Info);
            return null;
        }

        _settings.PythonVersion = latestVersion;
        SaveSettingsInternal("Pythonバージョン情報の保存", logSuccess: false);
        return latestVersion;
    }

    private static bool IsSupportedEmbeddedPythonVersion(Version version)
    {
        return version <= MaxEmbeddedPythonVersion;
    }

    #endregion
}
