using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

using static AudioSR.NET.Logger;

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
    private static readonly Version MaxEmbeddedPythonVersion = new(3, 12, 99);

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

    public long? Seed
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

    public bool OverwriteOutputFiles
    {
        get => _settings.OverwriteOutputFiles;
        set
        {
            if (_settings.OverwriteOutputFiles != value)
            {
                _settings.OverwriteOutputFiles = value;
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
                Logger.Log($"出力フォルダの作成に失敗しました: {ex.Message}", LogLevel.Info);
            }
        }

        // 組み込みPythonを検出して設定する
        _ = FindEmbeddedPythonAsync();
    }

    private void NotifyEmbeddedPythonFailure(string reason)
    {
        Logger.Log(reason, LogLevel.Info);
        UpdateStatus("ランタイムの準備に失敗しました。");
        MessageBox.Show(
            $"{reason}{Environment.NewLine}{Environment.NewLine}{BuildEmbeddedPythonGuidance()}",
            "Pythonランタイムの準備に失敗しました",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
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

    /// <summary>
    /// 組み込みPythonを検出するメソッド
    /// </summary>
    private async Task<bool> FindEmbeddedPythonAsync()
    {
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

            // 新しいディレクトリ構造: /lib/python
            var libPythonDir = Path.Combine(appDirectory, "lib", "python");
            Logger.Log($"新しい組み込みPythonを探します: {libPythonDir}", LogLevel.Info);

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
                        Logger.Log($"lib/python内の埋め込みPythonを検出しました: {PythonHome}", LogLevel.Info);
                        return true;
                    }
                }
            }

            // 古い方式: 直接pythonディレクトリを確認（後方互換性のため）
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

            // 直接埋め込みPythonフォルダが存在するか確認
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
    }

    /// <summary>
    /// 公式のembeddable Pythonをダウンロードして展開する
    /// </summary>
    private async Task<bool> DownloadEmbeddedPythonAsync()
    {
        var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var basePythonDir = Path.Combine(appDirectory, "lib", "python");
        var embeddedPythonPath = Path.Combine(basePythonDir, "python-embed");

        try
        {
            Logger.Log("DownloadEmbeddedPythonAsync 開始", LogLevel.Info);
            using var httpClient = new HttpClient(new System.Net.Http.SocketsHttpHandler { AutomaticDecompression = System.Net.DecompressionMethods.All });
            httpClient.Timeout = TimeSpan.FromSeconds(10);
            Logger.Log("ResolveDownloadablePythonVersionAsync を実行中...", LogLevel.Info);
            var targetVersion = await ResolveDownloadablePythonVersionAsync(httpClient);
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
                UpdateStatus("準備完了");
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
                UpdateStatus("埋め込みPythonをダウンロード中...");
                Logger.Log($"埋め込みPythonをダウンロードします: {downloadUrl}", LogLevel.Info);

                using var response = await httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                await using var contentStream = await response.Content.ReadAsStreamAsync();
                await using var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await contentStream.CopyToAsync(fileStream);
            }
            else
            {
                Logger.Log($"ダウンロード済みのZIPを使用します: {zipPath}", LogLevel.Info);
            }

            UpdateStatus("埋め込みPythonを展開中...");
            Logger.Log("埋め込みPythonの展開を開始します。", LogLevel.Info);
            ZipFile.ExtractToDirectory(zipPath, embeddedPythonPath, true);

            UpdateStatus("埋め込みPythonを確認中...");
            if (!IsEmbeddedPythonReady(embeddedPythonPath))
            {
                throw new InvalidOperationException("埋め込みPythonの展開に失敗しました。python.exeまたはpython DLLが見つかりません。");
            }

            WriteEmbeddedPythonVersion(versionFilePath, targetVersion);
            if (!string.Equals(_settings.PythonVersion, targetVersion, StringComparison.OrdinalIgnoreCase))
            {
                _settings.PythonVersion = targetVersion;
                SaveSettingsWithNotification("埋め込みPythonのバージョン更新", logSuccess: false, updateModelSelection: false);
            }

            PythonHome = embeddedPythonPath;
            Logger.Log($"埋め込みPythonの準備が完了しました: {PythonHome}", LogLevel.Info);
            UpdateStatus("準備完了");
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
        {
            return false;
        }

        var pythonExePath = Path.Combine(embeddedPythonPath, "python.exe");
        if (!File.Exists(pythonExePath))
        {
            return false;
        }

        return Directory.GetFiles(embeddedPythonPath, "python*.dll").Length > 0;
    }

    private static string GetEmbeddedPythonVersionFilePath(string embeddedPythonPath)
    {
        return Path.Combine(embeddedPythonPath, "python-embed-version.txt");
    }

    private static string? ReadEmbeddedPythonVersion(string versionFilePath)
    {
        if (!File.Exists(versionFilePath))
        {
            return null;
        }

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
                    Logger.Log($"IsEmbeddedPythonDownloadAvailableAsync をチェック中...", LogLevel.Info);
                    using var httpClient = new HttpClient();
                    if (await IsEmbeddedPythonDownloadAvailableAsync(httpClient, _settings.PythonVersion))
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
        SaveSettingsWithNotification("Pythonバージョン情報の保存", logSuccess: false, updateModelSelection: false);
        return latestVersion;
    }

    private async Task<string?> GetLatestStablePythonVersionAsync()
    {
        try
        {
            Logger.Log("GetLatestStablePythonVersionAsync 開始: Python.orgからバージョン一覧を取得中...", LogLevel.Info);
            using var httpClient = new HttpClient(new System.Net.Http.SocketsHttpHandler { AutomaticDecompression = System.Net.DecompressionMethods.All });
            httpClient.Timeout = TimeSpan.FromSeconds(10);
            Logger.Log("python.org/ftp/python/ へのアクセス中（タイムアウト10秒）...", LogLevel.Info);
            var indexContent = await httpClient.GetStringAsync("https://www.python.org/ftp/python/");
            Logger.Log($"python.orgから応答を取得しました（{indexContent.Length}バイト）", LogLevel.Info);
            var versions = new System.Collections.Generic.List<Version>();
            var matches = System.Text.RegularExpressions.Regex.Matches(indexContent, @"href=""(?<version>\d+\.\d+\.\d+)/""");
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
            Logger.Log($"最新バージョン: {versions[0]}", LogLevel.Info);

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
                    if (await IsEmbeddedPythonDownloadAvailableAsync(httpClient, version.ToString()))
                    {
                        Logger.Log($"✓ ダウンロード可能なPythonバージョンを検出しました: {version}", LogLevel.Info);
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

    private async Task<bool> IsEmbeddedPythonDownloadAvailableAsync(HttpClient httpClient, string version)
    {
        var zipFileName = $"python-{version}-embed-amd64.zip";
        var downloadUrl = new Uri($"https://www.python.org/ftp/python/{version}/{zipFileName}");
        Logger.Log($"ダウンロードURL確認中: {downloadUrl}", LogLevel.Info);

        try
        {
            using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(3));
            using var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Head, downloadUrl);
            Logger.Log($"HEADリクエスト送信中（タイムアウト3秒）...", LogLevel.Info);
            using var response = await httpClient.SendAsync(request, cts.Token);
            Logger.Log($"HEADリクエスト応答: {response.StatusCode}", LogLevel.Info);
            if (response.IsSuccessStatusCode)
            {
                Logger.Log($"✓ ファイルは利用可能です（HEAD成功）", LogLevel.Info);
                return true;
            }

            if (response.StatusCode != HttpStatusCode.MethodNotAllowed)
            {
                Logger.Log($"ファイルは利用不可です（状態: {response.StatusCode}）", LogLevel.Info);
                return false;
            }

            Logger.Log($"HEADがサポートされていません。GETで確認中...", LogLevel.Info);
            using var getRequest = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, downloadUrl);
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
        SaveSettingsWithNotification("Pythonバージョン情報の保存", logSuccess: false, updateModelSelection: false);
        return latestVersion;
    }

    private static bool IsSupportedEmbeddedPythonVersion(Version version)
    {
        return version <= MaxEmbeddedPythonVersion;
    }

    #region イベントハンドラ

    /// <summary>
    /// ウィンドウロード時のイベントハンドラ
    /// </summary>
    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        Logger.Log("アプリケーションが起動しました。", LogLevel.Info);
        Logger.Log("MainWindowがロードされました", LogLevel.Info);

        // アプリ起動時に自動的に初期化処理を開始
        InitializeAudioSRAsync();
    }

    /// <summary>
    /// AudioSRを初期化する（アプリ起動時に自動実行）
    /// </summary>
    private async void InitializeAudioSRAsync()
    {
        // ロックで複数スレッド/複数呼び出しからの同時初期化を防止
        lock (_initializationLock)
        {
            // 既に初期化済みまたは初期化中の場合はスキップ
            if (_audioSrInstance != null || _initializationInProgress)
            {
                Logger.Log("AudioSRは既に初期化されているか、初期化中です", LogLevel.Info);
                return;
            }

            _initializationInProgress = true;
        }

        try
        {
            // PythonHomeが設定されていない場合は取得
            if (string.IsNullOrEmpty(PythonHome))
            {
                var pythonReady = await FindEmbeddedPythonAsync();
                if (!pythonReady || string.IsNullOrEmpty(PythonHome))
                {
                    Logger.Log("Pythonホームディレクトリが設定されていません。", LogLevel.Info);
                    Logger.Log("Pythonホームディレクトリが設定されていません", LogLevel.Warning);
                    return;
                }
            }

            // 初期化UIを表示
            ShowInitializeUI();

            // 初期化実行
            try
            {
                _audioSrInstance = new AudioSrWrapper(PythonHome);
                await _audioSrInstance.InitializeAsync((step, totalSteps, message) =>
                {
                    UpdateInitializeProgress(step, totalSteps, message);
                });
                Logger.Log("AudioSRの初期化が完了しました", LogLevel.Info);
            }
            catch (Exception ex)
            {
                Logger.Log($"初期化エラー: {ex.Message}", LogLevel.Info);
                Logger.LogException("AudioSR初期化中にエラーが発生しました", ex);
                _audioSrInstance = null;
                throw; // 外側の catch で共通の例外処理（MessageBox表示など）を行う
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"初期化中に致命的なエラーが発生しました: {ex.Message}", LogLevel.Error);
            MessageBox.Show("初期化に失敗しました。詳細はログを確認してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            // 初期化UIを非表示
            HideInitializeUI();

            lock (_initializationLock)
            {
                _initializationInProgress = false;
            }
        }
    }

    /// <summary>
    /// 初期化UIを表示
    /// </summary>
    private void ShowInitializeUI()
    {
        _syncContext.Post(_ =>
        {
            if (FindName("InitializeOverlayPanel") is Grid overlayPanel)
            {
                overlayPanel.Visibility = Visibility.Visible;
            }
        }, null);
    }

    /// <summary>
    /// 初期化UIを非表示
    /// </summary>
    private void HideInitializeUI()
    {
        _syncContext.Post(_ =>
        {
            if (FindName("InitializeOverlayPanel") is Grid overlayPanel)
            {
                overlayPanel.Visibility = Visibility.Collapsed;
            }
        }, null);
    }

    /// <summary>
    /// 初期化進捗を更新
    /// </summary>
    private void UpdateInitializeProgress(int step, int totalSteps, string message)
    {
        _syncContext.Post(_ =>
        {
            if (FindName("InitializeProgressBar") is System.Windows.Controls.ProgressBar progressBar && FindName("InitializeStatusText") is TextBlock statusText && FindName("InitializeProgressText") is TextBlock progressText)
            {
                progressBar.Value = (double)step / totalSteps * 100;
                statusText.Text = message;
                progressText.Text = $"{step}/{totalSteps}";
            }
        }, null);
    }

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
        SaveSettingsWithNotification("手動保存", logSuccess: true);
    }

    private async void RetryEmbeddedPython_Click(object sender, RoutedEventArgs e)
    {
        Logger.Log("埋め込みPythonの再試行を開始します。", LogLevel.Info);
        var pythonReady = await FindEmbeddedPythonAsync();
        if (pythonReady && !string.IsNullOrEmpty(PythonHome))
        {
            MessageBox.Show(
                $"埋め込みPythonの準備が完了しました: {PythonHome}",
                "再試行完了",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }

    private void ShowEmbeddedPythonHelp_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            BuildEmbeddedPythonGuidance(),
            "Pythonランタイムのヘルプ",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    /// <summary>
    /// グローバル AudioSrWrapper インスタンス
    /// </summary>
    private AudioSrWrapper? _audioSrInstance = null;

    /// <summary>
    /// 初期化処理中フラグ
    /// </summary>
    private bool _initializationInProgress = false;

    /// <summary>
    /// 初期化ロック（複数スレッドからの同時初期化を防止）
    /// </summary>
    private static readonly object _initializationLock = new object();

    /// <summary>
    /// 処理開始ボタンのクリックイベント
    /// </summary>
    private async void StartProcessing_Click(object sender, RoutedEventArgs e)
    {
        SaveSettingsWithNotification("処理開始時の自動保存", logSuccess: false);

        if (_fileList.Count == 0)
        {
            Logger.Log("処理するファイルがありません。ファイルを追加してください。", LogLevel.Info);
            return;
        }

        // 既に実行中の場合は何もしない
        if (_cancellationTokenSource != null)
        {
            Logger.Log("既に処理を実行中です。", LogLevel.Info);
            return;
        }

        // AudioSRが未初期化の場合は初期化実行
        if (_audioSrInstance == null)
        {
            lock (_initializationLock)
            {
                // ダブルチェック：ロック内で再度確認
                if (_audioSrInstance != null)
                {
                    // 別のスレッドが初期化を完了したので処理を続行
                }
                else if (_initializationInProgress)
                {
                    Logger.Log("AudioSRを初期化中です。少しお待ちください...", LogLevel.Info);
                    return;
                }
                else
                {
                    // 初期化を開始
                    _initializationInProgress = true;
                }
            }

            if (_audioSrInstance == null && !_initializationInProgress)
            {
                return; // 別スレッドが初期化中
            }

            if (_audioSrInstance == null)
            {
                // 初期化UIを表示
                ShowInitializeUI();

                // 初期化実行
                try
                {
                    _audioSrInstance = new AudioSrWrapper(PythonHome);
                    await _audioSrInstance.InitializeAsync((step, totalSteps, message) =>
                    {
                        UpdateInitializeProgress(step, totalSteps, message);
                    });
                    Logger.Log("AudioSRの初期化が完了しました", LogLevel.Info);
                }
                catch (Exception ex)
                {
                    Logger.Log($"初期化エラー: {ex.Message}", LogLevel.Info);
                    Logger.LogException("AudioSR初期化中にエラーが発生しました", ex);
                    _audioSrInstance = null;
                }
                finally
                {
                    lock (_initializationLock)
                    {
                        _initializationInProgress = false;
                    }
                    // 初期化UIを非表示
                    HideInitializeUI();
                }
            }
        }

        // AudioSRが初期化されたことを確認
        if (_audioSrInstance == null)
        {
            Logger.Log("AudioSRの初期化に失敗しました。", LogLevel.Info);
            return;
        }

        // 処理の開始
        _cancellationTokenSource = new CancellationTokenSource();
        StartAudioProcessing(_audioSrInstance, _cancellationTokenSource.Token);
    }

    /// <summary>
    /// 処理停止ボタンのクリックイベント
    /// </summary>
    private void StopProcessing_Click(object sender, RoutedEventArgs e)
    {
        if (_cancellationTokenSource != null)
        {
            _cancellationTokenSource.Cancel();
            Logger.Log("処理を中断しています...", LogLevel.Info);
        }
    }

    /// <summary>
    /// アプリ終了時のイベントハンドラ
    /// </summary>
    private void Window_Closing(object sender, CancelEventArgs e)
    {
        SaveSettingsWithNotification("アプリ終了時の自動保存", logSuccess: false);

        // AudioSrWrapper の終了処理を呼び出す
        _audioSrInstance?.Dispose();
    }

    #endregion

    #region ヘルパーメソッド

    private void ApplyModelSelection()
    {
        if (cmbModel.SelectedIndex == 0)
        {
            ModelName = "basic";
        }
        else if (cmbModel.SelectedIndex == 1)
        {
            ModelName = "speech";
        }
    }

    private bool SaveSettingsWithNotification(string reason, bool logSuccess, bool updateModelSelection = true)
    {
        if (updateModelSelection)
        {
            ApplyModelSelection();
        }

        if (_settings.Save())
        {
            if (logSuccess)
            {
                Logger.Log("設定を保存しました。", LogLevel.Info);
            }

            return true;
        }

        Logger.Log($"警告: 設定の保存に失敗しました（{reason}）。", LogLevel.Info);
        MessageBox.Show(
            "設定の保存に失敗しました。ディスクの空き容量や書き込み権限を確認してください。",
            "設定保存エラー",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
        return false;
    }

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
                Logger.Log($"警告: {Path.GetFileName(filePath)} は音声ファイルではない可能性があります。", LogLevel.Info);
            }

            // ファイルが存在しない場合もスキップ
            if (!fileItem.Exists())
            {
                Logger.Log($"エラー: ファイルが存在しません: {filePath}", LogLevel.Info);
                skippedCount++;
                continue;
            }

            _fileList.Add(fileItem);
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
    private async void StartAudioProcessing(AudioSrWrapper? audioSr, CancellationToken cancellationToken)
    {
        Logger.Log("処理を開始します...", LogLevel.Info);
        Progress = 0;

        // 処理中の状態表示用
        UpdateStatus("処理中...");

        try
        {
            if (audioSr == null)
            {
                Logger.Log("エラー: AudioSR ラッパーが初期化されていません。", LogLevel.Info);
                return;
            }

            Logger.Log("AudioSRを初期化しました。", LogLevel.Info);

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
                    Logger.Log(statusMsg, LogLevel.Info);

                    // 出力ファイルパスを生成
                    var outputFile = GetAvailableOutputFilePath(Path.Combine(OutputFolder, fileName));

                    // AudioSRで処理（非同期で実行）
                    await audioSr!.ProcessFileAsync(
                        fileItem.Path,
                        outputFile,
                        ModelName,
                        DdimSteps,
                        GuidanceScale,
                        UseRandomSeed ? null : Seed,
                        (currentStep, totalSteps) =>
                        {
                            var percent = totalSteps > 0 ? (double)currentStep / totalSteps * 100 : 0;
                            _syncContext.Post(_ => { Progress = percent; }, null);
                        }
                    );

                    // 処理完了のステータスを更新（UIスレッドで実行）
                    _syncContext.Post(_ => { fileItem.Status = "完了"; }, null);
                    Logger.Log($"{fileName} の処理が完了しました。", LogLevel.Info);
                }
                catch (Exception ex)
                {
                    // エラー時のステータスを更新（UIスレッドで実行）
                    _syncContext.Post(_ => { fileItem.Status = "エラー"; }, null);
                    Logger.Log($"{Path.GetFileName(fileItem.Path)} の処理中にエラーが発生しました: {ex.Message}", LogLevel.Info);
                }

                // 進捗を更新
                processedFiles++;
                var progressPercent = (double)processedFiles / totalFiles * 100;
                Progress = progressPercent;
                UpdateStatus($"処理進捗: {processedFiles}/{totalFiles} ({progressPercent:F1}%)");
            }

            UpdateStatus("処理完了");
            Logger.Log("すべての処理が完了しました。", LogLevel.Info);
        }
        catch (OperationCanceledException)
        {
            UpdateStatus("処理が中断されました");
            Logger.Log("処理が中断されました。", LogLevel.Info);
        }
        catch (Exception ex)
        {
            UpdateStatus("エラーが発生しました");
            Logger.Log($"エラーが発生しました: {ex.Message}", LogLevel.Info);
            var userMessage = "処理中にエラーが発生しました。Python設定や入出力フォルダを確認のうえ、再試行してください。";
            var invalidOperation = ex as InvalidOperationException ?? ex.InnerException as InvalidOperationException;
            if (invalidOperation != null && !string.IsNullOrWhiteSpace(invalidOperation.Message))
            {
                userMessage = invalidOperation.Message;
            }
            MessageBox.Show(
                userMessage,
                "処理エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
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

        return long.TryParse(text, NumberStyles.Integer, cultureInfo, out var seed) && seed >= 0
            ? ValidationResult.ValidResult
            : new ValidationResult(false, "0〜9223372036854775807の整数を入力してください。");
    }
}

