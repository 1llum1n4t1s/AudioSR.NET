using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

using static AudioSR.NET.Logger;

namespace AudioSR.NET;

/// <summary>
/// AudioSRラッパークラス
/// </summary>
public partial class AudioSrWrapper : IDisposable
{
    /// <summary>
    /// 組み込みPythonのホームディレクトリパス
    /// </summary>
    private readonly string _pythonHome;

    /// <summary>
    /// Python実行ファイルのパス
    /// </summary>
    private readonly string _pythonExePath;

    /// <summary>
    /// ワーカースクリプトのパス
    /// </summary>
    private readonly string _workerScriptPath;

    /// <summary>
    /// Pythonワーカープロセス
    /// </summary>
    private Process? _workerProcess;

    /// <summary>
    /// ワーカープロセスへの入力ストリーム
    /// </summary>
    private StreamWriter? _workerInput;

    /// <summary>
    /// ワーカープロセスからの出力ストリーム
    /// </summary>
    private StreamReader? _workerOutput;

    /// <summary>
    /// 初期化済みフラグ
    /// </summary>
    private bool _initialized;

    /// <summary>
    /// 初期化実行中フラグ
    /// </summary>
    private bool _initializationInProgress;

    /// <summary>
    /// 初期化失敗フラグ
    /// </summary>
    private bool _initializationFailed;

    /// <summary>
    /// リソース破棄済みフラグ
    /// </summary>
    private bool _disposed;

    /// <summary>
    /// 現在使用中のデバイス（cuda, vulkan, cpu）
    /// </summary>
    private string? _currentDevice;

    /// <summary>
    /// プロセス通信用のセマフォ
    /// </summary>
    private readonly SemaphoreSlim _processSemaphore = new(1, 1);

    /// <summary>
    /// 依存がインストール済みかどうかを示すマーカーファイルのパス
    /// </summary>
    private string _depsInstalledMarkerFile => Path.Combine(_pythonHome, ".audiosr_deps_installed");

    /// <summary>
    /// 現在実行中の処理の進捗報告用コールバック
    /// </summary>
    private Action<int, int>? _currentProcessingProgress;

    /// <summary>
    /// tqdm の進捗出力をパースするための正規表現
    /// サンプルステップ: 10%|██        | 10/100 [00:01<00:09, 85.23it/s]
    /// ダウンロード: 1.23GB/4.56GB [00:05<00:15, 123MB/s]
    /// </summary>
    [System.Text.RegularExpressions.GeneratedRegex(@"(?:(?<percent>\d+)%\|.*?\|\s*)?(?<current>[\d\.]+)(?<unit>[kMG]B)?/(?<total>[\d\.]+)(?<total_unit>[kMG]B)?|(?<current_only>\d+)it \[")]
    private static partial System.Text.RegularExpressions.Regex TqdmProgressRegex();

    /// <summary>
    /// AudioSrWrapperの新しいインスタンスを初期化します
    /// </summary>
    /// <param name="pythonHome">組み込みPythonのパス</param>
    public AudioSrWrapper(string pythonHome)
    {
        // コンストラクタ呼び出しをログに記録
        var message = $"AudioSrWrapper constructor called with pythonHome: {pythonHome}";
        Log(message, LogLevel.Debug);

        // パスのバリデーション
        ArgumentException.ThrowIfNullOrEmpty(pythonHome);

        if (!Directory.Exists(pythonHome))
        {
            throw new DirectoryNotFoundException($"Python home directory not found: {pythonHome}");
        }

        // フィールドを初期化
        _pythonHome = pythonHome;
        _pythonExePath = Path.Combine(_pythonHome, "python.exe");
        if (!File.Exists(_pythonExePath))
        {
            throw new FileNotFoundException("python.exe が見つかりません。", _pythonExePath);
        }

        // ワーカースクリプトのパスを解決
        _workerScriptPath = ResolveWorkerScriptPath();
        _initialized = false;
        _initializationInProgress = false;
        _initializationFailed = false;
    }

    /// <summary>
    /// ワーカースクリプトのパスを解決します
    /// </summary>
    /// <returns>解決されたスクリプトのフルパス</returns>
    private static string ResolveWorkerScriptPath()
    {
        // アプリケーションのベースディレクトリを取得
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        
        // candidate 1: python/audiosr_worker.py
        var candidate = Path.Combine(baseDir, "python", "audiosr_worker.py");
        if (File.Exists(candidate))
        {
            return candidate;
        }

        // candidate 2: audiosr_worker.py (fallback)
        var fallback = Path.Combine(baseDir, "audiosr_worker.py");
        if (File.Exists(fallback))
        {
            return fallback;
        }

        // いずれも見つからない場合は例外
        throw new FileNotFoundException("AudioSR ワーカースクリプトが見つかりません。", candidate);
    }

        /// <summary>
        /// GPUに応じた最適なデバイス文字列を取得します
        /// </summary>
        /// <returns>デバイス文字列（cuda, vulkan, cpu）</returns>
        private static string DetectOptimalDevice()
        {
            try
            {
                Log("GPU検出を開始します...", LogLevel.Debug);

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController"))
                {
                    List<string> gpuNames = [];
                    foreach (var obj in searcher.Get())
                    {
                        var name = obj["Name"]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(name))
                        {
                            gpuNames.Add(name.ToLower());
                            Log($"検出されたGPU: {name}", LogLevel.Debug);
                        }
                        obj?.Dispose();
                    }

                    // GeForce搭載 -> CUDA
                    if (gpuNames.Any(g => g.Contains("geforce") || g.Contains("nvidia")))
                    {
                        Log("✓ NVIDIA GeForce搭載。CUDA を使用します。", LogLevel.Info);
                        return "cuda";
                    }

                    // AMD/Intel 等の場合は、標準の torch では GPU 支援が限定的なため CPU を推奨
                    // (torch-directml 等を入れれば directml デバイスが使えるが、現在は標準 torch のみ)
                    if (gpuNames.Any(g => g.Contains("radeon") || g.Contains("amd") || g.Contains("intel")))
                    {
                        Log("✓ RADEON/Intel搭載。安全のため CPU モードを使用します。", LogLevel.Info);
                        return "cpu";
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GPU検出中にエラーが発生しました: {ex.Message}", LogLevel.Warning);
                Log("CPUモードにフォールバックします。", LogLevel.Warning);
            }

            Log("GPU検出失敗。CPU を使用します。", LogLevel.Info);
            return "cpu";
        }

        /// <summary>
        /// Pythonの.pthファイルに import site を有効化します
        /// </summary>
        private void EnsurePythonPthConfiguration()
        {
            var pythonPrefix = GetPythonPrefix(AppSettings.Load().PythonVersion);
            var pthFilePath = ResolvePythonPthFilePath(_pythonHome, pythonPrefix);
            if (!File.Exists(pthFilePath))
            {
                return;
            }

            var content = File.ReadAllText(pthFilePath);
            if (!content.Contains("import site") || content.Contains("#import site"))
            {
                Log("Python._pthに 'import site' を追加します", LogLevel.Debug);
                content = content.Replace("#import site", "import site");
                if (!content.Contains("import site"))
                {
                    content += "\r\nimport site";
                }
            }

            var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
            if (!content.Contains("site-packages", StringComparison.OrdinalIgnoreCase))
            {
                content += $"\r\n{sitePackagesPath}";
            }

            File.WriteAllText(pthFilePath, content);
        }

        /// <summary>
        /// audiosrパッケージを埋め込みPythonのsite-packagesにのみインストールします
        /// </summary>
        /// <param name="onProgress">インストール進捗を報告するコールバック（段階番号, 総段階数, メッセージ）</param>
        private async Task EnsureAudioSrInstalledAsync(Action<int, int, string>? onProgress = null)
        {
            // site-packages ディレクトリのパスを取得
            var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
            try
            {
                // 1/10: site-packages ディレクトリを確認
                onProgress?.Invoke(1, 10, "site-packages ディレクトリを確認中...");
                Directory.CreateDirectory(sitePackagesPath);

                // 2/10: インストール済みマーカーを確認
                onProgress?.Invoke(2, 10, "インストール状態を確認中...");
                const string markerContent = "installed_v5"; // torchcodec 削除 & パッチ対応のため v5 に更新
                if (File.Exists(_depsInstalledMarkerFile) && File.ReadAllText(_depsInstalledMarkerFile) == markerContent)
                {
                    Log("✓ 依存インストール済みマーカーファイルが存在します。初期化をスキップします。", LogLevel.Info);
                    onProgress?.Invoke(10, 10, "インストール済み（スキップ）");
                    return;
                }

                // 3/10: pip の可用性を確認
                onProgress?.Invoke(3, 10, "pip を確認中...");
                if (!await IsPipAvailableAsync(sitePackagesPath))
                {
                    // 4-5/10: pip のインストール
                    await InstallPipAsync(sitePackagesPath, onProgress);
                }

                // 6/10: ビルド用の基本ツールをインストール
                onProgress?.Invoke(6, 10, "ビルドツールをインストール中...");
                await InstallPackagesAsync(sitePackagesPath, ["setuptools", "wheel"]);

                // 7/10: AudioSR と依存パッケージをインストール
                onProgress?.Invoke(7, 10, "AudioSR をインストール中...");
                await InstallPackagesAsync(sitePackagesPath, ["torch", "torchaudio", "audiosr", "matplotlib", "soundfile"]);

                // 10/10: 完了マーカーを作成
                File.WriteAllText(_depsInstalledMarkerFile, markerContent);
                onProgress?.Invoke(10, 10, "インストール完了");
                Log("依存パッケージのインストールが完了しました。", LogLevel.Info);
            }
            catch (Exception ex)
            {
                // エラーログを記録して再送
                Log($"EnsureAudioSrInstalled でエラーが発生しました: {ex.Message}", LogLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// pip が利用可能か確認します
        /// </summary>
        /// <param name="sitePackagesPath">確認対象の site-packages パス</param>
        /// <returns>利用可能な場合は true</returns>
        private async Task<bool> IsPipAvailableAsync(string sitePackagesPath)
        {
            // -m pip --version を実行して確認
            var result = await RunPythonCommandAsync(sitePackagesPath, "-m pip --version", null);
            return result.ExitCode == 0;
        }

        /// <summary>
        /// pip をインストールします
        /// </summary>
        /// <param name="sitePackagesPath">インストール先の site-packages パス</param>
        /// <param name="onProgress">進捗報告用のコールバック</param>
        private async Task InstallPipAsync(string sitePackagesPath, Action<int, int, string>? onProgress)
        {
            // 進捗を報告
            onProgress?.Invoke(4, 10, "pip をインストール中...");
            // ログを記録
            Log("pip をインストールします...", LogLevel.Info);

            // get-pip.py の保存パスを決定
            var getPipPath = Path.Combine(_pythonHome, "get-pip.py");
            
            // get-pip.py をダウンロード
            // 注意: 埋め込みPython環境では、SSL証明書の問題を避けるため
            // 信頼されたホストからのダウンロードを明示的に指定する必要がある場合があります
            DownloadFile("https://bootstrap.pypa.io/get-pip.py", getPipPath);

            // pip インストール用の引数を構築
            // --disable-pip-version-check: pip自身の更新チェックを無効化
            // --no-warn-script-location: スクリプトパスの警告を抑制
            // -v: 詳細なログを出力
            // --target: インストール先のディレクトリを明示的に指定
            var args = $"\"{getPipPath}\" --disable-pip-version-check --no-warn-script-location -v --target \"{sitePackagesPath}\"";
            
            // Python コマンドを実行
            var result = await RunPythonCommandAsync(sitePackagesPath, args, null);
            
            // 終了コードが 0 でない場合は例外を投げる
            if (result.ExitCode != 0)
            {
                // インストール失敗時に詳細なエラー情報をログに記録
                Log("=== pip インストール失敗詳細 ===", LogLevel.Error);
                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                {
                    Log($"STDOUT:\n{result.StandardOutput}", LogLevel.Error);
                }
                if (!string.IsNullOrWhiteSpace(result.StandardError))
                {
                    Log($"STDERR:\n{result.StandardError}", LogLevel.Error);
                }
                throw new InvalidOperationException($"pip のインストールに失敗しました。終了コード: {result.ExitCode}");
            }

            // 一時ファイルを削除
            if (File.Exists(getPipPath))
            {
                File.Delete(getPipPath);
            }
            // 完了を報告
            onProgress?.Invoke(5, 10, "pip をインストールしました");
        }

        /// <summary>
        /// 依存パッケージをインストールします
        /// </summary>
        /// <param name="sitePackagesPath">インストール先の site-packages パス</param>
        /// <param name="packages">インストールするパッケージのリスト</param>
        private async Task InstallPackagesAsync(string sitePackagesPath, IEnumerable<string> packages)
        {
            // パッケージリストをスペース区切りの文字列に変換
            var packageList = string.Join(" ", packages);
            
            // pip install 用の引数を構築
            // --upgrade: 既存パッケージがある場合はアップグレード
            // --no-warn-script-location: スクリプトパスの警告を抑制
            // --target: 埋め込みPython環境用のディレクトリを指定
            // --no-cache-dir: キャッシュを使用しない（クリーンな環境構築のため）
            var args = $"-m pip install --upgrade -v --no-warn-script-location --no-cache-dir --target \"{sitePackagesPath}\" {packageList}";
            
            // Python コマンドを実行
            var result = await RunPythonCommandAsync(sitePackagesPath, args, null);
            
            // 終了コードが 0 でない場合は例外を投げる
            if (result.ExitCode != 0)
            {
                // インストール失敗時に詳細なエラー情報をログに記録
                Log($"=== パッケージインストール失敗詳細 ({packageList}) ===", LogLevel.Error);
                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                {
                    Log($"STDOUT:\n{result.StandardOutput}", LogLevel.Error);
                }
                if (!string.IsNullOrWhiteSpace(result.StandardError))
                {
                    Log($"STDERR:\n{result.StandardError}", LogLevel.Error);
                }
                throw new InvalidOperationException($"依存パッケージのインストールに失敗しました。終了コード: {result.ExitCode}");
            }
            else
            {
                // 成功時も、確認のために最後の 500 文字をログに記録
                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                {
                    var output = result.StandardOutput.Trim();
                    var lastPart = output.Length > 500 ? output[^500..] : output;
                    Log($"pip インストール成功 (出力の末尾): ...{lastPart}", LogLevel.Info);
                }
            }
        }

        /// <summary>
        /// 指定された URL からファイルをダウンロードします
        /// </summary>
        /// <param name="url">ダウンロード元の URL</param>
        /// <param name="destination">保存先のパス</param>
        private static void DownloadFile(string url, string destination)
        {
            using var client = new HttpClient();
            var bytes = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
            File.WriteAllBytes(destination, bytes);
        }

        /// <summary>
        /// Python コマンドを実行し、標準出力を取得します
        /// </summary>
        /// <param name="sitePackagesPath">パッケージパス</param>
        /// <param name="arguments">引数</param>
        /// <param name="stdin">標準入力に送る文字列（nullの場合は指定なし）</param>
        /// <returns>終了コードと標準出力、標準エラー</returns>
        private async Task<(int ExitCode, string StandardOutput, string StandardError)> RunPythonCommandAsync(
            string sitePackagesPath,
            string arguments,
            string? stdin)
        {
            // 起動情報を設定
            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonExePath,
                Arguments = arguments,
                RedirectStandardInput = stdin != null,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            // 環境変数を構成
            ConfigurePythonEnvironment(startInfo, sitePackagesPath);

            // プロセスを開始
            using var process = new Process { StartInfo = startInfo };
            process.Start();

            // 標準入力があれば書き込み
            if (stdin != null)
            {
                await process.StandardInput.WriteAsync(stdin);
                process.StandardInput.Close();
            }

            // 出力を非同期で読み取り
            var stdOutTask = process.StandardOutput.ReadToEndAsync();
            var stdErrTask = process.StandardError.ReadToEndAsync();

            // 完了を待機
            await Task.WhenAll(stdOutTask, stdErrTask);
            await process.WaitForExitAsync();

            var stdOut = stdOutTask.Result;
            var stdErr = stdErrTask.Result;

            // ログに出力
            if (!string.IsNullOrWhiteSpace(stdOut))
            {
                Log(stdOut.Trim(), LogLevel.Debug);
            }

            if (!string.IsNullOrWhiteSpace(stdErr))
            {
                Log(stdErr.Trim(), LogLevel.Warning);
            }

            return (process.ExitCode, stdOut, stdErr);
        }

        /// <summary>
        /// Pythonの実行環境変数を設定します
        /// </summary>
        /// <param name="startInfo">プロセスの起動情報</param>
        /// <param name="sitePackagesPath">追加のパッケージディレクトリパス</param>
        private void ConfigurePythonEnvironment(ProcessStartInfo startInfo, string sitePackagesPath)
        {
            // Lib ディレクトリのパスを取得
            var pythonLibPath = Path.Combine(_pythonHome, "Lib");
            
            // 存在するディレクトリのみを抽出して PYTHONPATH を構築
            string[] searchPaths = [sitePackagesPath, pythonLibPath];
            var pythonPath = string.Join(";", searchPaths.Where(Directory.Exists));
            
            // 環境変数を設定
            startInfo.Environment["PYTHONHOME"] = _pythonHome;
            startInfo.Environment["PYTHONPATH"] = pythonPath;
            startInfo.Environment["PYTHONUTF8"] = "1";

            // PATH に PythonHome を追加（DLLのロードなどのため）
            var originalPath = startInfo.Environment["PATH"] ?? Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            if (!originalPath.Contains(_pythonHome, StringComparison.OrdinalIgnoreCase))
            {
                startInfo.Environment["PATH"] = $"{_pythonHome};{originalPath}";
            }
        }

        private async Task StartWorkerProcessAsync(string sitePackagesPath)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonExePath,
                Arguments = $"-u \"{_workerScriptPath}\"",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            ConfigurePythonEnvironment(startInfo, sitePackagesPath);

            _workerProcess = new Process { StartInfo = startInfo };
            _workerProcess.Start();

            _workerInput = _workerProcess.StandardInput;
            _workerOutput = _workerProcess.StandardOutput;

            _ = Task.Run(async () =>
            {
                try
                {
                    while (_workerProcess != null && !_workerProcess.HasExited)
                    {
                        var line = await _workerProcess.StandardError.ReadLineAsync();
                        if (line == null)
                        {
                            break;
                        }
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            // tqdm の進捗行かチェック
                            var match = TqdmProgressRegex().Match(line);
                            if (match.Success)
                            {
                                if (match.Groups["current"].Success && match.Groups["total"].Success)
                                {
                                    // 数値としてパース（単位は無視して比率として扱う）
                                    if (double.TryParse(match.Groups["current"].Value, out var current) &&
                                        double.TryParse(match.Groups["total"].Value, out var total))
                                    {
                                        // 進捗を現在の処理タスクに報告
                                        // 整数に丸めて通知
                                        _currentProcessingProgress?.Invoke((int)current, (int)total);
                                    }
                                }
                                else if (match.Groups["current_only"].Success)
                                {
                                    if (int.TryParse(match.Groups["current_only"].Value, out var current))
                                    {
                                        // 合計が不明な場合は 0 または直近の合計を使用（ここでは単に報告）
                                        // 合計がわからない場合は UI 側で 0% などの扱いになる可能性があるが、
                                        // 基本的には super_resolution は DDIM steps を合計とするはず
                                        _currentProcessingProgress?.Invoke(current, 0);
                                    }
                                }
                            }
                            else
                            {
                                // 進捗以外は通常ログに出力
                                Log($"Python: {line.Trim()}", LogLevel.Debug);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Python stderr 監視でエラーが発生しました: {ex.Message}", LogLevel.Debug);
                }
            });

            var pingResponse = await SendCommandAsync(new Dictionary<string, object?> { { "command", "ping" } }, TimeSpan.FromSeconds(30));
            if (!string.Equals(pingResponse.Status, "ok", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Python ワーカーの起動に失敗しました: {pingResponse.Message}");
            }
        }

        /// <summary>
        /// ワーカープロセスにコマンドを送信し、応答を待ちます
        /// </summary>
        /// <param name="payload">送信するペイロード</param>
        /// <param name="timeout">タイムアウト時間（nullの場合はデフォルト10分）</param>
        /// <param name="ct">キャンセル・トークン</param>
        /// <returns>ワーカープロセスからの応答</returns>
        private async Task<WorkerResponse> SendCommandAsync(Dictionary<string, object?> payload, TimeSpan? timeout = null, CancellationToken ct = default)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(AudioSrWrapper));
            }

            if (_workerInput == null || _workerOutput == null)
            {
                throw new InvalidOperationException("Python ワーカーが初期化されていません。");
            }

            var json = JsonSerializer.Serialize(payload);
            
            try
            {
                await _processSemaphore.WaitAsync(ct);
            }
            catch (ObjectDisposedException)
            {
                throw new TaskCanceledException("オブジェクト破棄のため通信がキャンセルされました。");
            }

            try
            {
                await _workerInput.WriteLineAsync(json);
                await _workerInput.FlushAsync();

                // タイムアウト付きで1行読み取り
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                cts.CancelAfter(timeout ?? TimeSpan.FromMinutes(10));

                WorkerResponse? response = null;
                while (!cts.IsCancellationRequested)
                {
                    var responseLine = await _workerOutput.ReadLineAsync(cts.Token);
                    if (string.IsNullOrWhiteSpace(responseLine))
                    {
                        throw new InvalidOperationException("Python ワーカーから応答がありません。");
                    }

                    // 行が JSON かどうか試行
                    try
                    {
                        response = JsonSerializer.Deserialize<WorkerResponse>(
                            responseLine,
                            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                        
                        if (response != null && (response.Status == "ok" || response.Status == "error"))
                        {
                            // 有効なレスポンスを受信
                            break;
                        }
                    }
                    catch (JsonException)
                    {
                        // JSON 以外の行（想定外の出力）はログに記録して次を待つ
                        Log($"Python stdout (非JSON): {responseLine}", LogLevel.Warning);
                    }
                }

                if (response == null)
                {
                    throw new InvalidOperationException("Python ワーカーから有効な応答を受信できませんでした。");
                }

                return response;
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException("Python ワーカーとの通信がタイムアウトしました。");
            }
            finally
            {
                try
                {
                    _processSemaphore.Release();
                }
                catch (ObjectDisposedException) { /* ignore */ }
            }
        }

        /// <summary>
        /// Python環境を初期化し、AudioSRをロードします
        /// </summary>
        /// <param name="onProgress">インストール進捗を報告するコールバック（段階番号, 総段階数, メッセージ）</param>
        public async Task InitializeAsync(Action<int, int, string>? onProgress = null)
        {
            if (_initialized)
            {
                Log("Python environment already initialized", LogLevel.Debug);
                return;
            }

            if (_initializationInProgress)
            {
                Log("初期化は既に実行中です。重複呼び出しを防止します。", LogLevel.Warning);
                return;
            }

            _initializationInProgress = true;
            var startTime = DateTime.Now;
            Log($"[{startTime:yyyy-MM-dd HH:mm:ss.fff}] Python環境を初期化します...", LogLevel.Debug);
            _initializationFailed = false;
            var initializationSucceeded = false;

            try
            {
                EnsurePythonPthConfiguration();

                onProgress?.Invoke(1, 10, "Python環境を確認中...");
                await EnsureAudioSrInstalledAsync(onProgress);

                onProgress?.Invoke(9, 10, "Pythonワーカーを起動中...");
                var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
                await StartWorkerProcessAsync(sitePackagesPath);

                initializationSucceeded = true;
                onProgress?.Invoke(10, 10, "初期化完了");
            }
            catch (Exception ex)
            {
                Log($"Python環境の初期化中にエラーが発生しました: {ex.Message}", LogLevel.Error);
                Log(ex.StackTrace ?? "スタックトレースなし", LogLevel.Error);
                throw;
            }
            finally
            {
                _initialized = initializationSucceeded;
                _initializationFailed = !initializationSucceeded;
                _initializationInProgress = false;
                var finalMsg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 初期化フラグを設定: _initialized = {_initialized}, _initializationFailed = {_initializationFailed}";
                Log(finalMsg, LogLevel.Debug);
            }

            var endTime = DateTime.Now;
            Log($"[{endTime:yyyy-MM-dd HH:mm:ss.fff}] Python environment successfully initialized in {(endTime - startTime).TotalSeconds:F2} seconds", LogLevel.Debug);
        }

        /// <summary>
        /// 単一のオーディオファイルを処理します
        /// </summary>
        /// <param name="inputFile">入力ファイルパス</param>
        /// <param name="outputFile">出力ファイルパス</param>
        /// <param name="modelName">使用するモデル名</param>
        /// <param name="ddimSteps">DDIMステップ数</param>
        /// <param name="guidanceScale">ガイダンススケール</param>
        /// <param name="seed">乱数シード（null=ランダム）</param>
        /// <param name="onProgress">処理進捗を報告するコールバック（現在のステップ, 総ステップ数）</param>
        public async Task ProcessFileAsync(string inputFile, string outputFile, string modelName, int ddimSteps, float guidanceScale, long? seed, Action<int, int>? onProgress = null)
        {
            Log($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ProcessFile 呼び出し: {inputFile} -> {outputFile}", LogLevel.Debug);

            if (!_initialized)
            {
                throw new InvalidOperationException("AudioSrWrapperが初期化されていません。StartProcessing_Click内で初期化を完了させてください。");
            }

            if (!File.Exists(inputFile))
            {
                throw new FileNotFoundException($"ファイルが見つかりません: {inputFile}");
            }

            var outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            _currentDevice ??= DetectOptimalDevice();

            var payload = new Dictionary<string, object?>
            {
                { "command", "process" },
                { "input", inputFile },
                { "output", outputFile },
                { "model_name", modelName },
                { "device", _currentDevice },
                { "ddim_steps", ddimSteps },
                { "guidance_scale", guidanceScale }
            };

            if (seed.HasValue)
            {
                payload["seed"] = seed.Value;
            }

            try
            {
                // 進捗報告用コールバックを設定
                _currentProcessingProgress = onProgress;

                var response = await SendCommandAsync(payload);
                if (!string.Equals(response.Status, "ok", StringComparison.OrdinalIgnoreCase))
                {
                    var detail = string.IsNullOrWhiteSpace(response.Traceback)
                        ? response.Message
                        : $"{response.Message}\n{response.Traceback}";
                    throw new InvalidOperationException($"AudioSR処理に失敗しました: {detail}");
                }
            }
            finally
            {
                // コールバックを解除
                _currentProcessingProgress = null;
            }

            if (!File.Exists(outputFile))
            {
                throw new InvalidOperationException("出力ファイルが生成されませんでした。");
            }

            onProgress?.Invoke(1, 1);
        }

        /// <summary>
        /// バッチファイルに記載された複数のオーディオファイルを処理します
        /// </summary>
        /// <param name="inputListFile">入力リストファイル</param>
        /// <param name="outputPath">出力ディレクトリパス</param>
        /// <param name="modelName">使用するモデル名</param>
        /// <param name="ddimSteps">DDIMステップ数</param>
        /// <param name="guidanceScale">ガイダンススケール</param>
        /// <param name="seed">乱数シード（null=ランダム）</param>
        public async Task ProcessBatchFileAsync(string inputListFile, string outputPath, string modelName, int ddimSteps, float guidanceScale, long? seed)
        {
            if (!File.Exists(inputListFile))
            {
                throw new FileNotFoundException($"入力リストファイルが見つかりません: {inputListFile}");
            }

            Directory.CreateDirectory(outputPath);
            var inputFiles = File.ReadAllLines(inputListFile).Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();
            Log($"Processing {inputFiles.Length} files from batch file: {inputListFile}", LogLevel.Debug);

            for (var i = 0; i < inputFiles.Length; i++)
            {
                var inputFile = inputFiles[i].Trim();
                if (!File.Exists(inputFile))
                {
                    Log($"警告: ファイルが見つかりません: {inputFile} - スキップします", LogLevel.Debug);
                    continue;
                }

                var outputFile = Path.Combine(outputPath, Path.GetFileName(inputFile));
                Log($"処理中 ({i + 1}/{inputFiles.Length}): {inputFile}", LogLevel.Debug);

                try
                {
                    await ProcessFileAsync(inputFile, outputFile, modelName, ddimSteps, guidanceScale, seed);
                }
                catch (Exception ex)
                {
                    Log($"ファイル処理エラー: {inputFile} - {ex.Message}", LogLevel.Debug);
                }
            }

            Log($"バッチ処理完了: {inputFiles.Length} ファイル", LogLevel.Debug);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                try
                {
                    if (_workerProcess != null && !_workerProcess.HasExited)
                    {
                        try
                        {
                            // スレッドセーフにシャットダウンコマンドを送信するため、セマフォを取得します。
                            // Disposeは同期メソッドのためWait()を使用しますが、デッドロックを避けるため
                            // タイムアウト付きで待機することが推奨されます。
                            if (_processSemaphore.Wait(TimeSpan.FromSeconds(1)))
                            {
                                try
                                {
                                    // 終了コマンド送信を試みる
                                    _workerInput?.WriteLine(JsonSerializer.Serialize(new Dictionary<string, object?> { { "command", "shutdown" } }));
                                    _workerInput?.Flush();
                                }
                                finally
                                {
                                    _processSemaphore.Release();
                                }
                            }
                        }
                        catch { /* ignore */ }

                        // 最大3秒待機し、終了しなければ強制終了
                        if (!_workerProcess.WaitForExit(3000))
                        {
                            _workerProcess.Kill();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Python ワーカー終了時にエラーが発生しました: {ex.Message}", LogLevel.Debug);
                }
                finally
                {
                    _workerProcess?.Dispose();
                    _workerProcess = null;
                    _processSemaphore.Dispose();
                    _initialized = false;
                }
            }

            _disposed = true;
        }

    /// <summary>
    /// 指定されたバージョンの Python 接頭辞（例: python311）を取得します
    /// </summary>
    /// <param name="versionText">バージョン文字列</param>
    /// <returns>Python 接頭辞。無効な場合は null</returns>
    private static string? GetPythonPrefix(string? versionText)
    {
        if (string.IsNullOrWhiteSpace(versionText))
        {
            return null;
        }

        if (!Version.TryParse(versionText, out var version))
        {
            return null;
        }

        return $"python{version.Major}{version.Minor}";
    }

    /// <summary>
    /// Python の .pth ファイルのパスを解決します
    /// </summary>
    /// <param name="pythonHome">Python のホームディレクトリ</param>
    /// <param name="pythonPrefix">Python の接頭辞</param>
    /// <returns>解決された .pth ファイルのフルパス</returns>
    private static string ResolvePythonPthFilePath(string pythonHome, string? pythonPrefix)
    {
        // 指定された接頭辞に基づくファイルを優先
        if (!string.IsNullOrEmpty(pythonPrefix))
        {
            var expected = Path.Combine(pythonHome, $"{pythonPrefix}._pth");
            if (File.Exists(expected))
            {
                return expected;
            }
        }

        // 見つからない場合はワイルドカードで検索、またはデフォルトを返す
        return Directory.GetFiles(pythonHome, "python*._pth").FirstOrDefault()
            ?? Path.Combine(pythonHome, "python._pth");
    }

    /// <summary>
    /// ワーカープロセスからの応答を表す内部クラス
    /// </summary>
    private sealed class WorkerResponse
    {
        /// <summary>
        /// ステータス（ok または error）
        /// </summary>
        public string Status { get; set; } = "error";

        /// <summary>
        /// メッセージ
        /// </summary>
        public string Message { get; set; } = "";

        /// <summary>
        /// エラー時のトレースバック
        /// </summary>
        public string? Traceback { get; set; }
    }
}
