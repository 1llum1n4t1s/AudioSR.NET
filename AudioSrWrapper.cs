using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

using static AudioSR.NET.Logger;

namespace AudioSR.NET
{
    /// <summary>
    /// AudioSRラッパークラス
    /// </summary>
    public class AudioSrWrapper : IDisposable
    {
        private readonly string _pythonHome;
        private readonly string _pythonExePath;
        private readonly string _workerScriptPath;
        private Process? _workerProcess;
        private StreamWriter? _workerInput;
        private StreamReader? _workerOutput;
        private bool _initialized;
        private bool _initializationInProgress;
        private bool _initializationFailed;
        private bool _disposed;
        private string? _currentDevice;
        private readonly object _processLock = new();

        /// <summary>
        /// 依存がインストール済みかどうかを示すマーカーファイルのパス
        /// </summary>
        private string _depsInstalledMarkerFile => Path.Combine(_pythonHome, ".audiosr_deps_installed");

        /// <summary>
        /// AudioSrWrapperの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="pythonHome">組み込みPythonのパス</param>
        public AudioSrWrapper(string pythonHome)
        {
            var message = $"AudioSrWrapper constructor called with pythonHome: {pythonHome}";
            Log(message, LogLevel.Debug);

            if (string.IsNullOrEmpty(pythonHome))
            {
                throw new ArgumentException("pythonHome cannot be null or empty", nameof(pythonHome));
            }

            if (!Directory.Exists(pythonHome))
            {
                throw new DirectoryNotFoundException($"Python home directory not found: {pythonHome}");
            }

            _pythonHome = pythonHome;
            _pythonExePath = Path.Combine(_pythonHome, "python.exe");
            if (!File.Exists(_pythonExePath))
            {
                throw new FileNotFoundException("python.exe が見つかりません。", _pythonExePath);
            }

            _workerScriptPath = ResolveWorkerScriptPath();
            _initialized = false;
            _initializationInProgress = false;
            _initializationFailed = false;
        }

        private static string ResolveWorkerScriptPath()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidate = Path.Combine(baseDir, "python", "audiosr_worker.py");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            var fallback = Path.Combine(baseDir, "audiosr_worker.py");
            if (File.Exists(fallback))
            {
                return fallback;
            }

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
                    var gpuNames = new List<string>();
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

                    // RADEON または Intel -> Vulkan
                    if (gpuNames.Any(g => g.Contains("radeon") || g.Contains("amd") || g.Contains("intel")))
                    {
                        Log("✓ RADEON/Intel搭載。Vulkan を使用します。", LogLevel.Info);
                        return "vulkan";
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
        private void EnsureAudioSrInstalled(Action<int, int, string>? onProgress = null)
        {
            var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
            try
            {
                onProgress?.Invoke(1, 10, "site-packages ディレクトリを確認中...");
                Directory.CreateDirectory(sitePackagesPath);

                onProgress?.Invoke(2, 10, "インストール状態を確認中...");
                if (File.Exists(_depsInstalledMarkerFile))
                {
                    Log("✓ 依存インストール済みマーカーファイルが存在します。初期化をスキップします。", LogLevel.Info);
                    onProgress?.Invoke(10, 10, "インストール済み（スキップ）");
                    return;
                }

                onProgress?.Invoke(3, 10, "pip を確認中...");
                if (!IsPipAvailable(sitePackagesPath))
                {
                    InstallPip(sitePackagesPath, onProgress);
                }

                onProgress?.Invoke(6, 10, "パッケージをインストール中...");
                InstallPackages(sitePackagesPath, new[] { "torch", "torchaudio", "audiosr" });

                File.WriteAllText(_depsInstalledMarkerFile, "installed");
                onProgress?.Invoke(10, 10, "インストール完了");
                Log("依存パッケージのインストールが完了しました。", LogLevel.Info);
            }
            catch (Exception ex)
            {
                Log($"EnsureAudioSrInstalled でエラーが発生しました: {ex.Message}", LogLevel.Error);
                throw;
            }
        }

        private bool IsPipAvailable(string sitePackagesPath)
        {
            var result = RunPythonCommand(sitePackagesPath, "-m pip --version", null);
            return result.ExitCode == 0;
        }

        /// <summary>
        /// pip をインストールします
        /// </summary>
        /// <param name="sitePackagesPath">インストール先の site-packages パス</param>
        /// <param name="onProgress">進捗報告用のコールバック</param>
        private void InstallPip(string sitePackagesPath, Action<int, int, string>? onProgress)
        {
            // 進捗を報告
            onProgress?.Invoke(4, 10, "pip をインストール中...");
            // ログを記録
            Log("pip をインストールします...", LogLevel.Info);

            // get-pip.py の保存パスを決定
            var getPipPath = Path.Combine(_pythonHome, "get-pip.py");
            // get-pip.py をダウンロード
            DownloadFile("https://bootstrap.pypa.io/get-pip.py", getPipPath);

            // pip インストール用の引数を構築（詳細出力のため -v を追加）
            var args = $"\"{getPipPath}\" --disable-pip-version-check --no-warn-script-location -v --target \"{sitePackagesPath}\"";
            // Python コマンドを実行
            var result = RunPythonCommand(sitePackagesPath, args, null);
            
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
        private void InstallPackages(string sitePackagesPath, IEnumerable<string> packages)
        {
            // パッケージリストをスペース区切りの文字列に変換
            var packageList = string.Join(" ", packages);
            // pip install 用の引数を構築（--quiet を削除し、詳細出力のため -v を追加）
            var args = $"-m pip install --upgrade -v --no-warn-script-location --target \"{sitePackagesPath}\" {packageList}";
            // Python コマンドを実行
            var result = RunPythonCommand(sitePackagesPath, args, null);
            
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

        private static void DownloadFile(string url, string destination)
        {
            using var client = new HttpClient();
            var bytes = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
            File.WriteAllBytes(destination, bytes);
        }

        private (int ExitCode, string StandardOutput, string StandardError) RunPythonCommand(
            string sitePackagesPath,
            string arguments,
            string? stdin)
        {
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

            ConfigurePythonEnvironment(startInfo, sitePackagesPath);

            using var process = new Process { StartInfo = startInfo };
            process.Start();

            if (stdin != null)
            {
                process.StandardInput.Write(stdin);
                process.StandardInput.Close();
            }

            var stdOut = process.StandardOutput.ReadToEnd();
            var stdErr = process.StandardError.ReadToEnd();
            process.WaitForExit();

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

        private void ConfigurePythonEnvironment(ProcessStartInfo startInfo, string sitePackagesPath)
        {
            var pythonLibPath = Path.Combine(_pythonHome, "Lib");
            var pythonPath = string.Join(";", new[] { sitePackagesPath, pythonLibPath }.Where(Directory.Exists));
            startInfo.Environment["PYTHONHOME"] = _pythonHome;
            startInfo.Environment["PYTHONPATH"] = pythonPath;
            startInfo.Environment["PYTHONUTF8"] = "1";

            var originalPath = startInfo.Environment["PATH"] ?? Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            if (!originalPath.Contains(_pythonHome, StringComparison.OrdinalIgnoreCase))
            {
                startInfo.Environment["PATH"] = $"{_pythonHome};{originalPath}";
            }
        }

        private void StartWorkerProcess(string sitePackagesPath)
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

            _ = Task.Run(() =>
            {
                try
                {
                    while (_workerProcess != null && !_workerProcess.HasExited)
                    {
                        var line = _workerProcess.StandardError.ReadLine();
                        if (line == null)
                        {
                            break;
                        }
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            Log($"Python: {line}", LogLevel.Debug);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Python stderr 監視でエラーが発生しました: {ex.Message}", LogLevel.Debug);
                }
            });

            var pingResponse = SendCommand(new Dictionary<string, object> { { "command", "ping" } });
            if (!string.Equals(pingResponse.Status, "ok", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Python ワーカーの起動に失敗しました: {pingResponse.Message}");
            }
        }

        private WorkerResponse SendCommand(Dictionary<string, object?> payload)
        {
            if (_workerInput == null || _workerOutput == null)
            {
                throw new InvalidOperationException("Python ワーカーが初期化されていません。");
            }

            var json = JsonSerializer.Serialize(payload);
            lock (_processLock)
            {
                _workerInput.WriteLine(json);
                _workerInput.Flush();

                var responseLine = _workerOutput.ReadLine();
                if (string.IsNullOrWhiteSpace(responseLine))
                {
                    throw new InvalidOperationException("Python ワーカーから応答がありません。");
                }

                var response = JsonSerializer.Deserialize<WorkerResponse>(
                    responseLine,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (response == null)
                {
                    throw new InvalidOperationException("Python ワーカーの応答を解析できません。");
                }

                return response;
            }
        }

        /// <summary>
        /// Python環境を初期化し、AudioSRをロードします
        /// </summary>
        /// <param name="onProgress">インストール進捗を報告するコールバック（段階番号, 総段階数, メッセージ）</param>
        public void Initialize(Action<int, int, string>? onProgress = null)
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
                EnsureAudioSrInstalled(onProgress);

                onProgress?.Invoke(9, 10, "Pythonワーカーを起動中...");
                var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
                StartWorkerProcess(sitePackagesPath);

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
        public void ProcessFile(string inputFile, string outputFile, string modelName, int ddimSteps, float guidanceScale, long? seed, Action<int, int>? onProgress = null)
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

            var response = SendCommand(payload);
            if (!string.Equals(response.Status, "ok", StringComparison.OrdinalIgnoreCase))
            {
                var detail = string.IsNullOrWhiteSpace(response.Traceback)
                    ? response.Message
                    : $"{response.Message}\n{response.Traceback}";
                throw new InvalidOperationException($"AudioSR処理に失敗しました: {detail}");
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
        public void ProcessBatchFile(string inputListFile, string outputPath, string modelName, int ddimSteps, float guidanceScale, long? seed)
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
                    ProcessFile(inputFile, outputFile, modelName, ddimSteps, guidanceScale, seed);
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
                        SendCommand(new Dictionary<string, object> { { "command", "shutdown" } });
                        _workerProcess.WaitForExit(3000);
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
                }
            }

            _disposed = true;
        }

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

        private static string ResolvePythonPthFilePath(string pythonHome, string? pythonPrefix)
        {
            if (!string.IsNullOrEmpty(pythonPrefix))
            {
                var expected = Path.Combine(pythonHome, $"{pythonPrefix}._pth");
                if (File.Exists(expected))
                {
                    return expected;
                }
            }

            return Directory.GetFiles(pythonHome, "python*._pth").FirstOrDefault()
                ?? Path.Combine(pythonHome, "python._pth");
        }

        private sealed class WorkerResponse
        {
            public string Status { get; set; } = "error";
            public string Message { get; set; } = "";
            public string? Traceback { get; set; }
        }
    }
}
