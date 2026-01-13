using Python.Runtime;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using static AudioSR.NET.Logger;

namespace AudioSR.NET
{
    /// <summary>
    /// Win32 API用のP/Invokeインポート
    /// </summary>
    internal static class Win32Native
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr LoadLibrary(string lpFileName);
        
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeLibrary(IntPtr hModule);
        
        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();
    }

    /// <summary>
    /// AudioSRラッパークラス
    /// </summary>
    public class AudioSrWrapper : IDisposable
    {
        private readonly string _pythonHome;
        private dynamic? _audiosr;
        private dynamic? _audiosrModel;
        private bool _initialized;
        private bool _initializationInProgress;
        private bool _initializationFailed;
        private bool _disposed;
        
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
            WriteDebugLog(message);

            if (string.IsNullOrEmpty(pythonHome))
            {
                throw new ArgumentException("pythonHome cannot be null or empty", nameof(pythonHome));
            }

            if (!Directory.Exists(pythonHome))
            {
                throw new DirectoryNotFoundException($"Python home directory not found: {pythonHome}");
            }

            _pythonHome = pythonHome;
            _initialized = false;
            _initializationInProgress = false;
            _initializationFailed = false;
            _audiosr = null;
            _audiosrModel = null;
        }

        /// <summary>
        /// デバッグログをファイルに書き込む
        /// </summary>
        private static void WriteDebugLog(string message)
        {
            Log(message, LogLevel.Debug);
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

                using (var searcher = new System.Management.ManagementObjectSearcher("SELECT * FROM Win32_VideoController"))
                {
                    var gpuNames = new System.Collections.Generic.List<string>();
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
                Log($"CPUモードにフォールバックします。", LogLevel.Warning);
            }

            Log("GPU検出失敗。CPU を使用します。", LogLevel.Info);
            return "cpu";
        }

        /// <summary>
        /// audiosrパッケージを埋め込みPythonのsite-packagesにのみインストールします
        /// </summary>
        /// <param name="onProgress">インストール進捗を報告するコールバック（段階番号, 総段階数, メッセージ）</param>
        private void EnsureAudioSrInstalled(Action<int, int, string>? onProgress = null)
        {
            try
            {
                var msg1 = "EnsureAudioSrInstalled 開始";
                Log(msg1, LogLevel.Debug);
                WriteDebugLog(msg1);

                // site-packages ディレクトリの再作成（埋め込み Python の site-packages が削除されるため）
                var ensureSitePackagesDir = Path.Combine(_pythonHome, "Lib", "site-packages");
                onProgress?.Invoke(1, 10, "site-packages ディレクトリを確認中...");
                if (!Directory.Exists(ensureSitePackagesDir))
                {
                    var msgSitePackages = $"site-packages ディレクトリを再作成中: {ensureSitePackagesDir}";
                    Log(msgSitePackages, LogLevel.Debug);
                    WriteDebugLog(msgSitePackages);
                    Directory.CreateDirectory(ensureSitePackagesDir);
                }

                // インストール済みマーカーファイルをチェック
                onProgress?.Invoke(2, 10, "インストール状態を確認中...");
                if (File.Exists(_depsInstalledMarkerFile))
                {
                    var msgMarker = "✓ 依存インストール済みマーカーファイルが存在します。初期化をスキップします。";
                    Log(msgMarker, LogLevel.Info);
                    WriteDebugLog(msgMarker);
                    onProgress?.Invoke(10, 10, "インストール済み（スキップ）");
                    return;
                }

                var msgStartInstall = "初回起動: 依存パッケージのインストールを開始します...";
                Log(msgStartInstall, LogLevel.Info);
                WriteDebugLog(msgStartInstall);
                onProgress?.Invoke(3, 10, "パッケージのインストールを開始します...");

                using (Py.GIL())
                {
                    // site-packages ディレクトリを Python パスに追加
                    var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
                    onProgress?.Invoke(3, 10, "Python パスを設定中...");
                    if (Directory.Exists(sitePackagesPath))
                    {
                        var msg_path = $"site-packages をPythonパスに追加: {sitePackagesPath}";
                        Log(msg_path, LogLevel.Debug);
                        WriteDebugLog(msg_path);

                        PythonEngine.Exec($"import sys; sys.path.insert(0, r'{sitePackagesPath}')");
                    }

                    // Python ランタイム内から get-pip.py をダウンロードして pip をセットアップ
                    onProgress?.Invoke(4, 10, "pip をインストール中...");
                    var getPipScript = $@"
import urllib.request
import sys
import os
import tempfile
import zipfile

try:
    import pip
    print('pip は既にインストールされています')
except ImportError:
    print('pip をインストール中...')
    try:
        # get-pip.py をダウンロード
        tmpdir = r'{sitePackagesPath}'
        get_pip_path = os.path.join(tmpdir, 'get-pip.py')
        print(f'get-pip.py をダウンロード中: {{get_pip_path}}')
        urllib.request.urlretrieve('https://bootstrap.pypa.io/get-pip.py', get_pip_path)
        
        # get-pip.py をテキストファイルとして読み込んで実行（subprocess なし）
        print(f'get-pip.py を実行中...')
        with open(get_pip_path, 'r', encoding='utf-8') as f:
            get_pip_code = f.read()
        
        # get-pip.py の環境変数を設定
        os_environ = os.environ.copy()
        os_environ['PIP_TARGET'] = tmpdir
        os_environ['PIP_NO_WARN_SCRIPT_LOCATION'] = '1'
        
        # get-pip.py を実行（exec を使用してプロセス起動をしない）
        exec_globals = {{'__name__': '__main__', '__file__': get_pip_path, '__doc__': None}}
        exec(get_pip_code, exec_globals)
        
        # 一時ファイルを削除
        os.remove(get_pip_path)
        print('✓ pip インストール成功')
    except Exception as e:
        print(f'✗ pip インストール失敗: {{e}}')
        raise
";
                    
                    try
                    {
                        var msgPip = "Python ランタイム内から pip をセットアップ中...";
                        Log(msgPip, LogLevel.Debug);
                        WriteDebugLog(msgPip);
                        PythonEngine.Exec(getPipScript);
                        var msgPipOk = "✓ pip セットアップ完了";
                        Log(msgPipOk, LogLevel.Debug);
                        WriteDebugLog(msgPipOk);
                    }
                    catch (Exception ex)
                    {
                        var msgPipErr = $"警告: pip セットアップ中にエラー（{ex.Message}）";
                        Log(msgPipErr, LogLevel.Debug);
                        WriteDebugLog(msgPipErr);
                        // 続行して audiosr をインポート試行
                    }

                    onProgress?.Invoke(5, 10, "audiosr のインストール状態を確認中...");
                    var msg2 = "Checking if audiosr is installed...";
                    Log(msg2, LogLevel.Debug);
                    WriteDebugLog(msg2);

                    try
                    {
                        Py.Import("audiosr");
                        var msg3 = "✓ audiosrは既にインストールされています";
                        Log(msg3, LogLevel.Debug);
                        WriteDebugLog(msg3);
                        onProgress?.Invoke(10, 10, "audiosr は既にインストール済みです");
                        return;
                    }
                    catch (Exception checkEx)
                    {
                        var msg4 = $"audiosrがインストールされていません: {checkEx.GetType().Name}: {checkEx.Message}";
                        Log(msg4, LogLevel.Debug);
                        WriteDebugLog(msg4);
                    }

                    onProgress?.Invoke(6, 10, "パッケージをインストール中...");
                    var msg5 = "パッケージのインストールを開始します...";
                    Log(msg5, LogLevel.Debug);
                    WriteDebugLog(msg5);

                    // Python内部からpipを使用してパッケージをインストール（subprocess使用禁止：AudioSR.NETが再起動される）
                    var installScript = $@"
import sys
import os

# インストール先のsite-packagesパスを設定
site_packages_path = r'{sitePackagesPath}'

def install_package(package, target_path):
    print(f'Installing {{package}} to {{target_path}}...')
    try:
        # pip の main 関数を使用
        from pip._internal.cli.main import main as pip_main

        # インストール実行（--targetオプションでインストール先を明示）
        exit_code = pip_main(['install', '--upgrade', '--target', target_path, package])

        if exit_code == 0:
            print(f'✓ {{package}} installed to {{target_path}}')
            return True
        else:
            print(f'✗ {{package}} failed with exit code {{exit_code}}')
            return False
    except Exception as e:
        print(f'✗ {{package}} failed: {{e}}')
        import traceback
        traceback.print_exc()
        return False

# sys.pathにsite-packagesを追加（まだ追加されていない場合）
if site_packages_path not in sys.path:
    sys.path.insert(0, site_packages_path)
    print(f'Added {{site_packages_path}} to sys.path')

# 必須パッケージをインストール
packages = ['torch', 'torchaudio', 'audiosr']
for pkg in packages:
    install_package(pkg, site_packages_path)

# インストール後、sys.pathを再度確認
print(f'sys.path after installation: {{sys.path[:3]}}')  # 最初の3つのパスを表示

print('Installation complete')
";

                    try
                    {
                        Log("パッケージインストールスクリプトを実行中...", LogLevel.Debug);
                        PythonEngine.Exec(installScript);
                        Log("パッケージのインストールが完了しました", LogLevel.Info);

                        // インストール成功を確認してマーカーファイルを作成
                        using (Py.GIL())
                        {
                            try
                            {
                                // sys.pathにsite-packagesが含まれていることを再確認
                                var verifyPathScript = $@"
import sys
site_packages_path = r'{sitePackagesPath}'
if site_packages_path not in sys.path:
    sys.path.insert(0, site_packages_path)
    print(f'Re-added {{site_packages_path}} to sys.path')
else:
    print(f'{{site_packages_path}} is already in sys.path')
print(f'Current sys.path (first 3): {{sys.path[:3]}}')
";
                                PythonEngine.Exec(verifyPathScript);

                                var audiosr = Py.Import("audiosr");
                                Log("✓ audiosrモジュルをインポート確認", LogLevel.Debug);
                                File.WriteAllText(_depsInstalledMarkerFile, "installed");
                                Log("✓ インストール完了マーカーファイルを作成しました", LogLevel.Info);
                            }
                            catch (Exception importEx)
                            {
                                Log($"警告: audiosrインポート確認に失敗: {importEx.Message}", LogLevel.Debug);
                                Log($"スタックトレース: {importEx.StackTrace}", LogLevel.Debug);
                            }
                        }

                        onProgress?.Invoke(9, 10, "インストール完了");
                    }
                    catch (Exception ex)
                    {
                        Log($"警告: パッケージインストール中にエラー: {ex.Message}", LogLevel.Debug);
                        WriteDebugLog($"パッケージインストールエラー: {ex.Message}");
                        if (ex.InnerException != null)
                        {
                            Log($"内部例外: {ex.InnerException.Message}", LogLevel.Debug);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                var msg17 = $"EnsureAudioSrInstalled例外: {ex.GetType().Name}: {ex.Message}";
                Log(msg17, LogLevel.Debug);
                WriteDebugLog(msg17);
                if (ex.InnerException != null)
                {
                    var msg18 = $"内部例外: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}";
                    Log(msg18, LogLevel.Debug);
                    WriteDebugLog(msg18);
                }
                Log($"Stack trace: {ex.StackTrace ?? "なし"}", LogLevel.Error);
                WriteDebugLog($"Stack trace: {ex.StackTrace}");
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

            // PythonEngine が既に初期化されている場合、直接 audiosr をインポートを試みる
            if (PythonEngine.IsInitialized && !_initialized)
            {
                Log("PythonEngine already initialized. Attempting to import audiosr...", LogLevel.Debug);
                try
                {
                    using (Py.GIL())
                    {
                        // site-packagesパスをsys.pathに追加
                        var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
                        Log($"site-packagesパスをsys.pathに追加中（既存ランタイム）: {sitePackagesPath}", LogLevel.Debug);
                        var addPathScript = $@"
import sys
site_packages_path = r'{sitePackagesPath}'
if site_packages_path not in sys.path:
    sys.path.insert(0, site_packages_path)
    print(f'Added {{site_packages_path}} to sys.path (existing runtime)')
";
                        PythonEngine.Exec(addPathScript);

                        _audiosr = Py.Import("audiosr");
                        Log("audiosrモジュールのインポートに成功しました（既に初期化済みランタイムで）", LogLevel.Info);
                        _initialized = true;
                        onProgress?.Invoke(10, 10, "初期化完了");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log($"既に初期化済みランタイムでのaudiosrインポート失敗: {ex.Message}", LogLevel.Debug);
                    // 失敗した場合は、通常のインストール処理に進む
                }
            }

            _initializationInProgress = true;

            var startTime = DateTime.Now;
            Log($"[{startTime:yyyy-MM-dd HH:mm:ss.fff}] Initializing Python environment...", LogLevel.Debug);
            _initializationFailed = false;
            var initializationSucceeded = false;

            try
            {
                var appSettings = AppSettings.Load();
                var pythonPrefix = GetPythonPrefix(appSettings.PythonVersion);

                // Pythonホームディレクトリの存在確認
                if (!Directory.Exists(_pythonHome))
                {
                    throw new DirectoryNotFoundException($"Python home directory not found: {_pythonHome}");
                }

                // PYTHONHOMEを明示的に設定
                Log($"Setting PYTHONHOME to: {_pythonHome}", LogLevel.Debug);
                Environment.SetEnvironmentVariable("PYTHONHOME", _pythonHome);

                // 既存のPATH環境変数を取得
                var originalPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                Log($"Original PATH: {originalPath}", LogLevel.Debug);

                // PATHにPython DLLディレクトリを先頭に追加（既に含まれていなければ）
                var newPath = originalPath;
                if (!originalPath.Contains(_pythonHome))
                {
                    newPath = $"{_pythonHome};{originalPath}";
                    Environment.SetEnvironmentVariable("PATH", newPath);
                    Log($"Updated PATH: {newPath}", LogLevel.Debug);
                }

                // PYTHONPATHを設定
                var pythonHomeDir = Path.GetDirectoryName(_pythonHome);
                if (pythonHomeDir != null)
                {
                    var libDir = Path.Combine(pythonHomeDir, "lib");
                    if (Directory.Exists(libDir))
                    {
                        Log($"Setting PYTHONPATH to include lib directory: {libDir}", LogLevel.Debug);
                        Environment.SetEnvironmentVariable("PYTHONPATH", libDir);
                    }
                }

                // Python._pthファイルのimport site設定を有効化する
                var pthFilePath = ResolvePythonPthFilePath(_pythonHome, pythonPrefix);
                if (File.Exists(pthFilePath))
                {
                    Log($"Checking {Path.GetFileName(pthFilePath)} for import site configuration", LogLevel.Debug);
                    var content = File.ReadAllText(pthFilePath);
                    if (!content.Contains("import site") || content.Contains("#import site"))
                    {
                        Log("Adding 'import site' to Python._pth file", LogLevel.Debug);
                        content = content.Replace("#import site", "import site");
                        if (!content.Contains("import site"))
                        {
                            content += "\r\nimport site";
                        }
                        File.WriteAllText(pthFilePath, content);
                    }
                    else
                    {
                        Log("Python._pth already contains 'import site' configuration", LogLevel.Debug);
                    }
                }

                // カレントディレクトリを確認
                Log($"Current directory: {Directory.GetCurrentDirectory()}", LogLevel.Debug);

                // DLLファイルの存在確認
                Log("Python DLLファイル一覧:", LogLevel.Debug);
                foreach (var dllFile in Directory.GetFiles(_pythonHome, "python*.dll"))
                {
                    Log($"  - {Path.GetFileName(dllFile)}", LogLevel.Debug);
                }

                // Python DLLパスを明示的に設定（ランタイム未初期化時のみ）
                var pythonDll = ResolvePythonDllPath(_pythonHome, pythonPrefix);
                if (string.IsNullOrEmpty(pythonDll))
                {
                    throw new FileNotFoundException("Python DLLが見つかりません。", _pythonHome);
                }
                
                // ランタイムが未初期化の場合のみ PythonDLL を設定
                if (!PythonEngine.IsInitialized)
                {
                    Runtime.PythonDLL = pythonDll;
                    Log($"Setting Runtime.PythonDLL to: {Runtime.PythonDLL}", LogLevel.Debug);
                }
                else
                {
                    Log($"Runtime already initialized, skipping PythonDLL setting. Current DLL: {Runtime.PythonDLL}", LogLevel.Debug);
                }

                var resolvedVersion = ResolvePythonVersion(appSettings.PythonVersion, pythonDll);
                if (resolvedVersion != null && resolvedVersion >= new Version(3, 14))
                {
                    throw new InvalidOperationException(
                        $"Python {resolvedVersion.Major}.{resolvedVersion.Minor} は未対応です。Python 3.13 以前を指定してください。");
                }

                // DLLが実際にロード可能かテスト
                var dllHandle = Win32Native.LoadLibrary(Runtime.PythonDLL);
                if (dllHandle == IntPtr.Zero)
                {
                    var errorCode = Win32Native.GetLastError();
                    throw new DllNotFoundException($"Python DLLをロードできません: {Runtime.PythonDLL}, エラーコード: {errorCode}");
                }
                else
                {
                    Win32Native.FreeLibrary(dllHandle);
                    Log($"Successfully loaded Python DLL: {Runtime.PythonDLL}", LogLevel.Debug);
                }

                // Python.Runtimeの初期化
                Log("Python.Runtime初期化開始: " + DateTime.Now.ToString("HH:mm:ss.fff"), LogLevel.Debug);
                Log($"Runtime.PythonDLL before Initialize: {Runtime.PythonDLL}", LogLevel.Debug);
                try
                {
                    PythonEngine.Initialize();
                    Log("PythonEngine successfully initialized", LogLevel.Debug);

                    // audiosrパッケージを自動インストール
                    Log("EnsureAudioSrInstalled を実行中...", LogLevel.Debug);
                    EnsureAudioSrInstalled(onProgress);
                    Log("EnsureAudioSrInstalled 完了", LogLevel.Debug);

                    using (Py.GIL())
                    {
                        try
                        {
                            // site-packagesパスをsys.pathに確実に追加
                            var sitePackagesPath = Path.Combine(_pythonHome, "Lib", "site-packages");
                            Log($"site-packagesパスをsys.pathに追加中: {sitePackagesPath}", LogLevel.Debug);
                            var addPathScript = $@"
import sys
site_packages_path = r'{sitePackagesPath}'
if site_packages_path not in sys.path:
    sys.path.insert(0, site_packages_path)
    print(f'Added {{site_packages_path}} to sys.path before import')
else:
    print(f'{{site_packages_path}} is already in sys.path')

# sys.pathの最初の3つを表示してデバッグ
print(f'sys.path before audiosr import: {{sys.path[:3]}}')
";
                            PythonEngine.Exec(addPathScript);

                            // audiosrモジュールをインポート
                            Log("audiosrモジュールをインポート中...", LogLevel.Debug);
                            _audiosr = Py.Import("audiosr");
                            Log("audiosrモジュールのインポートに成功しました", LogLevel.Info);

                            // インストール成功マーカーを作成
                            try
                            {
                                onProgress?.Invoke(10, 10, "初期化完了");
                                File.WriteAllText(_depsInstalledMarkerFile, "installed");
                                Log("依存インストール完了マーカーを作成しました", LogLevel.Info);
                            }
                            catch (Exception ex)
                            {
                                Log($"警告: マーカーファイル作成失敗（{ex.Message}）", LogLevel.Debug);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"audiosrのインポートに失敗しました: {ex.GetType().Name}: {ex.Message}", LogLevel.Error);
                            Log($"詳細スタックトレース: {ex.StackTrace}", LogLevel.Debug);

                            // デバッグ情報: sys.pathの内容を出力
                            try
                            {
                                var debugScript = @"
import sys
print('=== DEBUG: sys.path ===')
for i, p in enumerate(sys.path[:10]):
    print(f'{i}: {p}')
";
                                PythonEngine.Exec(debugScript);
                            }
                            catch { }

                            throw new InvalidOperationException("audiosrモジュールのインポートに失敗しました。必要なパッケージが正しくインストールされているか確認してください。", ex);
                        }
                    }
                }
                catch (NotSupportedException ex)
                {
                    Log($"Python ABIが未対応のため初期化に失敗しました: {ex.Message}", LogLevel.Debug);
                    throw new InvalidOperationException(
                        "PythonのABI互換性がないため初期化できません。Python 3.13 以前を指定してください。",
                        ex);
                }
                catch (Exception ex)
                {
                    Log($"Failed to initialize Python: {ex.Message}", LogLevel.Debug);
                    Log($"Exception type: {ex.GetType().Name}", LogLevel.Debug);
                    Log($"Stack trace: {ex.StackTrace ?? "なし"}", LogLevel.Error);
                    throw new InvalidOperationException("Python環境の初期化に失敗しました", ex);
                }

                initializationSucceeded = true;
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
                var finalMsg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 初期化フラグを設定: _initialized = {_initialized}, _initializationFailed = {_initializationFailed}, _audiosr = {(_audiosr != null ? "有効" : "null")}";
                Log(finalMsg, LogLevel.Debug);
                WriteDebugLog(finalMsg);
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
                Log($"[{DateTime.Now:HH:mm:ss.fff}] AudioSrWrapperが初期化されていません。", LogLevel.Debug);
                throw new InvalidOperationException("AudioSrWrapperが初期化されていません。StartProcessing_Click内で初期化を完了させてください。");
            }

            if (!File.Exists(inputFile)) 
            {
                Log($"[{DateTime.Now:HH:mm:ss.fff}] 入力ファイルが見つかりません: {inputFile}", LogLevel.Debug);
                throw new FileNotFoundException($"ファイルが見つかりません: {inputFile}");
            }
            
            Log($"[{DateTime.Now:HH:mm:ss.fff}] 処理開始: {inputFile} -> {outputFile}", LogLevel.Debug);
            Log($"[{DateTime.Now:HH:mm:ss.fff}] Parameters: modelName={modelName}, ddimSteps={ddimSteps}, guidanceScale={guidanceScale}, seed={seed}", LogLevel.Debug);
            Log($"[{DateTime.Now:HH:mm:ss.fff}] 現在の状態: audiosr={(_audiosr != null ? "有効" : "null")}", LogLevel.Debug);

            // 注：上部で既にファイルの存在チェックをしているため、二重チェックは不要

            Log($"Processing file: {inputFile} -> {outputFile}", LogLevel.Debug);
            
            // 出力ディレクトリの作成
            var outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir)) 
            {
                Directory.CreateDirectory(outputDir);
            }
            
            using (Py.GIL())
            {
                try
                {
                    Log($"audiosr参照: {_audiosr != null}, model参照: {_audiosrModel != null}", LogLevel.Debug);
                    Log($"パラメータ: modelName={modelName}, ddimSteps={ddimSteps}, guidanceScale={guidanceScale}, seed={seed}", LogLevel.Debug);

                    if (_audiosr == null)
                    {
                        throw new InvalidOperationException("AudioSRが初期化されていません。Initialize()メソッドを呼び出してください。");
                    }

                    if (_audiosrModel == null)
                    {
                        if (_audiosr == null)
                        {
                            throw new InvalidOperationException("AudioSRモジュールが初期化されていません。");
                        }

                        Log($"AudioSRモデルを初期化中... (model_name={modelName})", LogLevel.Debug);
                        var optimalDevice = DetectOptimalDevice();
                        try
                        {
                            var buildModel = _audiosr!.build_model;
                            using (var kwargs = new PyDict())
                            {
                                kwargs.SetItem("model_name", new PyString(modelName));
                                kwargs.SetItem("device", new PyString(optimalDevice));
                                Log($"デバイス {optimalDevice} を使用してモデルを初期化します。", LogLevel.Info);
                                _audiosrModel = buildModel(kwargs);
                            }
                            Log($"AudioSRモデルの初期化が完了しました。デバイス: {optimalDevice}", LogLevel.Info);
                        }
                        catch (Exception ex)
                        {
                            Log($"モデル初期化エラー: {ex.Message}", LogLevel.Debug);
                            Log("代替方法でモデル初期化を試行します...", LogLevel.Debug);

                            try
                            {
                                var buildModel = _audiosr!.build_model;
                                _audiosrModel = buildModel(modelName, optimalDevice);
                                Log($"AudioSRモデルの初期化が完了しました（代替方法）。デバイス: {optimalDevice}", LogLevel.Debug);
                            }
                            catch (Exception ex2)
                            {
                                Log($"モデル初期化失敗（代替方法も失敗）: {ex2.Message}", LogLevel.Debug);
                                throw new InvalidOperationException("AudioSRモデルの初期化に失敗しました。", ex2);
                            }
                        }
                    }

                    // 超解像処理を実行
                    Log($"AudioSR超解像処理を開始: {inputFile} -> {outputFile}", LogLevel.Debug);
                    try
                    {
                        var superResolution = _audiosrModel.super_resolution;

                        // キーワード引数を準備
                        using (var kwargs = new PyDict())
                        {
                            kwargs.SetItem("ddim_steps", new PyInt(ddimSteps));
                            kwargs.SetItem("guidance_scale", new PyFloat(guidanceScale));
                            if (seed.HasValue)
                            {
                                kwargs.SetItem("seed", new PyInt(seed.Value));
                            }

                            // super_resolutionメソッドを呼び出し
                            superResolution(inputFile, outputFile, kwargs);
                        }

                        Log($"AudioSR処理が完了しました: {outputFile}", LogLevel.Debug);
                    }
                    catch (PythonException pex)
                    {
                        Log($"AudioSR Python エラー: {pex.Message}", LogLevel.Debug);
                        Log($"Python traceback: {pex.StackTrace ?? "なし"}", LogLevel.Error);

                        // 別の呼び出し方を試す（位置引数）
                        Log("代替方法でsuper_resolution呼び出しを試行...", LogLevel.Debug);
                        try
                        {
                            var superResolution = _audiosrModel.super_resolution;

                            if (seed.HasValue)
                            {
                                superResolution(inputFile, outputFile, ddimSteps, guidanceScale, seed.Value);
                            }
                            else
                            {
                                superResolution(inputFile, outputFile, ddimSteps, guidanceScale);
                            }

                            Log($"AudioSR処理が完了しました（代替方法）: {outputFile}", LogLevel.Debug);
                        }
                        catch (Exception ex2)
                        {
                            Log($"代替方法も失敗: {ex2.Message}", LogLevel.Debug);
                            throw new Exception($"AudioSR処理エラー: {pex.Message}", pex);
                        }
                    }

                    // ファイルが正しく生成されたかチェック
                    if (File.Exists(outputFile))
                    {
                        var inputInfo = new FileInfo(inputFile);
                        var outputInfo = new FileInfo(outputFile);
                        Log($"処理完了: 入力サイズ={inputInfo.Length}, 出力サイズ={outputInfo.Length}", LogLevel.Debug);
                    }
                    else
                    {
                        Log("出力ファイルが生成されませんでした。", LogLevel.Warning);
                        throw new InvalidOperationException("出力ファイルが生成されませんでした。");
                    }
                }
                catch (PythonException pex)
                {
                    Log($"Python error: {pex.Message}", LogLevel.Debug);
                    Log($"Stack trace: {pex.StackTrace}", LogLevel.Debug);
                    throw new Exception($"Python処理エラー: {pex.Message}", pex);
                }
                catch (Exception ex)
                {
                    Log($"Error processing file: {ex.Message}", LogLevel.Debug);
                    Log($"Stack trace: {ex.StackTrace ?? "なし"}", LogLevel.Error);
                    throw;
                }
            }
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

        private static Version? ResolvePythonVersion(string? versionText, string? pythonDllPath)
        {
            if (Version.TryParse(versionText, out var parsed))
            {
                return parsed;
            }

            return TryParsePythonVersionFromDllName(pythonDllPath);
        }

        private static Version? TryParsePythonVersionFromDllName(string? pythonDllPath)
        {
            if (string.IsNullOrWhiteSpace(pythonDllPath))
            {
                return null;
            }

            var fileName = Path.GetFileNameWithoutExtension(pythonDllPath);
            if (string.IsNullOrWhiteSpace(fileName) || string.Equals(fileName, "python", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var match = Regex.Match(fileName, @"^python(?<major>\d)(?<minor>\d{1,2})$");
            if (!match.Success)
            {
                return null;
            }

            var major = int.Parse(match.Groups["major"].Value);
            var minor = int.Parse(match.Groups["minor"].Value);
            return new Version(major, minor);
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

        private static string? ResolvePythonDllPath(string pythonHome, string? pythonPrefix)
        {
            if (!string.IsNullOrEmpty(pythonPrefix))
            {
                var expected = Path.Combine(pythonHome, $"{pythonPrefix}.dll");
                if (File.Exists(expected))
                {
                    return expected;
                }
            }

            var candidates = Directory.GetFiles(pythonHome, "python*.dll");
            if (candidates.Length == 0)
            {
                return null;
            }

            var exact = candidates.FirstOrDefault(path => string.Equals(Path.GetFileName(path), "python.dll", StringComparison.OrdinalIgnoreCase));
            return exact ?? candidates[0];
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
                Log($"処理中 ({i+1}/{inputFiles.Length}): {inputFile}", LogLevel.Debug);
                
                try
                {
                    ProcessFile(inputFile, outputFile, modelName, ddimSteps, guidanceScale, seed);
                }
                catch (Exception ex)
                {
                    Log($"ファイル処理エラー: {inputFile} - {ex.Message}", LogLevel.Debug);
                    // バッチ処理は続行
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
                return;

            if (disposing)
            {
                // マネージドリソースの解放
                if (_initialized && PythonEngine.IsInitialized)
                {
                    try
                    {
                        PythonEngine.Shutdown();
                        Log("Python engine shut down successfully", LogLevel.Debug);
                    }
                    catch (Exception ex)
                    {
                        Log($"Error shutting down Python engine: {ex.Message}", LogLevel.Debug);
                    }
                }
            }

            // アンマネージドリソースの解放
            _disposed = true;
        }
    }
}




