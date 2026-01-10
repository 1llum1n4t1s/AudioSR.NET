using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using Python.Runtime;

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
        private dynamic? _audiosrModel; // AudioSRモデルインスタンス
        private bool _initialized;
        private bool _initializationFailed;
        private bool _disposed;
        private bool _testMode; // テストモード（audiosrが利用できない場合）
        
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
            Debug.WriteLine(message);
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
            _initializationFailed = false;
            _audiosr = null;
            _audiosrModel = null;
            _testMode = false;
        }

        /// <summary>
        /// デバッグログをファイルに書き込む
        /// </summary>
        private static void WriteDebugLog(string message)
        {
            try
            {
                var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "audiosr_debug.log");
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                File.AppendAllText(logPath, $"[{timestamp}] {message}{Environment.NewLine}");
            }
            catch
            {
                // ログ出力失敗は無視
            }
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
                Debug.WriteLine(msg1);
                WriteDebugLog(msg1);

                // site-packages ディレクトリの再作成（埋め込み Python の site-packages が削除されるため）
                var ensureSitePackagesDir = Path.Combine(_pythonHome, "Lib", "site-packages");
                onProgress?.Invoke(1, 10, "site-packages ディレクトリを確認中...");
                if (!Directory.Exists(ensureSitePackagesDir))
                {
                    var msgSitePackages = $"site-packages ディレクトリを再作成中: {ensureSitePackagesDir}";
                    Debug.WriteLine(msgSitePackages);
                    WriteDebugLog(msgSitePackages);
                    Directory.CreateDirectory(ensureSitePackagesDir);
                }

                // インストール済みマーカーファイルをチェック
                onProgress?.Invoke(2, 10, "インストール状態を確認中...");
                if (File.Exists(_depsInstalledMarkerFile))
                {
                    var msgMarker = "✓ 依存インストール済みマーカーファイルが存在します。初期化をスキップします。";
                    Debug.WriteLine(msgMarker);
                    WriteDebugLog(msgMarker);
                    onProgress?.Invoke(10, 10, "インストール済み（スキップ）");
                    return;
                }

                var msgStartInstall = "初回起動: 依存パッケージのインストールを開始します...";
                Debug.WriteLine(msgStartInstall);
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
                        Debug.WriteLine(msg_path);
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
                        Debug.WriteLine(msgPip);
                        WriteDebugLog(msgPip);
                        PythonEngine.Exec(getPipScript);
                        var msgPipOk = "✓ pip セットアップ完了";
                        Debug.WriteLine(msgPipOk);
                        WriteDebugLog(msgPipOk);
                    }
                    catch (Exception ex)
                    {
                        var msgPipErr = $"警告: pip セットアップ中にエラー（{ex.Message}）";
                        Debug.WriteLine(msgPipErr);
                        WriteDebugLog(msgPipErr);
                        // 続行して audiosr をインポート試行
                    }

                    onProgress?.Invoke(5, 10, "audiosr のインストール状態を確認中...");
                    var msg2 = "Checking if audiosr is installed...";
                    Debug.WriteLine(msg2);
                    WriteDebugLog(msg2);

                    try
                    {
                        Py.Import("audiosr");
                        var msg3 = "✓ audiosrは既にインストールされています";
                        Debug.WriteLine(msg3);
                        WriteDebugLog(msg3);
                        onProgress?.Invoke(10, 10, "audiosr は既にインストール済みです");
                        return;
                    }
                    catch (Exception checkEx)
                    {
                        var msg4 = $"audiosrがインストールされていません: {checkEx.GetType().Name}: {checkEx.Message}";
                        Debug.WriteLine(msg4);
                        WriteDebugLog(msg4);
                    }

                    onProgress?.Invoke(6, 10, "パッケージをインストール中...");
                    var msg5 = "パッケージのインストールを開始します...";
                    Debug.WriteLine(msg5);
                    WriteDebugLog(msg5);

                    // Python内部からpipを使用してパッケージをインストール
                    var installScript = @"
import sys
import subprocess

def install_package(package):
    print(f'Installing {package}...')
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', package],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f'✓ {package} installed')
        return True
    except Exception as e:
        print(f'✗ {package} failed: {e}')
        return False

# 必須パッケージをインストール
packages = ['torch', 'torchaudio', 'audiosr']
for pkg in packages:
    install_package(pkg)

print('Installation complete')
";

                    try
                    {
                        Debug.WriteLine("パッケージインストールスクリプトを実行中...");
                        PythonEngine.Exec(installScript);
                        Debug.WriteLine("✓ パッケージのインストールが完了しました");
                        onProgress?.Invoke(9, 10, "インストール完了");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"警告: パッケージインストール中にエラー: {ex.Message}");
                        WriteDebugLog($"パッケージインストールエラー: {ex.Message}");
                    }

                }
            }
            catch (Exception ex)
            {
                var msg17 = $"EnsureAudioSrInstalled例外: {ex.GetType().Name}: {ex.Message}";
                Debug.WriteLine(msg17);
                WriteDebugLog(msg17);
                if (ex.InnerException != null)
                {
                    var msg18 = $"内部例外: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}";
                    Debug.WriteLine(msg18);
                    WriteDebugLog(msg18);
                }
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
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
                Debug.WriteLine("Python environment already initialized");
                return;
            }

            var startTime = DateTime.Now;
            Debug.WriteLine($"[{startTime:yyyy-MM-dd HH:mm:ss.fff}] Initializing Python environment...");
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
                Debug.WriteLine($"Setting PYTHONHOME to: {_pythonHome}");
                Environment.SetEnvironmentVariable("PYTHONHOME", _pythonHome);

                // 既存のPATH環境変数を取得
                var originalPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                Debug.WriteLine($"Original PATH: {originalPath}");

                // PATHにPython DLLディレクトリを先頭に追加（既に含まれていなければ）
                var newPath = originalPath;
                if (!originalPath.Contains(_pythonHome))
                {
                    newPath = $"{_pythonHome};{originalPath}";
                    Environment.SetEnvironmentVariable("PATH", newPath);
                    Debug.WriteLine($"Updated PATH: {newPath}");
                }

                // PYTHONPATHを設定
                var pythonHomeDir = Path.GetDirectoryName(_pythonHome);
                if (pythonHomeDir != null)
                {
                    var libDir = Path.Combine(pythonHomeDir, "lib");
                    if (Directory.Exists(libDir))
                    {
                        Debug.WriteLine($"Setting PYTHONPATH to include lib directory: {libDir}");
                        Environment.SetEnvironmentVariable("PYTHONPATH", libDir);
                    }
                }

                // Python._pthファイルのimport site設定を有効化する
                var pthFilePath = ResolvePythonPthFilePath(_pythonHome, pythonPrefix);
                if (File.Exists(pthFilePath))
                {
                    Debug.WriteLine($"Checking {Path.GetFileName(pthFilePath)} for import site configuration");
                    var content = File.ReadAllText(pthFilePath);
                    if (!content.Contains("import site") || content.Contains("#import site"))
                    {
                        Debug.WriteLine("Adding 'import site' to Python._pth file");
                        content = content.Replace("#import site", "import site");
                        if (!content.Contains("import site"))
                        {
                            content += "\r\nimport site";
                        }
                        File.WriteAllText(pthFilePath, content);
                    }
                    else
                    {
                        Debug.WriteLine("Python._pth already contains 'import site' configuration");
                    }
                }

                // カレントディレクトリを確認
                Debug.WriteLine($"Current directory: {Directory.GetCurrentDirectory()}");

                // DLLファイルの存在確認
                Debug.WriteLine("Python DLLファイル一覧:");
                foreach (var dllFile in Directory.GetFiles(_pythonHome, "python*.dll"))
                {
                    Debug.WriteLine($"  - {Path.GetFileName(dllFile)}");
                }

                // Python DLLパスを明示的に設定
                var pythonDll = ResolvePythonDllPath(_pythonHome, pythonPrefix);
                if (string.IsNullOrEmpty(pythonDll))
                {
                    throw new FileNotFoundException("Python DLLが見つかりません。", _pythonHome);
                }
                Runtime.PythonDLL = pythonDll;
                Debug.WriteLine($"Setting Runtime.PythonDLL to: {Runtime.PythonDLL}");

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
                    Debug.WriteLine($"Successfully loaded Python DLL: {Runtime.PythonDLL}");
                }

                // Python.Runtimeの初期化
                Debug.WriteLine("Python.Runtime初期化開始: " + DateTime.Now.ToString("HH:mm:ss.fff"));
                Debug.WriteLine($"Runtime.PythonDLL before Initialize: {Runtime.PythonDLL}");
                try
                {
                    PythonEngine.Initialize();
                    Debug.WriteLine("PythonEngine successfully initialized");

                    // audiosrパッケージを自動インストール
                    Debug.WriteLine("EnsureAudioSrInstalled を実行中...");
                    EnsureAudioSrInstalled(onProgress);
                    Debug.WriteLine("EnsureAudioSrInstalled 完了");

                    using (Py.GIL())
                    {
                        try
                        {
                            // audiosrモジュールをインポート
                            Debug.WriteLine("audiosrモジュールをインポート中...");
                            _audiosr = Py.Import("audiosr");
                            _testMode = false;
                            Debug.WriteLine("✓ audiosrモジュールのインポートに成功しました");

                            // インストール成功マーカーを作成
                            try
                            {
                                onProgress?.Invoke(10, 10, "初期化完了");
                                File.WriteAllText(_depsInstalledMarkerFile, "installed");
                                Debug.WriteLine("✓ 依存インストール完了マーカーを作成しました");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"警告: マーカーファイル作成失敗（{ex.Message}）");
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"audiosrのインポートに失敗しました: {ex.GetType().Name}: {ex.Message}");
                            WriteDebugLog($"audiosrインポートエラー: {ex.Message}");

                            // テストモードに切り替え（警告のみ）
                            _testMode = true;
                            _audiosr = null;
                            Debug.WriteLine("警告: テストモードに切り替えました。実際のAudioSR処理は実行できません。");
                        }
                    }
                }
                catch (NotSupportedException ex)
                {
                    Debug.WriteLine($"Python ABIが未対応のため初期化に失敗しました: {ex.Message}");
                    throw new InvalidOperationException(
                        "PythonのABI互換性がないため初期化できません。Python 3.13 以前を指定してください。",
                        ex);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to initialize Python: {ex.Message}");
                    Debug.WriteLine($"Exception type: {ex.GetType().Name}");
                    Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    throw new InvalidOperationException("Python環境の初期化に失敗しました", ex);
                }

                initializationSucceeded = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Python環境の初期化中にエラーが発生しました: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                _testMode = true;
                throw;
            }
            finally
            {
                _initialized = initializationSucceeded;
                _initializationFailed = !initializationSucceeded;
                var finalMsg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 初期化フラグを設定: _initialized = {_initialized}, _initializationFailed = {_initializationFailed}, _testMode = {_testMode}, _audiosr = {(_audiosr != null ? "有効" : "null")}";
                Debug.WriteLine(finalMsg);
                WriteDebugLog(finalMsg);
            }

            var endTime = DateTime.Now;
            Debug.WriteLine($"[{endTime:yyyy-MM-dd HH:mm:ss.fff}] Python environment successfully initialized in {(endTime - startTime).TotalSeconds:F2} seconds");
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
        public void ProcessFile(string inputFile, string outputFile, string modelName, int ddimSteps, float guidanceScale, long? seed)
        {
            Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ProcessFile 呼び出し: {inputFile} -> {outputFile}");
            
            if (!_initialized) 
            {
                if (_initializationFailed)
                {
                    Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 前回の初期化が失敗しているため再初期化を試行します。");
                }
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] AudioSrWrapperが初期化されていません。初期化を行います。");
                Initialize();
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Initialize後の状態: initialized={_initialized}, testMode={_testMode}, audiosr={(_audiosr != null ? "取得済み" : "null")}");
                if (!_initialized)
                {
                    Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 初期化に失敗したため処理を中断します。");
                    throw new InvalidOperationException("AudioSrWrapperの初期化に失敗しました。");
                }
            }

            if (!File.Exists(inputFile)) 
            {
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 入力ファイルが見つかりません: {inputFile}");
                throw new FileNotFoundException($"ファイルが見つかりません: {inputFile}");
            }
            
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 処理開始: {inputFile} -> {outputFile}");
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Parameters: modelName={modelName}, ddimSteps={ddimSteps}, guidanceScale={guidanceScale}, seed={seed}");
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] 現在の状態: testMode={_testMode}, audiosr={(_audiosr != null ? "有効" : "null")}");

            // 注：上部で既にファイルの存在チェックをしているため、二重チェックは不要

            Debug.WriteLine($"Processing file: {inputFile} -> {outputFile}");
            
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
                    Debug.WriteLine($"テストモード状態: {_testMode}, audiosr参照: {_audiosr != null}, model参照: {_audiosrModel != null}");
                    Debug.WriteLine($"パラメータ: modelName={modelName}, ddimSteps={ddimSteps}, guidanceScale={guidanceScale}, seed={seed}");

                    // テストモードの場合は単純にコピー
                    if (_testMode || _audiosr == null)
                    {
                        Debug.WriteLine("警告: テストモードまたはaudiosrが未初期化のため、ファイルをコピーのみ行います。");
                        File.Copy(inputFile, outputFile, true);
                        return;
                    }

                    // モデルが未初期化の場合は初期化
                    if (_audiosrModel == null)
                    {
                        Debug.WriteLine($"AudioSRモデルを初期化中... (model_name={modelName})");
                        try
                        {
                            // build_model関数を呼び出してモデルをロード
                            var buildModel = _audiosr.build_model;
                            using (var kwargs = new PyDict())
                            {
                                kwargs.SetItem("model_name", new PyString(modelName));
                                kwargs.SetItem("device", new PyString("auto"));
                                _audiosrModel = buildModel(kwargs);
                            }
                            Debug.WriteLine("AudioSRモデルの初期化が完了しました。");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"モデル初期化エラー: {ex.Message}");
                            Debug.WriteLine("代替方法でモデル初期化を試行します...");

                            // 別の方法を試す（位置引数）
                            try
                            {
                                var buildModel = _audiosr.build_model;
                                _audiosrModel = buildModel(modelName, "auto");
                                Debug.WriteLine("AudioSRモデルの初期化が完了しました（代替方法）。");
                            }
                            catch (Exception ex2)
                            {
                                Debug.WriteLine($"モデル初期化失敗（代替方法も失敗）: {ex2.Message}");
                                throw new InvalidOperationException("AudioSRモデルの初期化に失敗しました。", ex2);
                            }
                        }
                    }

                    // 超解像処理を実行
                    Debug.WriteLine($"AudioSR超解像処理を開始: {inputFile} -> {outputFile}");
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

                        Debug.WriteLine($"AudioSR処理が完了しました: {outputFile}");
                    }
                    catch (PythonException pex)
                    {
                        Debug.WriteLine($"AudioSR Python エラー: {pex.Message}");
                        Debug.WriteLine($"Python traceback: {pex.StackTrace}");

                        // 別の呼び出し方を試す（位置引数）
                        Debug.WriteLine("代替方法でsuper_resolution呼び出しを試行...");
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

                            Debug.WriteLine($"AudioSR処理が完了しました（代替方法）: {outputFile}");
                        }
                        catch (Exception ex2)
                        {
                            Debug.WriteLine($"代替方法も失敗: {ex2.Message}");
                            throw new Exception($"AudioSR処理エラー: {pex.Message}", pex);
                        }
                    }

                    // ファイルが正しく生成されたかチェック
                    if (File.Exists(outputFile))
                    {
                        var inputInfo = new FileInfo(inputFile);
                        var outputInfo = new FileInfo(outputFile);
                        Debug.WriteLine($"処理完了: 入力サイズ={inputInfo.Length}, 出力サイズ={outputInfo.Length}");
                    }
                    else
                    {
                        Debug.WriteLine("警告: 出力ファイルが生成されませんでした。");
                        throw new InvalidOperationException("出力ファイルが生成されませんでした。");
                    }
                }
                catch (PythonException pex)
                {
                    Debug.WriteLine($"Python error: {pex.Message}");
                    Debug.WriteLine($"Stack trace: {pex.StackTrace}");
                    throw new Exception($"Python処理エラー: {pex.Message}", pex);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing file: {ex.Message}");
                    Debug.WriteLine($"Stack trace: {ex.StackTrace}");
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
            Debug.WriteLine($"Processing {inputFiles.Length} files from batch file: {inputListFile}");

            for (var i = 0; i < inputFiles.Length; i++)
            {
                var inputFile = inputFiles[i].Trim();
                if (!File.Exists(inputFile))
                {
                    Debug.WriteLine($"警告: ファイルが見つかりません: {inputFile} - スキップします");
                    continue;
                }

                var outputFile = Path.Combine(outputPath, Path.GetFileName(inputFile));
                Debug.WriteLine($"処理中 ({i+1}/{inputFiles.Length}): {inputFile}");
                
                try
                {
                    ProcessFile(inputFile, outputFile, modelName, ddimSteps, guidanceScale, seed);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ファイル処理エラー: {inputFile} - {ex.Message}");
                    // バッチ処理は続行
                }
            }

            Debug.WriteLine($"バッチ処理完了: {inputFiles.Length} ファイル");
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
                        Debug.WriteLine("Python engine shut down successfully");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error shutting down Python engine: {ex.Message}");
                    }
                }
            }

            // アンマネージドリソースの解放
            _disposed = true;
        }
    }
}
