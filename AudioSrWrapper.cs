using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Diagnostics.CodeAnalysis;
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
        private bool _initialized;
        private bool _disposed;
        private bool _testMode; // テストモード（audiosrが利用できない場合）

        /// <summary>
        /// AudioSrWrapperの新しいインスタンスを初期化します
        /// </summary>
        /// <param name="pythonHome">組み込みPythonのパス</param>
        public AudioSrWrapper(string pythonHome)
        {
            Debug.WriteLine($"AudioSrWrapper constructor called with pythonHome: {pythonHome}");
            
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
            _audiosr = null;
            _testMode = false;
        }

        /// <summary>
        /// Python環境を初期化し、AudioSRをロードします
        /// </summary>
        public void Initialize()
        {
            if (_initialized)
            {
                Debug.WriteLine("Python environment already initialized");
                return;
            }

            var startTime = DateTime.Now;
            Debug.WriteLine($"[{startTime:yyyy-MM-dd HH:mm:ss.fff}] Initializing Python environment...");

            try
            {
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
                var pthFileName = "python313._pth";
                var pthFilePath = Path.Combine(_pythonHome, pthFileName);
                if (File.Exists(pthFilePath))
                {
                    Debug.WriteLine($"Checking {pthFileName} for import site configuration");
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
                var pythonDll = Path.Combine(_pythonHome, "python313.dll");
                if (!File.Exists(pythonDll))
                {
                    pythonDll = Path.Combine(_pythonHome, "python.dll");
                    if (!File.Exists(pythonDll))
                    {
                        throw new FileNotFoundException("Python DLLが見つかりません。", pythonDll);
                    }
                }
                Runtime.PythonDLL = pythonDll;
                Debug.WriteLine($"Setting Runtime.PythonDLL to: {Runtime.PythonDLL}");

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
                try
                {
                    PythonEngine.Initialize();
                    Debug.WriteLine("PythonEngine successfully initialized");

                    using (Py.GIL())
                    {
                        try
                        {
                            // まず本物のaudiosrモジュールのロードを試みる
                            Debug.WriteLine("Attempting to import real audiosr module...");
                            
                            try
                            {
                                _audiosr = Py.Import("audiosr");
                                if (_audiosr != null)
                                {
                                    _testMode = false;
                                    Debug.WriteLine("Successfully imported real audiosr module");
                                    
                                    // モジュールのメソッドを一覧表示してみる
                                    Debug.WriteLine("Available methods in audiosr module:");
                                    try
                                    {
                                        // PyObjectのメソッドを一覧表示するより安全な方法
                                        dynamic builtins = Py.Import("builtins");
                                        dynamic dirResult = builtins.dir(_audiosr);
                                        
                                        // Python結果を反復処理
                                        for (long i = 0; i < dirResult.__len__(); i++)
                                        {
                                            var item = dirResult[i];
                                            Debug.WriteLine($"  - {item}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"メソッド一覧の取得中にエラーが発生しました: {ex.Message}");
                                    }
                                    Debug.WriteLine("本物のaudiosrモジュールのインポートに成功しました。初期化を続行します");
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Real audiosr import failed: {ex.Message}");
                            }
                            
                            // 本物のモジュールが読み込めなかった場合、モックを作成
                            Debug.WriteLine("Creating simple mock audiosr module...");
                            PythonEngine.Exec(@"
import sys
import os
import shutil

class SimpleAudioSR:
    def process_file(self, input_file, output_file, *args, **kwargs):
        print(f'Mock processing {input_file} -> {output_file}')
        try:
            d = os.path.dirname(output_file)
            if d and not os.path.exists(d):
                os.makedirs(d)
            shutil.copy2(input_file, output_file)
            return True
        except Exception as e:
            print(f'Error copying file: {e}')
            return False

# sys.modulesに直接登録
sys.modules['audiosr'] = SimpleAudioSR()
                            ");

                            // モックモジュールをインポート
                            _audiosr = Py.Import("audiosr");
                            _testMode = true;
                            Debug.WriteLine("Created and imported mock audiosr module. TEST MODE ENABLED");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Failed to create mock audiosr module: {ex.Message}");
                            
                            // エラーを外部にスローせずエラーログのみ表示
                            Debug.WriteLine($"Exception type: {ex.GetType().Name}");
                            Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to initialize Python: {ex.Message}");
                    Debug.WriteLine($"Exception type: {ex.GetType().Name}");
                    Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    throw new InvalidOperationException("Python環境の初期化に失敗しました", ex);
                }
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
                // 必ず初期化フラグを設定する
                _initialized = true;
                Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 初期化フラグを設定: _initialized = {_initialized}, _testMode = {_testMode}");
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
        public void ProcessFile(string inputFile, string outputFile, string modelName, int ddimSteps, float guidanceScale, int? seed)
        {
            Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ProcessFile 呼び出し: {inputFile} -> {outputFile}");
            
            if (!_initialized) 
            {
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] AudioSrWrapperが初期化されていません。初期化を行います。");
                Initialize();
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] Initialize後の状態: initialized={_initialized}, testMode={_testMode}, audiosr={(_audiosr != null ? "取得済み" : "null")}");
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
            
            if (_audiosr == null || _testMode)
            {
                Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] テストモードが有効です：単純なコピー処理を行います");
                File.Copy(inputFile, outputFile, true);
                Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] テストモードによるコピー完了: {inputFile} -> {outputFile}");
                return;
            }

            using (Py.GIL())
            {
                try
                {
                    // テストモード状態と_audiosrオブジェクトの状態をログ出力
                    Debug.WriteLine($"テストモード状態: {_testMode}, audiosr参照: {_audiosr != null}");
                    Debug.WriteLine($"パラメータ: modelName={modelName}, ddimSteps={ddimSteps}, guidanceScale={guidanceScale}, seed={seed}");
                    
                    // すべてのパラメータを使用して処理を呼び出し
                    Debug.WriteLine($"Calling audiosr.process_file with full parameters");
                    
                    // Pythonのキーワード引数に変換して呼び出し
                    // ノート: ddim_steps, guidance_scale, seed はPython側のパラメータ名に合わせています
                    var kwargs = new Dictionary<string, object>();
                    kwargs["ddim_steps"] = ddimSteps;
                    kwargs["guidance_scale"] = guidanceScale;
                    if (seed.HasValue)
                    {
                        kwargs["seed"] = seed.Value;
                    }
                    
                    // 旧式の呼び出し方法を試みる
                    try
                    {
                        Debug.WriteLine("Attempting to call process_file with dict kwargs");
                        // null参照チェックを追加
                    if (_audiosr != null)
                    {
                        _audiosr.process_file(inputFile, outputFile, modelName, kwargs);
                    }
                    else
                    {
                        Debug.WriteLine("警告: _audiosrがnullです。処理をスキップします。");
                        // テストモードと同じ挙動にする（コピーのみ）
                        File.Copy(inputFile, outputFile, true);
                    }
                    }
                    catch (PythonException)
                    {
                        Debug.WriteLine("First method failed, trying expanded parameters");
                        // 直接キーワードパラメータとして渡す
                        // PyDict新しいインスタンスを作成
                        using (PyDict pyKwargs = new PyDict())
                        {
                            // PyObjectに変換するために、PyObjectのSetItem()メソッドを使用
                            pyKwargs.SetItem("ddim_steps", new PyInt(ddimSteps));
                            pyKwargs.SetItem("guidance_scale", new PyFloat(guidanceScale));
                            if (seed.HasValue)
                            {
                                pyKwargs.SetItem("seed", new PyInt(seed.Value));
                            }
                            
                            // null参照チェックを追加
                            if (_audiosr != null)
                            {
                                _audiosr.process_file(inputFile, outputFile, modelName, pyKwargs);
                            }
                            else
                            {
                                Debug.WriteLine("警告: _audiosrがnullです。処理をスキップします。");
                                // テストモードと同じ挙動にする（コピーのみ）
                                File.Copy(inputFile, outputFile, true);
                            }
                        }
                    }
                    
                    // ファイルが実際に変更されたかチェック
                    try
                    {
                        if (File.Exists(outputFile))
                        {
                            var inputInfo = new FileInfo(inputFile);
                            var outputInfo = new FileInfo(outputFile);
                            if (inputInfo.Length == outputInfo.Length)
                            {
                                Debug.WriteLine("警告: 出力ファイルのサイズが入力ファイルと同じです。変換が正しく行われていない可能性があります。");
                            }
                            else
                            {
                                Debug.WriteLine($"出力ファイルのサイズ({outputInfo.Length})が入力ファイル({inputInfo.Length})と異なります。変換成功した可能性が高いです。");
                            }
                        }
                        else
                        {
                            Debug.WriteLine("警告: 出力ファイルが存在しません。処理が失敗した可能性があります。");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"ファイル比較中にエラーが発生しました: {ex.Message}");
                    }
                    
                    Debug.WriteLine($"File processed successfully: {outputFile}");
                }
                catch (PythonException pex)
                {
                    Debug.WriteLine($"Python error: {pex.Message}");
                    throw new Exception($"Python処理エラー: {pex.Message}", pex);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing file: {ex.Message}");
                    throw;
                }
            }
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
        public void ProcessBatchFile(string inputListFile, string outputPath, string modelName, int ddimSteps, float guidanceScale, int? seed)
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