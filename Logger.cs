using log4net;
using log4net.Config;
using System.IO;
using System.Reflection;

namespace AudioSR.NET;

/// <summary>
/// ログレベルを表す列挙型（互換性維持用）
/// </summary>
public enum LogLevel
{
    Debug,
    Info,
    Warning,
    Error
}

/// <summary>
/// ログ出力機能を提供するクラス（log4net使用）
/// </summary>
public static class Logger
{
    /// <summary>
    /// log4netロガーインスタンス
    /// </summary>
    private static readonly ILog _logger = LogManager.GetLogger(MethodBase.GetCurrentMethod()?.DeclaringType);

    /// <summary>
    /// 初期化フラグ
    /// </summary>
    private static bool _initialized = false;

    /// <summary>
    /// 初期化を行う
    /// </summary>
    private static void Initialize()
    {
        if (_initialized)
            return;

        // log4net設定ファイルの読み込み
        var configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log4net.config");
        if (File.Exists(configFile))
        {
            XmlConfigurator.Configure(new FileInfo(configFile));
        }
        else
        {
            // 設定ファイルがない場合は基本設定を使用
            BasicConfigurator.Configure();
        }

        _initialized = true;
    }

    /// <summary>
    /// ログを出力する
    /// </summary>
    /// <param name="message">ログメッセージ</param>
    /// <param name="level">ログレベル（デフォルト: Info）</param>
    public static void Log(string message, LogLevel level = LogLevel.Info)
    {
        Initialize();

        switch (level)
        {
            case LogLevel.Debug:
                _logger.Debug(message);
                break;
            case LogLevel.Info:
                _logger.Info(message);
                break;
            case LogLevel.Warning:
                _logger.Warn(message);
                break;
            case LogLevel.Error:
                _logger.Error(message);
                break;
        }
    }

    /// <summary>
    /// 複数行のログを出力する
    /// </summary>
    /// <param name="messages">ログメッセージの配列</param>
    /// <param name="level">ログレベル（デフォルト: Info）</param>
    public static void LogLines(string[] messages, LogLevel level = LogLevel.Info)
    {
        foreach (var message in messages)
        {
            Log(message, level);
        }
    }

    /// <summary>
    /// 例外情報を含むログを出力する（常にErrorレベル）
    /// </summary>
    /// <param name="message">ログメッセージ</param>
    /// <param name="exception">例外オブジェクト</param>
    public static void LogException(string message, Exception exception)
    {
        Initialize();
        _logger.Error(message, exception);
    }

    /// <summary>
    /// アプリケーション起動時のログを出力する（Debugレベル）
    /// </summary>
    /// <param name="args">コマンドライン引数</param>
    public static void LogStartup(string[] args)
    {
        var messages = new List<string>
        {
            "=== AudioSR.NET 起動ログ ===",
            $"起動時刻: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}",
            $"実行ファイルパス: {Environment.ProcessPath}",
            $"コマンドライン引数の数: {args.Length}",
            "コマンドライン引数:"
        };

        for (var i = 0; i < args.Length; i++)
        {
            messages.Add($"  [{i}]: {args[i]}");
        }

        LogLines(messages.ToArray(), LogLevel.Debug);
    }
}
