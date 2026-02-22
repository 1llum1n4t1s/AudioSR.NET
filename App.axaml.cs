using System;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Velopack;

using static AudioSR.NET.Logger;

namespace AudioSR.NET;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            Log("アプリケーション起動開始", LogLevel.Info);
            Logger.LogStartup(desktop.Args ?? []);

            try
            {
                VelopackApp.Build().Run();
            }
            catch (Exception ex)
            {
                Log($"Velopackの初期化中にエラーが発生しました: {ex.Message}", LogLevel.Warning);
            }

            desktop.MainWindow = new MainWindow();

            desktop.ShutdownRequested += (_, _) =>
            {
                Log("アプリケーション終了開始", LogLevel.Info);
            };

            desktop.Exit += (_, _) =>
            {
                try
                {
                    Log("プロセスをクリーンアップ中...", LogLevel.Debug);
                }
                catch
                {
                    // ログが失敗しても継続
                }
            };
        }

        base.OnFrameworkInitializationCompleted();
    }
}
