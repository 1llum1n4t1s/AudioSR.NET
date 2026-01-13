using System;
using System.Windows;
using Velopack;
using Velopack.Sources;

using static AudioSR.NET.Logger;

namespace AudioSR.NET
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// アプリケーション起動時のイベント
        /// </summary>
        protected override async void OnStartup(StartupEventArgs e)
        {
            Log($"アプリケーション起動開始", LogLevel.Info);
            Logger.LogStartup(e.Args);

            try
            {
                VelopackApp.Build().Run();
            }
            catch (Exception ex)
            {
                Log($"Velopackの初期化中にエラーが発生しました: {ex.Message}", LogLevel.Warning);
            }

            base.OnStartup(e);
        }

        /// <summary>
        /// アプリケーション終了時のイベント
        /// </summary>
        protected override void OnExit(ExitEventArgs e)
        {
            Log($"アプリケーション終了開始", LogLevel.Info);

            try
            {
                base.OnExit(e);
            }
            finally
            {
                try
                {
                    Log($"プロセスをクリーンアップ中...", LogLevel.Debug);
                }
                catch
                {
                    // ログが失敗しても継続
                }

                System.Threading.Thread.Sleep(100);
                Environment.Exit(0);
            }
        }
    }
}
