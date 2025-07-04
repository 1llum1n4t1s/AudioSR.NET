using System;
using System.Windows;

namespace AudioSR.NET
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // コマンドライン引数の処理
            if (e.Args.Length > 0)
            {
                // コマンドラインモードで実行する場合の処理
                RunConsoleMode(e.Args);
                Shutdown();
                return;
            }

            // GUIモードで続行（StartupUriによりMainWindowが起動する）
        }

        /// <summary>
        /// コマンドラインモードで実行
        /// </summary>
        private void RunConsoleMode(string[] args)
        {
            try
            {
                // Program.Mainと同等の処理を実行
                Program.Main(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラー: {ex.Message}");
            }
        }
    }
}
