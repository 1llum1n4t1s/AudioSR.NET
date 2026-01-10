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

            // GUI専用アプリのため、コマンドライン引数はサポートしない
            if (e.Args.Length > 0)
            {
                MessageBox.Show(
                    "このアプリはGUI専用です。コマンドラインからの実行はサポートしていません。",
                    "AudioSR.NET",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            // GUIモードで続行（StartupUriによりMainWindowが起動する）
        }
    }
}
