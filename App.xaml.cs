using System;
using System.Windows;
using Velopack;
using Velopack.Sources;

namespace AudioSR.NET
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            try
            {
                VelopackApp.Build().Run();

                var updateManager = new UpdateManager(
                    new GithubSource("https://github.com/1llum1n4t1s/AudioSR.NET", null, false));
                var updateInfo = await updateManager.CheckForUpdatesAsync();
                if (updateInfo is not null)
                {
                    await updateManager.DownloadUpdatesAsync(updateInfo);
                    updateManager.ApplyUpdatesAndRestart(updateInfo);
                    return;
                }
            }
            catch (Exception ex)
            {
                // Velopackがインストール環境でない場合やネットワークエラーなどを無視して起動
                System.Diagnostics.Debug.WriteLine($"更新チェック中にエラーが発生しました: {ex.Message}");
            }

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
