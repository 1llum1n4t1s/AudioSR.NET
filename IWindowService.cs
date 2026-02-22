using System.Threading.Tasks;

namespace AudioSR.NET;

/// <summary>
/// View依存操作を抽象化するインターフェース
/// </summary>
public interface IWindowService
{
    /// <summary>
    /// メッセージダイアログを表示します
    /// </summary>
    Task ShowMessageAsync(string title, string message);

    /// <summary>
    /// ファイル選択ダイアログを表示します
    /// </summary>
    /// <returns>選択されたファイルパスの配列</returns>
    Task<string[]> OpenFilePickerAsync();

    /// <summary>
    /// フォルダ選択ダイアログを表示します
    /// </summary>
    /// <returns>選択されたフォルダパス（キャンセル時はnull）</returns>
    Task<string?> OpenFolderPickerAsync();
}
