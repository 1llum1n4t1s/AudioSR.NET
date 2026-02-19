using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;

namespace AudioSR.NET;

/// <summary>
/// 処理対象のファイル情報を保持するクラス
/// </summary>
public class FileItem : INotifyPropertyChanged
{
    private string _path = "";
    private string _status = "待機中";

    /// <summary>
    /// ファイルの完全パス
    /// </summary>
    public string Path
    {
        get => _path;
        set
        {
            if (_path != value)
            {
                _path = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(Name));
            }
        }
    }

    /// <summary>
    /// ファイル名のみを取得
    /// </summary>
    public string Name => System.IO.Path.GetFileName(Path);

    /// <summary>
    /// 処理ステータス（待機中、処理中、完了、エラー）
    /// </summary>
    public string Status
    {
        get => _status;
        set
        {
            if (_status != value)
            {
                _status = value;
                OnPropertyChanged();
            }
        }
    }

    /// <summary>
    /// ファイルが存在するかを確認
    /// </summary>
    /// <returns>ファイルが存在すれば true</returns>
    public bool Exists() => File.Exists(Path);

    /// <summary>
    /// ファイルが音声ファイルかどうかを判断
    /// </summary>
    /// <returns>音声ファイルなら true</returns>
    public bool IsAudioFile()
    {
        var ext = System.IO.Path.GetExtension(Path);
        return ext.Equals(".wav", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".mp3", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".ogg", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".flac", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".aac", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".mp4", StringComparison.OrdinalIgnoreCase);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}