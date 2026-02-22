using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AudioSR.NET;

/// <summary>
/// AppSettings用のJSON Source Generator (Native AOT対応)
/// </summary>
[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(AppSettings))]
internal partial class AppSettingsJsonContext : JsonSerializerContext;

/// <summary>
/// アプリケーション設定を保存・読み込みするクラス
/// </summary>
public class AppSettings
{
    // Pythonの設定
    public string PythonHome { get; set; } = "";
    public string PythonVersion { get; set; } = "";
        
    // モデルの設定
    public string ModelName { get; set; } = "basic";
        
    // 処理パラメータ
    public int DdimSteps { get; set; } = 100;
    public float GuidanceScale { get; set; } = 3.5f;
    public long? Seed { get; set; } = null;
    public bool UseRandomSeed { get; set; } = true;
        
    // 出力設定
    // 既定の出力フォルダ: %USERPROFILE%/Documents
    public string OutputFolder { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Documents");
    public bool OverwriteOutputFiles { get; set; } = false;
        
    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "AudioSR.NET", 
        "settings.json");
            
    /// <summary>
    /// 設定をファイルから読み込みます
    /// </summary>
    /// <returns>読み込んだ設定、ファイルがない場合はデフォルト設定</returns>
    public static AppSettings Load()
    {
        try
        {
            // 設定ディレクトリが存在しない場合は作成
            var settingsDir = Path.GetDirectoryName(SettingsFilePath) ?? "";
            if (!Directory.Exists(settingsDir))
            {
                Directory.CreateDirectory(settingsDir);
            }
                
            // 設定ファイルが存在する場合は読み込み
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                var settings = JsonSerializer.Deserialize(json, AppSettingsJsonContext.Default.AppSettings);
                return settings ?? new AppSettings();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"設定の読み込み中にエラーが発生しました: {ex.Message}");
        }
            
        // デフォルト設定を返す
        return new AppSettings();
    }
        
    /// <summary>
    /// 設定をファイルに保存します
    /// </summary>
    /// <returns>保存に成功したかどうか</returns>
    public bool Save()
    {
        try
        {
            // 設定ディレクトリが存在しない場合は作成
            var settingsDir = Path.GetDirectoryName(SettingsFilePath) ?? "";
            if (!Directory.Exists(settingsDir))
            {
                Directory.CreateDirectory(settingsDir);
            }
                
            // 設定をJSON形式でアトミックに保存（一時ファイルへ書き込んでからリネーム）
            var json = JsonSerializer.Serialize(this, AppSettingsJsonContext.Default.AppSettings);
            var tempPath = SettingsFilePath + ".tmp";
            File.WriteAllText(tempPath, json);
            File.Replace(tempPath, SettingsFilePath, null);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"設定の保存中にエラーが発生しました: {ex.Message}");
            return false;
        }
    }
}
