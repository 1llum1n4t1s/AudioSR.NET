using System;
using System.IO;

namespace AudioSR.NET;

public class Program
{
    public static void Main(string[] args)
    {
        Console.WriteLine("AudioSR.NET - オーディオ超解像処理ツール");
        Console.WriteLine("========================================");

        if (args.Length < 1)
        {
            ShowUsage();
            return;
        }

        try
        {
            // コマンドライン引数の解析
            string? inputFile = null;
            string? inputListFile = null;
            var outputPath = Path.Combine(Environment.CurrentDirectory, "output");
            var modelName = "basic";
            var ddimSteps = 50;
            var guidanceScale = 3.5f;
            int? seed = null;

            for (var i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-i":
                    case "--input":
                        if (i + 1 < args.Length) inputFile = args[++i];
                        break;
                    case "-il":
                    case "--input-list":
                        if (i + 1 < args.Length) inputListFile = args[++i];
                        break;
                    case "-o":
                    case "--output":
                        if (i + 1 < args.Length) outputPath = args[++i];
                        break;
                    case "-m":
                    case "--model":
                        if (i + 1 < args.Length) modelName = args[++i];
                        break;
                    case "--ddim-steps":
                        if (i + 1 < args.Length) int.TryParse(args[++i], out ddimSteps);
                        break;
                    case "-gs":
                    case "--guidance-scale":
                        if (i + 1 < args.Length) float.TryParse(args[++i], out guidanceScale);
                        break;
                    case "--seed":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out var seedValue))
                            seed = seedValue;
                        break;
                    case "-h":
                    case "--help":
                        ShowUsage();
                        return;
                }
            }

            // Pythonのホームディレクトリを環境変数から取得
            var pythonHome = Environment.GetEnvironmentVariable("PYTHONHOME");
            if (string.IsNullOrEmpty(pythonHome))
            {
                Console.WriteLine("エラー: PYTHONHOME環境変数が設定されていません。");
                Console.WriteLine("Pythonのインストールディレクトリを指定して、環境変数を設定してください。");
                return;
            }

            using (var audioSr = new AudioSrWrapper(pythonHome))
            {
                if (!string.IsNullOrEmpty(inputFile))
                {
                    // 単一ファイルの処理
                    var outputFile = Path.Combine(outputPath, Path.GetFileName(inputFile));
                    Console.WriteLine($"処理中: {inputFile} -> {outputFile}");
                    audioSr.ProcessFile(inputFile, outputFile, modelName, ddimSteps, guidanceScale, seed);
                }
                else if (!string.IsNullOrEmpty(inputListFile))
                {
                    // バッチ処理
                    Console.WriteLine($"バッチ処理中: {inputListFile} -> {outputPath}");
                    audioSr.ProcessBatchFile(inputListFile, outputPath, modelName, ddimSteps, guidanceScale, seed);
                }
                else
                {
                    Console.WriteLine("エラー: 入力ファイルまたは入力リストファイルを指定してください。");
                    ShowUsage();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラーが発生しました: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"詳細: {ex.InnerException.Message}");
            }
        }
    }

    static void ShowUsage()
    {
        Console.WriteLine("使用方法:");
        Console.WriteLine("  AudioSR.NET -i <入力ファイル> [オプション]");
        Console.WriteLine("  AudioSR.NET -il <入力リストファイル> [オプション]");
        Console.WriteLine();
        Console.WriteLine("オプション:");
        Console.WriteLine("  -i, --input <ファイル>       処理する入力オーディオファイル");
        Console.WriteLine("  -il, --input-list <ファイル>  処理するオーディオファイルのリストを含むファイル");
        Console.WriteLine("  -o, --output <パス>          出力ディレクトリのパス（デフォルト: ./output）");
        Console.WriteLine("  -m, --model <名前>           使用するモデル名（basic または speech、デフォルト: basic）");
        Console.WriteLine("  --ddim-steps <数値>          DDIMステップ数（デフォルト: 50）");
        Console.WriteLine("  -gs, --guidance-scale <数値>  ガイダンススケール（デフォルト: 3.5）");
        Console.WriteLine("  --seed <数値>                乱数シード（指定しない場合はランダム）");
        Console.WriteLine("  -h, --help                   このヘルプメッセージを表示");
        Console.WriteLine();
        Console.WriteLine("注意: このプログラムを実行する前に以下の条件が必要です：");
        Console.WriteLine("  1. Python 3.9がインストールされていること");
        Console.WriteLine("  2. PYTHONHOME環境変数がPythonのインストールディレクトリを指すこと");
        Console.WriteLine("  3. AudioSRのPythonパッケージがインストールされていること（pip install audiosr==0.0.7）");
    }
}