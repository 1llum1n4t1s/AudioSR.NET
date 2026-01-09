# AudioSR.NET

AudioSR.NETは、[versatile_audio_super_resolution](https://github.com/haoheliu/versatile_audio_super_resolution)のC#ラッパーです。このライブラリを使用することで、C#アプリケーションから簡単にオーディオの超解像処理を行うことができます。

## 機能

- 任意のサンプリングレートのオーディオファイルを48kHzの高品質に変換
- バッチ処理対応
- 様々な音源（音楽、スピーチ、環境音など）に対応
- ドラッグ＆ドロップによる簡単操作のGUIインターフェース
- 設定の保存と読み込み機能

## 必要条件

- .NET 9.0
- Python 3.9
- AudioSRのPythonパッケージ（`pip install audiosr==0.0.7`）

## 使用方法

### GUIモード

AudioSR.NETをダブルクリックして起動すると、ドラッグ＆ドロップでファイルを変換できるGUIモードが起動します。

1. まず「Pythonホーム」に正しいPythonのインストールディレクトリを設定します
2. 変換に使用するモデルやパラメータを設定します
3. オーディオファイルをウィンドウにドラッグ＆ドロップするか、「ファイル追加」ボタンでファイルを選択します
4. 「処理開始」ボタンをクリックして変換を実行します
5. 変換結果は指定した出力フォルダに保存されます

### コマンドラインモード

```csharp
using AudioSR.NET;

// 単一ファイルの処理
var audioSr = new AudioSrWrapper();
audioSr.ProcessFile("input.wav", "output.wav");

// バッチ処理
audioSr.ProcessBatch(new List<string> { "file1.wav", "file2.wav" }, "output_directory");
```

## ライセンス

このプロジェクトはMITライセンスの下で提供されています。詳細は `LICENSE` を参照してください。

## 連絡先

- 名前: ゆろち
- 連絡先: https://github.com/1llum1n4t1s
