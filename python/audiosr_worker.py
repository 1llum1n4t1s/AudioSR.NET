import json
import logging
import os
import sys
import traceback
from typing import Any, Dict


# stdout を一時的に保存し、sys.stdout を stderr にリダイレクトして
# 外部ライブラリの print 等が JSON 通信を汚染しないようにする
_original_stdout = sys.stdout
sys.stdout = sys.stderr


def send(payload: Dict[str, Any]) -> None:
    """JSON 形式で結果を stdout (保存しておいた元の方) に出力する"""
    _original_stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    _original_stdout.flush()


def load_model(audiosr_module, requested_model: str, requested_device: str):
    try:
        return audiosr_module.build_model(model_name=requested_model, device=requested_device)
    except TypeError:
        return audiosr_module.build_model(requested_model, requested_device)


def main() -> int:
    audiosr_module = None
    model = None
    model_name = None
    device = None

    # Hugging Face 関連の環境変数を設定してダウンロードを安定させる
    # Windows でのシンボリックリンク作成エラーを回避
    os.environ["HF_HUB_DISABLE_SYMLINKS"] = "1"
    # 進捗バー (tqdm) の出力を強制する
    os.environ["TQDM_MININTERVAL"] = "1"
    # バッファリングを無効化
    os.environ["PYTHONUNBUFFERED"] = "1"

    # エンコーディングの再設定
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    # _original_stdout も再設定
    if hasattr(_original_stdout, "reconfigure"):
        _original_stdout.reconfigure(encoding="utf-8")

    # ロギングの設定（すべて stderr に流す）
    logging.basicConfig(level=logging.INFO, stream=sys.stderr)

    for line in sys.stdin:
        if not line:
            break
        line = line.strip()
        if not line:
            continue

        try:
            message = json.loads(line)
        except json.JSONDecodeError as exc:
            send({"status": "error", "message": f"JSON解析に失敗しました: {exc}"})
            continue

        command = message.get("command")
        if command == "ping":
            send({"status": "ok", "message": "ready"})
            continue

        if command == "shutdown":
            send({"status": "ok", "message": "shutdown"})
            break

        if command != "process":
            send({"status": "error", "message": f"未対応のコマンドです: {command}"})
            continue

        try:
            if audiosr_module is None:
                import audiosr as audiosr_module  # type: ignore
                import torchaudio
                import soundfile as sf
                import torch

                # torchaudio.load が FFmpeg や torchcodec を要求して失敗するのを防ぐため、
                # soundfile を使用するようにモンキーパッチを当てる
                def patched_load(filepath, **kwargs):
                    data, samplerate = sf.read(filepath, dtype="float32")
                    # numpy array を torch tensor (C, T) に変換
                    # torchaudio.load は float32 を返すため、dtype を合わせる
                    tensor = torch.from_numpy(data)
                    if len(tensor.shape) == 1:
                        # (T,) -> (1, T)
                        tensor = tensor.unsqueeze(0)
                    else:
                        # (T, C) -> (C, T)
                        tensor = tensor.T
                    return tensor, samplerate

                torchaudio.load = patched_load

            requested_model = message.get("model_name")
            requested_device = message.get("device")
            if not requested_model or not requested_device:
                raise ValueError("model_name または device が指定されていません")

            if model is None or model_name != requested_model or device != requested_device:
                model = load_model(audiosr_module, requested_model, requested_device)
                model_name = requested_model
                device = requested_device

            input_path = message.get("input")
            output_path = message.get("output")
            if not input_path or not output_path:
                raise ValueError("input または output が指定されていません")

            kwargs: Dict[str, Any] = {
                "ddim_steps": int(message.get("ddim_steps", 100)),
                "guidance_scale": float(message.get("guidance_scale", 3.5)),
            }
            if "seed" in message:
                kwargs["seed"] = int(message["seed"])

            # 推論実行
            # audiosr.super_resolution 関数を使用する（model のメソッドではない）
            # 返り値は (sampling_rate, waveform)
            sr_rate, waveform = audiosr_module.super_resolution(model, input_path, **kwargs)

            # 結果を保存（soundfile パッケージが必要）
            import soundfile as sf
            # waveform が torch tensor の場合は numpy に変換する
            if hasattr(waveform, "numpy"):
                waveform = waveform.cpu().numpy()
            # waveform は通常 (1, T) または (T,) の形式
            # soundfile.write は (T, C) を期待するので、必要に応じて転置する
            if len(waveform.shape) > 1 and waveform.shape[0] < waveform.shape[1]:
                waveform = waveform.T

            sf.write(output_path, waveform, sr_rate)

            send({"status": "ok", "message": "done"})
        except Exception as exc:  # noqa: BLE001
            trace = traceback.format_exc()
            send({"status": "error", "message": str(exc), "traceback": trace})

    return 0


if __name__ == "__main__":
    # システムエラー終了時も 0 を返さないようにする
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
