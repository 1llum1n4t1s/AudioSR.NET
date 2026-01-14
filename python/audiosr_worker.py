import json
import sys
import traceback
from typing import Any, Dict


def send(payload: Dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


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

    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

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

            model.super_resolution(input_path, output_path, **kwargs)

            send({"status": "ok", "message": "done"})
        except Exception as exc:  # noqa: BLE001
            trace = traceback.format_exc()
            send({"status": "error", "message": str(exc), "traceback": trace})

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
