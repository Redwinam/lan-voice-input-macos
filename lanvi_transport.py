import asyncio
import json
import threading

from flask import Flask, jsonify, request, send_file
from websockets.exceptions import ConnectionClosedOK, ConnectionClosedError, ConnectionClosed
import websockets


class ClientCounter:
    def __init__(self):
        self.count = 0
        self.lock = threading.Lock()

    def inc(self):
        with self.lock:
            self.count += 1
            return self.count

    def dec(self):
        with self.lock:
            self.count -= 1
            return self.count

    def value(self):
        with self.lock:
            return self.count


def create_http_app(resource_path, get_ports, get_qr_url, input_service):
    app = Flask(__name__)

    @app.route("/")
    def index():
        path = resource_path("index.html")
        response = send_file(path)
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

    @app.route("/config")
    def config():
        http_port, ws_port = get_ports()
        return jsonify({"ws_port": ws_port, "http_port": http_port, "url": get_qr_url()})

    @app.route("/health")
    def health():
        http_port, ws_port = get_ports()
        return jsonify({"ok": True, "ws_port": ws_port, "http_port": http_port})

    @app.route("/send", methods=["POST"])
    def send_http():
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            data = {}
        msg_type = (data.get("type") or request.form.get("type") or "text").strip()
        content = data.get("string")
        if content is None:
            content = request.form.get("string")
        text = str(content or "").strip()
        if not text:
            return jsonify({"ok": False, "message": "empty"}), 400

        if msg_type == "cmd":
            result = input_service.execute_command(text)
            output = result.output if isinstance(result.output, dict) else {"ok": False, "message": result.display_text}
            return jsonify({
                "type": "cmd_result",
                "string": text,
                "ok": bool(output.get("ok")),
                "message": output.get("message"),
            })

        input_service.handle_text(text)
        return jsonify({"ok": True})

    return app


def make_ws_handler(input_service, notify, get_ports, client_counter):
    async def ws_handler(websocket):
        c = client_counter.inc()
        http_port, ws_port = get_ports()
        notify("手机已连接", f"连接数：{c}（HTTP:{http_port} WS:{ws_port}）")
        try:
            async for msg in websocket:
                msg = msg.strip()
                if not msg:
                    continue
                print("收到：", msg)
                msg_type = "text"
                content = msg
                if msg.startswith("{"):
                    try:
                        payload = json.loads(msg)
                        if isinstance(payload, dict):
                            msg_type = (payload.get("type") or "text").strip()
                            content = payload.get("string")
                    except Exception:
                        msg_type = "text"
                        content = msg

                if msg_type == "cmd":
                    result = input_service.execute_command(str(content or "").strip())
                    resp = {
                        "type": "cmd_result",
                        "string": str(content or "").strip(),
                        "ok": bool(result.output.get("ok")) if isinstance(result.output, dict) else False,
                        "message": result.output.get("message") if isinstance(result.output, dict) else result.display_text,
                    }
                    await websocket.send(json.dumps(resp, ensure_ascii=False))
                else:
                    input_service.handle_text(str(content or ""))

        except (ConnectionClosedOK, ConnectionClosedError, ConnectionClosed, ConnectionResetError, OSError):
            pass
        finally:
            c = client_counter.dec()
            notify("手机已断开", f"连接数：{c}")

    return ws_handler


def ws_thread_main(ws_handler, ws_port, ping_interval, ping_timeout, ready_evt, set_state):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    async def _start_ws_server():
        return await websockets.serve(
            ws_handler, "0.0.0.0", ws_port,
            ping_interval=ping_interval,
            ping_timeout=ping_timeout,
            max_size=1_000_000,
            max_queue=32,
            compression=None,
        )

    server = loop.run_until_complete(_start_ws_server())
    set_state(loop, server)
    ready_evt.set()
    try:
        loop.run_forever()
    finally:
        try:
            if server:
                server.close()
                loop.run_until_complete(server.wait_closed())
        except Exception:
            pass
        try:
            pending = asyncio.all_tasks(loop)
            for task in pending:
                task.cancel()
            if pending:
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        except Exception:
            pass
        try:
            loop.close()
        except Exception:
            pass
