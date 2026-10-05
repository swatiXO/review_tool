"""Stand-in models and a tiny fake Ollama server, so the fallback is testable offline."""
import json
import re
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

from course_review.llm import LLMError

LISTING = re.compile(r"^\[(\d+)\] (.*)$")
ROMAN = re.compile(r"^(?:i|ii|iii|iv|v|vi|vii|viii|ix|x)[.)]\s", re.I)


class _Base:
    model = "fake-model"
    url = "fake://model"
    name = "fake-model"

    def __init__(self):
        self.calls = 0
        self.seconds = 0.0
        self.prompts = []

    def paragraphs(self, user):
        out = []
        for line in user.splitlines():
            m = LISTING.match(line)
            if m:
                out.append((int(m.group(1)), m.group(2)))
        return out


class RomanModel(_Base):
    """A 'correct' model for roman-numeral questions, which the rule-based parser does not know."""

    def chat_json(self, system, user):
        self.calls += 1
        self.prompts.append(user)
        qs = [{"para": i, "quote": t[:12]} for i, t in self.paragraphs(user) if ROMAN.match(t)]
        return {"questions": qs}


class HallucinatingModel(_Base):
    """Proposes paragraphs that do not exist and quotes that are not in the paragraph."""

    def chat_json(self, system, user):
        self.calls += 1
        paras = self.paragraphs(user)
        real = [{"para": paras[0][0], "quote": paras[0][1][:10]}] if paras else []
        return {"questions": real + [{"para": 9999, "quote": "invented"}, {"para": paras[-1][0], "quote": "not in the text"}]}


class DownModel(_Base):
    def chat_json(self, system, user):
        self.calls += 1
        raise LLMError("could not reach fake://model: connection refused")


class BadShapeModel(_Base):
    def chat_json(self, system, user):
        self.calls += 1
        return {"answer": "there are six questions"}


class FakeOllamaHandler(BaseHTTPRequestHandler):
    models = ["qwen3:14b"]
    chat_reply = '{"questions": []}'
    html_page = False
    seen = []

    def log_message(self, *a):
        pass

    def _send(self, body, status=200, ctype="application/json"):
        data = body.encode("utf8")
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.headers.get("ngrok-skip-browser-warning"):
            FakeOllamaHandler.seen.append(("GET", self.path, dict(self.headers), None))
        if FakeOllamaHandler.html_page:
            return self._send("<!DOCTYPE html><html>tunnel offline</html>", ctype="text/html")
        if self.path == "/api/tags":
            return self._send(json.dumps({"models": [{"name": m} for m in FakeOllamaHandler.models]}))
        self._send("{}", 404)

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(n) or b"{}")
        FakeOllamaHandler.seen.append(("POST", self.path, dict(self.headers), payload))
        if FakeOllamaHandler.html_page:
            return self._send("<html>offline</html>", ctype="text/html")
        if self.path == "/api/chat":
            return self._send(json.dumps({"message": {"role": "assistant", "content": FakeOllamaHandler.chat_reply}}))
        self._send("{}", 404)


def start_fake_server():
    FakeOllamaHandler.models = ["qwen3:14b"]
    FakeOllamaHandler.chat_reply = '{"questions": []}'
    FakeOllamaHandler.html_page = False
    FakeOllamaHandler.seen = []
    server = HTTPServer(("127.0.0.1", 0), FakeOllamaHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server, f"http://127.0.0.1:{server.server_address[1]}"
