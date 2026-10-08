"""Client for an Ollama server, local or reached through an ngrok link.

Configuration (command line flags win over environment):
  OLLAMA_URL    full base URL, e.g. https://abcd1234.ngrok-free.dev or http://localhost:11434
  OLLAMA_MODEL  model name, default qwen3.5:9b (the model the team runs)
No cloud service is involved; the only network call is to this server.
"""
import json
import os
import re
import time
import urllib.error
import urllib.request

DEFAULT_URL = "http://localhost:11434"
DEFAULT_MODEL = "qwen3.5:9b"


class LLMError(Exception):
    pass


def _strip_fences(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    return re.sub(r"\s*```$", "", text).strip()


class OllamaClient:
    def __init__(self, url=None, model=None, timeout=900, num_ctx=None, seed=7):
        url = url or os.environ.get("OLLAMA_URL")
        if not url and os.environ.get("OLLAMA_HOST"):  # same variable the lesson generator uses
            url = "https://" + os.environ["OLLAMA_HOST"]
        self.url = (url or DEFAULT_URL).rstrip("/")
        self.model = model or os.environ.get("OLLAMA_MODEL") or DEFAULT_MODEL
        self.timeout, self.seed = timeout, seed
        # 8192 tokens hold ~12000 characters of Urdu (measured 2.35 chars/token) plus the answer, and keep a 9B
        # model fully on an 8 GB GPU; a larger context spilled part of it to the CPU.
        self.num_ctx = num_ctx or int(os.environ.get("OLLAMA_NUM_CTX", "8192"))
        self.calls = 0
        self.seconds = 0.0

    @property
    def name(self) -> str:
        return self.model

    def _request(self, path, payload=None, timeout=None):
        data = json.dumps(payload).encode("utf8") if payload is not None else None
        req = urllib.request.Request(self.url + path, data=data, method="POST" if data else "GET", headers={
            "Content-Type": "application/json",
            "ngrok-skip-browser-warning": "true",   # skips ngrok's free-tier interstitial page
        })
        try:
            with urllib.request.urlopen(req, timeout=timeout or self.timeout) as r:
                body = r.read().decode("utf8")
        except urllib.error.HTTPError as e:
            raise LLMError(f"{self.url}{path} returned HTTP {e.code}") from e
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            raise LLMError(f"could not reach {self.url}: {e}") from e
        try:
            return json.loads(body)
        except json.JSONDecodeError as e:
            # an offline ngrok tunnel answers with an HTML error page
            raise LLMError(f"{self.url}{path} did not return JSON (is the ngrok tunnel up?)") from e

    def available(self):
        """(ok, message). Checks the server answers and the model is installed."""
        try:
            tags = self._request("/api/tags", timeout=15)
        except LLMError as e:
            return False, str(e)
        names = [m.get("name", "") for m in tags.get("models", [])]
        if not any(n == self.model or n.split(":")[0] == self.model.split(":")[0] and n.startswith(self.model) for n in names):
            return False, f"model '{self.model}' is not installed on {self.url} (found: {', '.join(names[:6]) or 'none'})"
        return True, f"{self.model} at {self.url}"

    def chat_vision(self, prompt: str, image_b64: str) -> str:
        """Plain-text answer about one image (for OCR with a vision model)."""
        payload = {
            "model": self.model, "stream": False, "think": False,
            "options": {"temperature": 0, "seed": self.seed, "num_ctx": self.num_ctx},
            "messages": [{"role": "user", "content": prompt, "images": [image_b64]}],
        }
        t0 = time.time()
        reply = self._request("/api/chat", payload)
        self.calls += 1
        self.seconds += time.time() - t0
        return ((reply.get("message") or {}).get("content") or "").strip()

    def chat_json(self, system: str, user: str) -> dict:
        payload = {
            "model": self.model, "stream": False, "format": "json", "think": False,
            "options": {"temperature": 0, "seed": self.seed, "num_ctx": self.num_ctx},
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
        }
        t0 = time.time()
        reply = self._request("/api/chat", payload)
        self.calls += 1
        self.seconds += time.time() - t0
        content = (reply.get("message") or {}).get("content", "")
        try:
            return json.loads(_strip_fences(content))
        except json.JSONDecodeError as e:
            raise LLMError(f"model did not return valid JSON: {content[:120]!r}") from e
