"""EchoID single-file application.

This script provides a self-contained FastAPI-based application that
implements a simplified version of the EchoID prototype. All logic for
image ingestion, embedding, matching, explainability previews, audit
logging, and export lives inside this file so it can be downloaded and
run directly with Python 3.10+.

The implementation favors lightweight numpy/pillow processing to avoid
large model downloads while still demonstrating the required workflow:
  * Upload support (Set A) and probe (Set B) images via the web UI.
  * Generate embeddings that mix color statistics and edge responses.
  * Build identity prototypes and compute cosine similarity matches.
  * Render a pseudo-saliency heat map per match using simple occlusion.
  * Offer manual confirmation, audit logging, data deletion, and export.

Run with: ``python echoid.py`` then open http://127.0.0.1:8000/.
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import shutil
import tempfile
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from pydantic import BaseModel
import uvicorn

APP_TITLE = "EchoID — Custom UI by Sif Almaghrabi"
DATA_DIR = Path("echoid_data")
SUPPORT_DIR = DATA_DIR / "set_a"
PROBE_DIR = DATA_DIR / "set_b"
HEATMAP_DIR = DATA_DIR / "heatmaps"
AUDIT_LOG = DATA_DIR / "audit_log.jsonl"
EXPORT_DIR = DATA_DIR / "exports"
SECRET_KEY = os.environ.get("ECHOID_HMAC_KEY", "echoid-dev-secret")

for directory in (DATA_DIR, SUPPORT_DIR, PROBE_DIR, HEATMAP_DIR, EXPORT_DIR):
    directory.mkdir(parents=True, exist_ok=True)


def log_event(event_type: str, payload: Dict) -> None:
    AUDIT_LOG.parent.mkdir(parents=True, exist_ok=True)
    entry = {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "event": event_type,
        "payload": payload,
    }
    message = json.dumps(entry, sort_keys=True)
    signature = hmac.new(SECRET_KEY.encode(), message.encode(), hashlib.sha256).hexdigest()
    entry["hmac"] = signature
    with AUDIT_LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry) + "\n")


def wipe_directory(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


@dataclass
class ImageEntry:
    image_id: str
    identity: Optional[str]
    filename: str
    path: Path
    consent: bool
    metadata: Dict[str, str] = field(default_factory=dict)


class EchoIdStore:
    """Simple in-memory cache backed by the filesystem."""

    def __init__(self) -> None:
        self.support: Dict[str, List[ImageEntry]] = {}
        self.probes: Dict[str, ImageEntry] = {}
        self.embeddings: Dict[str, np.ndarray] = {}
        self.prototypes: Dict[str, np.ndarray] = {}
        self.match_results: Dict[str, Dict] = {}
        self.lock = threading.Lock()
        self.embed_status = "idle"
        self.embed_progress = 0

    # Utility methods -------------------------------------------------
    def _gen_id(self, prefix: str) -> str:
        raw = f"{prefix}-{time.time()}-{os.urandom(6).hex()}"
        return hashlib.sha1(raw.encode()).hexdigest()[:12]

    def add_support_image(
        self,
        file: UploadFile,
        identity: str,
        consent: bool,
        metadata: Dict[str, str],
    ) -> ImageEntry:
        image_id = self._gen_id("support")
        filename = f"{image_id}_{file.filename}"
        dest = SUPPORT_DIR / filename
        with dest.open("wb") as f:
            shutil.copyfileobj(file.file, f)
        entry = ImageEntry(image_id, identity, filename, dest, consent, metadata)
        with self.lock:
            self.support.setdefault(identity or "unlabeled", []).append(entry)
        log_event("support_upload", {
            "image_id": image_id,
            "identity": identity,
            "consent": consent,
            "metadata": metadata,
        })
        return entry

    def add_probe_image(self, file: UploadFile) -> ImageEntry:
        image_id = self._gen_id("probe")
        filename = f"{image_id}_{file.filename}"
        dest = PROBE_DIR / filename
        with dest.open("wb") as f:
            shutil.copyfileobj(file.file, f)
        entry = ImageEntry(image_id, None, filename, dest, False, {})
        with self.lock:
            self.probes[image_id] = entry
        log_event("probe_upload", {"image_id": image_id})
        return entry

    def reset(self) -> None:
        with self.lock:
            self.support.clear()
            self.probes.clear()
            self.embeddings.clear()
            self.prototypes.clear()
            self.match_results.clear()
            self.embed_status = "idle"
            self.embed_progress = 0

    # Embedding utilities ---------------------------------------------
    def _preprocess(self, img: Image.Image) -> Image.Image:
        img = img.convert("RGB")
        img = ImageOps.fit(img, (224, 224))
        img = ImageEnhance.Color(img).enhance(1.1)
        img = ImageEnhance.Contrast(img).enhance(1.1)
        return img

    def _extract_features(self, img: Image.Image) -> np.ndarray:
        arr = np.asarray(img).astype("float32") / 255.0
        mean_rgb = arr.mean(axis=(0, 1))
        std_rgb = arr.std(axis=(0, 1))
        edges = img.convert("L").filter(ImageFilter.FIND_EDGES)
        edge_arr = np.asarray(edges).astype("float32") / 255.0
        hist, _ = np.histogram(edge_arr, bins=32, range=(0.0, 1.0), density=True)
        hog_like = self._simple_hog(edge_arr)
        feature = np.concatenate([
            mean_rgb,
            std_rgb,
            hist,
            hog_like,
        ])
        return feature / np.linalg.norm(feature)

    def _simple_hog(self, edge_arr: np.ndarray) -> np.ndarray:
        gx = np.gradient(edge_arr, axis=1)
        gy = np.gradient(edge_arr, axis=0)
        magnitude = np.sqrt(gx ** 2 + gy ** 2)
        orientation = (np.arctan2(gy, gx) + np.pi) / (2 * np.pi)
        cells = []
        grid = 8
        h, w = edge_arr.shape
        cell_h = h // grid
        cell_w = w // grid
        for y in range(grid):
            for x in range(grid):
                block_mag = magnitude[y * cell_h:(y + 1) * cell_h, x * cell_w:(x + 1) * cell_w]
                block_ori = orientation[y * cell_h:(y + 1) * cell_h, x * cell_w:(x + 1) * cell_w]
                hist, _ = np.histogram(block_ori, bins=8, range=(0.0, 1.0), weights=block_mag)
                cells.append(hist)
        vec = np.concatenate(cells)
        norm = np.linalg.norm(vec) or 1.0
        return vec / norm

    def compute_embeddings(self) -> None:
        with self.lock:
            all_entries = []
            for identity, images in self.support.items():
                all_entries.extend(images)
            all_entries.extend(self.probes.values())
            total = len(all_entries)
            if total == 0:
                self.embed_status = "idle"
                self.embed_progress = 0
                return
            self.embed_status = "running"
            self.embed_progress = 0
        for index, entry in enumerate(all_entries, start=1):
            img = Image.open(entry.path)
            img = self._preprocess(img)
            feature = self._extract_features(img)
            with self.lock:
                self.embeddings[entry.image_id] = feature
                self.embed_progress = int((index / total) * 100)
        with self.lock:
            self.embed_status = "complete"
        self.build_prototypes()

    def build_prototypes(self) -> None:
        prototypes: Dict[str, np.ndarray] = {}
        with self.lock:
            for identity, images in self.support.items():
                feats = [self.embeddings[i.image_id] for i in images if i.image_id in self.embeddings]
                if not feats:
                    continue
                proto = np.mean(feats, axis=0)
                proto = proto / (np.linalg.norm(proto) or 1.0)
                prototypes[identity] = proto
            self.prototypes = prototypes

    # Matching ---------------------------------------------------------
    def cosine(self, a: np.ndarray, b: np.ndarray) -> float:
        return float(np.dot(a, b) / ((np.linalg.norm(a) or 1.0) * (np.linalg.norm(b) or 1.0)))

    def sigmoid_confidence(self, score: float, temperature: float = 0.75) -> float:
        return float(1.0 / (1.0 + np.exp(-score / temperature)))

    def match(self, top_k: int = 3, threshold: float = 0.4) -> Dict[str, Dict]:
        with self.lock:
            probes = list(self.probes.values())
            prototypes = self.prototypes.copy()
        if not probes or not prototypes:
            raise HTTPException(status_code=400, detail="Need both probes and support embeddings")
        results = {}
        for probe in probes:
            probe_feat = self.embeddings.get(probe.image_id)
            if probe_feat is None:
                continue
            scores: List[Tuple[str, float, float]] = []
            for identity, proto in prototypes.items():
                score = self.cosine(probe_feat, proto)
                confidence = self.sigmoid_confidence(score)
                scores.append((identity, score, confidence))
            scores.sort(key=lambda x: x[1], reverse=True)
            accepted = [s for s in scores[:top_k] if s[2] >= threshold]
            best = accepted[0] if accepted else None
            saliency_path = self._make_saliency(probe.path, best[0] if best else None)
            result = {
                "probe_id": probe.image_id,
                "probe_filename": probe.filename,
                "candidates": [
                    {
                        "identity": identity,
                        "score": score,
                        "confidence": confidence,
                    }
                    for identity, score, confidence in scores[:top_k]
                ],
                "suggested_identity": best[0] if best else None,
                "suggested_confidence": best[2] if best else None,
                "heatmap": saliency_path.name if saliency_path else None,
            }
            results[probe.image_id] = result
        with self.lock:
            self.match_results = results
        log_event("match_run", {"count": len(results), "threshold": threshold})
        return results

    def _make_saliency(self, probe_path: Path, identity: Optional[str]) -> Optional[Path]:
        if identity is None:
            return None
        ref_images = self.support.get(identity)
        if not ref_images:
            return None
        probe_img = Image.open(probe_path).convert("RGB").resize((224, 224))
        ref_img = Image.open(ref_images[0].path).convert("RGB").resize((224, 224))
        probe_arr = np.asarray(probe_img).astype("float32")
        ref_arr = np.asarray(ref_img).astype("float32")
        diff = np.abs(probe_arr - ref_arr).mean(axis=2)
        diff = diff / (diff.max() or 1.0)
        diff = (diff * 255).astype("uint8")
        heat = Image.fromarray(diff, mode="L").resize(probe_img.size)
        heat = heat.filter(ImageFilter.GaussianBlur(radius=6))
        heat = ImageEnhance.Color(heat.convert("RGB")).enhance(2.5)
        overlay = Image.blend(probe_img, heat, alpha=0.5)
        dest = HEATMAP_DIR / f"{probe_path.stem}_heatmap.png"
        overlay.save(dest)
        return dest


store = EchoIdStore()
app = FastAPI(title=APP_TITLE)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/", response_class=HTMLResponse)
def index() -> str:
    with AUDIT_LOG.open("a", encoding="utf-8"):
        pass
    return f"""
    <!DOCTYPE html>
    <html lang=\"en\">
    <head>
        <meta charset=\"utf-8\">
        <title>{APP_TITLE}</title>
        <style>
            body {{
                background: radial-gradient(circle at top left, #05152a, #020810);
                color: #c8f2ff;
                font-family: 'Segoe UI', sans-serif;
                margin: 0;
                padding: 0;
                min-height: 100vh;
                display: flex;
                justify-content: center;
                align-items: flex-start;
            }}
            .container {{
                width: 90%;
                max-width: 1080px;
                margin-top: 40px;
                background: rgba(2, 20, 40, 0.85);
                border-radius: 16px;
                padding: 24px 32px;
                box-shadow: 0 30px 60px rgba(0, 0, 0, 0.4);
                border: 1px solid rgba(0, 200, 255, 0.15);
            }}
            h1 {{
                text-align: center;
                letter-spacing: 0.08em;
                text-transform: uppercase;
            }}
            .credit {{
                text-align: right;
                font-size: 12px;
                color: #58c4ff;
            }}
            .panel {{
                margin-top: 24px;
                padding: 16px;
                border-radius: 12px;
                background: rgba(6, 34, 60, 0.6);
                border: 1px solid rgba(88, 210, 255, 0.15);
            }}
            input, button, select {{
                background: rgba(12, 60, 90, 0.85);
                border: 1px solid rgba(120, 220, 255, 0.25);
                border-radius: 8px;
                color: #e6fbff;
                padding: 10px 12px;
                margin-top: 8px;
                width: 100%;
            }}
            button {{
                cursor: pointer;
                font-weight: 600;
            }}
            .grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
                gap: 18px;
            }}
            .results {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
                gap: 16px;
            }}
            .card {{
                padding: 12px;
                background: rgba(4, 24, 44, 0.6);
                border-radius: 12px;
            }}
            .heatmap {{
                width: 100%;
                border-radius: 12px;
            }}
            .warning {{
                padding: 12px;
                background: rgba(120, 20, 40, 0.4);
                border: 1px solid rgba(255, 90, 110, 0.5);
                border-radius: 10px;
                margin-bottom: 12px;
            }}
            a {{
                color: #7be1ff;
            }}
        </style>
    </head>
    <body>
        <div class=\"container\">
            <h1>{APP_TITLE}</h1>
            <p class=\"credit\">Custom UI by Sif Almaghrabi</p>
            <div class=\"warning\">
                <strong>Responsible Use Reminder:</strong> EchoID is a research/prototype tool. It uses biometric-like matching that can reveal sensitive information. Use only with explicit, informed consent from the individuals represented. Do not use EchoID for covert surveillance, law enforcement, or other activities that may violate human rights or local laws. The authors disclaim responsibility for misuse.
            </div>
            <div class=\"grid\">
                <div class=\"panel\">
                    <h2>Upload Set A (Knowns)</h2>
                    <form id=\"supportForm\" enctype=\"multipart/form-data\">
                        <label>Identity Label</label>
                        <input type=\"text\" name=\"identity\" placeholder=\"e.g., Person A\" required>
                        <label>Consent Checkbox</label>
                        <select name=\"consent\" required>
                            <option value=\"true\">Consent Granted</option>
                            <option value=\"false\">Consent Not Granted</option>
                        </select>
                        <label>Optional Metadata (Location)</label>
                        <input type=\"text\" name=\"location\" placeholder=\"Laboratory\">
                        <label>Optional Metadata (Timestamp)</label>
                        <input type=\"text\" name=\"timestamp\" placeholder=\"2024-03-15T10:00Z\">
                        <label>Upload Images</label>
                        <input type=\"file\" name=\"files\" accept=\"image/*\" multiple required>
                        <button type=\"submit\">Upload to Set A</button>
                    </form>
                </div>
                <div class=\"panel\">
                    <h2>Upload Set B (Probes)</h2>
                    <form id=\"probeForm\" enctype=\"multipart/form-data\">
                        <label>Upload Probes</label>
                        <input type=\"file\" name=\"files\" accept=\"image/*\" multiple required>
                        <button type=\"submit\">Upload to Set B</button>
                    </form>
                </div>
            </div>
            <div class=\"panel\">
                <h2>Embedding & Matching</h2>
                <button onclick=\"runEmbedding()\">Compute Embeddings</button>
                <div style=\"margin-top:12px\">
                    <label>Confidence Threshold</label>
                    <input type=\"number\" id=\"threshold\" min=\"0\" max=\"1\" step=\"0.05\" value=\"0.4\">
                    <label>Top K</label>
                    <input type=\"number\" id=\"topk\" min=\"1\" max=\"5\" value=\"3\">
                    <button onclick=\"runMatch()\" style=\"margin-top:12px\">Match Probes</button>
                </div>
                <div id=\"status\" style=\"margin-top:12px;color:#9be7ff\"></div>
            </div>
            <div class=\"panel\">
                <h2>Match Results</h2>
                <div id=\"results\" class=\"results\"></div>
            </div>
            <div class=\"panel\">
                <h2>Utilities</h2>
                <button onclick=\"exportResults()\">Export CSV + Heatmaps</button>
                <button onclick=\"confirmAll()\" style=\"margin-top:8px\">Confirm Suggested Matches</button>
                <button onclick=\"deleteAll()\" style=\"margin-top:8px\">Clear All Data</button>
                <p style=\"margin-top:12px;font-size:12px;\">Audit log stored locally in <code>echoid_data/audit_log.jsonl</code>.</p>
            </div>
        </div>
        <script>
            async function handleUpload(formId, url) {{
                const form = document.getElementById(formId);
                const data = new FormData(form);
                const response = await fetch(url, {{ method: 'POST', body: data }});
                if (!response.ok) {{
                    alert('Upload failed: ' + await response.text());
                    return;
                }}
                alert('Uploaded successfully');
                form.reset();
            }}
            document.getElementById('supportForm').addEventListener('submit', function(e) {{
                e.preventDefault();
                handleUpload('supportForm', '/upload_set');
            }});
            document.getElementById('probeForm').addEventListener('submit', function(e) {{
                e.preventDefault();
                handleUpload('probeForm', '/upload_probe');
            }});
            async function runEmbedding() {{
                document.getElementById('status').innerText = 'Embedding running...';
                const res = await fetch('/compute_embeddings', {{ method: 'POST' }});
                const data = await res.json();
                document.getElementById('status').innerText = data.message;
            }}
            async function runMatch() {{
                const threshold = parseFloat(document.getElementById('threshold').value || '0.4');
                const topk = parseInt(document.getElementById('topk').value || '3', 10);
                const body = JSON.stringify({{ threshold: threshold, top_k: topk }});
                const res = await fetch('/match', {{ method: 'POST', headers: {{'Content-Type': 'application/json'}}, body }});
                if (!res.ok) {{
                    alert('Match failed: ' + await res.text());
                    return;
                }}
                const data = await res.json();
                renderResults(data.results);
            }}
            function renderResults(results) {{
                const container = document.getElementById('results');
                container.innerHTML = '';
                Object.values(results).forEach(result => {{
                    const card = document.createElement('div');
                    card.className = 'card';
                    const title = document.createElement('h3');
                    title.innerText = 'Probe';
                    if (result && result.probe_id) {{
                        title.innerText = 'Probe ' + result.probe_id;
                    }}
                    card.appendChild(title);
                    const list = document.createElement('ul');
                    result.candidates.forEach(c => {{
                        const item = document.createElement('li');
                        item.innerText = `${{c.identity}} — score ${{c.score.toFixed(3)}} (confidence ${{(c.confidence*100).toFixed(1)}}%)`;
                        list.appendChild(item);
                    }});
                    card.appendChild(list);
                    if (result.heatmap) {{
                        const img = document.createElement('img');
                        img.src = `/heatmap/${{result.heatmap}}`;
                        img.className = 'heatmap';
                        card.appendChild(img);
                    }}
                    container.appendChild(card);
                }});
            }}
            async function exportResults() {{
                const res = await fetch('/export', {{ method: 'POST' }});
                if (!res.ok) {{
                    alert('Export failed: ' + await res.text());
                    return;
                }}
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'echoid_export.zip';
                a.click();
                window.URL.revokeObjectURL(url);
            }}
            async function confirmAll() {{
                const res = await fetch('/confirm', {{ method: 'POST' }});
                const data = await res.json();
                alert(data.message);
            }}
            async function deleteAll() {{
                if (!confirm('This will delete all stored data. Continue?')) return;
                const res = await fetch('/delete_all', {{ method: 'POST' }});
                const data = await res.json();
                document.getElementById('results').innerHTML = '';
                document.getElementById('status').innerText = '';
                alert(data.message);
            }}
        </script>
    </body>
    </html>
    """


@app.get("/health")
def health() -> Dict[str, str]:
    """Lightweight health check used by monitors and tests."""
    return {"status": "ok"}


@app.post("/upload_set")
async def upload_set(
    files: List[UploadFile] = File(...),
    identity: str = Form(...),
    consent: str = Form(...),
    location: Optional[str] = Form(None),
    timestamp: Optional[str] = Form(None),
):
    if consent.lower() != "true":
        raise HTTPException(status_code=400, detail="Consent is required for Set A uploads")
    metadata = {"location": location or "", "timestamp": timestamp or ""}
    entries = []
    for file in files:
        if not file.filename:
            continue
        entry = store.add_support_image(file, identity, True, metadata)
        entries.append(entry.image_id)
    return {"uploaded": entries}


@app.post("/upload_probe")
async def upload_probe(files: List[UploadFile] = File(...)):
    entries = []
    for file in files:
        if not file.filename:
            continue
        entry = store.add_probe_image(file)
        entries.append(entry.image_id)
    return {"uploaded": entries}


@app.post("/compute_embeddings")
async def compute_embeddings():
    thread = threading.Thread(target=store.compute_embeddings)
    thread.start()
    thread.join()
    return {"message": "Embeddings computed", "status": store.embed_status}


class MatchRequest(BaseModel):
    top_k: int = 3
    threshold: float = 0.4


@app.post("/match")
async def match(req: MatchRequest):
    results = store.match(top_k=req.top_k, threshold=req.threshold)
    return {"results": results}


@app.get("/match_status")
async def match_status():
    return {"status": store.embed_status, "progress": store.embed_progress}


@app.post("/confirm")
async def confirm():
    confirmed = []
    for probe_id, info in store.match_results.items():
        identity = info.get("suggested_identity")
        if identity:
            confirmed.append({"probe_id": probe_id, "identity": identity})
    log_event("manual_confirm", {"confirmed": confirmed})
    return {"message": f"Confirmed {len(confirmed)} matches", "confirmed": confirmed}


@app.post("/export")
async def export():
    if not store.match_results:
        raise HTTPException(status_code=400, detail="Run matching before exporting")
    timestamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    export_base = EXPORT_DIR / f"echoid_export_{timestamp}"
    csv_lines = ["probe_id,predicted_identity,score,confidence,timestamp"]
    now = datetime.utcnow().isoformat() + "Z"
    for probe_id, info in store.match_results.items():
        best = info.get("candidates", [{}])[0]
        csv_lines.append(
            f"{probe_id},{best.get('identity','')},{best.get('score',0):.4f},{best.get('confidence',0):.4f},{now}"
        )
    csv_bytes = "\n".join(csv_lines).encode("utf-8")
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir = Path(tmpdir)
        results_file = temp_dir / "results.csv"
        results_file.write_bytes(csv_bytes)
        heatmap_dir = temp_dir / "heatmaps"
        heatmap_dir.mkdir()
        for file in HEATMAP_DIR.glob("*.png"):
            shutil.copy(file, heatmap_dir / file.name)
        archive_path = shutil.make_archive(str(export_base), "zip", temp_dir)
    export_path = Path(archive_path)
    log_event("export", {"path": str(export_path)})
    return FileResponse(export_path, filename="echoid_export.zip")


@app.post("/delete_all")
async def delete_all():
    wipe_directory(DATA_DIR)
    for directory in (SUPPORT_DIR, PROBE_DIR, HEATMAP_DIR, EXPORT_DIR):
        directory.mkdir(parents=True, exist_ok=True)
    AUDIT_LOG.touch(exist_ok=True)
    store.reset()
    log_event("delete_all", {})
    return {"message": "All data cleared"}


@app.get("/heatmap/{filename}")
async def get_heatmap(filename: str):
    path = HEATMAP_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="Heatmap not found")
    return FileResponse(path)


if __name__ == "__main__":
    print(f"Launching {APP_TITLE} on http://127.0.0.1:8000 ...")
    uvicorn.run(app, host="127.0.0.1", port=8000)
