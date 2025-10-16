# EchoID Single-File Edition

This repository now contains a single Python file, `echoid.py`, that implements
an end-to-end prototype of the EchoID workflow. The app runs entirely locally
and serves a lightweight web interface for uploading Set A (support) and Set B
(probe) images, computing embeddings, matching, reviewing saliency maps, and
exporting results.

## Requirements

- Python 3.10 or newer
- pip packages: `fastapi`, `uvicorn[standard]`, `python-multipart`, `pillow`,
  `numpy`

Install dependencies with:

```bash
pip install fastapi "uvicorn[standard]" python-multipart pillow numpy
```

## Usage

```bash
python echoid.py
```

Then open <http://127.0.0.1:8000> in a browser. Upload at least one consented
support image for each identity, add probe images, compute embeddings, and run
matching. You can confirm matches, export CSV + heatmaps, and securely delete
all stored data from the interface.

All temporary files live under `echoid_data/` beside the script. The audit log
(`echoid_data/audit_log.jsonl`) keeps an HMAC-signed record of actions.
