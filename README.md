# Offline Face Matcher

A premium Tkinter experience that lets you curate a reference gallery and hunt
for the same person across massive photo dumps—fully offline. Under the hood it
combines facenet-pytorch's MTCNN detector with an InceptionResnetV1 embedding
network and custom augmentation/ensembling to stay accurate even when your
person ages, changes style, or appears under harsh lighting.

## Headline Features

- **Multiple Pic A photos.** Drop in as many reference shots as you own; the app
  learns an ensemble profile and even adds mirrored variants for resilience.
- **Smart comparison scanning.** Each candidate photo can contain many faces—the
  strongest similarity against the learned profile is tracked and logged.
- **Confidence controls.** Tune minimum detector confidences and similarity
  thresholds to balance recall vs precision for your dataset.
- **Beautiful command center.** Custom-branded interface proudly states "Custom
  made by Sif Almaghrabi", with live progress, scrollable galleries, and a rich
  event feed.
- **Curation tools.** Remove any reference/comparison image after upload or
  clear the whole list in one click.
- **Insightful exports.** Annotated matches land in a timestamped session folder
  (optional) alongside a human-readable summary of every candidate image.

## Requirements

- Python 3.8 or newer.
- System packages that provide Tk support (often included with Python on
  Windows/macOS; install `python3-tk` on Linux if needed).
- Python dependencies listed in `requirements.txt`.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
pip install -r requirements.txt
```

> **Note:** The first time you run the program it will download pretrained
> weights for the face detector/encoder (about 100 MB). After the download is
> cached the application works fully offline.

## Usage

```bash
python face_matcher_app.py
```

1. Click **Add target photos** under *Pic A — Reference gallery* and select one
   or more reference portraits of the person you're tracking.
2. Click **Add comparison photos** to load every candidate image you want to
   scan.
3. Adjust the similarity/confidence controls or leave the recommended defaults
   for a balanced search.
4. Choose your output directory and optionally keep the "Create timestamped
   session folder" toggle on to isolate each run.
5. Press **Launch matching**. Progress updates in real time and a detailed
   summary (plus annotated images) is written to the output folder.

## How It Works

1. MTCNN detects and aligns faces across all reference images, filtering by the
   confidence you specify.
2. Each face—and an optional mirrored augmentation—is embedded using
   InceptionResnetV1. Their normalized vectors are averaged to create an extra
   "ensemble" profile that stabilizes long-term changes in appearance.
3. Candidate faces go through the same embedding pipeline. Cosine similarity vs
   every reference vector is computed; any face beating the similarity
   threshold is annotated and saved.
4. The session summary file records per-image stats, highest scores, and the
   reference photo that each match most closely resembles.

## Tips for Better Results

- Start with several high-quality reference shots spanning different ages,
  angles, and expressions.
- Lower the similarity threshold (e.g. 0.55) if you're missing true matches, or
  raise it (e.g. 0.70) to reduce false positives.
- Increase the minimum confidence filters if you get noisy detections.
- Run everything offline after the first model download—no data ever leaves your
  machine.

## Limitations

- Extremely low-resolution or heavily occluded faces remain challenging.
- The UI processes images sequentially; large batches may take time without GPU
  acceleration.
- Remember that the tool is for organizing personal collections—it performs no
  automated identity verification beyond similarity scoring.
