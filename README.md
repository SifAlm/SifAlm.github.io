# Offline Face Matcher

A ready-to-run Tkinter application that detects faces in photos using
`facenet-pytorch`'s MTCNN detector and highlights people who match a selected
face across a collection of images. The app works entirely offline once model
weights are available locally.

## Features

- Select a single **target** image and multiple **candidate** images using a GUI.
- Detect faces in all photos and compute facial embeddings locally.
- Compare candidate faces to the target face and annotate matches.
- Save annotated copies to a chosen output directory.

## Requirements

- Python 3.8 or newer.
- System packages that provide Tk support (often included with Python on Windows/macOS; install `python3-tk` on Linux if needed).
- Python dependencies listed in `requirements.txt`.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\\Scripts\\activate
pip install -r requirements.txt
```

> **Note:** The first time you run the program it will download pretrained
> weights for the face detector/encoder (about 100 MB). After the download is
> cached the application works fully offline.

## Usage

```bash
python face_matcher_app.py
```

1. Click **Browse** next to *Target image* and choose the photo that contains the
   person you want to match.
2. Click **Add images** and choose one or more candidate photos.
3. (Optional) Adjust the similarity threshold. Higher values (e.g., `0.7`) are
   stricter; lower values (e.g., `0.5`) allow more matches.
4. Pick an output directory where annotated copies should be saved. The default
   `annotated_output` folder will be created automatically.
5. Click **Match Faces**. Progress and results appear in the log window. When a
   match is found an annotated PNG is written to the output directory.

## How It Works

1. Faces are detected using MTCNN.
2. The face with the highest detection probability in the target image is used
   to compute an embedding with InceptionResnetV1.
3. Each candidate face is compared to the target embedding using cosine
   similarity.
4. Faces that meet or exceed the similarity threshold are highlighted with green
   bounding boxes and labeled with their similarity score.

## Limitations

- Only the most confident face in the target image is used for matching.
- Lighting, pose, and occlusion can affect similarity scores. Adjust the
  threshold to fine-tune results.
- Model downloads require an initial internet connection; afterwards the cached
  weights allow offline operation.
