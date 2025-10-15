"""
Face Matcher Application

This module provides a Tkinter GUI that allows a user to select a target image
and a collection of candidate images. It leverages facenet-pytorch's MTCNN for
face detection and InceptionResnetV1 for embeddings to identify faces in the
candidate images that match the target face. Matched faces are outlined and
saved as annotated copies in the chosen output directory.
"""
from __future__ import annotations

import threading
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

import cv2
import numpy as np
import torch
from facenet_pytorch import InceptionResnetV1, MTCNN
from PIL import Image, ImageDraw, ImageFont
from tkinter import END, Button, DoubleVar, Entry, Label, StringVar, Text, Tk, filedialog, messagebox


@dataclass
class DetectionResult:
    """Stores metadata for a detected face."""

    box: Sequence[float]
    probability: float
    similarity: float


class FaceMatcherApp:
    """Tkinter-based GUI application for matching faces across images."""

    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("Offline Face Matcher")
        self.root.geometry("720x520")

        self.device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        self.mtcnn: Optional[MTCNN] = None
        self.encoder: Optional[InceptionResnetV1] = None
        self.target_embedding: Optional[torch.Tensor] = None

        self.target_path = StringVar()
        self.output_dir = StringVar(value=str(Path.cwd() / "annotated_output"))
        self.threshold = DoubleVar(value=0.6)
        self.candidate_paths: List[str] = []

        self._build_widgets()

    # ------------------------------------------------------------------
    # GUI Construction
    def _build_widgets(self) -> None:
        Label(self.root, text="Target image:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
        Entry(self.root, textvariable=self.target_path, width=60).grid(row=0, column=1, padx=10, pady=(10, 0))
        Button(self.root, text="Browse", command=self._select_target).grid(row=0, column=2, padx=10, pady=(10, 0))

        Label(self.root, text="Candidate images:").grid(row=1, column=0, sticky="nw", padx=10, pady=10)
        self.candidate_display = Text(self.root, height=6, width=60, state="disabled")
        self.candidate_display.grid(row=1, column=1, padx=10, pady=10)
        Button(self.root, text="Add images", command=self._select_candidates).grid(row=1, column=2, padx=10, pady=10)

        Label(self.root, text="Output directory:").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        Entry(self.root, textvariable=self.output_dir, width=60).grid(row=2, column=1, padx=10, pady=10)
        Button(self.root, text="Browse", command=self._select_output_dir).grid(row=2, column=2, padx=10, pady=10)

        Label(self.root, text="Similarity threshold (0-1):").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        Entry(self.root, textvariable=self.threshold, width=10).grid(row=3, column=1, sticky="w", padx=10, pady=10)

        Button(self.root, text="Match Faces", command=self._run_async).grid(row=4, column=0, columnspan=3, pady=10)

        Label(self.root, text="Log:").grid(row=5, column=0, sticky="nw", padx=10)
        self.log = Text(self.root, height=12, width=90, state="disabled")
        self.log.grid(row=5, column=0, columnspan=3, padx=10, pady=(0, 10))

    # ------------------------------------------------------------------
    # GUI Event Handlers
    def _select_target(self) -> None:
        path = filedialog.askopenfilename(title="Select target image", filetypes=self._image_filetypes())
        if path:
            self.target_path.set(path)
            self._log(f"Selected target image: {path}")

    def _select_candidates(self) -> None:
        paths = filedialog.askopenfilenames(title="Select candidate images", filetypes=self._image_filetypes())
        if not paths:
            return
        self.candidate_paths.extend(path for path in paths if path not in self.candidate_paths)
        self._refresh_candidate_display()
        self._log(f"Added {len(paths)} candidate images.")

    def _select_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_dir.set(directory)
            self._log(f"Output directory set to: {directory}")

    def _refresh_candidate_display(self) -> None:
        self.candidate_display.configure(state="normal")
        self.candidate_display.delete("1.0", END)
        for path in self.candidate_paths:
            self.candidate_display.insert(END, f"{path}\n")
        self.candidate_display.configure(state="disabled")

    # ------------------------------------------------------------------
    def _run_async(self) -> None:
        thread = threading.Thread(target=self._run_matching, daemon=True)
        thread.start()

    def _run_matching(self) -> None:
        if not self.target_path.get():
            messagebox.showerror("Missing target", "Please select a target image before running.")
            return
        if not self.candidate_paths:
            messagebox.showerror("Missing candidates", "Please add at least one candidate image.")
            return

        try:
            threshold = float(self.threshold.get())
        except (TypeError, ValueError):
            messagebox.showerror("Invalid threshold", "Similarity threshold must be a numeric value between 0 and 1.")
            return
        if not 0 <= threshold <= 1:
            messagebox.showerror("Invalid threshold", "Similarity threshold must be between 0 and 1.")
            return

        output_dir = Path(self.output_dir.get()).expanduser()
        output_dir.mkdir(parents=True, exist_ok=True)

        self._log("Loading models... This may take a moment.")
        self._load_models()
        if self.mtcnn is None or self.encoder is None:
            messagebox.showerror("Model error", "Failed to initialize face detection or embedding models.")
            return

        self._log("Extracting target face embedding...")
        target_embedding = self._extract_target_embedding(self.target_path.get())
        if target_embedding is None:
            messagebox.showerror("Target not found", "No face detected in the target image.")
            return
        self.target_embedding = target_embedding

        total_matches = 0
        for candidate_path in self.candidate_paths:
            matches = self._process_candidate(candidate_path, output_dir, threshold)
            total_matches += matches

        self._log(f"Finished processing. Total matches found: {total_matches}")

    # ------------------------------------------------------------------
    def _load_models(self) -> None:
        if self.mtcnn is None:
            self.mtcnn = MTCNN(image_size=160, margin=20, keep_all=True, device=self.device)
        if self.encoder is None:
            self.encoder = InceptionResnetV1(pretrained="vggface2").eval().to(self.device)

    def _extract_target_embedding(self, image_path: str) -> Optional[torch.Tensor]:
        assert self.mtcnn is not None
        assert self.encoder is not None

        image = Image.open(image_path).convert("RGB")
        faces, probs = self.mtcnn(image, return_prob=True)
        if faces is None or probs is None or len(probs) == 0:
            self._log("No faces detected in target image.")
            return None

        best_idx = int(np.argmax(probs))
        face_tensor = faces[best_idx].unsqueeze(0).to(self.device)
        with torch.no_grad():
            embedding = self.encoder(face_tensor)
        embedding = torch.nn.functional.normalize(embedding, p=2, dim=1)
        self._log(f"Target face detected with probability {probs[best_idx]:.2f}.")
        return embedding.squeeze(0)

    def _process_candidate(self, image_path: str, output_dir: Path, threshold: float) -> int:
        assert self.mtcnn is not None
        assert self.encoder is not None
        assert self.target_embedding is not None

        image = Image.open(image_path).convert("RGB")
        boxes, probs = self.mtcnn.detect(image)
        if boxes is None or probs is None:
            self._log(f"No faces detected in candidate: {image_path}")
            return 0

        faces_tensor = self.mtcnn.extract(image, boxes, None)
        if faces_tensor is None:
            self._log(f"Failed to align faces for candidate: {image_path}")
            return 0

        faces_tensor = faces_tensor.to(self.device)
        with torch.no_grad():
            embeddings = self.encoder(faces_tensor)
        embeddings = torch.nn.functional.normalize(embeddings, p=2, dim=1)

        similarities = torch.nn.functional.cosine_similarity(embeddings, self.target_embedding.unsqueeze(0))
        matches: List[DetectionResult] = []
        for box, prob, sim in zip(boxes, probs, similarities.cpu().numpy()):
            if prob is None:
                continue
            if sim >= threshold:
                matches.append(DetectionResult(box=box, probability=float(prob), similarity=float(sim)))

        if matches:
            annotated_path = self._annotate_and_save(image, matches, image_path, output_dir)
            self._log(
                f"{len(matches)} match(es) found in {image_path}. Saved annotated image to {annotated_path}."
            )
        else:
            self._log(f"No matches found in {image_path}.")

        return len(matches)

    def _annotate_and_save(
        self, image: Image.Image, matches: List[DetectionResult], original_path: str, output_dir: Path
    ) -> Path:
        draw = ImageDraw.Draw(image)
        font = self._load_font()
        for detection in matches:
            x1, y1, x2, y2 = detection.box
            draw.rectangle(((x1, y1), (x2, y2)), outline="lime", width=3)
            label = f"sim {detection.similarity:.2f}"
            text_size = draw.textbbox((0, 0), label, font=font)
            if text_size:
                _, _, text_w, text_h = text_size
            else:
                text_w, text_h = 0, 0
            draw.rectangle(((x1, y1 - text_h - 4), (x1 + text_w + 4, y1)), fill="lime")
            draw.text((x1 + 2, y1 - text_h - 2), label, fill="black", font=font)

        image_array = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        output_filename = Path(original_path).stem + "_annotated.png"
        output_path = output_dir / output_filename
        cv2.imwrite(str(output_path), image_array)
        return output_path

    def _load_font(self) -> ImageFont.ImageFont:
        try:
            return ImageFont.truetype("arial.ttf", size=16)
        except OSError:
            return ImageFont.load_default()

    # ------------------------------------------------------------------
    def _image_filetypes(self) -> Tuple[Tuple[str, str], ...]:
        return (
            ("Image files", "*.png *.jpg *.jpeg *.bmp"),
            ("All files", "*.*"),
        )

    def _log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(END, message + "\n")
        self.log.see(END)
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = FaceMatcherApp()
    app.run()
