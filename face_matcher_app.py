"""Face Matcher Application with enhanced accuracy and rich UI."""
from __future__ import annotations

import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

import cv2
import numpy as np
import torch
from facenet_pytorch import InceptionResnetV1, MTCNN
from PIL import Image, ImageDraw, ImageFont
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    BooleanVar,
    DoubleVar,
    IntVar,
    Listbox,
    Scrollbar,
    StringVar,
    Tk,
    filedialog,
    messagebox,
)
from tkinter import ttk


@dataclass
class TargetFace:
    """Metadata for an enrolled target face embedding."""

    source_path: str
    probability: float
    embedding: torch.Tensor
    variant: str


@dataclass
class DetectionResult:
    """Stores metadata for a detected face."""

    box: Sequence[float]
    probability: float
    similarity: float
    matched_target: TargetFace


@dataclass
class CandidateSummary:
    """Aggregated statistics about a processed candidate image."""

    path: str
    best_similarity: float
    best_target: Optional[str]
    match_count: int


class FaceMatcherApp:
    """Tkinter-based GUI application for matching faces across images."""

    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("Signature Face Intelligence Studio")
        self.root.geometry("980x720")
        self.root.configure(bg="#0b0d17")

        self.device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        self.mtcnn: Optional[MTCNN] = None
        self.encoder: Optional[InceptionResnetV1] = None
        self.target_faces: List[TargetFace] = []
        self.target_bank: Optional[torch.Tensor] = None

        self.target_paths: List[str] = []
        self.candidate_paths: List[str] = []
        self.output_dir = StringVar(value=str(Path.cwd() / "annotated_output"))
        self.threshold = DoubleVar(value=0.63)
        self.min_target_prob = DoubleVar(value=0.75)
        self.min_candidate_prob = DoubleVar(value=0.75)
        self.enable_flip_aug = BooleanVar(value=True)
        self.create_session_dir = BooleanVar(value=True)
        self.max_candidates_preview = IntVar(value=5)

        self._build_widgets()

    # ------------------------------------------------------------------
    # GUI Construction
    def _build_widgets(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#0b0d17")
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), foreground="#4deeea", background="#0b0d17")
        style.configure("SubTitle.TLabel", font=("Segoe UI", 11), foreground="#ffffff", background="#0b0d17")
        style.configure("TLabel", foreground="#d0d2d6", background="#0b0d17", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabelframe", background="#131a2a", foreground="#4deeea", font=("Segoe UI", 11, "bold"))
        style.configure("TLabelframe.Label", foreground="#4deeea", background="#131a2a")
        style.configure("Horizontal.TProgressbar", background="#4deeea")

        header = ttk.Frame(self.root, padding=20)
        header.pack(fill=BOTH)
        ttk.Label(header, text="Custom made by Sif Almaghrabi", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Upload multiple target portraits and compare against unlimited photo collections."
            " Advanced similarity fusion keeps track of your person across decades.",
            style="SubTitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        main = ttk.Frame(self.root, padding=(20, 0))
        main.pack(fill=BOTH, expand=True)

        lists_frame = ttk.Frame(main)
        lists_frame.pack(fill=BOTH, expand=True)

        # Target images panel
        target_frame = ttk.Labelframe(lists_frame, text="Pic A — Reference gallery", padding=12)
        target_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
        self.target_listbox = self._build_list_panel(target_frame)

        target_btn_row = ttk.Frame(target_frame)
        target_btn_row.pack(fill=BOTH, pady=(8, 0))
        ttk.Button(target_btn_row, text="Add target photos", command=self._select_targets).pack(side=LEFT, padx=(0, 6))
        ttk.Button(target_btn_row, text="Remove selected", command=self._remove_selected_targets).pack(side=LEFT, padx=(0, 6))
        ttk.Button(target_btn_row, text="Clear", command=self._clear_targets).pack(side=LEFT)

        # Candidate images panel
        candidate_frame = ttk.Labelframe(lists_frame, text="Comparison photos", padding=12)
        candidate_frame.pack(side=RIGHT, fill=BOTH, expand=True)
        self.candidate_listbox = self._build_list_panel(candidate_frame)

        candidate_btn_row = ttk.Frame(candidate_frame)
        candidate_btn_row.pack(fill=BOTH, pady=(8, 0))
        ttk.Button(candidate_btn_row, text="Add comparison photos", command=self._select_candidates).pack(side=LEFT, padx=(0, 6))
        ttk.Button(candidate_btn_row, text="Remove selected", command=self._remove_selected_candidates).pack(side=LEFT, padx=(0, 6))
        ttk.Button(candidate_btn_row, text="Clear", command=self._clear_candidates).pack(side=LEFT)

        # Options frame
        options = ttk.Labelframe(main, text="Intelligence controls", padding=12)
        options.pack(fill=BOTH, expand=True, pady=(16, 0))

        ttk.Label(options, text="Output directory:").grid(row=0, column=0, sticky="w")
        self.output_entry = ttk.Entry(options, textvariable=self.output_dir, width=60)
        self.output_entry.grid(row=0, column=1, sticky="we", padx=(6, 0))
        ttk.Button(options, text="Browse", command=self._select_output_dir).grid(row=0, column=2, padx=(8, 0))
        options.columnconfigure(1, weight=1)

        ttk.Label(options, text="Similarity threshold (0-1):").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Spinbox(options, from_=0.0, to=1.0, increment=0.01, textvariable=self.threshold, width=8).grid(
            row=1, column=1, sticky="w", padx=(6, 0), pady=(10, 0)
        )

        ttk.Label(options, text="Min target face confidence:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Spinbox(options, from_=0.0, to=1.0, increment=0.01, textvariable=self.min_target_prob, width=8).grid(
            row=2, column=1, sticky="w", padx=(6, 0), pady=(10, 0)
        )

        ttk.Label(options, text="Min comparison face confidence:").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Spinbox(options, from_=0.0, to=1.0, increment=0.01, textvariable=self.min_candidate_prob, width=8).grid(
            row=3, column=1, sticky="w", padx=(6, 0), pady=(10, 0)
        )

        ttk.Checkbutton(options, text="Mirror augment faces for decades-proof matching", variable=self.enable_flip_aug).grid(
            row=1, column=2, columnspan=2, sticky="w", padx=(12, 0)
        )
        ttk.Checkbutton(options, text="Create timestamped session folder", variable=self.create_session_dir).grid(
            row=2, column=2, columnspan=2, sticky="w", padx=(12, 0)
        )

        ttk.Label(options, text="Preview top matches (count):").grid(row=3, column=2, sticky="w", padx=(12, 0))
        ttk.Spinbox(options, from_=1, to=50, textvariable=self.max_candidates_preview, width=5).grid(
            row=3, column=3, sticky="w", padx=(6, 0)
        )

        # Controls
        controls = ttk.Frame(main, padding=(0, 16))
        controls.pack(fill=BOTH)
        self.progress = ttk.Progressbar(controls, mode="determinate")
        self.progress.pack(fill=BOTH, expand=True, padx=(0, 12), side=LEFT)
        self.match_button = ttk.Button(controls, text="Launch matching", style="Accent.TButton", command=self._run_async)
        self.match_button.pack(side=RIGHT)

        # Log area
        log_frame = ttk.Labelframe(main, text="Event log", padding=12)
        log_frame.pack(fill=BOTH, expand=True)
        self.log = Listbox(log_frame, height=12, background="#10192e", foreground="#eaf6ff", activestyle="none")
        self.log.pack(fill=BOTH, expand=True, side=LEFT)
        log_scroll = Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        log_scroll.pack(side=RIGHT, fill="y")
        self.log.configure(yscrollcommand=log_scroll.set)

    def _build_list_panel(self, container: ttk.Labelframe) -> Listbox:
        frame = ttk.Frame(container)
        frame.pack(fill=BOTH, expand=True)
        listbox = Listbox(
            frame,
            height=10,
            selectmode="extended",
            background="#10192e",
            foreground="#eaf6ff",
            activestyle="none",
        )
        listbox.pack(fill=BOTH, expand=True, side=LEFT)
        scrollbar = Scrollbar(frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side=RIGHT, fill="y")
        listbox.configure(yscrollcommand=scrollbar.set)
        return listbox

    # ------------------------------------------------------------------
    # GUI Event Handlers
    def _select_targets(self) -> None:
        paths = filedialog.askopenfilenames(title="Select reference images", filetypes=self._image_filetypes())
        if not paths:
            return
        added = 0
        for path in paths:
            if path not in self.target_paths:
                self.target_paths.append(path)
                added += 1
        if added:
            self._refresh_listbox(self.target_listbox, self.target_paths)
            self._log(f"Added {added} reference image(s).")

    def _remove_selected_targets(self) -> None:
        self._remove_selected(self.target_listbox, self.target_paths, "reference")

    def _clear_targets(self) -> None:
        if self.target_paths:
            self.target_paths.clear()
            self._refresh_listbox(self.target_listbox, self.target_paths)
            self._log("Cleared all reference images.")

    def _select_candidates(self) -> None:
        paths = filedialog.askopenfilenames(title="Select comparison images", filetypes=self._image_filetypes())
        if not paths:
            return
        added = 0
        for path in paths:
            if path not in self.candidate_paths:
                self.candidate_paths.append(path)
                added += 1
        if added:
            self._refresh_listbox(self.candidate_listbox, self.candidate_paths)
            self._log(f"Added {added} comparison image(s).")

    def _remove_selected_candidates(self) -> None:
        self._remove_selected(self.candidate_listbox, self.candidate_paths, "comparison")

    def _clear_candidates(self) -> None:
        if self.candidate_paths:
            self.candidate_paths.clear()
            self._refresh_listbox(self.candidate_listbox, self.candidate_paths)
            self._log("Cleared all comparison images.")

    def _remove_selected(self, listbox: Listbox, storage: List[str], label: str) -> None:
        selection = listbox.curselection()
        if not selection:
            messagebox.showinfo("No selection", f"Select one or more {label} images to remove.")
            return
        removed = 0
        for index in reversed(selection):
            if 0 <= index < len(storage):
                storage.pop(index)
                removed += 1
        self._refresh_listbox(listbox, storage)
        self._log(f"Removed {removed} {label} image(s).")

    def _refresh_listbox(self, listbox: Listbox, items: Sequence[str]) -> None:
        listbox.delete(0, END)
        for path in items:
            listbox.insert(END, path)

    def _select_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_dir.set(directory)
            self._log(f"Output directory set to: {directory}")

    # ------------------------------------------------------------------
    def _run_async(self) -> None:
        if getattr(self, "_worker", None) and self._worker.is_alive():
            messagebox.showinfo("Processing", "Matching is already running. Please wait for it to finish.")
            return
        self._worker = threading.Thread(target=self._run_matching, daemon=True)
        self._worker.start()

    def _run_matching(self) -> None:
        if not self.target_paths:
            self._error("Missing reference", "Please add at least one Pic A reference image before running.")
            return
        if not self.candidate_paths:
            self._error("Missing comparisons", "Please add at least one comparison image.")
            return

        try:
            threshold = float(self.threshold.get())
            min_target_prob = float(self.min_target_prob.get())
            min_candidate_prob = float(self.min_candidate_prob.get())
        except (TypeError, ValueError):
            self._error("Invalid settings", "Thresholds must be numeric.")
            return
        if not 0 <= threshold <= 1:
            self._error("Invalid threshold", "Similarity threshold must be between 0 and 1.")
            return
        if not 0 <= min_target_prob <= 1 or not 0 <= min_candidate_prob <= 1:
            self._error("Invalid confidence", "Confidence filters must be between 0 and 1.")
            return

        output_dir = Path(self.output_dir.get()).expanduser()
        if self.create_session_dir.get():
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = output_dir / f"session_{stamp}"
        output_dir.mkdir(parents=True, exist_ok=True)

        self._toggle_run_state(True)
        self._log("Loading neural backbones…")
        self._load_models()
        if self.mtcnn is None or self.encoder is None:
            self._error("Model error", "Failed to initialize face detection or embedding models.")
            self._toggle_run_state(False)
            return

        self._log("Profiling target gallery…")
        target_faces = self._extract_target_profile(min_target_prob)
        if not target_faces:
            self._error("No faces", "No suitable faces were detected in the reference gallery.")
            self._toggle_run_state(False)
            return
        self.target_faces = target_faces
        self.target_bank = torch.stack([face.embedding for face in target_faces]).to(self.device)
        self._log(f"Enrolled {len(target_faces)} reference face variants.")

        total_matches = 0
        summaries: List[CandidateSummary] = []
        self._set_progress(maximum=len(self.candidate_paths), value=0)

        for index, candidate_path in enumerate(self.candidate_paths, start=1):
            matches, summary = self._process_candidate(candidate_path, output_dir, threshold, min_candidate_prob)
            total_matches += len(matches)
            summaries.append(summary)
            self._set_progress(value=index)

        self._log(f"Finished processing {len(self.candidate_paths)} comparison image(s). Total matches: {total_matches}")
        self._present_summary(output_dir, summaries)
        self._toggle_run_state(False)

    # ------------------------------------------------------------------
    def _load_models(self) -> None:
        if self.mtcnn is None:
            self.mtcnn = MTCNN(image_size=160, margin=20, keep_all=True, device=self.device)
        if self.encoder is None:
            self.encoder = InceptionResnetV1(pretrained="vggface2").eval().to(self.device)

    def _extract_target_profile(self, min_prob: float) -> List[TargetFace]:
        assert self.mtcnn is not None
        assert self.encoder is not None

        collected: List[TargetFace] = []
        for path in self.target_paths:
            try:
                image = Image.open(path).convert("RGB")
            except Exception as exc:
                self._log(f"Failed to open {path}: {exc}")
                continue
            faces, probs = self.mtcnn(image, return_prob=True)
            if faces is None or probs is None:
                self._log(f"No faces detected in reference image: {path}")
                continue

            for idx, (face_tensor, prob) in enumerate(zip(faces, probs)):
                if prob is None or prob < min_prob:
                    continue
                embedding = self._encode_face(face_tensor.unsqueeze(0))
                collected.append(TargetFace(path, float(prob), embedding, variant=f"face_{idx+1}"))
                if self.enable_flip_aug.get():
                    flipped = torch.flip(face_tensor.unsqueeze(0), dims=[3])
                    embedding_flip = self._encode_face(flipped)
                    collected.append(TargetFace(path, float(prob), embedding_flip, variant=f"face_{idx+1}_mirror"))

        if not collected:
            return []

        # Normalize combined embeddings for stable comparison
        stacked = torch.stack([face.embedding for face in collected])
        mean_embedding = torch.nn.functional.normalize(stacked.mean(dim=0, keepdim=True), p=2, dim=1)
        # Add mean profile for robustness
        collected.append(TargetFace("Average profile", 1.0, mean_embedding.squeeze(0), variant="ensemble"))
        return collected

    def _encode_face(self, face_tensor: torch.Tensor) -> torch.Tensor:
        assert self.encoder is not None
        face_tensor = face_tensor.to(self.device)
        with torch.no_grad():
            embedding = self.encoder(face_tensor)
        if self.enable_flip_aug.get():
            flipped = torch.flip(face_tensor, dims=[3])
            with torch.no_grad():
                flipped_embedding = self.encoder(flipped)
            embedding = (embedding + flipped_embedding) / 2.0
        embedding = torch.nn.functional.normalize(embedding, p=2, dim=1)
        return embedding.squeeze(0).cpu()

    def _process_candidate(
        self,
        image_path: str,
        output_dir: Path,
        threshold: float,
        min_prob: float,
    ) -> Tuple[List[DetectionResult], CandidateSummary]:
        assert self.mtcnn is not None
        assert self.encoder is not None
        assert self.target_bank is not None

        try:
            image = Image.open(image_path).convert("RGB")
        except Exception as exc:
            self._log(f"Failed to open {image_path}: {exc}")
            return [], CandidateSummary(image_path, 0.0, None, 0)

        boxes, probs = self.mtcnn.detect(image)
        if boxes is None or probs is None:
            self._log(f"No faces detected in candidate: {image_path}")
            return [], CandidateSummary(image_path, 0.0, None, 0)

        valid_indices = [i for i, prob in enumerate(probs) if prob is not None and prob >= min_prob]
        if not valid_indices:
            self._log(f"All detected faces were below confidence filter in {image_path}")
            return [], CandidateSummary(image_path, 0.0, None, 0)

        faces_tensor = self.mtcnn.extract(image, boxes[valid_indices], None)
        if faces_tensor is None:
            self._log(f"Failed to align faces for candidate: {image_path}")
            return [], CandidateSummary(image_path, 0.0, None, 0)

        faces_tensor = faces_tensor.to(self.device)
        with torch.no_grad():
            embeddings = self.encoder(faces_tensor)
            if self.enable_flip_aug.get():
                flipped = torch.flip(faces_tensor, dims=[3])
                flipped_embeddings = self.encoder(flipped)
                embeddings = (embeddings + flipped_embeddings) / 2.0
        embeddings = torch.nn.functional.normalize(embeddings, p=2, dim=1)

        similarities = torch.matmul(embeddings, self.target_bank.T)
        best_similarities, best_indices = torch.max(similarities, dim=1)

        matches: List[DetectionResult] = []
        for idx, face_idx in enumerate(valid_indices):
            similarity = float(best_similarities[idx].item())
            if similarity >= threshold:
                matched_target = self.target_faces[int(best_indices[idx].item())]
                matches.append(
                    DetectionResult(
                        box=boxes[face_idx],
                        probability=float(probs[face_idx]),
                        similarity=similarity,
                        matched_target=matched_target,
                    )
                )

        if matches:
            annotated_path = self._annotate_and_save(image, matches, image_path, output_dir)
            self._log(
                f"{len(matches)} match(es) found in {image_path}. Saved annotated image to {annotated_path}."
            )
        else:
            top_similarity = float(best_similarities.max().item()) if best_similarities.numel() else 0.0
            self._log(
                f"No matches above threshold in {image_path}. Highest similarity: {top_similarity:.2f}."
            )

        best_similarity = float(best_similarities.max().item()) if best_similarities.numel() else 0.0
        best_target = None
        if best_similarities.numel():
            best_target = self.target_faces[int(best_indices[best_similarities.argmax()].item())].source_path

        summary = CandidateSummary(
            path=image_path,
            best_similarity=best_similarity,
            best_target=best_target,
            match_count=len(matches),
        )
        return matches, summary

    def _annotate_and_save(
        self, image: Image.Image, matches: List[DetectionResult], original_path: str, output_dir: Path
    ) -> Path:
        draw = ImageDraw.Draw(image)
        font = self._load_font()
        for detection in matches:
            x1, y1, x2, y2 = detection.box
            draw.rectangle(((x1, y1), (x2, y2)), outline="#4deeea", width=4)
            label = (
                f"sim {detection.similarity:.2f} | ref: {Path(detection.matched_target.source_path).name}"
                f" [{detection.matched_target.variant}]"
            )
            text_bbox = draw.textbbox((0, 0), label, font=font)
            if text_bbox:
                _, _, text_w, text_h = text_bbox
            else:
                text_w, text_h = 0, 0
            draw.rectangle(((x1, y1 - text_h - 6), (x1 + text_w + 6, y1)), fill="#4deeea")
            draw.text((x1 + 3, y1 - text_h - 3), label, fill="#04121f", font=font)

        image_array = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        output_filename = f"{Path(original_path).stem}_annotated.png"
        output_path = output_dir / output_filename
        cv2.imwrite(str(output_path), image_array)
        return output_path

    def _load_font(self) -> ImageFont.ImageFont:
        for font_name in ("segoeui.ttf", "arial.ttf"):
            try:
                return ImageFont.truetype(font_name, size=18)
            except OSError:
                continue
        return ImageFont.load_default()

    def _present_summary(self, output_dir: Path, summaries: Sequence[CandidateSummary]) -> None:
        preview_count = max(1, int(self.max_candidates_preview.get()))
        top_matches = sorted(summaries, key=lambda s: s.best_similarity, reverse=True)[:preview_count]
        self._log("Top similarity scores across the collection:")
        for summary in top_matches:
            if summary.best_target:
                self._log(
                    f" - {summary.path} → {summary.best_similarity:.3f} vs {Path(summary.best_target).name}"
                )
            else:
                self._log(f" - {summary.path} → {summary.best_similarity:.3f} (no confident reference)")

        summary_path = output_dir / "match_summary.txt"
        with summary_path.open("w", encoding="utf-8") as handle:
            handle.write("Face Matching Session Summary\n")
            handle.write(f"Total candidates: {len(summaries)}\n")
            handle.write(f"Reference faces enrolled: {len(self.target_faces)}\n\n")
            for summary in summaries:
                handle.write(f"Image: {summary.path}\n")
                handle.write(f"  Matches above threshold: {summary.match_count}\n")
                if summary.best_target:
                    handle.write(
                        f"  Best similarity: {summary.best_similarity:.3f} vs {summary.best_target}\n"
                    )
                else:
                    handle.write(f"  Best similarity: {summary.best_similarity:.3f} (no reference)\n")
                handle.write("\n")
        self._log(f"Session summary exported to {summary_path}")

    # ------------------------------------------------------------------
    def _image_filetypes(self) -> Tuple[Tuple[str, str], ...]:
        return (("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*"))

    def _log(self, message: str) -> None:
        def append() -> None:
            self.log.insert(END, message)
            self.log.see(END)

        self.root.after(0, append)

    def _error(self, title: str, message: str) -> None:
        self.root.after(0, lambda: messagebox.showerror(title, message))

    def _set_progress(self, value: Optional[int] = None, maximum: Optional[int] = None) -> None:
        def update() -> None:
            if maximum is not None:
                self.progress["maximum"] = maximum
            if value is not None:
                self.progress["value"] = value

        self.root.after(0, update)

    def _toggle_run_state(self, running: bool) -> None:
        def update() -> None:
            if running:
                self.match_button.configure(state="disabled")
            else:
                self.match_button.configure(state="normal")
                self.progress["value"] = 0

        self.root.after(0, update)

    # ------------------------------------------------------------------
    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = FaceMatcherApp()
    app.run()
