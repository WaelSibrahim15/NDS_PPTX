"""NDS — export the narrated deck as an MP4 video.

Fully self-contained pipeline: slides rendered as PNG frames with Pillow
(app.renderer), audio muxed with the ffmpeg binary bundled by imageio-ffmpeg.
No LibreOffice, PowerPoint, or system ffmpeg required.
"""
import shutil
import subprocess
from pathlib import Path
from typing import Callable, List, Optional

from .models import DeckPlan
from .renderer import render_frames
from .tts import audio_duration_seconds

W, H = 1920, 1080
TAIL_SILENCE = 1.0  # breathing room after each narration, in seconds


def _ffmpeg() -> str:
    import imageio_ffmpeg

    return imageio_ffmpeg.get_ffmpeg_exe()


def export_mp4(
    plan: DeckPlan,
    audio_files: List[Optional[Path]],
    out_path: Path,
    template: str = "niq",
    progress: Callable[[int, str], None] = lambda pct, msg: None,
) -> Path:
    ffmpeg = _ffmpeg()
    work = out_path.parent / "video_work"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)

    progress(5, "Rendering slide frames…")
    frames = render_frames(plan, work, template=template)

    segments = []
    n = len(frames)
    for i, frame in enumerate(frames):
        progress(10 + int(75 * i / n), f"Encoding slide {i + 1} of {n}…")
        audio = audio_files[i] if i < len(audio_files) else None
        seg = work / f"seg_{i + 1:03d}.mp4"
        if audio is not None and Path(audio).exists():
            dur = audio_duration_seconds(Path(audio)) + TAIL_SILENCE
            cmd = [
                ffmpeg, "-y", "-loop", "1", "-i", str(frame), "-i", str(audio),
                "-t", f"{dur:.2f}",
                "-c:v", "libx264", "-tune", "stillimage", "-r", "30",
                "-pix_fmt", "yuv420p",
                "-c:a", "aac", "-b:a", "160k", "-ar", "44100", "-ac", "2",
                "-af", "apad",
                str(seg),
            ]
        else:  # slide without narration: 5 seconds of silence
            cmd = [
                ffmpeg, "-y", "-loop", "1", "-i", str(frame),
                "-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo",
                "-t", "5",
                "-c:v", "libx264", "-tune", "stillimage", "-r", "30",
                "-pix_fmt", "yuv420p",
                "-c:a", "aac", "-b:a", "160k", "-ar", "44100", "-ac", "2",
                str(seg),
            ]
        result = subprocess.run(cmd, capture_output=True, timeout=600)
        if result.returncode != 0:
            raise RuntimeError(f"ffmpeg failed on slide {i + 1}: {result.stderr.decode()[-400:]}")
        segments.append(seg)

    progress(90, "Joining the video…")
    concat_list = work / "list.txt"
    concat_list.write_text("".join(f"file '{s.name}'\n" for s in segments))
    result = subprocess.run(
        [ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", str(concat_list), "-c", "copy", str(out_path)],
        capture_output=True, timeout=600, cwd=str(work),
    )
    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg concat failed: {result.stderr.decode()[-400:]}")

    shutil.rmtree(work, ignore_errors=True)
    return out_path
