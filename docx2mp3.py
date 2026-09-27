#!/usr/bin/env python3
"""Convert a Word (.docx) document into a single MP3 that reads the whole text aloud.

The document body is read in order (paragraphs and tables), the table of contents
is skipped, and each paragraph is voiced in Korean or English depending on its
script. Headings become MP3 chapter markers.

Engines:
  supertonic  Supertone Supertonic neural TTS, run locally with ONNX Runtime
              (natural; downloads the model from huggingface.co on first use)
  edge        Microsoft Edge neural voices via `edge-tts` (needs speech.platform.bing.com)
  espeak      espeak-ng (offline; robotic but always available)
  auto        supertonic if installed, otherwise espeak (default)
"""
import argparse
import asyncio
import os
import re
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

HANGUL = re.compile(r"[가-힣]")
SKIP_STYLES = ("toc",)
SKIP_TEXTS = {"목차", "표목차", "그림목차", "Contents", "Table of Contents"}
PART_STYLES = ("Heading 1", "목차 제목")  # start a new file with --split-dir
CHAPTER_STYLES = ("Heading 2",)
PART_TEXTS = {"참고문헌", "References"}
ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
EDGE_VOICES = {"ko": "ko-KR-SunHiNeural", "en": "en-US-AriaNeural"}
ESPEAK_VOICES = {"ko": "ko", "en": "en-us"}
SAMPLE_RATE = 24000
MAX_CHUNK = 1500
KO_LETTERS = dict(zip(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    ["에이", "비", "씨", "디", "이", "에프", "지", "에이치", "아이", "제이", "케이", "엘", "엠",
     "엔", "오", "피", "큐", "알", "에스", "티", "유", "브이", "더블유", "엑스", "와이", "지"],
))


def detect_lang(text):
    # Any Hangul means Korean: Korean paragraphs often carry English terms in parentheses.
    return "ko" if HANGUL.search(text) else "en"


def normalize(text):
    text = text.replace("\t", " ").replace(" ", " ")
    text = re.sub(r"[“”]", '"', text)
    text = re.sub(r"[‘’]", "'", text)
    text = re.sub(r"<(표|그림)\s*(\d+)>", r"\1 \2.", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def speakable(text, lang):
    """Rewrite symbols and acronyms the neural voices tend to misread."""
    unit = "달러" if lang == "ko" else " dollars"
    text = re.sub(r"\$\s?(\d[\d,]*(?:\.\d+)?)", lambda m: m.group(1) + unit, text)
    text = text.replace("·", ", ").replace("→", ", ")
    roman = "|".join(sorted(ROMAN, key=len, reverse=True))
    chapter = (lambda n: f"제{n}장") if lang == "ko" else (lambda n: f"Chapter {n}")
    # Chapter numbers: "II. 이론적 배경" -> "제2장. 이론적 배경", "제 IV장" -> "제4장".
    text = re.sub(rf"^({roman})\.\s", lambda m: chapter(ROMAN[m.group(1)]) + ". ", text)
    text = re.sub(rf"제\s?({roman})\s?장", lambda m: f"제{ROMAN[m.group(1)]}장", text)
    if lang == "ko":
        text = re.sub(r"(?<![A-Za-z])TV(?![A-Za-z])", "티비", text)
        # Spell upper-case acronyms (ECM, ODM, SER-M, LG) with Korean letter names.
        text = re.sub(
            r"(?<![A-Za-z])[A-Z]{2,}(?:-[A-Z]+)?(?![a-z])",
            lambda m: " ".join("".join(KO_LETTERS[c] for c in part)
                               for part in m.group(0).split("-")),
            text,
        )
    return text


def table_text(table):
    rows = [[normalize(c.text) for c in r.cells] for r in table.rows]
    if not rows:
        return []
    header, lines = rows[0], []
    for row in rows[1:]:
        cells = []
        for i, cell in enumerate(row):
            if not cell or cell == "-" or (i and cell == row[i - 1]):
                continue
            name = header[i] if i < len(header) else ""
            cells.append(f"{name}: {cell}" if i and name else cell)
        if cells:
            lines.append(", ".join(cells) + ".")
    return lines


def extract_segments(path):
    """Return a list of (text, kind) in document order; kind is "part", "chapter" or None."""
    document = docx.Document(path)
    segments = []
    for el in document.element.body:
        if el.tag == qn("w:p"):
            p = Paragraph(el, document)
            style = p.style.name if p.style is not None else ""
            if style.lower().startswith(SKIP_STYLES):
                continue
            text = normalize(p.text)
            if text and text not in SKIP_TEXTS:
                if style in PART_STYLES or text in PART_TEXTS:
                    kind = "part"
                elif style in CHAPTER_STYLES:
                    kind = "chapter"
                else:
                    kind = None
                segments.append((text, kind))
        elif el.tag == qn("w:tbl"):
            segments.extend((line, None) for line in table_text(Table(el, document)))
    return segments


def split_chunks(text, limit=MAX_CHUNK):
    if len(text) <= limit:
        return [text]
    sentences = re.split(r"(?<=[.!?。])\s+", text)
    chunks, cur = [], ""
    for s in sentences:
        if cur and len(cur) + len(s) + 1 > limit:
            chunks.append(cur)
            cur = s
        else:
            cur = f"{cur} {s}".strip()
    if cur:
        chunks.append(cur)
    return chunks


def to_wav(src, dst):
    subprocess.run(
        ["ffmpeg", "-loglevel", "error", "-y", "-i", src, "-ac", "1", "-ar", str(SAMPLE_RATE), dst],
        check=True,
    )


def synth_espeak(text, lang, out_wav, rate):
    raw = out_wav + ".raw.wav"
    subprocess.run(
        ["espeak-ng", "-v", ESPEAK_VOICES[lang], "-s", str(rate), "-w", raw, text], check=True
    )
    to_wav(raw, out_wav)
    os.remove(raw)


_SUPERTONIC = None


def synth_supertonic(text, lang, out_wav, voice, speed):
    global _SUPERTONIC
    if _SUPERTONIC is None:
        from supertonic import TTS

        tts = TTS(auto_download=True)
        _SUPERTONIC = (tts, {})
    tts, styles = _SUPERTONIC
    if voice not in styles:
        styles[voice] = tts.get_voice_style(voice_name=voice)
    wav, _ = tts.synthesize(text, voice_style=styles[voice], lang=lang, speed=speed)
    raw = out_wav + ".raw.wav"
    tts.save_audio(wav, raw)
    to_wav(raw, out_wav)
    os.remove(raw)


def supertonic_available():
    try:
        import supertonic  # noqa: F401
        return True
    except ImportError:
        return False


async def _edge_save(text, voice, out, rate):
    import edge_tts

    await edge_tts.Communicate(text, voice, rate=rate).save(out)


def synth_edge(text, lang, out_wav, rate, retries=4):
    mp3 = out_wav + ".mp3"
    for attempt in range(retries):
        try:
            asyncio.run(_edge_save(text, EDGE_VOICES[lang], mp3, rate))
            break
        except Exception:
            if attempt == retries - 1:
                raise
    to_wav(mp3, out_wav)
    os.remove(mp3)


def edge_available():
    try:
        synth_edge("테스트", "ko", os.path.join(tempfile.gettempdir(), "edge_probe.wav"), "+0%", 1)
        return True
    except Exception:
        return False


def silence(path, seconds):
    subprocess.run(
        ["ffmpeg", "-loglevel", "error", "-y", "-f", "lavfi", "-i",
         f"anullsrc=r={SAMPLE_RATE}:cl=mono", "-t", str(seconds), path],
        check=True,
    )


def duration(path):
    out = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", path],
        capture_output=True, text=True, check=True,
    )
    return float(out.stdout.strip())


def assemble(items, output, work, short_gap, long_gap):
    """Concatenate synthesized chunks into one chaptered MP3; return its length in seconds."""
    base = os.path.splitext(os.path.basename(output))[0]
    concat = os.path.join(work, f"list_{base}.txt")
    meta_path = os.path.join(work, f"meta_{base}.txt")
    chapters, t = [], 0.0
    with open(concat, "w") as f:
        for (_, _, heading, _), wav, d in items:
            if heading:
                f.write(f"file '{long_gap}'\n")
                t += 1.2
                chapters.append((t, heading))
            f.write(f"file '{wav}'\nfile '{long_gap if heading else short_gap}'\n")
            t += d + (1.2 if heading else 0.5)
    meta = [";FFMETADATA1", f"title={base}"]
    ends = [c[0] for c in chapters[1:]] + [t]
    for (start, title), end in zip(chapters, ends):
        meta += ["[CHAPTER]", "TIMEBASE=1/1000", f"START={int(start * 1000)}",
                 f"END={int(end * 1000)}", f"title={title[:120]}"]
    with open(meta_path, "w", encoding="utf-8") as f:
        f.write("\n".join(meta) + "\n")
    subprocess.run(
        ["ffmpeg", "-loglevel", "error", "-y", "-f", "concat", "-safe", "0", "-i", concat,
         "-i", meta_path, "-map_metadata", "1", "-map_chapters", "1",
         "-c:a", "libmp3lame", "-b:a", "64k", "-id3v2_version", "3", output],
        check=True,
    )
    return t


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help="input .docx file")
    ap.add_argument("output", nargs="?", help="output .mp3 (default: input name with .mp3)")
    ap.add_argument("--engine", choices=["auto", "supertonic", "edge", "espeak"], default="auto")
    ap.add_argument("--voice", default="F1", help="supertonic voice: F1-F5, M1-M5")
    ap.add_argument("--speed", type=float, default=1.05, help="supertonic speed (1.0 = normal)")
    ap.add_argument("--edge-rate", default="+0%", help="edge-tts speed, e.g. +10%%")
    ap.add_argument("--espeak-rate", type=int, default=165, help="espeak-ng words per minute")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--split-dir", help="also write one MP3 per part (Heading 1 / summary / references)")
    ap.add_argument("--text-out", help="also write the text that is read aloud to this file")
    args = ap.parse_args()

    output = args.output or os.path.splitext(args.input)[0] + ".mp3"
    engine = args.engine
    if engine == "auto":
        engine = "supertonic" if supertonic_available() else "espeak"
    print(f"engine: {engine}", file=sys.stderr)

    segments = extract_segments(args.input)
    if args.text_out:
        with open(args.text_out, "w", encoding="utf-8") as f:
            f.write("\n".join(speakable(t, detect_lang(t)) for t, _ in segments) + "\n")

    jobs = []  # (text, lang, heading_title or None, kind or None)
    for text, kind in segments:
        lang = detect_lang(text)
        for i, chunk in enumerate(split_chunks(speakable(text, lang))):
            first = kind and i == 0
            jobs.append((chunk, lang, text if first else None, kind if first else None))

    work = tempfile.mkdtemp(prefix="docx2mp3_")
    try:
        def run(idx):
            text, lang, _, _ = jobs[idx]
            out = os.path.join(work, f"{idx:05d}.wav")
            if engine == "supertonic":
                synth_supertonic(text, lang, out, args.voice, args.speed)
            elif engine == "edge":
                synth_edge(text, lang, out, args.edge_rate)
            else:
                synth_espeak(text, lang, out, args.espeak_rate)
            return out

        # Supertonic already uses every core through ONNX Runtime.
        workers = 1 if engine == "supertonic" else args.workers
        with ThreadPoolExecutor(workers) as pool:
            wavs = []
            for n, wav in enumerate(pool.map(run, range(len(jobs))), 1):
                wavs.append(wav)
                if n % 50 == 0 or n == len(jobs):
                    print(f"  {n}/{len(jobs)} chunks", file=sys.stderr)

        short_gap, long_gap = os.path.join(work, "gap_s.wav"), os.path.join(work, "gap_l.wav")
        silence(short_gap, 0.5)
        silence(long_gap, 1.2)
        durations = [duration(w) for w in wavs]
        items = list(zip(jobs, wavs, durations))

        t = assemble(items, output, work, short_gap, long_gap)
        if args.split_dir:
            os.makedirs(args.split_dir, exist_ok=True)
            starts = [i for i, ((_, _, h, kind), _, _) in enumerate(items) if kind == "part"]
            starts = [0] + [i for i in starts if i > 0]
            for n, (a, b) in enumerate(zip(starts, starts[1:] + [len(items)]), 1):
                title = next((h for (_, _, h, k), _, _ in items[a:b] if k == "part"), "part")
                name = re.sub(r"[^\w가-힣]+", "_", title).strip("_")[:60]
                path = os.path.join(args.split_dir, f"{n:02d}_{name}.mp3")
                d = assemble(items[a:b], path, work, short_gap, long_gap)
                print(f"  {path} ({d / 60:.1f} min)", file=sys.stderr)
        print(f"wrote {output} ({t / 60:.1f} min)", file=sys.stderr)
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
