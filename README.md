# docxToMp3

Word(.docx) 문서 전체를 소리 내어 읽는 MP3 한 개로 변환합니다.

- 본문 순서대로 문단과 표를 읽고, 목차(toc)는 건너뜁니다.
- 문단마다 한국어/영어를 자동 판별해 해당 언어 음성으로 읽습니다.
- Heading 1·2는 MP3 챕터 마커로 들어갑니다.

## 준비

```bash
pip install -r requirements.txt
sudo apt-get install -y ffmpeg espeak-ng   # espeak-ng은 오프라인 엔진용
```

## 사용

```bash
python3 docx2mp3.py 문서.docx 결과.mp3                 # auto: edge 가능하면 edge, 아니면 espeak
python3 docx2mp3.py 문서.docx 결과.mp3 --engine edge   # 자연스러운 신경망 음성 (네트워크 필요)
python3 docx2mp3.py 문서.docx 결과.mp3 --engine espeak # 오프라인 음성
python3 docx2mp3.py 문서.docx --edge-rate +15% --text-out 읽은내용.txt
```

`edge` 엔진은 `speech.platform.bing.com`에 접속할 수 있어야 합니다.
