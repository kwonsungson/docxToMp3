# docxToMp3

Word(.docx) 문서 전체를 소리 내어 읽는 MP3 한 개로 변환합니다.

- 본문 순서대로 문단과 표를 읽고, 목차(toc)는 건너뜁니다.
- 문단마다 한국어/영어를 자동 판별해 해당 언어 음성으로 읽습니다.
- Heading 1·2는 MP3 챕터 마커로 들어갑니다.

## 준비

```bash
pip install -r requirements.txt
sudo apt-get install -y ffmpeg              # espeak-ng는 선택(오프라인 예비 엔진)
```

## 사용

```bash
python3 docx2mp3.py 문서.docx 결과.mp3                 # auto: supertonic 설치돼 있으면 사용, 아니면 espeak
python3 docx2mp3.py 문서.docx --voice M1 --speed 1.1  # supertonic 목소리(F1~F5, M1~M5)와 속도
python3 docx2mp3.py 문서.docx 결과.mp3 --engine espeak # 오프라인 음성
python3 docx2mp3.py 문서.docx --edge-rate +15% --text-out 읽은내용.txt
```

기본 엔진인 Supertonic(신경망 음성)은 처음 실행할 때 huggingface.co에서 모델을 내려받고, 이후에는 CPU에서 로컬로 합성합니다.
`$299.99`는 "299.99달러"로, 한국어 문단의 대문자 약어(ECM, ODM)는 "이씨엠", "오디엠"처럼 읽도록 바꿔서 합성합니다.
