# 📄 PDF2PNG Converter

PDF 파일을 고품질 PNG 이미지로 간편하게 변환하는 윈도우 데스크탑 프로그램입니다.

## ✨ 주요 기능

- **🖱️ 드래그 앤 드롭** — PDF 파일을 창에 끌어다 놓으면 바로 추가
- **📂 다중 파일 지원** — 여러 PDF를 한 번에 변환
- **🎨 DPI 선택** — 150 (일반) / 300 (고화질) / 600 (인쇄용)
- **📑 페이지 병합** — 여러 페이지를 1장 PNG로 합치거나 페이지별 개별 PNG로 분리
- **📊 진행률 표시** — 실시간 변환 진행 상태 확인
- **🌙 다크 테마** — 눈이 편한 모던 다크 UI

## 📸 스크린샷

> 프로그램 실행 후 PDF 파일을 드래그하거나 클릭하여 추가합니다.

## 🚀 사용법

### EXE 실행 (권장)

1. [Releases](../../releases) 페이지에서 `PDF2PNG.exe`를 다운로드합니다
2. 더블클릭하여 실행합니다
3. PDF 파일을 드래그하거나 클릭하여 추가합니다
4. DPI와 페이지 병합 옵션을 설정합니다
5. **🚀 변환 시작** 버튼을 클릭합니다
6. 원본 PDF와 같은 폴더에 PNG 파일이 생성됩니다

> ⚠️ 첫 실행 시 Windows Defender SmartScreen 경고가 나타날 수 있습니다.  
> **"추가 정보" → "실행"** 을 눌러주세요.

### Python으로 실행

```bash
# 의존성 설치
pip install PyMuPDF Pillow customtkinter tkinterdnd2

# 실행
python pdf2png.py
```

## 🛠️ 직접 EXE 빌드

```bash
pip install pyinstaller

python -m PyInstaller --noconfirm --onefile --windowed \
    --collect-all tkinterdnd2 \
    --collect-all customtkinter \
    --name "PDF2PNG" pdf2png.py
```

빌드된 EXE는 `dist/PDF2PNG.exe`에 생성됩니다.

## 📋 기술 스택

| 기술 | 용도 |
|---|---|
| **Python 3.10+** | 코어 런타임 |
| **PyMuPDF (fitz)** | PDF 렌더링 |
| **Pillow** | 이미지 합성 (페이지 병합) |
| **CustomTkinter** | 모던 다크 테마 GUI |
| **TkinterDnD2** | 드래그 앤 드롭 지원 |
| **PyInstaller** | EXE 패키징 |

## 📄 라이선스

MIT License
