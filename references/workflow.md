# 논문 → 발표 자료 작업 흐름

## 1. 입력 확인

- 논문 PDF: arXiv면 `https://arxiv.org/pdf/<id>`에서 받습니다. 저자, venue, 연도를 확인해 `meta.source`와 표지 `citation`에 씁니다.
- 발표 형식: PYSR(분기 논문 리뷰, 영상), 연구실 세미나(여러 논문, 긴 발표), 자기 논문 발표 중 무엇인지 확인합니다. 형식에 따라 `meta.footer`와 분량이 달라집니다.
- 모르는 것만 묻습니다: 발표 길이, 청중 수준, 강조할 관점(예: 우리 연구와의 연결).

## 2. 읽기와 정리

논문을 끝까지 읽고 아래를 정리합니다. 이 정리가 구성안의 재료입니다.

- 문제: 무엇이 안 되는가, 왜 중요한가
- 핵심 아이디어: 한 문장
- 방법: 구성 요소와 흐름, 핵심 수식 1~3개
- 실험: 설정(모델, 데이터, 지표), 주요 결과 2~4개, ablation
- 한계와 논의: 논문이 말한 것과 발표자가 보는 것을 구분
- figure 목록: `crop_figure.py paper.pdf --list`

## 3. 구성안

기본 흐름 (PYSR, 약 20~30장):

| 구간 | 장 수 | 주로 쓰는 레이아웃 |
|---|---|---|
| 표지, 목차 | 2 | Title, Contents |
| Background: 문제와 기존 연구 | 3~5 | H(핵심 질문), A, E |
| Method | 5~8 | Contents(current), A, C, F |
| Experiments | 5~8 | Contents(current), D(표), B, G(차트) |
| Discussion: 의의, 한계, 우리 연구와의 연결 | 2~3 | E, H |
| Conclusion, 마무리 | 2 | E 또는 A, Closing |
| Appendix | 필요 시 | G, hidden |

구성안은 표로 보여주고 확인을 받은 뒤 만듭니다:

```
# | 레이아웃 | 제목(주장) | 핵심 메시지 | figure
3 | A | 정답 보상은 추론 과정을 보지 못함 | … | Fig.1
```

같은 레이아웃이 3장 넘게 이어지면 내용에 맞는 다른 레이아웃을 검토합니다.

## 4. 그림

```bash
python3 scripts/crop_figure.py paper.pdf --auto "Figure 2" --out figs/fig2.png
python3 scripts/crop_figure.py paper.pdf --page 5 --grid --out /tmp/p5.png   # 좌표 확인용
python3 scripts/crop_figure.py paper.pdf --page 5 --bbox 104 68 508 358 --out figs/fig2a.png
```

- `--list`의 bbox는 **캡션** 위치입니다. `--auto`는 캡션 위(표는 아래)의 그래픽을 모아 figure 영역을 추정합니다.
- `--grid` 렌더의 눈금은 PDF 포인트(1/72 in, 좌상단 원점)이고, 이 값을 그대로 `--bbox`에 씁니다.
- 자른 이미지는 반드시 열어서 확인합니다. 축 라벨이 잘렸거나 본문이 섞였으면 bbox로 다시 자릅니다.
- 여러 패널 figure는 슬라이드에서 말할 패널만 자릅니다 (`fig4a.png`, `fig4b.png`).

## 5. spec 작성과 생성

```bash
python3 scripts/build_deck.py deck.json "[2026Q3] ShortTitle.pptx"
python3 scripts/render_check.py "[2026Q3] ShortTitle.pptx" --out render
```

작업 폴더 예:

```
work/
  paper.pdf
  figs/fig1.png …
  deck.json
  render/contact.png
```

## 6. 검토 루프

1. `render_check.py` 경고를 모두 해결합니다. overflow 경고는 문장을 줄이거나 슬라이드를 나눠서 해결하고, 글자 크기는 줄이지 않습니다.
2. `render/contact.png`로 흐름과 반복을 봅니다. `render/slide-*.png`(10장 이상이면 `slide-01.png`처럼 0으로 채움)로 개별 슬라이드의 잘림과 정렬을 봅니다. `build_deck.py`의 `WARNING` 줄(작은 그림 등)도 확인합니다.
   - LibreOffice 렌더에서는 영문과 한글 조사 사이에 띄어쓰기처럼 보이는 간격이 생깁니다("Jacobian 으로"). PowerPoint에서는 나타나지 않으므로 고치지 않습니다.
3. 색 규칙을 점검합니다: 같은 도형에 다른 색이 없는지, 강조가 1~2개인지 봅니다.
4. 수치와 figure 번호를 논문과 대조합니다.

## 7. 전달

- 사용자의 발표 폴더에 `[YYYYQn] 짧은제목.pptx`로 저장합니다.
- 구성안에서 바뀐 점과 확인이 필요한 점(해석이 애매한 부분, 빠진 figure)을 짧게 알립니다.

## 리눅스 렌더링 폰트

LibreOffice가 `NanumSquareOTF`를 찾지 못하면 NanumSquare를 설치하고(`fonts-nanum`, `fonts-nanum-extra`) `/etc/fonts/local.conf`에 다음을 넣은 뒤 `fc-cache -f`를 실행합니다.

```xml
<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
 <match target="pattern"><test name="family"><string>NanumSquareOTF ExtraBold</string></test>
   <edit name="family" mode="assign" binding="strong"><string>NanumSquare</string></edit>
   <edit name="weight" mode="assign" binding="strong"><const>extrabold</const></edit></match>
 <match target="pattern"><test name="family"><string>NanumSquareOTF</string></test>
   <edit name="family" mode="assign" binding="strong"><string>NanumSquare</string></edit></match>
</fontconfig>
```
