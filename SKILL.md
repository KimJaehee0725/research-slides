---
name: research-slides
description: DSBA 연구실 템플릿(Pantone 팔레트, NanumSquare)으로 논문 리뷰·연구 발표 PowerPoint를 만든다. Papers You Should Read(PYSR), 연구실 세미나, 논문 리뷰 영상, DSBA 템플릿이나 research-slides로 표시된 .pptx를 만들거나 고칠 때 사용한다.
---

# research-slides

논문 한 편(또는 여러 편)을 연구실 발표 자료로 만드는 스킬입니다. 슬라이드를 도형 단위로 손으로 배치하지 않습니다. 레이아웃을 고르고 **이름 있는 placeholder 역할**을 JSON spec으로 채운 뒤 `scripts/build_deck.py`로 생성합니다. 색과 도형 스타일은 스크립트가 정하므로 덱 전체가 같은 규칙을 따릅니다.

경로는 모두 이 스킬 폴더(`SKILL.md`가 있는 곳) 기준입니다. 원본 repo: https://github.com/KimJaehee0725/research-slides

## 준비

```bash
pip install python-pptx pymupdf pillow matplotlib   # 필요 시 --break-system-packages
# 수식: pandoc (brew install pandoc / apt install pandoc, 또는 pip install pypandoc-binary)
# 렌더링 검토용: LibreOffice(soffice), poppler(pdftoppm), NanumSquare 폰트
```

리눅스에서 렌더링하면 `NanumSquareOTF` 이름이 대체 폰트로 바뀔 수 있습니다. 이때는 `references/workflow.md`의 fontconfig 설정을 적용합니다.

## 작업 순서

1. **논문 읽기.** PDF 전체를 읽고 문제, 핵심 아이디어, 방법, 주요 결과, 한계를 정리합니다. `python3 scripts/crop_figure.py paper.pdf --list`로 figure·table 목록을 뽑습니다.
2. **구성안 확인.** 슬라이드별 `레이아웃 · 제목 · 핵심 메시지 · 사용할 figure`를 표로 만들어 사용자에게 먼저 보여줍니다. 발표 길이나 대상을 모르면 이 단계에서 묻습니다. (세부: `references/workflow.md`)
3. **그림 추출.** `crop_figure.py --auto "Figure 3"`로 잘라낸 뒤 **반드시 이미지를 직접 열어 확인**합니다. 잘림이 있으면 `--page N --grid`로 좌표를 보고 `--bbox`로 다시 자릅니다.
4. **spec 작성.** `references/spec.md` 스키마를 따릅니다. 예시는 `examples/pysr_example.json`에 있습니다.
5. **생성.** `python3 scripts/build_deck.py deck.json out.pptx`
6. **검토.** `python3 scripts/render_check.py out.pptx --out render/`로 경고를 확인하고 `render/contact.png`와 개별 슬라이드를 봅니다. 경고가 없고 눈으로 봐도 문제가 없을 때까지 spec을 고쳐 다시 생성합니다.
7. **내레이션.** 영상용이면 각 슬라이드의 `notes`에 대본을 씁니다. (`references/narration.md`)

## 반드시 지킬 규칙

- 템플릿은 `assets/research-slides-template.pptx`입니다. 사용자가 다른 복사본을 주더라도 레이아웃 이름이 같으면 그 파일을 `--template`으로 씁니다.
- 텍스트는 placeholder 역할에만 넣습니다. 레이아웃에 역할이 있는 곳에 자유 텍스트 상자를 만들지 않습니다.
- 도형은 `shapes`의 `kind`로만 고르고, 색을 직접 지정하지 않습니다. **같은 종류의 도형은 같은 색**입니다. 역할이 다를 때만 다른 kind를 씁니다. 슬라이드 하나에 `key`/`emph`는 1~2개까지만 씁니다.
- **제목은 명사 위주 4단어 이내, 명사형으로 끝맺습니다** (예: `RLHF vs SFT`, `실험 결과`, `Training Dynamics`, `J-space의 용량`). 주장·결론은 제목이 아니라 밴드 1단에 씁니다. 빌드 시 규칙을 어기면 `WARNING`이 나옵니다.
- **수식은 PowerPoint 수식(OMML)으로 넣습니다.** 문장 안에서는 `$...$`, 단독 수식은 `shapes`의 `math`(LaTeX)입니다. 논문 수식을 이미지로 캡쳐하지 않습니다. pandoc이 필요합니다.
- **중앙 온점(·)과 긴 대시(—, –)는 슬라이드에 쓰지 않습니다.** 나열은 쉼표, 구분은 콜론(:)이나 `/`, 범위는 `-`(Fig. 30-31)로 씁니다. 빌드 시 경고가 나옵니다.
- **줄간격은 기본 1.5**입니다(레이아웃과 텍스트 상자에 적용됨). 공간이 모자랄 때만 슬라이드에서 `"line_spacing": {"explain": 1.2}`처럼 조정합니다. 글자 크기는 줄이지 않습니다.
- 글자 강조는 `**핵심어**`(Dusty Cedar 굵게)와 `__구조 용어__`(Navy 굵게)만 씁니다. 한 줄에 하나를 넘지 않게 합니다.
- 논문 그림은 원본 캡쳐를 비율을 유지한 채 씁니다. 다시 그리지 않고, 강조는 `highlights` 틀로만 표시합니다.
- 밴드: 1단은 주장, 2단(`> `)은 근거·수치입니다. 32pt 기준 한 줄 40자, 3~4줄까지이며 명사형으로 끝맺습니다(~함/~임).
- 수치, 인용, 출처는 논문에 있는 것만 씁니다. 근거 없는 수치나 익명 인용은 쓰지 않습니다.
- `source`는 `meta.source`에 한 번 적으면 모든 본문 슬라이드에 들어갑니다. 다른 논문을 인용하는 슬라이드만 따로 지정합니다.
- 템플릿 디자인을 바꿔야 하면 `scripts/rs_style.py`와 `scripts/make_template.py`를 고친 뒤 `python3 scripts/make_template.py`로 다시 생성합니다. pptx를 직접 편집하지 않습니다.

## 레이아웃 요약

| 레이아웃 | 용도 | 역할 |
|---|---|---|
| `Title` | 표지 | title, citation, presenter, date, venue_logo |
| `Contents` | 목차·섹션 구분 | toc {items, current} |
| `A_FigureFull_Band` | 그림 1개 + 밴드 | section, title, figure, band, source |
| `B_FigureTwo_Band` | 그림 2개 비교 + 밴드 | figure_left, figure_right, band |
| `C_FigureExplain_Band` | 그림 + 해설 + 밴드 | figure, explain, band |
| `D_TopBand_Wide` | 상단 밴드 + 표·차트 | band, wide(또는 shapes) |
| `E_TwoColumn` | 2단 카드 | header_left/right, body_left/right |
| `F_Diagram_Band` | 도형·수식 + 밴드 | band, shapes |
| `G_TitleOnly` | 자유 구성, 부록 | shapes |
| `H_Statement` | 핵심 질문·한 문장 | statement, support |
| `Closing` | 마무리 | title, subtitle, date |

좌표, 크기, 글자 수 한도는 `references/layouts.md`, 색·도형·차트 규칙은 `references/style.md`에 있습니다.

## 전달 전 점검

- [ ] `render_check.py` 경고 0개, contact sheet 육안 확인
- [ ] 제목이 명사형 4단어 이내인지 (빌드 WARNING 0개), 핵심 메시지는 밴드에 있는지
- [ ] 수식이 PowerPoint 수식인지 (이미지 아님)
- [ ] 중앙 온점/긴 대시 경고 0개, 줄간격 조정은 필요한 슬라이드에만
- [ ] figure 번호와 수치가 논문과 일치하는지
- [ ] 강조색이 역할대로만 쓰였는지 (같은 도형에 다른 색 없음)
- [ ] 영상용이면 모든 슬라이드에 notes
- [ ] 파일명: `[YYYYQn] 짧은제목.pptx` (예: `[2026Q3] J-space.pptx`)
