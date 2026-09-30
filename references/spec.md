# Deck spec (JSON) 스키마

`build_deck.py`의 입력 파일입니다. 상대 경로는 spec 파일 위치를 기준으로 합니다.

```json
{
  "meta": {
    "title": "파일 속성에 들어갈 제목",
    "date": "2026년 Q3",
    "footer": "Papers You Should Read  ·  DSBA Lab",
    "source": "Kim et al., NeurIPS 2026"
  },
  "slides": [ { "layout": "...", ... } ]
}
```

- `meta.date`: `Title`, `Closing`의 `date`에 기본값으로 들어갑니다.
- `meta.footer`: 하단 바 왼쪽 문구입니다. 세미나면 `"DSBA Seminar  ·  DSBA Lab"`처럼 바꿉니다.
- `meta.source`: 본문 레이아웃(A~H)의 `source` 기본값입니다. 특정 슬라이드에서 `"source": ""`로 두면 비웁니다.

## 슬라이드 공통 키

| 키 | 설명 |
|---|---|
| `layout` | 레이아웃 이름 (필수) |
| 역할 이름 | `title`, `band`, `figure` 등. `references/layouts.md` 참고 |
| `shapes` | 컴포넌트 목록 (아래) |
| `notes` | 발표자 노트 = 내레이션 대본 |
| `hidden` | `true`면 숨김 슬라이드 (부록, 백업) |
| `line_spacing` | 역할별 줄간격 조정. 예: `{"explain": 1.2}`. 기본은 1.5이며 공간이 모자랄 때만 씁니다 |

## 텍스트 값

- 문자열 또는 문자열 목록입니다. 목록의 원소 하나가 문단 하나입니다.
- `"> 근거"`는 2단(➢), `">> 세부"`는 3단(–)입니다. 1단은 ✓입니다. `band`, `explain`, `body_*`에만 단계가 적용됩니다.
- `**핵심어**`는 Dusty Cedar 굵게, `__용어__`는 Navy 굵게 표시됩니다.
- `$...$`는 PowerPoint 수식(OMML)으로 들어갑니다. 예: `"$r_t = b_t - b_{t-1}$를 보상으로 사용"`. 수식이 든 도형은 PowerPoint가 아닌 뷰어용 텍스트 대체본과 함께 저장됩니다.
- `title`은 명사 위주 4단어 이내, 명사형 종결입니다. 어기면 빌드 때 `WARNING`이 출력됩니다.
- 중앙 온점(·)과 긴 대시(—, –)를 쓰면 `WARNING`이 출력됩니다. 쉼표, 콜론, `/`, `-`로 바꿉니다.
- `presenter`처럼 줄만 나누는 역할도 목록으로 씁니다: `["서울대학교 산업공학과", "DSBA 연구실", "김재희"]`

## 그림 역할 (`figure`, `figure_left`, `figure_right`, `venue_logo`)

```json
"figure": "figs/fig2.png"
"figure": {"path": "figs/fig2.png",
           "crop": [0, 0, 0, 0.1],
           "align": "c",
           "highlights": [{"box": [0.55, 0.1, 0.4, 0.5], "label": "제안 방법"}]}
```

- 그림은 역할 영역 안에 **비율을 유지해 맞춥니다**(잘리지 않음).
- `crop`은 `[left, top, right, bottom]` 비율입니다. 캡쳐에 캡션이나 여백이 섞였을 때만 씁니다.
- `highlights.box`는 **표시된 그림 기준 비율**입니다. 그림 크기가 바뀌어도 위치가 유지됩니다.
- `label`은 틀 바로 아래에 18pt Dusty Cedar로 들어갑니다. 아래 공간이 없으면 틀 위로 올라가고, 그림 영역 밖으로는 나가지 않습니다. 짧게(10자 안팎) 씁니다.
- 그림이 영역의 45% 미만만 채우면 빌드 때 `WARNING`이 출력됩니다. 필요한 패널만 자르거나 레이아웃을 바꿉니다. (`references/layouts.md`의 비율 표)

## 목차 (`Contents`)

```json
{"layout": "Contents", "toc": {"items": ["Background", "Method", "Experiments", "Discussion", "Conclusion"], "current": 1}}
```

섹션이 시작될 때마다 `current`만 바꿔 다시 넣습니다. 0부터 셉니다.

## shapes

모든 항목에 `type`이 있고, 대부분 `box: [x, y, w, h]`가 있습니다. `"frame": "zone"`이면 레이아웃 자유 영역 대비 비율이고, 생략하면 인치 단위 절대 좌표입니다.

| type | 키 | 설명 |
|---|---|---|
| `box` | text, kind, size(24), align(l/c/r), shape(round/rect/pill/oval) | 기본 도형 |
| `flow` | items, key[], emph[], vertical, gap, size | 화살표로 이은 상자 행렬 |
| `chevrons` | items, key[], size | 단계 파이프라인 |
| `arrow` | from[x,y], to[x,y], kind(default/key/emph), head(end/start/both) | 연결선 |
| `highlight` | box | 강조 틀 (Dusty Cedar 3pt) |
| `text` | text, style(body/muted/key/struct/caption), size, align, anchor(t/m/b), bullets, line_spacing(1.5) | 보조 텍스트 |
| `image` | path, crop, align | 추가 이미지 (아이콘, 생성 이미지) |
| `math` | tex, size(36), align, render | PowerPoint 수식(OMML, 편집 가능). 대체 이미지 포함. `render: "image"`면 이미지로만 |
| `table` | rows, col_widths, size(20), row_h(0.6), ours_rows[], key_cells[[r,c]], first_col, align(c/l) | 표 (box의 h는 무시). 문장 열은 `align: "l"` |
| `chart` | categories, series[{name, values, role}], kind(bar/hbar/line), ylabel, labels, number_format | PowerPoint 네이티브 차트 |

### box `kind`

| kind | 용도 |
|---|---|
| `default` | 일반 모듈, 단계, 입력·출력 |
| `key` | 논문이 제안한 핵심 요소 (슬라이드당 1~2개) |
| `emph` | 문제점, 한계, 핵심 주장 |
| `ok` | 검증 통과, 정답 |
| `note` | 참고, 낮은 우선순위 |
| `muted` | baseline, 기존 방법, 비활성 |
| `outline` | 여러 도형을 묶는 틀 |

### chart series `role`

`ours` (Dusty Cedar), `baseline` (Ultimate Gray), `alt` (Provence), `alt2` (Navy), `ok`, `note`.
Ours가 둘 이상이면 가장 중요한 것 하나만 `ours`로 두고 나머지는 `alt`로 둡니다.

## 예시

`examples/pysr_example.json`(PYSR 흐름)과 `examples/template_showcase.json`(모든 레이아웃과 컴포넌트)을 참고합니다.
