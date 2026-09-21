# レイアウト関数リファレンス

使用例は `examples/sample_deck.js` にすべてある。迷ったらそちらをコピーして書き換えること。

```js
const { Deck } = require("/path/to/consulting-deck/scripts/deck_lib.js");
const d = new Deck({ title: "資料名", author: "作成者", status: "Draft" }); // statusは省略可
// ... スライドを追加 ...
await d.save("output.pptx");
```

## 固定グリッド（inch、16:9 = 13.333 × 7.5）

| 要素 | 位置 |
|---|---|
| セクション表示 | x0.5 y0.3（10pt グレー） |
| ステータス | 右上 x11.53 y0.25 w1.3 h0.34（枠線・文字 #7D7D7D、枠0.75pt、11pt 太字） |
| メッセージタイトル | x0.5 y0.62 w12.33 h0.85（20pt 太字、最大2行） |
| 上罫線 | y1.55、x0.5〜12.83（#DEDEDE、0.75pt） |
| 本文エリア | x0.5 y1.8 w12.33 h4.85 |
| 示唆ボックス（takeaway） | 本文エリア最下部 h0.6 |
| 下罫線 | y6.85、x0.5〜12.83（#DEDEDE、0.75pt） |
| 出典・注記 | x0.5 y6.9（9pt グレー） |
| ページ番号 | 右下 y6.9（9pt グレー、表紙以外に自動付番） |

## 塗りと枠線の定数

| 定数 | 値 | 用途 |
|---|---|---|
| `C.orange` | #D04A02 | 最強調の塗り（1スライド1まとまり）、強調の文字 |
| `C.tint` | #FBEAE0 | 強調エリアの塗り |
| `C.fill` | #F2F2F2 | 非強調の塗り（見出し・象限） |
| `C.line` | #DEDEDE | 非強調のバー、罫線・枠線の色 |
| `B.std` | #DEDEDE 0.75pt | 標準の枠線 |
| `B.thin` | #DEDEDE 0.5pt | リスト内の区切り線 |
| `B.key` | #D04A02 1.5pt | 最重要カードの枠 |
| `B.gap(w)` | 白 w pt | 隣接する塗り同士のすき間 |

## 間隔（近接のルール）

| 定数 | 値 | 使う場所 |
|---|---|---|
| `G.NEAR` | 0.1 inch | 関係する要素の間（見出し→本文、グラフ見出し→グラフ、表→補足） |
| `G.GAP` | 0.4 inch | まとまり同士の間（列、カード、結論と要点、本文と示唆ボックス） |
| 箇条書き | 親→子 3pt ／ 項目同士 8pt | ライブラリが自動適用 |

## 共通引数（コンテンツ系すべて）

- `title`：メッセージタイトル（必須。結論を1文で。`==強調==` 可）
- `section`：セクション表示（例 "現状と課題"）
- `source`：出典（「出典：」は自動で付く）
- `note`：注記（「注：」は自動で付く）
- `takeaway`：本文下部の示唆ボックス（1行。so what を書く）

`points`（箇条書き）は文字列の配列。2階層目は `{ text: "親", sub: ["子1", "子2"] }`。

## 各関数

### cover({ title, subtitle, date, author })

### executiveSummary({ title, lead, points: [{ head, body }] })
- `lead`：最重要の結論（最大2行、強調1か所）
- `points`：数は内容次第。head は短い見出し（「現状」「課題」など）、body は1〜2行

### agenda({ title, items: [...], current })
- `current`：強調するセクションの番号（0始まり）。セクションの冒頭で再掲するときに使う

### scope({ title, inScope, outOfScope, assumptions })
- `assumptions` は省略可（省略時は2列）

### columns({ title, columns: [{ label, points, key }] })
- 2〜4列。`key: true` の列見出しがオレンジになる

### flow({ title, columns: [{ label, points, key }] })
- 現状→課題→施策。既定で最後の列が強調。列数は2〜4

### textVisual({ title, points, visual, visualTitle, ratio })
- `visual: { chart: { type, data, highlight, options } }`
  - `type`：bar / line / pie / doughnut
  - `data`：pptxgenjs 形式 `[{ name, labels: [...], values: [...] }]`
  - `highlight`：注目箇所の index。指定したものだけオレンジ、他はグレーになる
    - 1系列の棒・円・ドーナツ → 項目（labels）の index
    - 複数系列の棒・折れ線 → 系列（data）の index
  - 凡例は自動で配置される（円：スライス外側に項目名＋% ／ ドーナツ：セグメント上に項目名＋% ／ 複数系列：グラフ上部 ／ 1系列の棒：軸ラベル＋visualTitle）
- `visual` を関数 `(slide, box, deck) => {}` にすると、右エリア box 内に自由描画できる（色は `C`、サイズは `F` の定数を使うこと）
- `ratio`：左の要点エリアの幅比率（既定 0.4）

### matrix2x2({ title, xAxis, yAxis, quadrants, keyQuadrant, items, points })
- `quadrants`：[左上, 右上, 左下, 右下] のラベル
- `keyQuadrant`：強調する象限（0〜3）
- `items`：`{ label, x: 0〜1, y: 0〜1, highlight }`（yは上が高い）

### comparison({ title, options, criteria: [{ name, values }], recommend, points })
- `recommend`：推奨案の index。列見出しがオレンジ、列の背景が淡色になる
- 評価記号は ◎○△× を使い、凡例を `note` に書く

### kpi({ title, kpis: [{ name, current, target, unit, note, key }], points })
- 1スライド4つまで。既定で1つ目が強調。`current` が不明なら省略

### roadmap({ title, periods, tracks: [{ name, tasks: [{ label, start, end, key }] }], milestones: [{ label, at }] })
- `start` / `end` は periods の index（end を含む）
- `key: true` のタスクがオレンジ

### nextSteps({ title, actions: [{ what, who, when }], decisions })
- `decisions`：ご判断いただきたい事項（省略すると表が全幅）

### text({ title, points })
- 箇条書きのみ。Appendix の算出根拠などに使う

### divider({ label, sub })
- 中扉。ページ番号は付く

## 出力
- `d.storyline()`：縦読み用のタイトル一覧（`save()` 時にも表示）
- `d.save(path)`：保存し、縦読みと要確認の警告を表示
