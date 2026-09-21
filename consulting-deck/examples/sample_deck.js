// サンプル: 全レイアウトの使用例（新しい資料を作るときはこのファイルをコピーして中身を差し替える）
// 実行: node sample_deck.js [出力パス]
const path = require("path");
const { Deck } = require(path.join(__dirname, "../scripts/deck_lib.js"));

(async () => {
  const d = new Deck({ title: "営業DX推進に向けた提言", author: "〇〇部", status: "Draft" });
  const SEC = ["現状と課題", "打ち手の検討", "提言", "実行計画"];

  d.cover({ client: "〇〇株式会社", title: "営業DX推進に向けた提言", subtitle: "SFA刷新と営業プロセス標準化による受注率改善", date: "2026年9月", author: "〇〇部 経営企画" });

  d.executiveSummary({
    title: "SFA刷新とプロセス標準化により、3年で受注率を5pt改善できる",
    lead: "既存SFAの刷新とプロセス標準化を一体で進め、==受注率5pt改善==を目指すべき",
    points: [
      { head: "現状", body: "案件情報が担当者ごとに分散し、受注率は業界平均を下回る水準で推移している" },
      { head: "課題", body: "SFAの入力負荷が高く定着していないため、パイプラインを組織として管理できていない" },
      { head: "打ち手", body: "3案を比較した結果、クラウドSFAへの刷新と営業プロセス標準化の組み合わせが最も効果が高い" },
      { head: "効果", body: "受注率5pt改善により、年間売上で約12億円の増収を見込む（要試算）" },
      { head: "次の一手", body: "10月の経営会議で方針を承認いただき、11月からパイロットを開始したい" },
    ],
  });

  d.agenda({ items: ["前提条件・スコープ", ...SEC] });

  d.scope({
    title: "本検討は国内法人営業部門を対象とし、海外拠点は対象外とする",
    inScope: ["国内法人営業部門（約300名）", "案件管理から受注までのプロセス", "SFA・名刺管理ツール"],
    outOfScope: ["海外拠点の営業部門", "受注後の請求・債権管理", "マーケティング施策"],
    assumptions: ["現行SFAの保守契約は2027年3月で終了", "IT予算は年3億円を上限とする", "人員構成は現状維持"],
  });

  d.flow({
    title: "案件情報の分散が、受注率の伸び悩みを招いている",
    section: SEC[0],
    columns: [
      { label: "現状", points: ["受注率は18%で業界平均を3pt下回る", "案件情報はExcelと個人メモに分散"] },
      { label: "課題", points: ["SFAの入力項目が多く定着率は40%", "失注理由が蓄積されず改善につながらない"] },
      { label: "施策", points: ["入力負荷の低いSFAへ刷新", "営業プロセスを標準化し管理指標を統一"] },
    ],
    takeaway: "課題の根本はツールではなく、入力されない仕組みにある",
    source: "社内SFAログ（2025年度）、業界団体調査",
  });

  d.textVisual({
    title: "SFA定着率の低い部門ほど、受注率が低い傾向にある",
    section: SEC[0],
    points: ["定着率が60%を超える部門は、受注率が平均を上回る", "**第3営業部**は定着率・受注率ともに最低水準", "定着率の向上が受注率改善の前提条件となる"],
    visualTitle: "部門別の受注率（%）",
    visual: { chart: { type: "bar", data: [{ name: "受注率", labels: ["第1営業部", "第2営業部", "第3営業部", "第4営業部"], values: [22, 19, 12, 20] }], highlight: 2 } },
    source: "社内SFAログ（2025年度）",
  });

  d.matrix2x2({
    title: "効果と実現性の両面から、SFA刷新を最優先で進めるべき",
    section: SEC[1],
    xAxis: { label: "実現性", low: "低", high: "高" },
    yAxis: { label: "効果", low: "小", high: "大" },
    quadrants: ["中長期で検討", "最優先", "見送り", "早期に着手"],
    keyQuadrant: 1,
    items: [
      { label: "SFA刷新", x: 0.78, y: 0.82, highlight: true },
      { label: "プロセス標準化", x: 0.66, y: 0.7 },
      { label: "AI需要予測", x: 0.25, y: 0.75 },
      { label: "研修の拡充", x: 0.8, y: 0.3 },
      { label: "組織再編", x: 0.2, y: 0.25 },
    ],
    points: ["SFA刷新は効果・実現性ともに高い", "プロセス標準化はSFA刷新と同時に進めると効果が大きい", "AI需要予測はデータ整備後に再検討"],
    note: "効果・実現性は部門長ヒアリングに基づく定性評価",
  });

  d.comparison({
    title: "3案のうち、クラウドSFAへの刷新が費用対効果で最も優れる",
    section: SEC[1],
    options: ["A案：現行SFA改修", "B案：クラウドSFA刷新", "C案：自社開発"],
    criteria: [
      { name: "初期費用", values: ["1.0億円", "1.8億円", "4.5億円"] },
      { name: "運用負荷", values: ["△", "◎", "×"] },
      { name: "定着の見込み", values: ["△", "◎", "○"] },
      { name: "導入期間", values: ["6か月", "9か月", "18か月"] },
      { name: "総合評価", values: ["○", "==◎==", "△"] },
    ],
    recommend: 1,
    points: ["B案は初期費用がA案より高いが、運用負荷と定着の見込みで大きく上回る"],
    note: "評価は ◎＞○＞△＞× の4段階",
  });

  d.kpi({
    title: "施策の実行により、3年後に受注率5pt・売上12億円の改善を見込む",
    section: SEC[2],
    kpis: [
      { name: "受注率", current: "18", target: "23", unit: "%", note: "業界平均を2pt上回る水準" },
      { name: "SFA定着率", current: "40", target: "85", unit: "%", note: "週次の案件更新率で測定" },
      { name: "年間増収額", current: "―", target: "12", unit: "億円", note: "受注率改善分から試算（要精査）" },
    ],
    points: ["KPIは四半期ごとに経営会議で報告する"],
  });

  d.roadmap({
    title: "11月からパイロットを開始し、2027年度中に全部門へ展開する",
    section: SEC[3],
    periods: ["2026 Q4", "2027 Q1", "2027 Q2", "2027 Q3", "2027 Q4"],
    milestones: [{ label: "全社展開の判断", at: 1 }],
    tracks: [
      { name: "SFA刷新", tasks: [{ label: "要件定義・選定", start: 0, end: 0 }, { label: "パイロット導入", start: 1, end: 1, key: true }, { label: "全部門展開", start: 2, end: 4 }] },
      { name: "プロセス標準化", tasks: [{ label: "現行プロセス調査", start: 0, end: 0 }, { label: "標準プロセス策定・展開", start: 1, end: 3 }] },
      { name: "定着化", tasks: [{ label: "研修・効果測定", start: 2, end: 4 }] },
    ],
  });

  d.nextSteps({
    title: "10月の経営会議で方針承認をいただき、パイロットに着手したい",
    actions: [
      { what: "経営会議での方針承認", who: "経営企画", when: "10月20日" },
      { what: "SFAベンダー3社への提案依頼", who: "情報システム部", when: "10月末" },
      { what: "パイロット部門の選定", who: "営業本部", when: "10月末" },
    ],
    decisions: ["B案（クラウドSFA刷新）での推進可否", "初年度予算1.8億円の確保", "パイロット部門の指定"],
  });

  d.divider({ label: "Appendix", sub: "詳細データ・算出根拠" });

  d.text({
    title: "増収額は、受注率改善分に平均案件単価を掛けて試算した",
    section: "Appendix",
    points: ["年間商談数：約8,000件（2025年度実績）", "平均受注単価：約3,000万円", { text: "増収額の試算", sub: ["8,000件 × 5pt × 3,000万円 ＝ 約12億円", "商談数・単価は現状維持を前提とする"] }],
    source: "社内SFAログ（2025年度）",
  });

  d.textVisual({
    title: "商談数の3割を第2営業部が占め、SFA定着の効果が最も大きい",
    section: "Appendix",
    points: ["商談数の構成比は第1営業部に次いで第2営業部が大きい", "パイロット部門の候補として第2営業部を優先的に検討する"],
    visualTitle: "部門別の商談数構成比（2025年度）",
    visual: { chart: { type: "pie", highlight: 1, data: [{ name: "商談数", labels: ["第1営業部", "第2営業部", "第3営業部", "第4営業部"], values: [35, 30, 15, 20] }] } },
    source: "社内SFAログ（2025年度）",
  });

  await d.save(process.argv[2] || "sample_deck.pptx");
})();
