import { useState, useRef } from "react";

const CATEGORIES = [
  {
    id: "basic", label: "企業基本情報", color: "#1e3a8a",
    items: [
      { id: "company_overview",  label: "会社概要",              desc: "設立・本社・従業員数・上場区分" },
      { id: "business_products", label: "事業領域・主力商材",    desc: "取扱品・サービスラインナップ" },
      { id: "group_structure",   label: "グループ構造",          desc: "親会社・子会社・資本関係" },
      { id: "bases_network",     label: "拠点・物流ネットワーク", desc: "拠点数・倉庫所在地・配送エリア" },
    ],
  },
  {
    id: "market", label: "市場・競合", color: "#0f766e",
    items: [
      { id: "market_share",      label: "業界シェア・市場ポジション", desc: "業界内順位・シェア推移" },
      { id: "competitors",       label: "主要競合比較",               desc: "競合他社・差別化ポイント" },
      { id: "market_size",       label: "市場規模・成長率",           desc: "TAM・市場トレンド・予測" },
      { id: "market_structure",  label: "マーケット構造",             desc: "バリューチェーン・プレイヤー構造" },
    ],
  },
  {
    id: "finance", label: "財務", color: "#b45309",
    items: [
      { id: "pl_summary",        label: "売上・利益サマリー",  desc: "直近3期の売上・営業利益・率" },
      { id: "growth_profit",     label: "成長性・収益性評価",  desc: "CAGR・利益率トレンド・評価" },
      { id: "financial_health",  label: "財務健全性",          desc: "自己資本比率・有利子負債・格付" },
      { id: "capex",             label: "投資・設備動向",      desc: "CAPEX・設備投資方針・DX予算" },
    ],
  },
  {
    id: "strategy", label: "戦略・方針", color: "#7c3aed",
    items: [
      { id: "mid_term_plan",    label: "中期経営計画",         desc: "KPI・重点施策・投資計画" },
      { id: "dx_strategy",      label: "DX・自動化戦略",       desc: "IT投資方針・システム化動向" },
      { id: "hr_hiring",        label: "人材・採用動向",       desc: "採用状況・人手不足感" },
      { id: "sustainability",   label: "サステナビリティ方針", desc: "ESG・CO2削減・物流環境対応" },
    ],
  },
  {
    id: "logistics", label: "物流特性", color: "#be185d",
    items: [
      { id: "logistics_flow",   label: "物流フロー概要",   desc: "入荷〜保管〜出荷の基本フロー" },
      { id: "product_features", label: "取扱品の特性",     desc: "温度帯・サイズ・危険物・重量" },
      { id: "seasonality",      label: "繁閑・季節変動",   desc: "ピーク時期・変動幅・対応状況" },
      { id: "existing_systems", label: "既存システム環境", desc: "WMS・TMS・ERP・連携状況" },
    ],
  },
  {
    id: "contact", label: "担当者・組織", color: "#0369a1",
    items: [
      { id: "org_structure",  label: "部署・組織構成",           desc: "物流系部署の階層・組織図" },
      { id: "key_person",     label: "想定キーパーソン",         desc: "担当役職・意思決定者・窓口" },
      { id: "decision_flow",  label: "決裁フロー",               desc: "稟議プロセス・承認ライン・期間" },
      { id: "recent_news",    label: "直近のニュース・トピック", desc: "IR・プレスリリース・受賞歴" },
    ],
  },
];

const ALL_ITEMS = CATEGORIES.flatMap(c => c.items.map(i => ({ ...i, categoryColor: c.color, categoryLabel: c.label })));
const DEFAULT_SELECTED = new Set(["company_overview","business_products","market_share","competitors","pl_summary","growth_profit","mid_term_plan","logistics_flow","product_features","org_structure","key_person","recent_news"]);

const STATUS = {
  confirmed:   { label: "確認済み",   bg: "#dcfce7", color: "#15803d", border: "#bbf7d0" },
  partial:     { label: "一部のみ",   bg: "#fef9c3", color: "#a16207", border: "#fde68a" },
  unconfirmed: { label: "確認できず", bg: "#fee2e2", color: "#dc2626", border: "#fecaca" },
};

const FACT_ONLY = "公開情報・報道・公式サイトで確認できた事実のみ記載。推測・類推・一般論は禁止。不明な場合はstatusをunconfirmedにして空白にする。";

const PROMPTS = {
  company_overview:  (c) => `"${c}"の会社概要（設立・本社・従業員数・上場区分）を調査。${FACT_ONLY} JSONのみ返答: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報（なければ空文字）"}`,
  business_products: (c) => `"${c}"の事業領域・主力商材を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  group_structure:   (c) => `"${c}"のグループ構造（親会社・子会社・資本関係）を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  bases_network:     (c) => `"${c}"の拠点数・物流ネットワークを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  market_share:      (c) => `"${c}"の業界シェア・市場ポジションを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  competitors:       (c) => `"${c}"の主要競合と差別化ポイントを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  market_size:       (c) => `"${c}"が属する市場の規模・成長率を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  market_structure:  (c) => `"${c}"の業界バリューチェーン・プレイヤー構造を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  pl_summary:        (c) => `"${c}"の直近3期の売上・営業利益・率を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","revenue":"売上高","growth":"成長率","margin":"営業利益率","health":"財務評価","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  growth_profit:     (c) => `"${c}"の成長性・収益性トレンドを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  financial_health:  (c) => `"${c}"の財務健全性（自己資本比率・有利子負債）を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  capex:             (c) => `"${c}"の設備投資・CAPEX・DX予算動向を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  mid_term_plan:     (c) => `"${c}"の中期経営計画（KPI・重点施策）を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  dx_strategy:       (c) => `"${c}"のDX・自動化戦略・IT投資方針を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  hr_hiring:         (c) => `"${c}"の人材・採用動向・人手不足への対応を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  sustainability:    (c) => `"${c}"のESG・CO2削減・サステナビリティ方針を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  logistics_flow:    (c) => `"${c}"の物流フロー（入荷〜保管〜出荷）概要を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  product_features:  (c) => `"${c}"が扱う商品の特性（温度帯・サイズ・重量）を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  seasonality:       (c) => `"${c}"の繁閑・季節変動・ピーク時期を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  existing_systems:  (c) => `"${c}"が使用するWMS・TMS・ERP等の既存システムを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  org_structure:     (c) => `"${c}"の物流・DX系部署の組織構成を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","org_structure":"階層表記","dept":"担当部署","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  key_person:        (c) => `"${c}"の物流・システム系キーパーソン・担当役職を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","role":"想定役職","authority":"想定決裁権限","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  decision_flow:     (c) => `"${c}"の稟議・意思決定プロセス・承認ラインを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
  recent_news:       (c) => `"${c}"の直近1年のニュース・IR・プレスリリースを調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`,
};

// 環境変数からAPIキーを取得（Viteの場合はVITE_プレフィックス必須）
const API_KEY = import.meta.env.VITE_ANTHROPIC_API_KEY || "";

async function callClaude(system, user) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
  method: "POST",
  headers: {
    "Content-Type": "application/json",
    "x-api-key": API_KEY,
    "anthropic-version": "2023-06-01",
    "anthropic-dangerous-direct-browser-access": "true",
  },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      system,
      messages: [{ role: "user", content: user }],
      tools: [{ type: "web_search_20250305", name: "web_search" }],
    }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.map(b => b.text || "").filter(Boolean).join("\n") || "";
}

function extractJSON(text) {
  try {
    const clean = text.replace(/```json[\s\S]*?```|```/g, "").trim();
    const s = clean.indexOf("{"), e = clean.lastIndexOf("}");
    if (s !== -1 && e !== -1) return JSON.parse(clean.slice(s, e + 1));
  } catch (_) {}
  return null;
}

function StatusBadge({ status, size = "sm" }) {
  if (!status) return null;
  const cfg = STATUS[status] || STATUS.unconfirmed;
  return (
    <span style={{
      fontSize: size === "lg" ? 12 : 10, fontWeight: 600, color: cfg.color,
      background: cfg.bg, border: `1px solid ${cfg.border}`,
      padding: size === "lg" ? "3px 9px" : "1px 6px", borderRadius: 99, whiteSpace: "nowrap", flexShrink: 0,
    }}>{cfg.label}</span>
  );
}

function Checkbox({ checked, color, indeterminate }) {
  return (
    <div style={{
      width: 14, height: 14, borderRadius: 3, flexShrink: 0, marginTop: 1,
      border: `1.5px solid ${checked || indeterminate ? color : "#d1d5db"}`,
      background: checked ? color : "transparent",
      display: "flex", alignItems: "center", justifyContent: "center",
    }}>
      {checked && <span style={{ fontSize: 9, color: "#fff", fontWeight: 700, lineHeight: 1 }}>✓</span>}
      {indeterminate && !checked && <div style={{ width: 6, height: 2, background: color, borderRadius: 1 }} />}
    </div>
  );
}

function ResultCard({ item, result }) {
  const [open, setOpen] = useState(true);
  const st = result?.status || "unconfirmed";
  const cfg = STATUS[st];
  const borderColor = st === "unconfirmed" ? "#fecaca" : st === "partial" ? "#fde68a" : "#e5e7eb";
  const headerBg = st === "unconfirmed" ? "#fff5f5" : st === "partial" ? "#fffdf0" : "#fafafa";
  return (
    <div style={{ border: `1px solid ${borderColor}`, borderRadius: 10, marginBottom: 8, overflow: "hidden", background: "#fff" }}>
      <div onClick={() => setOpen(o => !o)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 14px", cursor: "pointer", background: headerBg }}>
        <div style={{ width: 3, height: 14, background: cfg.color, borderRadius: 2, flexShrink: 0 }} />
        <span style={{ fontSize: 13, fontWeight: 600, color: "#111827", flex: 1 }}>{item.label}</span>
        <StatusBadge status={st} />
        <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 4 }}>{open ? "▲" : "▼"}</span>
      </div>
      {open && (
        <div style={{ padding: "10px 14px 12px", borderTop: `1px solid ${cfg.border}` }}>
          <p style={{ margin: "0 0 8px", fontSize: 13, color: "#374151", lineHeight: 1.85 }}>{result?.summary || "情報を取得できませんでした。"}</p>
          {result?.missing && st !== "confirmed" && (
            <div style={{ display: "flex", alignItems: "flex-start", gap: 7, background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 7, padding: "7px 11px", marginTop: 6 }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: cfg.color, whiteSpace: "nowrap", marginTop: 1 }}>
                {st === "unconfirmed" ? "⚠ 要ヒアリング" : "△ 要確認"}
              </span>
              <span style={{ fontSize: 12, color: cfg.color, lineHeight: 1.6 }}>{result.missing}</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [company, setCompany] = useState("");
  const [visitDate, setVisitDate] = useState("");
  const [salesRep, setSalesRep] = useState("");
  const [priorInfo, setPriorInfo] = useState("");
  const [parsedPrior, setParsedPrior] = useState(null);
  const [selected, setSelected] = useState(new Set(DEFAULT_SELECTED));
  const [appStatus, setAppStatus] = useState("idle");
  const [progress, setProgress] = useState({ done: 0, total: 0, current: "" });
  const [results, setResults] = useState({});
  const [activeCategory, setActiveCategory] = useState(null);
  const [apiKeyMissing, setApiKeyMissing] = useState(!API_KEY);
  const abortRef = useRef(false);

  const toggle = (id) => setSelected(s => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleCat = (cat) => {
    const ids = cat.items.map(i => i.id);
    const allOn = ids.every(id => selected.has(id));
    setSelected(s => { const n = new Set(s); ids.forEach(id => allOn ? n.delete(id) : n.add(id)); return n; });
  };

  const statusCounts = Object.values(results).reduce((acc, r) => {
    const k = r?.status || "unconfirmed"; acc[k] = (acc[k] || 0) + 1; return acc;
  }, {});

  const hearingItems = ALL_ITEMS.filter(i => selected.has(i.id) && results[i.id]?.status === "unconfirmed");
  const partialItems = ALL_ITEMS.filter(i => selected.has(i.id) && results[i.id]?.status === "partial");

  const runResearch = async () => {
    if (!company.trim() || selected.size === 0 || !API_KEY) return;
    abortRef.current = false;
    setAppStatus("loading");
    setResults({});
    setParsedPrior(null);
    setActiveCategory(null);
    const items = ALL_ITEMS.filter(i => selected.has(i.id));
    const hasPrior = priorInfo.trim().length > 0;
    setProgress({ done: 0, total: items.length + (hasPrior ? 1 : 0), current: "" });

    const SYSTEM = `あなたはB2B営業支援AIです。Web検索で企業の公開情報を調査しJSONのみ返答。

【厳守ルール】
- summaryにはWeb上で実際に確認できた事実のみを記載する
- 推測・類推・一般論・「〜と思われる」「〜の可能性がある」「一般的に〜」「〜が考えられる」は一切書かない
- 情報が見つからない・不確かな場合はstatusをunconfirmedにし、summaryには「公開情報なし」と記載する
- 一部しか確認できない場合はstatusをpartialにし、確認できた事実のみsummaryに書く
- missingには「訪問時に直接確認すべき具体的な質問」を書く（推定ではなく確認事項として）
- 前置き・マークダウン不要。JSONのみ返答。`;

    let priorContext = "";
    if (hasPrior) {
      setProgress({ done: 0, total: items.length + 1, current: "事前情報を整理中" });
      try {
        const parseRaw = await callClaude(
          `あなたはB2B営業支援AIです。貼り付けられたヒアリングメモ・事前情報を整理しJSONのみ返答。前置き不要。`,
          `以下は"${company}"に関する事前情報です。構造化してください。\n---\n${priorInfo}\n---\nJSONのみ: {"summary":"全体サマリー3〜4文","key_points":["重要ポイント1","重要ポイント2"],"known_challenges":"判明している課題・ニーズ","known_contacts":"判明している担当者情報","known_systems":"判明しているシステム環境","other":"その他有用情報"}`
        );
        const pp = extractJSON(parseRaw);
        if (pp) {
          setParsedPrior(pp);
          priorContext = "\n\n【事前情報】" + JSON.stringify(pp);
        }
      } catch (_) {}
      setProgress({ done: 1, total: items.length + 1, current: "" });
    }

    const newResults = {};
    const offset = hasPrior ? 1 : 0;
    for (let i = 0; i < items.length; i++) {
      if (abortRef.current) break;
      const item = items[i];
      setProgress({ done: i + offset, total: items.length + offset, current: item.label });
      try {
        const base = PROMPTS[item.id]?.(company) || `"${company}"について「${item.label}」を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"不足情報"}`;
        const prompt = priorContext
          ? base + priorContext + "\n※事前情報に記載された内容はそのままsummaryに含めてよい。公開情報で追加確認できた事実のみ追記する。推測は書かない。事前情報で判明済みの内容はstatusをconfirmedかpartialに。"
          : base;
        const raw = await callClaude(SYSTEM, prompt);
        const parsed = extractJSON(raw);
        newResults[item.id] = parsed || { summary: raw.slice(0, 200), status: "partial", missing: "" };
      } catch {
        newResults[item.id] = { summary: "取得できませんでした。", status: "unconfirmed", missing: "デスクトップリサーチでは確認不可。訪問時に直接確認が必要です。" };
      }
      setResults({ ...newResults });
    }
    setProgress(p => ({ ...p, done: p.total, current: "" }));
    setAppStatus("done");
    setActiveCategory(CATEGORIES[0].id);
  };

  const hasResults = Object.keys(results).length > 0;
  const displayCat = CATEGORIES.find(c => c.id === activeCategory) || CATEGORIES[0];

  const inp = (label, val, setter, ph, type = "text") => (
    <div style={{ marginBottom: 7 }}>
      <div style={{ fontSize: 10, fontWeight: 600, color: "#6b7280", marginBottom: 3 }}>{label}</div>
      <input type={type} value={val} onChange={e => setter(e.target.value)} placeholder={ph}
        style={{ width: "100%", padding: "7px 9px", fontSize: 12, border: "1px solid #e5e7eb", borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none" }}
        onFocus={e => e.target.style.borderColor = "#1e3a8a"}
        onBlur={e => e.target.style.borderColor = "#e5e7eb"} />
    </div>
  );

  return (
    <div style={{ fontFamily: "'Helvetica Neue','Hiragino Kaku Gothic ProN',Meiryo,sans-serif", background: "#f8f9fc", minHeight: "100vh" }}>

      <div style={{ background: "#1e3a8a", padding: "12px 22px", display: "flex", alignItems: "center", gap: 11 }}>
        <div style={{ width: 28, height: 28, background: "#F96167", borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>📦</div>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#fff" }}>SBR 顧客リサーチ &amp; 社内報告ツール</div>
          <div style={{ fontSize: 11, color: "#93c5fd" }}>新規顧客の事前調査 → 社内報告PPT 一気通貫</div>
        </div>
      </div>

      {/* APIキー未設定の警告 */}
      {apiKeyMissing && (
        <div style={{ background: "#fef2f2", borderBottom: "1px solid #fecaca", padding: "10px 22px", fontSize: 13, color: "#dc2626" }}>
          ⚠ <strong>VITE_ANTHROPIC_API_KEY</strong> が設定されていません。Vercelの環境変数に設定してください。
        </div>
      )}

      <div style={{ display: "flex", height: `calc(100vh - ${apiKeyMissing ? 90 : 50}px)` }}>

        {/* 左パネル */}
        <div style={{ width: 292, background: "#fff", borderRight: "1px solid #e5e7eb", overflowY: "auto", display: "flex", flexDirection: "column", flexShrink: 0 }}>
          <div style={{ padding: "14px 14px 0", flex: 1 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: "#1e3a8a", marginBottom: 8, display: "flex", alignItems: "center", gap: 5 }}>
              <div style={{ width: 3, height: 11, background: "#1e3a8a", borderRadius: 2 }} />基本情報
            </div>
            {inp("顧客企業名 *", company, setCompany, "例：ニチレイロジグループ株式会社")}
            {inp("訪問予定日", visitDate, setVisitDate, "", "date")}
            {inp("担当者名", salesRep, setSalesRep, "例：田中 太郎")}

            <div style={{ height: 1, background: "#f3f4f6", margin: "10px 0" }} />

            <div style={{ fontSize: 11, fontWeight: 600, color: "#1e3a8a", marginBottom: 6, display: "flex", alignItems: "center", gap: 5 }}>
              <div style={{ width: 3, height: 11, background: "#1e3a8a", borderRadius: 2 }} />事前情報
              <span style={{ fontWeight: 400, color: "#9ca3af", fontSize: 10 }}>任意</span>
            </div>
            <div style={{ marginBottom: 4, fontSize: 11, color: "#6b7280", lineHeight: 1.5 }}>
              ヒアリングメモ・打合せ記録などをそのまま貼り付けてください。
            </div>
            <textarea
              value={priorInfo}
              onChange={e => setPriorInfo(e.target.value)}
              placeholder={"例：\n・担当の山田部長から「WMS刷新を検討中」と聞いた\n・現状はExcel管理でミスが多い"}
              rows={4}
              style={{ width: "100%", padding: "8px 9px", fontSize: 12, border: "1px solid #e5e7eb", borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none", resize: "vertical", lineHeight: 1.6 }}
              onFocus={e => e.target.style.borderColor = "#1e3a8a"}
              onBlur={e => e.target.style.borderColor = "#e5e7eb"}
            />
            {priorInfo.trim() && (
              <div style={{ fontSize: 11, color: "#15803d", marginBottom: 4, marginTop: 3 }}>
                ✓ 事前情報あり — リサーチ時に自動反映します
              </div>
            )}

            <div style={{ height: 1, background: "#f3f4f6", margin: "10px 0" }} />

            <div style={{ fontSize: 11, fontWeight: 600, color: "#1e3a8a", marginBottom: 8, display: "flex", alignItems: "center", gap: 5 }}>
              <div style={{ width: 3, height: 11, background: "#1e3a8a", borderRadius: 2 }} />
              調査項目　<span style={{ fontWeight: 400, color: "#9ca3af", fontSize: 10 }}>{selected.size}件選択</span>
            </div>

            {CATEGORIES.map(cat => {
              const ids = cat.items.map(i => i.id);
              const allOn = ids.every(id => selected.has(id));
              const partial = ids.some(id => selected.has(id)) && !allOn;
              return (
                <div key={cat.id} style={{ marginBottom: 9 }}>
                  <div onClick={() => toggleCat(cat)} style={{ display: "flex", alignItems: "center", gap: 7, padding: "4px 6px", borderRadius: 6, cursor: "pointer", marginBottom: 3 }}>
                    <Checkbox checked={allOn} indeterminate={partial} color={cat.color} />
                    <div style={{ width: 7, height: 7, borderRadius: "50%", background: cat.color, flexShrink: 0 }} />
                    <span style={{ fontSize: 11, fontWeight: 600, color: "#374151" }}>{cat.label}</span>
                  </div>
                  <div style={{ paddingLeft: 10 }}>
                    {cat.items.map(item => {
                      const on = selected.has(item.id);
                      const r = results[item.id];
                      return (
                        <label key={item.id} style={{ display: "flex", alignItems: "flex-start", gap: 7, padding: "4px 6px", borderRadius: 6, cursor: "pointer", marginBottom: 1 }}>
                          <Checkbox checked={on} color={cat.color} />
                          <input type="checkbox" checked={on} onChange={() => toggle(item.id)} style={{ display: "none" }} />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontSize: 12, color: on ? "#111827" : "#9ca3af", fontWeight: on ? 500 : 400, display: "flex", alignItems: "center", gap: 5, flexWrap: "wrap", lineHeight: 1.4 }}>
                              {item.label}
                              {r && <StatusBadge status={r.status} />}
                            </div>
                          </div>
                        </label>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>

          <div style={{ padding: "12px 14px", borderTop: "1px solid #f3f4f6" }}>
            {appStatus === "loading" && (
              <div style={{ marginBottom: 8 }}>
                <div style={{ background: "#e5e7eb", borderRadius: 99, height: 4, overflow: "hidden", marginBottom: 4 }}>
                  <div style={{ height: "100%", width: `${Math.round((progress.done / progress.total) * 100)}%`, background: "#1e3a8a", transition: "width .4s", borderRadius: 99 }} />
                </div>
                <div style={{ fontSize: 11, color: "#6b7280", textAlign: "center" }}>
                  {progress.current ? `「${progress.current}」調査中…` : "完了"}　{progress.done}/{progress.total}
                </div>
              </div>
            )}
            <button onClick={runResearch}
              disabled={appStatus === "loading" || !company.trim() || !selected.size || !API_KEY}
              style={{
                width: "100%", padding: "10px", fontSize: 13, fontWeight: 700, fontFamily: "inherit",
                background: (appStatus === "loading" || !company.trim() || !selected.size || !API_KEY) ? "#d1d5db" : "#1e3a8a",
                color: "#fff", border: "none", borderRadius: 8,
                cursor: (appStatus === "loading" || !company.trim() || !selected.size || !API_KEY) ? "not-allowed" : "pointer",
              }}>
              {appStatus === "loading" ? "調査中…" : "🔍 リサーチ開始"}
            </button>
          </div>
        </div>

        {/* 右パネル */}
        <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
          {!hasResults ? (
            <div style={{ height: "100%", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: "#d1d5db" }}>
              <div style={{ fontSize: 48, marginBottom: 12 }}>🔍</div>
              <div style={{ fontSize: 15, fontWeight: 600 }}>顧客名を入力してリサーチ開始</div>
              <div style={{ fontSize: 12, marginTop: 6 }}>調査項目を選んでAIが自動でWeb検索・分析します</div>
            </div>
          ) : (
            <div style={{ maxWidth: 820 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                <div>
                  <div style={{ fontSize: 20, fontWeight: 700, color: "#111827" }}>{company}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>
                    {visitDate && `訪問予定：${visitDate}　`}{salesRep && `担当：${salesRep}`}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", justifyContent: "flex-end" }}>
                  {["confirmed","partial","unconfirmed"].map(s => statusCounts[s] ? (
                    <div key={s} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <StatusBadge status={s} size="lg" />
                      <span style={{ fontSize: 12, color: "#6b7280" }}>{statusCounts[s]}件</span>
                    </div>
                  ) : null)}
                </div>
              </div>

              {parsedPrior && (
                <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 10, padding: "12px 16px", marginBottom: 12 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#15803d", marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                    <span>📋 事前情報（整理済み）</span>
                    <span style={{ fontSize: 10, fontWeight: 400, color: "#16a34a", background: "#dcfce7", border: "1px solid #bbf7d0", padding: "1px 7px", borderRadius: 99 }}>リサーチに反映済み</span>
                  </div>
                  <p style={{ margin: "0 0 8px", fontSize: 13, color: "#166534", lineHeight: 1.75 }}>{parsedPrior.summary}</p>
                  {parsedPrior.key_points?.length > 0 && (
                    <div style={{ marginBottom: 6 }}>
                      <div style={{ fontSize: 11, fontWeight: 600, color: "#15803d", marginBottom: 4 }}>重要ポイント</div>
                      {parsedPrior.key_points.map((pt, i) => (
                        <div key={i} style={{ display: "flex", gap: 6, fontSize: 12, color: "#166534", marginBottom: 2 }}>
                          <span style={{ flexShrink: 0 }}>•</span><span>{pt}</span>
                        </div>
                      ))}
                    </div>
                  )}
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "4px 16px", marginTop: 6 }}>
                    {[["判明している課題", parsedPrior.known_challenges], ["担当者情報", parsedPrior.known_contacts], ["システム環境", parsedPrior.known_systems], ["その他情報", parsedPrior.other]].filter(([, v]) => v).map(([k, v]) => (
                      <div key={k}><span style={{ fontSize: 10, fontWeight: 600, color: "#15803d" }}>{k}：</span><span style={{ fontSize: 12, color: "#166534" }}>{v}</span></div>
                    ))}
                  </div>
                </div>
              )}

              {appStatus === "done" && hearingItems.length > 0 && (
                <div style={{ background: "#fff5f5", border: "1px solid #fecaca", borderRadius: 10, padding: "12px 16px", marginBottom: 12 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#dc2626", marginBottom: 8 }}>
                    ⚠️ デスクトップリサーチで確認できなかった項目 ({hearingItems.length}件) — 初回訪問で確認
                  </div>
                  {hearingItems.map(item => (
                    <div key={item.id} style={{ display: "flex", gap: 8, fontSize: 12, marginBottom: 4 }}>
                      <span style={{ color: "#dc2626", flexShrink: 0 }}>•</span>
                      <span style={{ fontWeight: 600, color: "#111827", flexShrink: 0, minWidth: 120 }}>{item.label}</span>
                      <span style={{ color: "#6b7280" }}>{results[item.id]?.missing || "訪問時に確認"}</span>
                    </div>
                  ))}
                </div>
              )}

              {appStatus === "done" && partialItems.length > 0 && (
                <div style={{ background: "#fffdf0", border: "1px solid #fde68a", borderRadius: 10, padding: "10px 16px", marginBottom: 12 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: "#a16207", marginBottom: 6 }}>
                    △ 情報が一部のみ取得できた項目 ({partialItems.length}件) — 追加確認推奨
                  </div>
                  {partialItems.map(item => (
                    <div key={item.id} style={{ display: "flex", gap: 8, fontSize: 12, marginBottom: 3 }}>
                      <span style={{ color: "#a16207", flexShrink: 0 }}>•</span>
                      <span style={{ fontWeight: 600, color: "#111827", flexShrink: 0, minWidth: 120 }}>{item.label}</span>
                      <span style={{ color: "#78350f" }}>{results[item.id]?.missing || ""}</span>
                    </div>
                  ))}
                </div>
              )}

              <div style={{ display: "flex", gap: 4, marginBottom: 14, flexWrap: "wrap" }}>
                {CATEGORIES.filter(cat => cat.items.some(i => results[i.id])).map(cat => {
                  const catItems = cat.items.filter(i => results[i.id]);
                  const hasU = catItems.some(i => results[i.id]?.status === "unconfirmed");
                  const hasP = catItems.some(i => results[i.id]?.status === "partial");
                  const active = activeCategory === cat.id;
                  return (
                    <button key={cat.id} onClick={() => setActiveCategory(cat.id)} style={{
                      padding: "5px 12px", fontSize: 12, fontWeight: 600, borderRadius: 20, cursor: "pointer", fontFamily: "inherit",
                      background: active ? cat.color : "#fff", color: active ? "#fff" : "#374151",
                      border: `1.5px solid ${active ? cat.color : "#e5e7eb"}`,
                      display: "flex", alignItems: "center", gap: 5,
                    }}>
                      {cat.label}
                      {hasU && <span style={{ width: 7, height: 7, borderRadius: "50%", background: active ? "#fca5a5" : "#ef4444", display: "inline-block", flexShrink: 0 }} />}
                      {!hasU && hasP && <span style={{ width: 7, height: 7, borderRadius: "50%", background: active ? "#fde68a" : "#f59e0b", display: "inline-block", flexShrink: 0 }} />}
                    </button>
                  );
                })}
              </div>

              {displayCat && (
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: displayCat.color, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
                    <div style={{ width: 3, height: 14, background: displayCat.color, borderRadius: 2 }} />
                    {displayCat.label}
                  </div>
                  {displayCat.items.filter(i => results[i.id]).map(item => (
                    <ResultCard key={item.id} item={item} result={results[item.id]} />
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
