import { useState, useRef, useEffect } from "react";
import PptxGenJS from "pptxgenjs";

const C = {
  blue: "#0068B7", blueDark: "#00519A", blueLight: "#E8F4FF",
  white: "#FFFFFF", gray50: "#F5F7FA", gray100: "#F0F0F0",
  gray200: "#E0E0E0", gray400: "#B0B0B0", gray600: "#6B7280",
  gray800: "#374151", black: "#1A1A1A",
  green: "#15803d", greenBg: "#f0fdf4", greenBorder: "#bbf7d0",
  yellow: "#a16207", yellowBg: "#fef9c3", yellowBorder: "#fde68a",
  red: "#dc2626", redBg: "#fee2e2", redBorder: "#fecaca",
};

const CATEGORIES = [
  { id: "basic", label: "企業基本情報", color: C.blue, items: [
    { id: "company_overview", label: "会社概要", desc: "設立・本社・従業員数・上場区分" },
    { id: "business_products", label: "事業領域・主力商材", desc: "取扱品・サービスラインナップ" },
    { id: "group_structure", label: "グループ構造", desc: "親会社・子会社・資本関係" },
    { id: "bases_network", label: "拠点・物流ネットワーク", desc: "拠点数・倉庫所在地・配送エリア" },
  ]},
  { id: "market", label: "市場・競合", color: "#0f766e", items: [
    { id: "market_share", label: "業界シェア・市場ポジション", desc: "業界内順位・シェア推移" },
  ]},
  { id: "finance", label: "財務", color: "#b45309", items: [
    { id: "pl_summary", label: "売上・利益サマリー", desc: "直近3期の売上・営業利益・決算期" },
    { id: "growth_profit", label: "成長性・収益性評価", desc: "CAGR・利益率トレンド・評価" },
    { id: "financial_health", label: "財務健全性", desc: "自己資本比率・有利子負債・物流コスト比率" },
  ]},
  { id: "strategy", label: "戦略・方針", color: "#7c3aed", items: [
    { id: "mid_term_plan", label: "中期経営計画", desc: "KPI・重点施策・投資計画" },
    { id: "dx_strategy", label: "DX・自動化戦略", desc: "IR・ニュースからの引用ベース" },
    { id: "sustainability", label: "サステナビリティ方針", desc: "IR・ニュースからの引用ベース" },
  ]},
  { id: "logistics", label: "物流特性", color: "#be185d", items: [
    { id: "logistics_flow", label: "物流フロー概要", desc: "入荷〜保管〜出荷の基本フロー" },
    { id: "product_features", label: "取扱品の特性", desc: "温度帯・サイズ・危険物・重量" },
    { id: "existing_systems", label: "既存システム環境", desc: "WMS・TMS・ERP・連携状況" },
  ]},
  { id: "contact", label: "担当者・組織", color: "#0369a1", items: [
    { id: "org_structure", label: "部署・組織構成", desc: "物流系部署の階層・組織図" },
    { id: "recent_news", label: "直近のニュース・トピック", desc: "IR・プレスリリース・受賞歴" },
  ]},
];

const ALL_ITEMS = CATEGORIES.flatMap(c => c.items.map(i => ({ ...i, categoryColor: c.color, categoryLabel: c.label })));

const REPORT_ORDER = [
  "company_overview", "business_products",
  "pl_summary", "growth_profit", "financial_health",
  "mid_term_plan", "dx_strategy", "sustainability",
  "logistics_flow", "product_features", "existing_systems",
  "org_structure", "market_share",
  "group_structure", "bases_network", "recent_news",
];

const PRESETS = [
  { label: "初回訪問前", ids: new Set(["company_overview","business_products","pl_summary","mid_term_plan","logistics_flow","recent_news"]) },
  { label: "財務・IR重点", ids: new Set(["pl_summary","growth_profit","financial_health","mid_term_plan"]) },
  { label: "物流特性重点", ids: new Set(["logistics_flow","product_features","existing_systems","org_structure"]) },
];

const PPT_STYLES = [
  { id: "visual", label: "ビジュアル中心", desc: "カード・チャート・数値を大きく表示", icon: "🎨" },
  { id: "text", label: "テキスト中心", desc: "文字情報重視・シンプル・ソース明記", icon: "📝" },
];

const INDUSTRY_OPTIONS = ["食品・飲料・冷凍食品","小売・EC・通販","医薬品・医療機器","化学・素材・危険物","自動車・輸送機器","アパレル・雑貨","電子部品・精密機器","その他"];

const STATUS = {
  confirmed: { label: "確認済み", bg: C.greenBg, color: C.green, border: C.greenBorder },
  partial: { label: "一部のみ", bg: C.yellowBg, color: C.yellow, border: C.yellowBorder },
  unconfirmed: { label: "確認できず", bg: C.redBg, color: C.red, border: C.redBorder },
};

const TOOLTIPS = {
  basicInfo: "企業名はPPT表紙・全スライドのタイトルに使用されます。担当者部署はフッターに表示されます。",
  industry: "業界を指定すると業界シェア調査の精度が上がります。自動判定はAIが企業から業界を推定します。",
  priorInfo: "商談メモや名刺情報を貼り付けると、AIがリサーチ結果に自動反映します。",
  pptStyle: "ビジュアル中心はカラーブロック・数値強調。テキスト中心はシンプル・ソース明記。",
  items: "チェックした項目のみ調査します（1項目約1分）。1項目＝1スライドで出力されます。",
};

const FACT_ONLY = "公開情報・報道・公式サイトで確認できた事実のみ記載。推測・類推・一般論は禁止。不明な場合はstatusをunconfirmedにして空白にする。";

// トークン数を多く必要とする項目
const HIGH_TOKEN_ITEMS = new Set(["recent_news", "mid_term_plan", "group_structure", "sustainability", "dx_strategy"]);

function safeStr(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number") return String(v);
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

// 日本語金額文字列を億円単位の数値に変換
function parseJPY(str) {
  if (!str) return 0;
  const s = safeStr(str).replace(/,/g, "");
  let val = 0;
  const choMatch = s.match(/([\d.]+)\s*兆/);
  const okuMatch = s.match(/([\d.]+)\s*億/);
  const hyakuMatch = s.match(/([\d.]+)\s*百億/);
  if (choMatch) val += parseFloat(choMatch[1]) * 10000;
  if (hyakuMatch) val += parseFloat(hyakuMatch[1]) * 100;
  else if (okuMatch) val += parseFloat(okuMatch[1]);
  if (val === 0) {
    const plain = parseFloat(s.replace(/[^\d.]/g, ""));
    if (!isNaN(plain)) val = plain;
  }
  return Math.round(val);
}

function guessDomain(company) {
  const c = company.replace(/株式会社|有限会社|合同会社|ホールディングス|グループ|HD|Holdings/gi, "").replace(/[　\s]/g, "").toLowerCase();
  const map = { "subaru": "subaru.co.jp", "スバル": "subaru.co.jp", "toyota": "toyota.co.jp", "トヨタ": "toyota.co.jp", "honda": "honda.co.jp", "ホンダ": "honda.co.jp", "nissan": "nissan.co.jp", "日産": "nissan.co.jp", "nichirei": "nichirei.co.jp", "ニチレイ": "nichirei.co.jp", "yamato": "yamato-hd.co.jp", "ヤマト": "yamato-hd.co.jp", "sagawa": "sagawa-exp.co.jp", "佐川": "sagawa-exp.co.jp", "nipponexpress": "nipponexpress.com", "日本通運": "nipponexpress.com", "hitachi": "hitachi.co.jp", "日立": "hitachi.co.jp", "fujitsu": "fujitsu.com", "富士通": "fujitsu.com", "isuzu": "isuzu.co.jp", "いすゞ": "isuzu.co.jp" };
  for (const [key, domain] of Object.entries(map)) { if (c.includes(key.toLowerCase())) return domain; }
  return null;
}

const URL_PATTERNS = {
  company_overview: (d) => [`https://${d}`, `https://${d}/company/`, `https://${d}/corporate/`],
  business_products: (d) => [`https://${d}/products/`, `https://${d}/services/`, `https://${d}/business/`],
  group_structure: (d) => [`https://${d}/corporate/group/`, `https://${d}/company/group/`],
  bases_network: (d) => [`https://${d}/corporate/network/`, `https://${d}/company/bases/`],
  market_share: (d) => [`https://${d}/ir/`, `https://${d}/investor/`],
  pl_summary: (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/result/`, `https://${d}/ir/`],
  growth_profit: (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/result/`],
  financial_health: (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/`],
  mid_term_plan: (d) => [`https://${d}/ir/strategy/`, `https://${d}/ir/management/`],
  dx_strategy: (d) => [`https://${d}/news/`, `https://${d}/ir/`, `https://${d}/corporate/dx/`],
  sustainability: (d) => [`https://${d}/sustainability/`, `https://${d}/csr/`, `https://${d}/ir/`],
  logistics_flow: (d) => [`https://${d}/corporate/logistics/`, `https://${d}/business/logistics/`],
  product_features: (d) => [`https://${d}/products/`, `https://${d}/business/products/`],
  existing_systems: (d) => [`https://${d}/corporate/dx/`, `https://${d}/ir/strategy/`],
  org_structure: (d) => [`https://${d}/corporate/organization/`, `https://${d}/company/organization/`],
  recent_news: (d) => [`https://${d}/news/`, `https://${d}/press/`, `https://${d}/ir/news/`],
};

const SEARCH_QUERIES = {
  company_overview: (c) => [`${c} 会社概要 設立 従業員数 本社`, `${c} 企業情報`],
  business_products: (c) => [`${c} 事業内容 主力商品 サービス`, `${c} 事業領域`],
  group_structure: (c) => [`${c} グループ会社 子会社 組織`, `${c} グループ構成`],
  bases_network: (c) => [`${c} 拠点 物流センター 倉庫 所在地`, `${c} 配送網`],
  market_share: (c, ind) => [`${c} ${ind || ""}業界シェア 市場ポジション`, `${c} 業界順位`],
  pl_summary: (c) => [`${c} 決算 売上高 営業利益 決算期 2024 2025`, `${c} 業績 財務ハイライト 物流コスト`],
  growth_profit: (c) => [`${c} 業績推移 成長率 収益性`, `${c} CAGR 利益率トレンド`],
  financial_health: (c) => [`${c} 自己資本比率 有利子負債 財務健全性`, `${c} 物流コスト比率 財務指標`],
  mid_term_plan: (c) => [`${c} 中期経営計画 2025 2026 KPI`, `${c} 経営戦略 重点施策`],
  dx_strategy: (c) => [`${c} DX デジタル化 自動化 IR ニュース 2024 2025`, `${c} DX戦略 IT投資 プレスリリース`],
  sustainability: (c) => [`${c} サステナビリティ ESG CO2 IR ニュース 2024 2025`, `${c} 環境 脱炭素 プレスリリース`],
  logistics_flow: (c) => [`${c} 物流フロー 入荷 出荷 保管`, `${c} サプライチェーン 物流`],
  product_features: (c) => [`${c} 取扱品 商品特性 温度管理`, `${c} 商品カテゴリ`],
  existing_systems: (c) => [`${c} WMS TMS ERP システム`, `${c} 基幹システム IT`],
  org_structure: (c) => [`${c} 組織図 物流部門 SCM`, `${c} 部署構成`],
  recent_news: (c) => [`${c} プレスリリース ニュース 2025`, `${c} IR 新着 2024 2025`],
};

const RETRY_QUERIES = {
  company_overview: (c, h) => [`${c} 会社沿革 設立経緯 ${h || ""}`, `${c} annual report 企業概要`],
  business_products: (c, h) => [`${c} 製品 サービス 詳細 ${h || ""}`, `${c} 事業ポートフォリオ`],
  group_structure: (c, h) => [`${c} 有価証券報告書 グループ ${h || ""}`, `${c} 関係会社 出資`],
  bases_network: (c, h) => [`${c} 物流拠点 センター 一覧 ${h || ""}`, `${c} warehouse distribution center`],
  market_share: (c, h) => [`${c} 市場シェア 業界レポート ${h || ""}`, `${c} 競合比較 ポジショニング`],
  pl_summary: (c, h) => [`${c} 有価証券報告書 財務諸表 ${h || ""}`, `${c} 決算短信 業績`],
  growth_profit: (c, h) => [`${c} 過去5年 業績 推移 ${h || ""}`, `${c} 収益性 ROE ROA`],
  financial_health: (c, h) => [`${c} 貸借対照表 自己資本 ${h || ""}`, `${c} 財務健全性 格付け`],
  mid_term_plan: (c, h) => [`${c} 長期ビジョン 経営計画 ${h || ""}`, `${c} 投資計画 事業方針`],
  dx_strategy: (c, h) => [`${c} DX 取り組み 事例 ${h || ""}`, `${c} デジタル化 自動化 発表`],
  sustainability: (c, h) => [`${c} ESG 環境目標 ${h || ""}`, `${c} カーボンニュートラル 脱炭素`],
  logistics_flow: (c, h) => [`${c} 物流 倉庫管理 ${h || ""}`, `${c} 配送 在庫管理 フロー`],
  product_features: (c, h) => [`${c} 取扱商品 仕様 ${h || ""}`, `${c} 商品 温度帯 危険物`],
  existing_systems: (c, h) => [`${c} システム 導入 IT ${h || ""}`, `${c} WMS TMS 物流システム`],
  org_structure: (c, h) => [`${c} 組織 部門 体制 ${h || ""}`, `${c} 担当部署 物流管理`],
  recent_news: (c, h) => [`${c} 最新 ニュース トピック ${h || ""}`, `${c} 2025 発表 動向`],
};

const PROMPTS = {
  company_overview: (c) => `"${c}"の会社概要（設立・本社・従業員数・上場区分）。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  business_products: (c) => `"${c}"の事業領域・主力商材。事業別売上構成比（%）が分かる場合は記載。${FACT_ONLY} JSONのみ: {"summary":"3文以内","segments":[{"name":"事業名","ratio":"xx%"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  group_structure: (c) => `"${c}"のグループ構造。${FACT_ONLY} JSONのみ: {"summary":"3文以内","subsidiaries":[{"name":"子会社名","role":"役割"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  bases_network: (c) => `"${c}"の拠点・物流ネットワーク。主要拠点名と所在地を列挙。${FACT_ONLY} JSONのみ: {"summary":"3文以内","bases":[{"name":"拠点名","location":"都道府県","type":"工場|センター|本社|営業所"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  market_share: (c, ind) => `"${c}"の${ind ? `${ind}業界における` : ""}業界シェア・市場ポジション。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  pl_summary: (c) => `"${c}"の財務分析。①直近3期の売上高・営業利益・営業利益率 ②決算期（何月期か） ③物流コスト比率（判明する場合のみ）。結論を先に述べ、根拠となる数値を明記すること。${FACT_ONLY} JSONのみ: {"summary":"結論1文","detail":"根拠数値を含む詳細3文以内","revenue":"最新売上高（億円単位の数値のみ、例：47283）","growth":"成長率","margin":"営業利益率","fiscal_month":"決算月（例：3月期）","logistics_cost_ratio":"物流コスト比率（不明な場合は空文字）","yearly_data":[{"year":"2023/3","revenue_str":"3兆2081億円","revenue":32081,"profit_str":"2291億円","profit":2291,"margin":"7.1%"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  growth_profit: (c) => `"${c}"の成長性・収益性を分析。結論を先に述べ、根拠となる数値・トレンドを明記すること。${FACT_ONLY} JSONのみ: {"summary":"結論1文","detail":"根拠数値を含む詳細3文以内","conclusions":["結論1","結論2","結論3"],"margin_trend":[{"year":"2023/3","margin":6.5},{"year":"2024/3","margin":7.5},{"year":"2025/3","margin":8.7}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  financial_health: (c) => `"${c}"の財務健全性を分析。①自己資本比率 ②有利子負債 ③物流コスト比率（判明する場合）を調査。結論を先に、根拠数値を明記すること。${FACT_ONLY} JSONのみ: {"summary":"結論1文","detail":"根拠数値を含む詳細3文以内","equity_ratio":"自己資本比率（数値のみ、例：41.0）","roe":"ROE（数値のみ、例：13.4）","conclusions":["結論1","結論2","結論3"],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  mid_term_plan: (c) => `"${c}"の中期経営計画（KPI・重点施策）。計画名・期間・主要KPIを明記。${FACT_ONLY} JSONのみ: {"summary":"3文以内","plan_name":"計画名","period":"期間","kpis":[{"label":"KPI名","value":"目標値"}],"phases":[{"period":"期間","content":"施策"}],"quote":"IR引用文（あれば）","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  dx_strategy: (c) => `"${c}"のDX・自動化戦略について、IR・ニュース・プレスリリースから実際のコメントや発表内容を引用して調査してください。推測は禁止。${FACT_ONLY} JSONのみ: {"summary":"引用ベースの事実3文以内","quote":"引用文（あれば）","initiatives":[{"label":"施策名","status":"進行中|予定|完了"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  sustainability: (c) => `"${c}"のサステナビリティ・ESG方針について、IR・ニュース・プレスリリースから実際のコメントや発表内容を引用して調査してください。推測は禁止。${FACT_ONLY} JSONのみ: {"summary":"引用ベースの事実3文以内","quote":"引用文（あれば）","targets":[{"label":"目標","value":"数値・期限"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  logistics_flow: (c) => `"${c}"の物流フロー概要。入荷〜保管〜出荷の流れを具体的に。${FACT_ONLY} JSONのみ: {"summary":"3文以内","steps":[{"label":"ステップ名","detail":"説明"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  product_features: (c) => `"${c}"が扱う商品の特性。${FACT_ONLY} JSONのみ: {"summary":"3文以内","features":[{"label":"特性名","value":"内容"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  existing_systems: (c) => `"${c}"が使用するWMS・TMS・ERP等の既存システム。${FACT_ONLY} JSONのみ: {"summary":"3文以内","systems":[{"name":"システム名","type":"種別","status":"稼働中|移行中|予定"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  org_structure: (c) => `"${c}"の物流・DX系部署の組織構成。${FACT_ONLY} JSONのみ: {"summary":"3文以内","departments":[{"name":"部署名","role":"役割"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  recent_news: (c) => `"${c}"の直近1年のニュース・IR・プレスリリース。日付付きで列挙。${FACT_ONLY} JSONのみ: {"summary":"3文以内","news":[{"date":"yyyy/mm","title":"タイトル","detail":"詳細","url":"記事URL（あれば）"}],"status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
};

// ★修正①：maxTokens引数を追加
async function callClaudeNoSearch(system, user, maxTokens = 1500) {
  const res = await fetch("/api/chat", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ model: "claude-haiku-4-5-20251001", max_tokens: maxTokens, system, messages: [{ role: "user", content: user }] }) });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.map(b => b.text || "").filter(Boolean).join("\n") || "";
}

async function callClaude(system, user) {
  const res = await fetch("/api/chat", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ model: "claude-haiku-4-5-20251001", max_tokens: 1500, system, messages: [{ role: "user", content: user }], tools: [{ type: "web_search_20250305", name: "web_search" }] }) });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.map(b => b.text || "").filter(Boolean).join("\n") || "";
}

async function fetchUrl(url) {
  try {
    const res = await fetch("/api/fetch-url", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ url }) });
    const data = await res.json();
    return { text: data.text || "", url };
  } catch (_) { return { text: "", url }; }
}

function extractJSON(text) {
  try {
    const clean = text.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
    const s = clean.indexOf("{"), e = clean.lastIndexOf("}");
    if (s !== -1 && e !== -1) return JSON.parse(clean.slice(s, e + 1));
  } catch (_) {}
  return null;
}

// ★修正②：itemIdを引数に追加し、HIGH_TOKEN_ITEMSで分岐
async function deepResearchPremium(company, itemId, prompt, industry) {
  const collectedTexts = [], collectedUrls = [];
  const domain = guessDomain(company);
  if (domain) {
    const urlPatterns = URL_PATTERNS[itemId]?.(domain) || [`https://${domain}`];
    for (const url of urlPatterns.slice(0, 3)) {
      const { text } = await fetchUrl(url);
      if (text.length > 300) { collectedTexts.push(`【${url}】\n${text.slice(0, 3000)}`); collectedUrls.push(url); }
      await new Promise(r => setTimeout(r, 500));
    }
  }
  await new Promise(r => setTimeout(r, 3000));
  const queries = SEARCH_QUERIES[itemId]?.(company, industry) || [`${company} ${itemId}`];
  try {
    const s1 = await callClaude(`あなたは企業調査AIです。検索して最も信頼性の高い公式情報を取得してください。`, `「${queries[0]}」で検索し、最も信頼性の高い公式情報を取得してください。\nURL: [参照したURL]\n内容: [取得した情報の要約（500文字以内）]`);
    if (s1.length > 50) {
      collectedTexts.push(`【Web検索1回目】\n${s1.slice(0, 2000)}`);
      const m = s1.match(/URL:\s*(https?:\/\/[^\s\n]+)/);
      if (m) { collectedUrls.push(m[1]); const { text: dt } = await fetchUrl(m[1]); if (dt.length > 300) collectedTexts.push(`【${m[1]} 詳細】\n${dt.slice(0, 3000)}`); }
    }
  } catch (_) {}
  await new Promise(r => setTimeout(r, 5000));
  if (queries.length > 1) {
    try {
      const s2 = await callClaude(`あなたは企業調査AIです。検索して補完情報・最新データを取得してください。`, `「${queries[1]}」で検索し、最新の情報・数値データを取得してください。\nURL: [参照したURL]\n内容: [取得した情報の要約（500文字以内）]`);
      if (s2.length > 50) { collectedTexts.push(`【Web検索2回目】\n${s2.slice(0, 2000)}`); const m2 = s2.match(/URL:\s*(https?:\/\/[^\s\n]+)/); if (m2) collectedUrls.push(m2[1]); }
    } catch (_) {}
  }
  await new Promise(r => setTimeout(r, 2000));
  const allInfo = collectedTexts.join("\n\n---\n\n");
  const sourceUrls = [...new Set(collectedUrls)].slice(0, 3).join(", ");
  const finalPrompt = allInfo.length > 200 ? `${prompt}\n\n以下は複数の公式情報源から収集した情報です。source_urlには「${sourceUrls}」を記載してください：\n\n${allInfo.slice(0, 8000)}` : `${prompt}\n\nsource_urlには参照したURLを記載してください。情報が見つからない場合はstatusをunconfirmedにしてください。`;
  const maxTokens = HIGH_TOKEN_ITEMS.has(itemId) ? 3000 : 1500;
  return await callClaudeNoSearch(`あなたはB2B営業支援AIです。収集した複数の公式情報源を総合分析し、JSONのみ返答。【厳守ルール】- 収集した公式情報に記載された事実のみ記載する - 推測・類推・一般論は一切書かない - 財務項目は結論を先に述べ根拠数値を明記する - 数値データは必ず数値型で返す（文字列不可）- 情報が見つからない場合はstatusをunconfirmedにし「公開情報なし」と記載 - missingには「訪問時に直接確認すべき具体的な質問」を書く - source_urlには実際に参照したURLを記載する - 前置き・マークダウン不要。JSONのみ返答。`, finalPrompt, maxTokens);
}

// ★修正③：retryResearchも同様にmaxTokens分岐
async function retryResearch(company, itemId, prompt, industry, hint = "") {
  const collectedTexts = [], collectedUrls = [];
  const queries = RETRY_QUERIES[itemId]?.(company, hint) || [`${company} ${itemId} ${hint}`];
  await new Promise(r => setTimeout(r, 2000));
  try {
    const s1 = await callClaude(`あなたは企業調査AIです。前回の調査で情報が不十分だった項目を別のアプローチで再調査してください。`, `「${queries[0]}」で再調査してください。前回と異なるソースを優先してください。\nURL: [参照したURL]\n内容: [取得した情報の要約（500文字以内）]`);
    if (s1.length > 50) {
      collectedTexts.push(`【再調査1回目】\n${s1.slice(0, 2000)}`);
      const m = s1.match(/URL:\s*(https?:\/\/[^\s\n]+)/);
      if (m) { collectedUrls.push(m[1]); const { text: dt } = await fetchUrl(m[1]); if (dt.length > 300) collectedTexts.push(`【${m[1]} 詳細】\n${dt.slice(0, 3000)}`); }
    }
  } catch (_) {}
  await new Promise(r => setTimeout(r, 5000));
  try {
    const s2 = await callClaude(`あなたは企業調査AIです。追加情報を収集してください。`, `「${queries[1] || queries[0]}」でさらに深掘り調査してください。${hint ? `特に「${hint}」に関する情報を重点的に。` : ""}\nURL: [参照したURL]\n内容: [取得した情報の要約（500文字以内）]`);
    if (s2.length > 50) { collectedTexts.push(`【再調査2回目】\n${s2.slice(0, 2000)}`); const m2 = s2.match(/URL:\s*(https?:\/\/[^\s\n]+)/); if (m2) collectedUrls.push(m2[1]); }
  } catch (_) {}
  await new Promise(r => setTimeout(r, 2000));
  const allInfo = collectedTexts.join("\n\n---\n\n");
  const sourceUrls = [...new Set(collectedUrls)].slice(0, 3).join(", ");
  const hintText = hint ? `\n\nユーザーからの追加指示：「${hint}」この点を特に重視してください。` : "";
  const finalPrompt = allInfo.length > 200 ? `${prompt}${hintText}\n\n以下は再調査で収集した情報です。source_urlには「${sourceUrls}」を記載してください：\n\n${allInfo.slice(0, 8000)}` : `${prompt}${hintText}`;
  const maxTokens = HIGH_TOKEN_ITEMS.has(itemId) ? 3000 : 1500;
  return await callClaudeNoSearch(`あなたはB2B営業支援AIです。再調査で収集した情報を分析し、JSONのみ返答。【厳守ルール】- 収集した公式情報に記載された事実のみ記載する - 推測は禁止 - 財務項目は結論を先に根拠数値を明記 - 数値データは数値型で返す - source_urlには実際に参照したURLを記載 - 前置き・マークダウン不要。JSONのみ返答。`, finalPrompt, maxTokens);
}

// 1項目1スライドのHTML生成
function buildSlideHTML(company, dept, results, pptStyle, selectedIds) {
  const today = new Date().toLocaleDateString("ja-JP");
  const ss = safeStr;
  const SLIDE_COLORS = ["#0068B7","#BAD6F0","#185FA5","#E8F4FF","#00519A","#93c5fd"];

  // ドーナツチャートSVG生成
  function donutChart(segments, size = 180) {
    if (!segments || segments.length === 0) return "";
    const cx = size / 2, cy = size / 2, r = size * 0.38, ir = size * 0.22;
    let angle = -Math.PI / 2;
    const total = segments.reduce((s, seg) => s + (parseFloat(ss(seg.ratio)) || 0), 0);
    const paths = segments.map((seg, i) => {
      const pct = total > 0 ? (parseFloat(ss(seg.ratio)) || 0) / total : 1 / segments.length;
      const sweep = pct * 2 * Math.PI;
      const x1 = cx + r * Math.cos(angle), y1 = cy + r * Math.sin(angle);
      const x2 = cx + r * Math.cos(angle + sweep), y2 = cy + r * Math.sin(angle + sweep);
      const ix1 = cx + ir * Math.cos(angle), iy1 = cy + ir * Math.sin(angle);
      const ix2 = cx + ir * Math.cos(angle + sweep), iy2 = cy + ir * Math.sin(angle + sweep);
      const large = sweep > Math.PI ? 1 : 0;
      const path = `M${x1.toFixed(1)} ${y1.toFixed(1)} A${r} ${r} 0 ${large} 1 ${x2.toFixed(1)} ${y2.toFixed(1)} L${ix2.toFixed(1)} ${iy2.toFixed(1)} A${ir} ${ir} 0 ${large} 0 ${ix1.toFixed(1)} ${iy1.toFixed(1)}Z`;
      angle += sweep;
      return `<path d="${path}" fill="${SLIDE_COLORS[i % SLIDE_COLORS.length]}" stroke="#fff" stroke-width="2"/>`;
    }).join("");
    return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${paths}</svg>`;
  }

  // 棒グラフSVG生成（実データ使用）
  function barChart(yearlyData, width = 320, height = 160) {
    if (!yearlyData || yearlyData.length === 0) return `<div style="color:#6B7280;font-size:13px;padding:20px 0;">財務データ取得中</div>`;
    const maxRev = Math.max(...yearlyData.map(y => y.revenue || parseJPY(ss(y.revenue_str)) || 0), 1);
    const maxProf = Math.max(...yearlyData.map(y => y.profit || parseJPY(ss(y.profit_str)) || 0), 1);
    const barW = Math.floor((width - 60) / yearlyData.length / 2.5);
    const chartH = height - 30;
    const bars = yearlyData.map((y, i) => {
      const rev = y.revenue || parseJPY(ss(y.revenue_str)) || 0;
      const prof = y.profit || parseJPY(ss(y.profit_str)) || 0;
      const revH = Math.round((rev / maxRev) * chartH * 0.85);
      const profH = Math.round((prof / maxRev) * chartH * 0.85);
      const x = 40 + i * ((width - 60) / yearlyData.length);
      return `
        <rect x="${x}" y="${chartH - revH}" width="${barW}" height="${revH}" fill="#0068B7" rx="2"/>
        <rect x="${x + barW + 3}" y="${chartH - profH}" width="${barW}" height="${profH}" fill="#E8F4FF" stroke="#0068B7" stroke-width="1" rx="2"/>
        <text x="${x + barW}" y="${height - 5}" text-anchor="middle" font-size="10" fill="#6B7280">${ss(y.year)}</text>
        <text x="${x}" y="${chartH - revH - 4}" font-size="9" fill="#0068B7">${ss(y.revenue_str || "")}</text>
      `;
    }).join("");
    return `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <line x1="36" y1="0" x2="36" y2="${chartH}" stroke="#E0E0E0" stroke-width="1"/>
      <line x1="36" y1="${chartH}" x2="${width}" y2="${chartH}" stroke="#E0E0E0" stroke-width="1"/>
      ${bars}
      <rect x="${width - 100}" y="4" width="10" height="8" fill="#0068B7" rx="1"/>
      <text x="${width - 86}" y="12" font-size="9" fill="#6B7280">売上高</text>
      <rect x="${width - 100}" y="18" width="10" height="8" fill="#E8F4FF" stroke="#0068B7" stroke-width="1" rx="1"/>
      <text x="${width - 86}" y="26" font-size="9" fill="#6B7280">営業利益</text>
    </svg>`;
  }

  // 利益率推移グラフ
  function marginTrendChart(marginTrend, width = 300, height = 120) {
    if (!marginTrend || marginTrend.length === 0) return "";
    const maxM = Math.max(...marginTrend.map(m => parseFloat(m.margin) || 0), 1);
    const pts = marginTrend.map((m, i) => {
      const x = 30 + i * ((width - 60) / Math.max(marginTrend.length - 1, 1));
      const y = (height - 30) - Math.round((parseFloat(m.margin) || 0) / maxM * (height - 40));
      return { x, y, label: ss(m.year), val: m.margin };
    });
    const polyline = pts.map(p => `${p.x},${p.y}`).join(" ");
    return `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <polyline points="${polyline}" fill="none" stroke="#0068B7" stroke-width="2" stroke-linejoin="round"/>
      ${pts.map(p => `
        <circle cx="${p.x}" cy="${p.y}" r="4" fill="#0068B7"/>
        <text x="${p.x}" y="${p.y - 8}" text-anchor="middle" font-size="9" fill="#0068B7">${p.val}%</text>
        <text x="${p.x}" y="${height - 5}" text-anchor="middle" font-size="9" fill="#6B7280">${p.label}</text>
      `).join("")}
    </svg>`;
  }

  // 各スライドのHTMLを生成（1項目1スライド）
  function buildItemSlide(itemId, result, slideNum, totalSlides) {
    const item = ALL_ITEMS.find(i => i.id === itemId);
    const cat = CATEGORIES.find(c => c.items.some(i => i.id === itemId));
    const accentColor = cat?.color || C.blue;
    const ss2 = (v) => safeStr(v);
    const sourceUrls = result?.source_url && typeof result.source_url === "string"
      ? result.source_url.split(",").map(u => u.trim()).filter(u => u.startsWith("http"))
      : [];
    const stLabel = result?.status === "confirmed" ? "✓ 確認済み" : result?.status === "partial" ? "△ 一部のみ" : "⚠ 要確認";
    const stColor = result?.status === "confirmed" ? "#15803d" : result?.status === "partial" ? "#a16207" : "#dc2626";
    const stBg = result?.status === "confirmed" ? "#f0fdf4" : result?.status === "partial" ? "#fef9c3" : "#fee2e2";

    let body = "";

    // --- 項目ごとのレイアウト ---
    if (itemId === "company_overview") {
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col center">
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;width:100%;">
              ${["設立","本社所在地","従業員数","上場区分"].map(k => `
                <div style="background:#F5F7FA;border-radius:8px;padding:12px 14px;border-left:3px solid ${accentColor};">
                  <div style="font-size:11px;color:#6B7280;margin-bottom:4px;">${k}</div>
                  <div style="font-size:13px;font-weight:500;color:#1A1A1A;">—</div>
                </div>`).join("")}
            </div>
          </div>
        </div>`;

    } else if (itemId === "business_products") {
      const segments = Array.isArray(result?.segments) ? result.segments : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${segments.length > 0 ? `
              <div style="margin-top:16px;">
                <div class="section-label">事業セグメント</div>
                ${segments.map((s, i) => `
                  <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
                    <div style="width:12px;height:12px;background:${SLIDE_COLORS[i % SLIDE_COLORS.length]};border-radius:2px;flex-shrink:0;"></div>
                    <div style="flex:1;font-size:13px;color:#1A1A1A;">${ss2(s.name)}</div>
                    <div style="font-size:16px;font-weight:500;color:${accentColor};">${ss2(s.ratio)}</div>
                  </div>`).join("")}
              </div>` : ""}
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col center">
            ${segments.length > 0 ? `
              <div class="chart-label">事業別売上構成比</div>
              ${donutChart(segments, 200)}
              <div class="legend" style="margin-top:12px;">
                ${segments.map((s, i) => `<span class="legend-item"><span class="legend-dot" style="background:${SLIDE_COLORS[i % SLIDE_COLORS.length]}"></span>${ss2(s.name)}</span>`).join("")}
              </div>` : `<div style="color:#6B7280;font-size:14px;">事業構成比は非公開</div>`}
          </div>
        </div>`;

    } else if (itemId === "pl_summary") {
      const yearly = Array.isArray(result?.yearly_data) ? result.yearly_data : [];
      body = `
        <div style="display:flex;flex-direction:column;gap:16px;height:100%;">
          <div style="display:flex;flex-direction:column;gap:6px;">
            <div class="conclusion-item blue">▶ ${ss2(result?.summary) || "情報なし"}</div>
            ${result?.detail ? `<div style="font-size:13px;color:#374151;line-height:1.7;">${ss2(result.detail)}</div>` : ""}
          </div>
          <div style="display:flex;gap:12px;">
            <div class="kpi-card"><div class="kpi-label">売上高（最新期）</div><div class="kpi-num">${ss2(result?.revenue) || "—"}</div><div class="kpi-sub">${ss2(result?.growth) || ""}</div></div>
            <div class="kpi-card"><div class="kpi-label">営業利益率</div><div class="kpi-num">${ss2(result?.margin) || "—"}</div></div>
            <div class="kpi-card"><div class="kpi-label">決算期</div><div class="kpi-num" style="font-size:22px;">${ss2(result?.fiscal_month) || "—"}</div></div>
            <div class="kpi-card" style="border-top-color:${result?.logistics_cost_ratio ? accentColor : "#d97706"};"><div class="kpi-label">物流コスト比率</div><div class="kpi-num" style="font-size:${result?.logistics_cost_ratio ? "24px" : "16px"};color:${result?.logistics_cost_ratio ? accentColor : "#d97706"};">${ss2(result?.logistics_cost_ratio) || "要確認"}</div></div>
          </div>
          <div style="flex:1;">
            <div class="chart-label">売上高・営業利益の推移</div>
            ${barChart(yearly, 480, 180)}
          </div>
        </div>`;

    } else if (itemId === "growth_profit") {
      const conclusions = Array.isArray(result?.conclusions) ? result.conclusions : [];
      const marginTrend = Array.isArray(result?.margin_trend) ? result.margin_trend : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="section-label">分析結論</div>
            ${conclusions.length > 0
              ? conclusions.map(c => `<div class="conclusion-item blue" style="margin-bottom:8px;">▶ ${ss2(c)}</div>`).join("")
              : `<div class="conclusion-item blue">▶ ${ss2(result?.summary) || "情報なし"}</div>`}
            ${result?.detail ? `<div style="font-size:13px;color:#374151;line-height:1.7;margin-top:12px;">${ss2(result.detail)}</div>` : ""}
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:12px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col center">
            ${marginTrend.length > 0 ? `
              <div class="chart-label">営業利益率の推移</div>
              ${marginTrendChart(marginTrend, 320, 160)}` : ""}
          </div>
        </div>`;

    } else if (itemId === "financial_health") {
      const conclusions = Array.isArray(result?.conclusions) ? result.conclusions : [];
      const equityRatio = parseFloat(ss2(result?.equity_ratio)) || 0;
      const roe = parseFloat(ss2(result?.roe)) || 0;
      body = `
        <div class="two-col">
          <div class="col">
            <div class="section-label">分析結論</div>
            ${conclusions.length > 0
              ? conclusions.map(c => `<div class="conclusion-item blue" style="margin-bottom:8px;">▶ ${ss2(c)}</div>`).join("")
              : `<div class="conclusion-item blue">▶ ${ss2(result?.summary) || "情報なし"}</div>`}
            ${result?.detail ? `<div style="font-size:13px;color:#374151;line-height:1.7;margin-top:12px;">${ss2(result.detail)}</div>` : ""}
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:12px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            <div style="display:flex;flex-direction:column;gap:16px;">
              ${equityRatio > 0 ? `
                <div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                    <span style="font-size:12px;color:#6B7280;">自己資本比率</span>
                    <span style="font-size:20px;font-weight:500;color:${accentColor};">${equityRatio}%</span>
                  </div>
                  <div style="background:#E0E0E0;border-radius:99px;height:10px;overflow:hidden;">
                    <div style="height:100%;width:${Math.min(equityRatio, 100)}%;background:${accentColor};border-radius:99px;"></div>
                  </div>
                  <div style="font-size:11px;color:#6B7280;margin-top:4px;">業界平均目安：30〜40%</div>
                </div>` : ""}
              ${roe > 0 ? `
                <div>
                  <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
                    <span style="font-size:12px;color:#6B7280;">ROE</span>
                    <span style="font-size:20px;font-weight:500;color:${accentColor};">${roe}%</span>
                  </div>
                  <div style="background:#E0E0E0;border-radius:99px;height:10px;overflow:hidden;">
                    <div style="height:100%;width:${Math.min(roe * 3, 100)}%;background:${accentColor};border-radius:99px;"></div>
                  </div>
                </div>` : ""}
              <div class="kpi-card"><div class="kpi-label">自己資本比率</div><div class="kpi-num">${equityRatio > 0 ? equityRatio + "%" : ss2(result?.equity_ratio) || "—"}</div></div>
              <div class="kpi-card"><div class="kpi-label">ROE</div><div class="kpi-num">${roe > 0 ? roe + "%" : ss2(result?.roe) || "—"}</div></div>
            </div>
          </div>
        </div>`;

    } else if (itemId === "mid_term_plan") {
      const kpis = Array.isArray(result?.kpis) ? result.kpis : [];
      const phases = Array.isArray(result?.phases) ? result.phases : [];
      body = `
        <div class="two-col">
          <div class="col">
            ${result?.plan_name ? `<div class="plan-badge">${ss2(result.plan_name)}　${ss2(result?.period)}</div>` : ""}
            ${result?.quote ? `<div class="quote-box blue">"${ss2(result.quote)}"</div>` : ""}
            <div style="font-size:13px;color:#374151;line-height:1.7;">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:12px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${kpis.length > 0 ? `
              <div class="section-label">重点KPI</div>
              ${kpis.map(k => `
                <div class="kpi-mini"><span class="kpi-mini-label">${ss2(k.label)}</span><span class="kpi-mini-val">${ss2(k.value)}</span></div>`).join("")}` : ""}
            ${phases.length > 0 ? `
              <div class="section-label" style="margin-top:12px;">フェーズ別施策</div>
              ${phases.map((p, i) => `
                <div class="phase-item">
                  <div class="phase-dot" style="background:${i === 0 ? accentColor : "#BAD6F0"};"></div>
                  <div><div class="phase-period">${ss2(p.period)}</div><div class="phase-content">${ss2(p.content)}</div></div>
                </div>`).join("")}` : ""}
          </div>
        </div>`;

    } else if (itemId === "dx_strategy" || itemId === "sustainability") {
      const initiatives = Array.isArray(result?.initiatives) ? result.initiatives : [];
      const targets = Array.isArray(result?.targets) ? result.targets : [];
      const items2 = itemId === "dx_strategy" ? initiatives : targets;
      const quoteColor = itemId === "sustainability" ? "#15803d" : accentColor;
      const quoteBg = itemId === "sustainability" ? "#f0fdf4" : "#E8F4FF";
      body = `
        <div class="two-col">
          <div class="col">
            ${result?.quote ? `
              <div style="background:${quoteBg};border-left:4px solid ${quoteColor};padding:14px 16px;border-radius:0;margin-bottom:16px;font-size:14px;font-style:italic;color:${quoteColor};line-height:1.7;">
                "${ss2(result.quote)}"
              </div>` : ""}
            <div style="font-size:13px;color:#374151;line-height:1.7;">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:12px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${items2.length > 0 ? `
              <div class="section-label">${itemId === "dx_strategy" ? "主要施策" : "目標・コミットメント"}</div>
              ${items2.map(item2 => `
                <div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px;padding:10px 12px;background:#F5F7FA;border-radius:6px;">
                  <div style="width:8px;height:8px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:4px;"></div>
                  <div style="flex:1;">
                    <span style="font-size:13px;color:#1A1A1A;">${ss2(item2.label || item2.name || "")}</span>
                    ${item2.status ? `<span style="font-size:11px;color:${accentColor};margin-left:8px;background:#E8F4FF;padding:2px 6px;border-radius:3px;">${ss2(item2.status)}</span>` : ""}
                    ${item2.value ? `<div style="font-size:14px;font-weight:500;color:${accentColor};margin-top:2px;">${ss2(item2.value)}</div>` : ""}
                  </div>
                </div>`).join("")}` : ""}
          </div>
        </div>`;

    } else if (itemId === "logistics_flow") {
      const steps = Array.isArray(result?.steps) ? result.steps : [];
      body = `
        <div style="display:flex;flex-direction:column;gap:20px;height:100%;">
          ${steps.length > 0 ? `
            <div class="flow-row">
              ${steps.map((s, i) => `
                <div class="flow-step ${i === 0 ? "active" : ""}">
                  <div style="font-size:22px;margin-bottom:6px;">${["📦","🏢","🔍","🚚","🏪"][i] || "→"}</div>
                  <div class="flow-step-label">${ss2(s.label)}</div>
                  <div class="flow-step-detail">${ss2(s.detail)}</div>
                </div>
                ${i < steps.length - 1 ? `<div class="flow-arrow">→</div>` : ""}
              `).join("")}
            </div>` : `
            <div class="flow-row">
              <div class="flow-step active"><div style="font-size:22px;margin-bottom:6px;">📦</div><div class="flow-step-label">入荷・検品</div></div>
              <div class="flow-arrow">→</div>
              <div class="flow-step"><div style="font-size:22px;margin-bottom:6px;">🏢</div><div class="flow-step-label">保管・仕分け</div></div>
              <div class="flow-arrow">→</div>
              <div class="flow-step"><div style="font-size:22px;margin-bottom:6px;">🚚</div><div class="flow-step-label">出荷・配送</div></div>
            </div>`}
          <div style="flex:1;">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:12px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
        </div>`;

    } else if (itemId === "existing_systems") {
      const systems = Array.isArray(result?.systems) ? result.systems : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:16px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${systems.length > 0 ? `
              <div class="section-label">システム一覧</div>
              ${systems.map(s => `
                <div class="system-item ${s.status === "移行中" ? "active" : ""}">
                  <div><div class="system-name">${ss2(s.name)}</div><div class="system-type">${ss2(s.type)}</div></div>
                  <div class="system-status">${ss2(s.status)}</div>
                </div>`).join("")}` : ""}
          </div>
        </div>`;

    } else if (itemId === "recent_news") {
      const news = Array.isArray(result?.news) ? result.news : [];
      body = `
        <div style="display:flex;flex-direction:column;gap:12px;height:100%;">
          <div style="font-size:13px;color:#374151;line-height:1.7;">${ss2(result?.summary) || ""}</div>
          <div style="flex:1;display:flex;flex-direction:column;gap:8px;">
            ${news.slice(0, 5).map(n => `
              <div class="news-item">
                <div class="news-date">${ss2(n.date)}</div>
                <div class="news-content">
                  <div class="news-title">${ss2(n.title)}</div>
                  ${n.detail ? `<div style="font-size:12px;color:#6B7280;margin-top:2px;">${ss2(n.detail)}</div>` : ""}
                </div>
                ${n.url ? `<a href="${ss2(n.url)}" class="news-link" target="_blank">📎</a>` : ""}
              </div>`).join("")}
          </div>
        </div>`;

    } else if (itemId === "group_structure") {
      const subsidiaries = Array.isArray(result?.subsidiaries) ? result.subsidiaries : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:16px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${subsidiaries.length > 0 ? `
              <div class="section-label">主要グループ会社</div>
              ${subsidiaries.slice(0, 6).map(s => `
                <div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;padding:8px 12px;background:#F5F7FA;border-radius:6px;">
                  <div style="width:6px;height:6px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:5px;"></div>
                  <div><div style="font-size:13px;font-weight:500;color:#1A1A1A;">${ss2(s.name)}</div><div style="font-size:11px;color:#6B7280;">${ss2(s.role)}</div></div>
                </div>`).join("")}` : ""}
          </div>
        </div>`;

    } else if (itemId === "bases_network") {
      const bases = Array.isArray(result?.bases) ? result.bases : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:16px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${bases.length > 0 ? `
              <div class="section-label">主要拠点一覧</div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">
                ${bases.slice(0, 8).map(b => `
                  <div style="background:#F5F7FA;border-radius:6px;padding:8px 10px;">
                    <div style="font-size:12px;font-weight:500;color:#1A1A1A;">${ss2(b.name)}</div>
                    <div style="font-size:11px;color:#6B7280;">${ss2(b.location)} ${b.type ? `・${ss2(b.type)}` : ""}</div>
                  </div>`).join("")}
              </div>` : ""}
          </div>
        </div>`;

    } else if (itemId === "org_structure") {
      const departments = Array.isArray(result?.departments) ? result.departments : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:16px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${departments.length > 0 ? `
              <div class="section-label">物流・DX系部署</div>
              ${departments.slice(0, 6).map(d => `
                <div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;padding:8px 12px;background:#F5F7FA;border-radius:6px;">
                  <div style="width:6px;height:6px;border-radius:50%;background:${accentColor};flex-shrink:0;margin-top:5px;"></div>
                  <div><div style="font-size:13px;font-weight:500;color:#1A1A1A;">${ss2(d.name)}</div><div style="font-size:11px;color:#6B7280;">${ss2(d.role)}</div></div>
                </div>`).join("")}` : ""}
          </div>
        </div>`;

    } else if (itemId === "market_share") {
      body = `
        <div style="display:flex;flex-direction:column;gap:16px;height:100%;">
          <div class="conclusion-item blue">▶ ${ss2(result?.summary) || "情報なし"}</div>
          ${result?.detail ? `<div style="font-size:13px;color:#374151;line-height:1.7;">${ss2(result.detail)}</div>` : ""}
          ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box">△ ${ss2(result?.missing)}</div>` : ""}
        </div>`;

    } else if (itemId === "product_features") {
      const features = Array.isArray(result?.features) ? result.features : [];
      body = `
        <div class="two-col">
          <div class="col">
            <div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
            ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:16px;">△ ${ss2(result?.missing)}</div>` : ""}
          </div>
          <div class="col">
            ${features.length > 0 ? `
              <div class="section-label">商品特性</div>
              ${features.slice(0, 6).map(f => `
                <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:#F5F7FA;border-radius:6px;margin-bottom:6px;">
                  <span style="font-size:12px;color:#6B7280;">${ss2(f.label)}</span>
                  <span style="font-size:13px;font-weight:500;color:#1A1A1A;">${ss2(f.value)}</span>
                </div>`).join("")}` : ""}
          </div>
        </div>`;

    } else {
      // デフォルト
      body = `<div class="big-text">${ss2(result?.summary) || "情報なし"}</div>
        ${result?.missing && result?.status !== "confirmed" ? `<div class="warn-box" style="margin-top:16px;">△ ${ss2(result?.missing)}</div>` : ""}`;
    }

    return `
      <div class="slide content">
        <div class="slide-header" style="border-bottom-color:${accentColor};">
          <div class="header-accent" style="background:${accentColor};"></div>
          <div class="header-text">
            <div class="header-sub">${cat?.label || ""}</div>
            <div class="header-title">${item?.label || itemId}</div>
          </div>
          <div style="display:flex;align-items:center;gap:12px;margin-left:auto;">
            <div style="font-size:11px;font-weight:500;color:${stColor};background:${stBg};padding:3px 10px;border-radius:99px;">${stLabel}</div>
            <div class="header-company">${company}</div>
          </div>
        </div>
        <div class="slide-body">${body}</div>
        <div class="slide-footer">
          <span>${today}${dept ? " | " + dept : ""}${sourceUrls.length > 0 ? " | 📎 " + sourceUrls.map(u => { try { return new URL(u).hostname; } catch { return u; } }).join(", ") : ""}</span>
          <span>${slideNum} / ${totalSlides}</span>
        </div>
      </div>`;
  }

  // 表示順に並べてスライド生成
  const activeIds = REPORT_ORDER.filter(id => selectedIds.has(id) && results[id]);
  const totalSlides = 1 + activeIds.length;

  const slides = [`
    <div class="slide cover">
      <div class="cover-bar"></div>
      <div class="cover-body">
        <div class="cover-sub">顧客リサーチレポート　|　SoftBank Robotics Corp.</div>
        <div class="cover-company">${company}</div>
        <div class="cover-divider"></div>
        ${dept ? `<div class="cover-dept">担当部署：${dept}</div>` : ""}
        <div class="cover-date">調査日：${today}</div>
        <div class="cover-style">${PPT_STYLES.find(s => s.id === pptStyle)?.label || "ビジュアル中心"}　|　全${totalSlides}枚</div>
      </div>
    </div>
  `];

  activeIds.forEach((id, i) => {
    slides.push(buildItemSlide(id, results[id], i + 2, totalSlides));
  });

  const css = `
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Helvetica Neue', 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif; }
    body { background: #E8E8E8; }
    .slide-wrapper { width: 960px; margin: 0 auto; }
    .slide { width: 960px; min-height: 540px; background: #fff; position: relative; display: flex; flex-direction: column; margin-bottom: 2px; }
    .cover { background: #F5F7FA; }
    .cover-bar { position: absolute; left: 0; top: 0; width: 8px; height: 100%; background: #0068B7; }
    .cover-body { padding: 60px 60px 60px 80px; display: flex; flex-direction: column; justify-content: center; min-height: 540px; }
    .cover-sub { font-size: 14px; color: #6B7280; margin-bottom: 12px; }
    .cover-company { font-size: 52px; font-weight: 500; color: #1A1A1A; line-height: 1.1; margin-bottom: 16px; }
    .cover-divider { width: 60px; height: 3px; background: #0068B7; margin-bottom: 20px; }
    .cover-dept { font-size: 16px; color: #6B7280; margin-bottom: 8px; }
    .cover-date { font-size: 14px; color: #B0B0B0; margin-bottom: 6px; }
    .cover-style { font-size: 13px; color: #B0B0B0; }
    .slide-header { background: #fff; border-bottom: 3px solid #0068B7; padding: 14px 24px; display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
    .header-accent { width: 5px; height: 32px; border-radius: 0; flex-shrink: 0; }
    .header-sub { font-size: 11px; color: #6B7280; }
    .header-title { font-size: 22px; font-weight: 500; color: #1A1A1A; }
    .header-company { font-size: 11px; color: #B0B0B0; }
    .slide-body { flex: 1; padding: 24px 28px; overflow: hidden; }
    .slide-footer { height: 30px; background: #F5F7FA; border-top: 1px solid #E0E0E0; display: flex; align-items: center; justify-content: space-between; padding: 0 24px; flex-shrink: 0; font-size: 11px; color: #B0B0B0; }
    .two-col { display: flex; gap: 24px; height: 100%; }
    .col { flex: 1; display: flex; flex-direction: column; gap: 10px; min-width: 0; }
    .col.center { align-items: center; justify-content: center; }
    .section-label { font-size: 12px; font-weight: 500; color: #0068B7; border-bottom: 1px solid #E8F4FF; padding-bottom: 6px; margin-bottom: 8px; }
    .big-text { font-size: 15px; color: #374151; line-height: 1.85; }
    .warn-box { background: #fef9c3; border-left: 3px solid #d97706; padding: 8px 12px; font-size: 13px; color: #a16207; border-radius: 0; }
    .err-box { background: #fee2e2; border-left: 3px solid #dc2626; padding: 8px 12px; font-size: 13px; color: #dc2626; }
    .chart-label { font-size: 12px; color: #6B7280; margin-bottom: 8px; font-weight: 500; }
    .legend { display: flex; flex-wrap: wrap; gap: 10px; font-size: 12px; color: #374151; }
    .legend-item { display: flex; align-items: center; gap: 5px; }
    .legend-dot { width: 10px; height: 10px; border-radius: 2px; display: inline-block; }
    .kpi-card { background: #F5F7FA; border-top: 2px solid #0068B7; border-radius: 4px; padding: 12px 14px; flex: 1; }
    .kpi-label { font-size: 11px; color: #6B7280; margin-bottom: 6px; }
    .kpi-num { font-size: 26px; font-weight: 500; color: #0068B7; line-height: 1.1; }
    .kpi-sub { font-size: 11px; color: #B0B0B0; margin-top: 4px; }
    .conclusion-item { padding: 10px 16px; font-size: 13px; font-weight: 500; }
    .conclusion-item.blue { background: #E8F4FF; border-left: 4px solid #0068B7; color: #0C447C; }
    .plan-badge { background: #E8F4FF; border: 1px solid #BAD6F0; border-radius: 4px; padding: 8px 14px; font-size: 14px; font-weight: 500; color: #0068B7; margin-bottom: 12px; }
    .kpi-mini { display: flex; justify-content: space-between; align-items: center; padding: 8px 10px; background: #F5F7FA; border-radius: 4px; margin-bottom: 6px; }
    .kpi-mini-label { font-size: 12px; color: #6B7280; }
    .kpi-mini-val { font-size: 15px; font-weight: 500; color: #0068B7; }
    .phase-item { display: flex; gap: 10px; align-items: flex-start; margin-bottom: 10px; }
    .phase-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; margin-top: 4px; }
    .phase-period { font-size: 11px; font-weight: 500; color: #0068B7; margin-bottom: 2px; }
    .phase-content { font-size: 12px; color: #374151; line-height: 1.5; }
    .quote-box { font-size: 14px; font-style: italic; line-height: 1.7; padding: 12px 16px; margin-bottom: 14px; }
    .quote-box.blue { background: #E8F4FF; border-left: 4px solid #0068B7; color: #0C447C; }
    .flow-row { display: flex; align-items: stretch; gap: 6px; }
    .flow-step { flex: 1; background: #F5F7FA; border-radius: 8px; padding: 16px 12px; text-align: center; }
    .flow-step.active { background: #E8F4FF; border: 1px solid #BAD6F0; }
    .flow-step-label { font-size: 14px; font-weight: 500; color: #1A1A1A; margin-bottom: 4px; }
    .flow-step-detail { font-size: 12px; color: #6B7280; line-height: 1.5; }
    .flow-arrow { font-size: 20px; color: #0068B7; font-weight: 500; flex-shrink: 0; display: flex; align-items: center; }
    .system-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 14px; background: #F5F7FA; border-radius: 6px; margin-bottom: 6px; }
    .system-item.active { background: #E8F4FF; }
    .system-name { font-size: 14px; font-weight: 500; color: #1A1A1A; }
    .system-type { font-size: 11px; color: #6B7280; }
    .system-status { font-size: 11px; color: #0068B7; background: #E8F4FF; padding: 3px 8px; border-radius: 3px; }
    .news-item { display: flex; gap: 12px; align-items: flex-start; padding: 10px 14px; background: #F5F7FA; border-radius: 6px; }
    .news-date { background: #0068B7; color: #fff; font-size: 11px; padding: 3px 10px; border-radius: 3px; white-space: nowrap; flex-shrink: 0; }
    .news-content { flex: 1; }
    .news-title { font-size: 13px; color: #1A1A1A; font-weight: 500; }
    .news-link { font-size: 16px; text-decoration: none; flex-shrink: 0; }
  `;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${css}</style></head><body><div class="slide-wrapper">${slides.join("")}</div></body></html>`;
}

// PPT生成（1項目1スライド）
async function generatePPT(company, dept, results, pptStyle, selectedIds) {
  const prs = new PptxGenJS();
  prs.layout = "LAYOUT_WIDE";
  const today = new Date().toLocaleDateString("ja-JP");
  const W = 13.33, H = 7.5;
  const ss2 = safeStr;

  const addHeader = (slide, title, sub, accentColor = "0068B7") => {
    slide.addShape("rect", { x: 0, y: 0, w: W, h: 1.3, fill: { color: "FFFFFF" } });
    slide.addShape("rect", { x: 0, y: 1.28, w: W, h: 0.06, fill: { color: accentColor } });
    slide.addShape("rect", { x: 0, y: 0, w: 0.08, h: 1.3, fill: { color: accentColor } });
    if (sub) slide.addText(sub, { x: 0.2, y: 0.08, w: 10, h: 0.35, fontSize: 10, color: "6B7280" });
    slide.addText(title, { x: 0.2, y: 0.42, w: 9.5, h: 0.72, fontSize: 22, color: "1A1A1A", bold: true });
    slide.addText(company, { x: 10.8, y: 0.08, w: 2.4, h: 0.35, fontSize: 10, color: "B0B0B0", align: "right" });
  };
  const addFooter = (slide, pageNum, total, sourceUrl = "") => {
    slide.addShape("rect", { x: 0, y: H - 0.45, w: W, h: 0.45, fill: { color: "F5F7FA" } });
    slide.addShape("rect", { x: 0, y: H - 0.47, w: W, h: 0.04, fill: { color: "E0E0E0" } });
    const footerLeft = `${today}${dept ? " | " + dept : ""}${sourceUrl ? " | 📎 " + sourceUrl : ""}`;
    slide.addText(footerLeft, { x: 0.3, y: H - 0.38, w: 10, h: 0.3, fontSize: 9, color: "B0B0B0" });
    slide.addText(`${pageNum} / ${total}`, { x: W - 1.5, y: H - 0.38, w: 1.2, h: 0.3, fontSize: 9, color: "B0B0B0", align: "right" });
  };
  const tr = (v, len = 150) => { const s = ss2(v); return s.length > len ? s.slice(0, len) + "…" : s; };

  const activeIds = REPORT_ORDER.filter(id => selectedIds.has(id) && results[id]);
  const totalSlides = 1 + activeIds.length;

  // 表紙
  const cover = prs.addSlide();
  cover.background = { color: "F5F7FA" };
  cover.addShape("rect", { x: 0, y: 0, w: 0.12, h: H, fill: { color: "0068B7" } });
  cover.addText("顧客リサーチレポート　|　SoftBank Robotics Corp.", { x: 0.4, y: 1.4, w: 12, h: 0.5, fontSize: 13, color: "6B7280" });
  cover.addText(company, { x: 0.4, y: 2.0, w: 12, h: 1.4, fontSize: 44, color: "1A1A1A", bold: true });
  cover.addShape("rect", { x: 0.4, y: 3.6, w: 0.9, h: 0.06, fill: { color: "0068B7" } });
  if (dept) cover.addText(`担当部署：${dept}`, { x: 0.4, y: 3.9, w: 8, h: 0.45, fontSize: 15, color: "6B7280" });
  cover.addText(`調査日：${today}`, { x: 0.4, y: 4.5, w: 6, h: 0.4, fontSize: 13, color: "B0B0B0" });
  cover.addText(`${PPT_STYLES.find(s => s.id === pptStyle)?.label || "ビジュアル中心"}　|　全${totalSlides}枚`, { x: 0.4, y: 5.0, w: 6, h: 0.35, fontSize: 12, color: "B0B0B0" });

  // 1項目1スライド
  activeIds.forEach((itemId, idx) => {
    const result = results[itemId];
    const item = ALL_ITEMS.find(i => i.id === itemId);
    const cat = CATEGORIES.find(c => c.items.some(i => i.id === itemId));
    const accentHex = (cat?.color || C.blue).replace("#", "");
    const sourceUrls = result?.source_url && typeof result.source_url === "string"
      ? result.source_url.split(",").map(u => u.trim()).filter(u => u.startsWith("http"))
      : [];
    const sourceHost = sourceUrls.length > 0 ? (() => { try { return new URL(sourceUrls[0]).hostname; } catch { return sourceUrls[0]; } })() : "";
    const pageNum = idx + 2;
    const stLabel = result?.status === "confirmed" ? "✓ 確認済み" : result?.status === "partial" ? "△ 一部のみ" : "⚠ 要確認";
    const stColor = result?.status === "confirmed" ? "15803d" : result?.status === "partial" ? "a16207" : "dc2626";

    const slide = prs.addSlide();
    slide.background = { color: "FFFFFF" };
    addHeader(slide, item?.label || itemId, cat?.label || "", accentHex);
    addFooter(slide, pageNum, totalSlides, sourceHost);

    // ステータスバッジ
    slide.addText(stLabel, { x: 9.5, y: 0.45, w: 2.0, h: 0.4, fontSize: 11, color: stColor, align: "right" });

    const bodyY = 1.45, bodyH = H - 1.45 - 0.5;

    // 項目ごとのPPTレイアウト
    if (["pl_summary","growth_profit","financial_health"].includes(itemId)) {
      // 財務系：結論→KPI→グラフ
      const conclusions = Array.isArray(result?.conclusions) ? result.conclusions :
        result?.summary ? [result.summary] : [];
      let y = bodyY;
      conclusions.slice(0, 3).forEach(c => {
        slide.addShape("rect", { x: 0.3, y, w: W - 0.6, h: 0.4, fill: { color: "E8F4FF" } });
        slide.addShape("rect", { x: 0.3, y, w: 0.05, h: 0.4, fill: { color: accentHex } });
        slide.addText(`▶ ${tr(c, 90)}`, { x: 0.45, y: y + 0.06, w: W - 0.8, h: 0.28, fontSize: 11, color: "0C447C" });
        y += 0.46;
      });
      y += 0.1;
      if (itemId === "pl_summary") {
        const kpis = [
          { label: "売上高（最新期）", value: ss2(result?.revenue) || "—", sub: ss2(result?.growth) },
          { label: "営業利益率", value: ss2(result?.margin) || "—" },
          { label: "決算期", value: ss2(result?.fiscal_month) || "—" },
          { label: "物流コスト比率", value: ss2(result?.logistics_cost_ratio) || "要確認" },
        ];
        kpis.forEach((kpi, i) => {
          const x = 0.3 + i * 3.25;
          slide.addShape("rect", { x, y, w: 3.1, h: 1.1, fill: { color: "F5F7FA" } });
          slide.addShape("rect", { x, y, w: 3.1, h: 0.05, fill: { color: accentHex } });
          slide.addText(kpi.label, { x: x + 0.1, y: y + 0.08, w: 2.9, h: 0.3, fontSize: 9, color: "6B7280" });
          slide.addText(kpi.value, { x: x + 0.1, y: y + 0.38, w: 2.9, h: 0.55, fontSize: kpi.value.length > 8 ? 16 : 22, color: accentHex, bold: true });
          if (kpi.sub) slide.addText(kpi.sub, { x: x + 0.1, y: y + 0.9, w: 2.9, h: 0.18, fontSize: 9, color: "B0B0B0" });
        });
      } else if (itemId === "financial_health") {
        const eq = parseFloat(ss2(result?.equity_ratio)) || 0;
        const roe = parseFloat(ss2(result?.roe)) || 0;
        if (eq > 0) {
          slide.addText("自己資本比率", { x: 0.3, y, w: 4, h: 0.3, fontSize: 11, color: "6B7280" });
          slide.addText(`${eq}%`, { x: 4.5, y, w: 2, h: 0.3, fontSize: 18, color: accentHex, bold: true, align: "right" });
          slide.addShape("rect", { x: 0.3, y: y + 0.35, w: W - 0.6, h: 0.14, fill: { color: "E0E0E0" } });
          slide.addShape("rect", { x: 0.3, y: y + 0.35, w: Math.min(eq / 100, 1) * (W - 0.6), h: 0.14, fill: { color: accentHex } });
          y += 0.65;
        }
        if (roe > 0) {
          slide.addText("ROE", { x: 0.3, y, w: 4, h: 0.3, fontSize: 11, color: "6B7280" });
          slide.addText(`${roe}%`, { x: 4.5, y, w: 2, h: 0.3, fontSize: 18, color: accentHex, bold: true, align: "right" });
          slide.addShape("rect", { x: 0.3, y: y + 0.35, w: W - 0.6, h: 0.14, fill: { color: "E0E0E0" } });
          slide.addShape("rect", { x: 0.3, y: y + 0.35, w: Math.min(roe / 30, 1) * (W - 0.6), h: 0.14, fill: { color: accentHex } });
        }
      }
      if (result?.detail) {
        slide.addText(tr(result.detail, 200), { x: 0.3, y: H - 1.5, w: W - 0.6, h: 0.9, fontSize: 11, color: "374151", wrap: true });
      }

    } else if (itemId === "mid_term_plan") {
      const kpis = Array.isArray(result?.kpis) ? result.kpis : [];
      const phases = Array.isArray(result?.phases) ? result.phases : [];
      const colW = (W - 0.9) / 2;
      // 左列：計画名・引用・サマリー
      if (result?.plan_name) {
        slide.addShape("rect", { x: 0.3, y: bodyY, w: colW, h: 0.5, fill: { color: "E8F4FF" } });
        slide.addText(`${ss2(result.plan_name)}　${ss2(result?.period)}`, { x: 0.4, y: bodyY + 0.1, w: colW - 0.2, h: 0.3, fontSize: 12, color: accentHex, bold: true });
      }
      if (result?.quote) {
        slide.addShape("rect", { x: 0.3, y: bodyY + 0.6, w: 0.05, h: 0.8, fill: { color: accentHex } });
        slide.addText(`"${tr(ss2(result.quote), 80)}"`, { x: 0.45, y: bodyY + 0.65, w: colW - 0.3, h: 0.7, fontSize: 11, color: "0C447C", italic: true, wrap: true });
      }
      slide.addText(tr(ss2(result?.summary), 150), { x: 0.3, y: bodyY + 1.55, w: colW, h: 1.5, fontSize: 11, color: "374151", wrap: true });
      // 右列：KPI・フェーズ
      let ry = bodyY;
      if (kpis.length > 0) {
        slide.addText("重点KPI", { x: 0.45 + colW, y: ry, w: colW, h: 0.28, fontSize: 11, color: accentHex, bold: true });
        ry += 0.32;
        kpis.slice(0, 4).forEach(k => {
          slide.addShape("rect", { x: 0.45 + colW, y: ry, w: colW, h: 0.38, fill: { color: "F5F7FA" } });
          slide.addText(ss2(k.label), { x: 0.55 + colW, y: ry + 0.06, w: colW * 0.6, h: 0.26, fontSize: 10, color: "6B7280" });
          slide.addText(ss2(k.value), { x: 0.55 + colW + colW * 0.6, y: ry + 0.04, w: colW * 0.38, h: 0.3, fontSize: 13, color: accentHex, bold: true, align: "right" });
          ry += 0.42;
        });
        ry += 0.1;
      }
      if (phases.length > 0) {
        slide.addText("フェーズ別施策", { x: 0.45 + colW, y: ry, w: colW, h: 0.28, fontSize: 11, color: accentHex, bold: true });
        ry += 0.32;
        phases.slice(0, 4).forEach((p, i) => {
          slide.addShape("rect", { x: 0.45 + colW, y: ry, w: 0.1, h: 0.1, fill: { color: i === 0 ? accentHex : "BAD6F0" } });
          slide.addText(ss2(p.period), { x: 0.65 + colW, y: ry, w: colW - 0.25, h: 0.2, fontSize: 10, color: accentHex });
          slide.addText(tr(ss2(p.content), 60), { x: 0.65 + colW, y: ry + 0.2, w: colW - 0.25, h: 0.28, fontSize: 10, color: "374151", wrap: true });
          ry += 0.55;
        });
      }

    } else if (itemId === "dx_strategy" || itemId === "sustainability") {
      const items2 = Array.isArray(result?.initiatives) ? result.initiatives : Array.isArray(result?.targets) ? result.targets : [];
      const quoteHex = itemId === "sustainability" ? "15803d" : accentHex;
      const quoteBgHex = itemId === "sustainability" ? "f0fdf4" : "E8F4FF";
      const colW = (W - 0.9) / 2;
      if (result?.quote) {
        slide.addShape("rect", { x: 0.3, y: bodyY, w: colW, h: 0.9, fill: { color: quoteBgHex } });
        slide.addShape("rect", { x: 0.3, y: bodyY, w: 0.05, h: 0.9, fill: { color: quoteHex } });
        slide.addText(`"${tr(ss2(result.quote), 100)}"`, { x: 0.45, y: bodyY + 0.1, w: colW - 0.25, h: 0.7, fontSize: 11, color: quoteHex, italic: true, wrap: true });
      }
      slide.addText(tr(ss2(result?.summary), 200), { x: 0.3, y: bodyY + (result?.quote ? 1.0 : 0), w: colW, h: 2.0, fontSize: 12, color: "374151", wrap: true });
      if (items2.length > 0) {
        let ry = bodyY;
        slide.addText(itemId === "dx_strategy" ? "主要施策" : "目標・コミットメント", { x: 0.45 + colW, y: ry, w: colW, h: 0.28, fontSize: 11, color: accentHex, bold: true });
        ry += 0.32;
        items2.slice(0, 6).forEach(it => {
          slide.addShape("rect", { x: 0.45 + colW, y: ry, w: colW, h: 0.48, fill: { color: "F5F7FA" } });
          slide.addShape("rect", { x: 0.45 + colW, y: ry, w: 0.04, h: 0.48, fill: { color: accentHex } });
          slide.addText(tr(ss2(it.label || it.name || ""), 40), { x: 0.55 + colW, y: ry + 0.06, w: colW * 0.65, h: 0.28, fontSize: 11, color: "1A1A1A" });
          if (it.status || it.value) slide.addText(ss2(it.status || it.value), { x: 0.55 + colW + colW * 0.65, y: ry + 0.1, w: colW * 0.3, h: 0.28, fontSize: 10, color: accentHex, align: "right" });
          ry += 0.52;
        });
      }

    } else if (itemId === "logistics_flow") {
      const steps = Array.isArray(result?.steps) ? result.steps : [];
      if (steps.length > 0) {
        const stepW = (W - 0.8 - (steps.length - 1) * 0.2) / steps.length;
        steps.forEach((s, i) => {
          const x = 0.3 + i * (stepW + 0.2);
          slide.addShape("rect", { x, y: bodyY, w: stepW, h: 1.4, fill: { color: i === 0 ? "E8F4FF" : "F5F7FA" }, line: i === 0 ? { color: accentHex, width: 1 } : undefined });
          slide.addText(ss2(s.label), { x: x + 0.1, y: bodyY + 0.3, w: stepW - 0.2, h: 0.4, fontSize: 13, color: "1A1A1A", bold: true, align: "center" });
          slide.addText(tr(ss2(s.detail), 40), { x: x + 0.1, y: bodyY + 0.75, w: stepW - 0.2, h: 0.55, fontSize: 10, color: "6B7280", align: "center", wrap: true });
          if (i < steps.length - 1) {
            slide.addShape("rect", { x: x + stepW + 0.02, y: bodyY + 0.65, w: 0.16, h: 0.1, fill: { color: accentHex } });
          }
        });
      }
      slide.addText(tr(ss2(result?.summary), 250), { x: 0.3, y: bodyY + 1.6, w: W - 0.6, h: 1.5, fontSize: 12, color: "374151", wrap: true });

    } else if (itemId === "recent_news") {
      const news = Array.isArray(result?.news) ? result.news : [];
      let ny = bodyY;
      news.slice(0, 5).forEach(n => {
        slide.addShape("rect", { x: 0.3, y: ny, w: W - 0.6, h: 0.7, fill: { color: "F5F7FA" } });
        slide.addShape("rect", { x: 0.3, y: ny, w: 0.9, h: 0.7, fill: { color: accentHex } });
        slide.addText(ss2(n.date), { x: 0.32, y: ny + 0.2, w: 0.86, h: 0.3, fontSize: 10, color: "FFFFFF", align: "center" });
        slide.addText(tr(ss2(n.title), 70), { x: 1.3, y: ny + 0.12, w: W - 1.7, h: 0.3, fontSize: 12, color: "1A1A1A", bold: true });
        if (n.detail) slide.addText(tr(ss2(n.detail), 80), { x: 1.3, y: ny + 0.42, w: W - 1.7, h: 0.22, fontSize: 10, color: "6B7280" });
        ny += 0.76;
      });

    } else {
      // デフォルト：サマリー＋構造化データ
      slide.addText(tr(ss2(result?.summary), 300), { x: 0.3, y: bodyY, w: W - 0.6, h: 2.0, fontSize: 13, color: "374151", wrap: true });
      if (result?.missing && result?.status !== "confirmed") {
        slide.addShape("rect", { x: 0.3, y: bodyY + 2.1, w: W - 0.6, h: 0.6, fill: { color: "fef9c3" } });
        slide.addShape("rect", { x: 0.3, y: bodyY + 2.1, w: 0.05, h: 0.6, fill: { color: "d97706" } });
        slide.addText(`△ ${tr(ss2(result.missing), 150)}`, { x: 0.45, y: bodyY + 2.2, w: W - 0.8, h: 0.4, fontSize: 11, color: "a16207", wrap: true });
      }
    }
  });

  await prs.writeFile({ fileName: `SBR_リサーチ_${company}_${today}.pptx` });
}

// スライドプレビュー（iframeフルサイズ）
function SlidePreview({ company, dept, results, pptStyle, selectedIds }) {
  const [currentSlide, setCurrentSlide] = useState(0);
  const [generating, setGenerating] = useState(false);
  const iframeRef = useRef(null);
  const blobUrlRef = useRef(null);

  const activeIds = REPORT_ORDER.filter(id => selectedIds.has(id) && results[id]);
  const totalSlides = 1 + activeIds.length;

  useEffect(() => {
    const html = buildSlideHTML(company, dept, results, pptStyle, selectedIds);
    if (blobUrlRef.current) URL.revokeObjectURL(blobUrlRef.current);
    const blob = new Blob([html], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    blobUrlRef.current = url;
    if (iframeRef.current) iframeRef.current.src = url;
    return () => { if (blobUrlRef.current) URL.revokeObjectURL(blobUrlRef.current); };
  }, [company, dept, results, pptStyle, selectedIds]);

  useEffect(() => {
    const iframe = iframeRef.current;
    if (!iframe) return;
    const scrollToSlide = () => {
      try {
        const doc = iframe.contentDocument;
        if (!doc) return;
        const slides = doc.querySelectorAll(".slide");
        if (slides[currentSlide]) slides[currentSlide].scrollIntoView({ behavior: "smooth", block: "start" });
      } catch (_) {}
    };
    iframe.addEventListener("load", scrollToSlide);
    scrollToSlide();
    return () => iframe.removeEventListener("load", scrollToSlide);
  }, [currentSlide]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div style={{ flex: 1, overflow: "hidden", background: "#E8E8E8" }}>
        <iframe ref={iframeRef} style={{ width: "100%", height: "100%", border: "none", display: "block" }} sandbox="allow-same-origin allow-popups" title="スライドプレビュー" />
      </div>
      <div style={{ padding: "10px 16px", background: C.white, borderTop: `1px solid ${C.gray200}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexShrink: 0 }}>
        <button onClick={() => setCurrentSlide(s => Math.max(0, s - 1))} disabled={currentSlide === 0}
          style={{ padding: "6px 14px", borderRadius: 6, border: `1px solid ${C.gray200}`, background: currentSlide === 0 ? C.gray50 : C.white, cursor: currentSlide === 0 ? "not-allowed" : "pointer", fontSize: 12, color: currentSlide === 0 ? C.gray400 : C.black, fontFamily: "inherit" }}>
          ← 前へ
        </button>
        <div style={{ display: "flex", gap: 4, alignItems: "center", flexWrap: "wrap", maxWidth: 500, justifyContent: "center" }}>
          {Array.from({ length: totalSlides }).map((_, i) => (
            <div key={i} onClick={() => setCurrentSlide(i)}
              style={{ width: i === currentSlide ? 16 : 6, height: 6, borderRadius: 99, background: i === currentSlide ? C.blue : C.gray200, cursor: "pointer", transition: "all 0.2s", flexShrink: 0 }} />
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ fontSize: 11, color: C.gray600, whiteSpace: "nowrap" }}>{currentSlide + 1} / {totalSlides}</span>
          <button onClick={() => setCurrentSlide(s => Math.min(totalSlides - 1, s + 1))} disabled={currentSlide === totalSlides - 1}
            style={{ padding: "6px 14px", borderRadius: 6, border: `1px solid ${C.gray200}`, background: currentSlide === totalSlides - 1 ? C.gray50 : C.white, cursor: currentSlide === totalSlides - 1 ? "not-allowed" : "pointer", fontSize: 12, color: currentSlide === totalSlides - 1 ? C.gray400 : C.black, fontFamily: "inherit" }}>
            次へ →
          </button>
          <button onClick={async () => { setGenerating(true); try { await generatePPT(company, dept, results, pptStyle, selectedIds); } catch (e) { alert("PPT生成エラー: " + e.message); } setGenerating(false); }}
            disabled={generating}
            style={{ padding: "7px 18px", borderRadius: 6, border: "none", background: generating ? C.gray200 : C.blue, color: C.white, cursor: generating ? "not-allowed" : "pointer", fontSize: 12, fontWeight: 700, fontFamily: "inherit" }}>
            {generating ? "生成中..." : "📥 PPTダウンロード"}
          </button>
        </div>
      </div>
    </div>
  );
}

function WarehouseProgress({ done, total, current }) {
  const pct = total > 0 ? done / total : 0;
  const remaining = (total - done) * 60;
  const mins = Math.floor(remaining / 60), secs = remaining % 60;
  const timeStr = mins > 0 ? `約${mins}分` : `約${secs}秒`;
  const [tick, setTick] = useState(0);
  useEffect(() => { const id = setInterval(() => setTick(t => t + 1), 100); return () => clearInterval(id); }, []);
  const cycle = tick % 200, CONV_END = 80, ARM_END = 140;
  const boxConvX = cycle < CONV_END ? 15 + (cycle / CONV_END) * 100 : 115;
  const boxOnConv = cycle < CONV_END;
  const armPhase = cycle >= CONV_END && cycle < ARM_END ? (cycle - CONV_END) / (ARM_END - CONV_END) : 0;
  const armDown = Math.sin(armPhase * Math.PI);
  const armGripping = armPhase > 0.3 && armPhase < 0.85;
  const boxOnArm = armPhase > 0.35 && armPhase < 0.9;
  const forkX = 30 + pct * 155, beltOffset = (tick * 1.5) % 20;
  return (
    <div style={{ background: C.gray50, border: `1px solid ${C.gray200}`, borderRadius: 10, padding: "10px 14px", marginBottom: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <span style={{ fontSize: 11, color: C.gray600, fontWeight: 500 }}>{current ? `「${current}」調査中...` : done === total && total > 0 ? "完了！" : "準備中"}</span>
        <span style={{ fontSize: 11, color: C.gray400 }}>{done}/{total}</span>
      </div>
      <svg viewBox="0 0 300 82" width="100%" style={{ display: "block", background: C.gray100, borderRadius: 8 }}>
        <rect x="0" y="65" width="300" height="17" fill={C.gray200} /><rect x="0" y="63" width="300" height="3" fill={C.gray200} />
        <rect x="8" y="50" width="120" height="14" rx="3" fill="#334155" />
        <clipPath id="wbc9"><rect x="8" y="50" width="120" height="14" /></clipPath>
        <g clipPath="url(#wbc9)">{Array.from({ length: 9 }).map((_, i) => <line key={i} x1={8 + ((i * 16 - beltOffset + 120) % 120)} y1="50" x2={8 + ((i * 16 - beltOffset + 120) % 120)} y2="64" stroke="#475569" strokeWidth="1.5" />)}</g>
        <rect x="8" y="50" width="120" height="14" rx="3" fill="none" stroke="#475569" strokeWidth="1" />
        <circle cx="11" cy="57" r="5" fill="#1e293b" stroke="#475569" strokeWidth="1" /><circle cx="125" cy="57" r="5" fill="#1e293b" stroke="#475569" strokeWidth="1" />
        {boxOnConv && <g><rect x={boxConvX} y="40" width="16" height="12" rx="2" fill={C.blue} stroke={C.blueDark} strokeWidth="0.8" /></g>}
        <g transform="translate(145,8)">
          <rect x="-10" y="0" width="20" height="5" rx="2" fill="#475569" /><rect x="-4" y="-3" width="8" height="5" rx="1" fill="#64748b" />
          <rect x="-3" y="4" width="6" height={20 + armDown * 16} rx="2" fill="#334155" />
          <g transform={`translate(0,${24 + armDown * 16})`}>
            <rect x="-6" y="0" width="12" height="4" rx="1" fill="#1e293b" />
            <rect x={armGripping ? -7 : -9} y="3" width="5" height="8" rx="1" fill="#94a3b8" />
            <rect x={armGripping ? 2 : 4} y="3" width="5" height="8" rx="1" fill="#94a3b8" />
            {boxOnArm && <g transform="translate(-7,9)"><rect x="0" y="0" width="14" height="11" rx="2" fill={C.blue} stroke={C.blueDark} strokeWidth="0.6" /></g>}
          </g>
          <rect x="-10" y="3" width="7" height="9" rx="2" fill={C.blue} /><circle cx="-6" cy="7" r="2.5" fill={C.blueLight} />
        </g>
        <rect x="175" y="58" width="50" height="6" rx="1" fill="#92400e" /><rect x="179" y="54" width="42" height="5" rx="1" fill="#78350f" />
        {[180, 193, 206, 219].map(x => <rect key={x} x={x} y="63" width="4" height="2" fill="#78350f" />)}
        {Array.from({ length: Math.min(done, 8) }).map((_, i) => { const col = i % 4, row = Math.floor(i / 4); return (<g key={i}><rect x={180 + col * 10} y={46 - row * 11} width="9" height="9" rx="1" fill={C.blue} stroke={C.blueDark} strokeWidth="0.5" /></g>); })}
        <g transform={`translate(${forkX},28)`}>
          <rect x="14" y="-2" width="4" height="30" rx="1" fill="#374151" /><rect x="17" y="20" width="22" height="2.5" rx="1" fill="#94a3b8" /><rect x="17" y="24" width="22" height="2.5" rx="1" fill="#94a3b8" />
          <rect x="-13" y="10" width="29" height="18" rx="3" fill="#1A1A1A" /><rect x="-9" y="4" width="15" height="10" rx="2" fill="#374151" /><rect x="-7" y="5" width="11" height="7" rx="1" fill={C.gray200} opacity="0.85" />
          <ellipse cx="14" cy="14" rx="2.5" ry="2" fill={C.gray100} /><rect x="-15" y="16" width="6" height="10" rx="2" fill={C.blue} />
          <ellipse cx="8" cy="29" rx="5" ry="4" fill="#1e293b" /><ellipse cx="8" cy="29" rx="3" ry="2.5" fill="#334155" /><circle cx="8" cy="29" r="1" fill="#64748b" />
          <ellipse cx="-8" cy="29" rx="4" ry="3.5" fill="#1e293b" /><ellipse cx="-8" cy="29" rx="2.5" ry="2" fill="#334155" />
        </g>
        <rect x="258" y="40" width="38" height="25" rx="2" fill={C.gray50} stroke={C.gray200} strokeWidth="1" strokeDasharray="3,2" />
        <text x="277" y="53" fontSize="7" fill={C.blue} textAnchor="middle" fontWeight="600">GOAL</text>
        <text x="277" y="62" fontSize="6" fill={C.gray400} textAnchor="middle">出荷済</text>
      </svg>
      <div style={{ background: C.gray200, borderRadius: 99, height: 5, overflow: "hidden", marginTop: 8 }}>
        <div style={{ height: "100%", width: `${Math.round(pct * 100)}%`, background: C.blue, transition: "width .5s", borderRadius: 99 }} />
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
        <span style={{ fontSize: 10, color: C.gray600 }}>残り{timeStr}</span>
        <span style={{ fontSize: 10, color: C.gray600, fontWeight: 600 }}>{Math.round(pct * 100)}%</span>
      </div>
    </div>
  );
}

function CompanyModal({ company, candidates, onConfirm }) {
  const [selected, setSelected] = useState(candidates[0] || "");
  const [custom, setCustom] = useState("");
  const isOther = selected === "__other__";
  const finalName = isOther ? custom : selected;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
      <div style={{ background: C.white, borderRadius: 12, padding: "24px", width: 380, maxWidth: "90vw", boxShadow: "0 20px 60px rgba(0,0,0,0.15)" }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.black, marginBottom: 6 }}>企業名を確認してください</div>
        <div style={{ fontSize: 12, color: C.gray600, marginBottom: 16 }}>「{company}」に該当する企業候補：</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
          {candidates.map(c => (
            <label key={c} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 8, border: `1.5px solid ${selected === c ? C.blue : C.gray200}`, background: selected === c ? C.blueLight : C.white, cursor: "pointer" }}>
              <input type="radio" name="company" value={c} checked={selected === c} onChange={() => setSelected(c)} style={{ accentColor: C.blue }} />
              <span style={{ fontSize: 13, color: C.black, fontWeight: selected === c ? 600 : 400 }}>{c}</span>
            </label>
          ))}
          <label style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 8, border: `1.5px solid ${isOther ? C.blue : C.gray200}`, background: isOther ? C.blueLight : C.white, cursor: "pointer" }}>
            <input type="radio" name="company" value="__other__" checked={isOther} onChange={() => setSelected("__other__")} style={{ accentColor: C.blue }} />
            <span style={{ fontSize: 13, color: C.black }}>その他</span>
          </label>
          {isOther && <input autoFocus type="text" value={custom} onChange={e => setCustom(e.target.value)} placeholder="企業名を入力" style={{ padding: "9px 12px", fontSize: 13, border: `1.5px solid ${C.blue}`, borderRadius: 8, outline: "none", fontFamily: "inherit" }} />}
        </div>
        <button onClick={() => finalName.trim() && onConfirm(finalName.trim())} disabled={!finalName.trim()}
          style={{ width: "100%", padding: "11px", fontSize: 13, fontWeight: 700, background: finalName.trim() ? C.blue : C.gray200, color: C.white, border: "none", borderRadius: 8, cursor: finalName.trim() ? "pointer" : "not-allowed", fontFamily: "inherit" }}>
          確定してリサーチ開始
        </button>
      </div>
    </div>
  );
}

function RetryModal({ item, onConfirm, onClose }) {
  const [hint, setHint] = useState("");
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
      <div style={{ background: C.white, borderRadius: 12, padding: "24px", width: 380, maxWidth: "90vw", boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.black, marginBottom: 4 }}>「{item.label}」を追加調査</div>
        <div style={{ fontSize: 12, color: C.gray600, marginBottom: 14 }}>前回と異なるクエリ・ソースで再調査します</div>
        <div style={{ fontSize: 10, fontWeight: 600, color: C.gray600, marginBottom: 4 }}>深掘りしたい内容（任意）</div>
        <textarea value={hint} onChange={e => setHint(e.target.value)} placeholder="例：2025年3月期の数値が知りたい" rows={3}
          style={{ width: "100%", padding: "8px 10px", fontSize: 12, border: `1px solid ${C.gray200}`, borderRadius: 8, fontFamily: "inherit", boxSizing: "border-box", outline: "none", resize: "vertical", lineHeight: 1.6, marginBottom: 14 }}
          onFocus={e => e.target.style.borderColor = C.blue} onBlur={e => e.target.style.borderColor = C.gray200} />
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "9px", fontSize: 12, fontWeight: 600, fontFamily: "inherit", background: C.white, color: C.gray600, border: `1px solid ${C.gray200}`, borderRadius: 8, cursor: "pointer" }}>キャンセル</button>
          <button onClick={() => onConfirm(hint)} style={{ flex: 2, padding: "9px", fontSize: 13, fontWeight: 700, fontFamily: "inherit", background: C.blue, color: C.white, border: "none", borderRadius: 8, cursor: "pointer" }}>🔍 調査開始</button>
        </div>
      </div>
    </div>
  );
}

function StatusBadge({ status, size = "sm" }) {
  if (!status) return null;
  const cfg = STATUS[status] || STATUS.unconfirmed;
  return <span style={{ fontSize: size === "lg" ? 12 : 10, fontWeight: 600, color: cfg.color, background: cfg.bg, border: `1px solid ${cfg.border}`, padding: size === "lg" ? "3px 9px" : "1px 6px", borderRadius: 99, whiteSpace: "nowrap", flexShrink: 0 }}>{cfg.label}</span>;
}

function Checkbox({ checked, color, indeterminate }) {
  return (
    <div style={{ width: 14, height: 14, borderRadius: 3, flexShrink: 0, marginTop: 1, border: `1.5px solid ${checked || indeterminate ? color : C.gray200}`, background: checked ? color : "transparent", display: "flex", alignItems: "center", justifyContent: "center" }}>
      {checked && <span style={{ fontSize: 9, color: "#fff", fontWeight: 700, lineHeight: 1 }}>✓</span>}
      {indeterminate && !checked && <div style={{ width: 6, height: 2, background: color, borderRadius: 1 }} />}
    </div>
  );
}

function SectionLabel({ label, tooltip, sub }) {
  return (
    <div style={{ fontSize: 11, fontWeight: 600, color: C.blue, marginBottom: sub ? 4 : 8, display: "flex", alignItems: "center", gap: 5 }}>
      <div style={{ width: 3, height: 11, background: C.blue, borderRadius: 2 }} />
      {label}
      {tooltip && <Tooltip text={tooltip} />}
      {sub && <span style={{ fontWeight: 400, color: C.gray400, fontSize: 10 }}>{sub}</span>}
    </div>
  );
}

function Tooltip({ text }) {
  const [show, setShow] = useState(false);
  return (
    <div style={{ position: "relative", display: "inline-flex", alignItems: "center" }} onMouseEnter={() => setShow(true)} onMouseLeave={() => setShow(false)}>
      <div style={{ width: 14, height: 14, borderRadius: "50%", background: C.gray200, display: "flex", alignItems: "center", justifyContent: "center", cursor: "help", fontSize: 9, color: C.gray600, fontWeight: 700 }}>?</div>
      {show && (
        <div style={{ position: "absolute", left: 20, top: -4, width: 220, background: C.black, color: C.gray100, fontSize: 11, lineHeight: 1.6, padding: "8px 10px", borderRadius: 8, zIndex: 100, boxShadow: "0 4px 16px rgba(0,0,0,0.2)", whiteSpace: "pre-wrap" }}>
          {text}
          <div style={{ position: "absolute", left: -5, top: 8, width: 0, height: 0, borderTop: "5px solid transparent", borderBottom: "5px solid transparent", borderRight: `5px solid ${C.black}` }} />
        </div>
      )}
    </div>
  );
}

function ResultCardVisual({ item, result, onRetry, retrying }) {
  const [open, setOpen] = useState(true);
  const st = result?.status || "unconfirmed";
  const cfg = STATUS[st];
  const sourceUrls = result?.source_url && typeof result.source_url === "string" ? result.source_url.split(",").map(u => u.trim()).filter(u => u.startsWith("http")) : [];
  const isFinance = ["pl_summary", "growth_profit", "financial_health"].includes(item.id);
  return (
    <div style={{ borderRadius: 10, marginBottom: 10, overflow: "hidden", background: C.white, boxShadow: "0 1px 4px rgba(0,0,0,0.06)", border: `1px solid ${C.gray200}` }}>
      <div onClick={() => setOpen(o => !o)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 14px", cursor: "pointer", background: C.gray50, borderBottom: open ? `1px solid ${C.gray200}` : "none" }}>
        <div style={{ width: 4, height: 16, background: C.blue, borderRadius: 2, flexShrink: 0 }} />
        <span style={{ fontSize: 13, fontWeight: 700, color: C.black, flex: 1 }}>{item.label}</span>
        <StatusBadge status={st} />
        <span style={{ fontSize: 11, color: C.gray400, marginLeft: 4 }}>{open ? "▲" : "▼"}</span>
      </div>
      {open && (
        <div style={{ padding: "12px 14px 14px" }}>
          {isFinance && result?.detail ? (
            <>
              <div style={{ background: C.blueLight, border: `1px solid #BAD6F0`, borderRadius: 8, padding: "8px 12px", marginBottom: 8 }}>
                <div style={{ fontSize: 10, fontWeight: 600, color: C.blue, marginBottom: 3 }}>結論</div>
                <div style={{ fontSize: 13, fontWeight: 600, color: C.black, lineHeight: 1.7 }}>{safeStr(result.summary)}</div>
              </div>
              <div style={{ fontSize: 12, color: C.gray800, lineHeight: 1.85, marginBottom: 8 }}>{safeStr(result.detail)}</div>
              {(result.revenue || result.margin || result.fiscal_month) && (
                <div style={{ display: "flex", gap: 6, marginBottom: 8, flexWrap: "wrap" }}>
                  {result.fiscal_month && <div style={{ background: C.gray100, borderRadius: 6, padding: "4px 10px", fontSize: 10, color: C.gray800 }}><span style={{ color: C.gray600 }}>決算期</span> {safeStr(result.fiscal_month)}</div>}
                  {result.revenue && <div style={{ background: C.gray100, borderRadius: 6, padding: "4px 10px", fontSize: 10, color: C.gray800 }}><span style={{ color: C.gray600 }}>売上高</span> {safeStr(result.revenue)}</div>}
                  {result.margin && <div style={{ background: C.gray100, borderRadius: 6, padding: "4px 10px", fontSize: 10, color: C.gray800 }}><span style={{ color: C.gray600 }}>営業利益率</span> {safeStr(result.margin)}</div>}
                  {result.logistics_cost_ratio && <div style={{ background: C.gray100, borderRadius: 6, padding: "4px 10px", fontSize: 10, color: C.gray800 }}><span style={{ color: C.gray600 }}>物流コスト比率</span> {safeStr(result.logistics_cost_ratio)}</div>}
                </div>
              )}
            </>
          ) : (
            <p style={{ margin: "0 0 8px", fontSize: 13, color: C.gray800, lineHeight: 1.85 }}>{safeStr(result?.summary) || "情報を取得できませんでした。"}</p>
          )}
          {result?.quote && (
            <div style={{ background: C.gray50, border: `1px solid ${C.gray200}`, borderLeft: `3px solid ${C.blue}`, borderRadius: 4, padding: "8px 12px", marginBottom: 8, fontSize: 12, color: C.gray800, fontStyle: "italic", lineHeight: 1.7 }}>
              "{safeStr(result.quote)}"
            </div>
          )}
          {result?.missing && st !== "confirmed" && (
            <div style={{ display: "flex", alignItems: "flex-start", gap: 7, background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 7, padding: "7px 11px", marginBottom: 8 }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: cfg.color, whiteSpace: "nowrap", marginTop: 1 }}>{st === "unconfirmed" ? "⚠ 要ヒアリング" : "△ 要確認"}</span>
              <span style={{ fontSize: 12, color: cfg.color, lineHeight: 1.6 }}>{safeStr(result.missing)}</span>
            </div>
          )}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 6 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 5, flexWrap: "wrap" }}>
              {sourceUrls.map((url, i) => { let host = url; try { host = new URL(url).hostname; } catch { } return (<a key={i} href={url} target="_blank" rel="noopener noreferrer" style={{ fontSize: 10, color: C.blue, textDecoration: "none" }}>📎 {host}</a>); })}
            </div>
            <button onClick={onRetry} disabled={retrying} style={{ padding: "4px 10px", fontSize: 10, fontWeight: 600, fontFamily: "inherit", background: retrying ? C.gray100 : C.white, color: retrying ? C.gray400 : C.blue, border: `1px solid ${retrying ? C.gray200 : C.blue}`, borderRadius: 6, cursor: retrying ? "not-allowed" : "pointer", whiteSpace: "nowrap" }}>
              {retrying ? "調査中..." : "🔍 追加調査"}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function ResultCardText({ item, result, onRetry, retrying }) {
  const [open, setOpen] = useState(true);
  const st = result?.status || "unconfirmed";
  const cfg = STATUS[st];
  const sourceUrls = result?.source_url && typeof result.source_url === "string" ? result.source_url.split(",").map(u => u.trim()).filter(u => u.startsWith("http")) : [];
  const isFinance = ["pl_summary", "growth_profit", "financial_health"].includes(item.id);
  return (
    <div style={{ borderRadius: 8, marginBottom: 8, overflow: "hidden", background: C.white, border: `1px solid ${C.gray200}` }}>
      <div onClick={() => setOpen(o => !o)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 14px", cursor: "pointer" }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: C.black, flex: 1 }}>{item.label}</span>
        <StatusBadge status={st} />
        <span style={{ fontSize: 11, color: C.gray400, marginLeft: 4 }}>{open ? "▲" : "▼"}</span>
      </div>
      {open && (
        <div style={{ padding: "8px 14px 12px", borderTop: `1px solid ${C.gray100}` }}>
          {isFinance && result?.detail ? (
            <>
              <p style={{ margin: "0 0 4px", fontSize: 13, fontWeight: 600, color: C.black, lineHeight: 1.7 }}>{safeStr(result.summary)}</p>
              <p style={{ margin: "0 0 8px", fontSize: 12, color: C.gray800, lineHeight: 1.85 }}>{safeStr(result.detail)}</p>
            </>
          ) : (
            <p style={{ margin: "0 0 8px", fontSize: 12, color: C.gray800, lineHeight: 1.85 }}>{safeStr(result?.summary) || "情報を取得できませんでした。"}</p>
          )}
          {result?.quote && <p style={{ margin: "0 0 8px", fontSize: 12, color: C.gray600, borderLeft: `2px solid ${C.gray200}`, paddingLeft: 8, fontStyle: "italic" }}>"{safeStr(result.quote)}"</p>}
          {result?.missing && st !== "confirmed" && <p style={{ margin: "0 0 8px", fontSize: 11, color: cfg.color }}>{st === "unconfirmed" ? "⚠ 要ヒアリング：" : "△ 要確認："}{safeStr(result.missing)}</p>}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <div style={{ fontSize: 10, color: C.gray400 }}>
              {sourceUrls.length > 0 && `📎 ${sourceUrls.map(u => { try { return new URL(u).hostname; } catch { return u; } }).join(", ")}`}
            </div>
            <button onClick={onRetry} disabled={retrying} style={{ padding: "3px 8px", fontSize: 10, fontWeight: 600, fontFamily: "inherit", background: C.white, color: retrying ? C.gray400 : C.blue, border: `1px solid ${retrying ? C.gray200 : C.blue}`, borderRadius: 5, cursor: retrying ? "not-allowed" : "pointer" }}>
              {retrying ? "調査中..." : "🔍 追加調査"}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [company, setCompany] = useState("");
  const [dept, setDept] = useState("");
  const [priorInfo, setPriorInfo] = useState("");
  const [parsedPrior, setParsedPrior] = useState(null);
  const [selected, setSelected] = useState(new Set(PRESETS[0].ids));
  const [industryMode, setIndustryMode] = useState("auto");
  const [industrySelect, setIndustrySelect] = useState("食品・飲料・冷凍食品");
  const [industryCustom, setIndustryCustom] = useState("");
  const [pptStyle, setPptStyle] = useState("visual");
  const [appStatus, setAppStatus] = useState("idle");
  const [progress, setProgress] = useState({ done: 0, total: 0, current: "" });
  const [results, setResults] = useState({});
  const [activeTab, setActiveTab] = useState("data");
  const [modal, setModal] = useState(null);
  const [retryModal, setRetryModal] = useState(null);
  const [retryingIds, setRetryingIds] = useState(new Set());
  const [bulkRetrying, setBulkRetrying] = useState(false);
  const abortRef = useRef(false);

  const getIndustry = () => { if (industryMode === "auto") return ""; if (industrySelect === "その他") return industryCustom; return industrySelect; };
  const activePreset = PRESETS.findIndex(p => { if (p.ids.size !== selected.size) return false; for (const id of p.ids) if (!selected.has(id)) return false; return true; });
  const toggle = (id) => setSelected(s => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleCat = (cat) => { const ids = cat.items.map(i => i.id); const allOn = ids.every(id => selected.has(id)); setSelected(s => { const n = new Set(s); ids.forEach(id => allOn ? n.delete(id) : n.add(id)); return n; }); };
  const applyPreset = (preset) => setSelected(new Set(preset.ids));

  const orderedResults = REPORT_ORDER.filter(id => selected.has(id) && results[id]);
  const confirmedResults = orderedResults.filter(id => results[id]?.status !== "unconfirmed");
  const unconfirmedResults = orderedResults.filter(id => results[id]?.status === "unconfirmed");
  const statusCounts = Object.values(results).reduce((acc, r) => { const k = r?.status || "unconfirmed"; acc[k] = (acc[k] || 0) + 1; return acc; }, {});

  const checkCompany = async (name) => {
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        if (attempt > 0) await new Promise(r => setTimeout(r, attempt * 3000));
        const raw = await callClaude(`あなたは企業名の曖昧さを判定するAIです。企業名が曖昧・略称・グループ名の場合のみ候補リストを返してください。明確な企業名の場合はambiguous:falseを返してください。JSONのみ返答。`, `企業名「${name}」は曖昧ですか？曖昧な場合は日本の代表的な該当企業を最大5件リストアップ。JSONのみ: {"ambiguous":true|false,"candidates":["企業名1","企業名2"]}`);
        const parsed = extractJSON(raw);
        if (parsed?.ambiguous && parsed?.candidates?.length > 0) return parsed.candidates;
        return null;
      } catch (_) { if (attempt === 2) return null; }
    }
    return null;
  };

  const startResearch = async (confirmedCompany) => {
    setModal(null); setCompany(confirmedCompany); abortRef.current = false;
    setAppStatus("loading"); setResults({}); setParsedPrior(null); setActiveTab("data");
    const items = ALL_ITEMS.filter(i => selected.has(i.id));
    const hasPrior = priorInfo.trim().length > 0;
    const industry = getIndustry();
    setProgress({ done: 0, total: items.length + (hasPrior ? 1 : 0), current: "" });
    let priorContext = "";
    if (hasPrior) {
      setProgress({ done: 0, total: items.length + 1, current: "事前情報を整理中" });
      try {
        const parseRaw = await callClaudeNoSearch(`あなたはB2B営業支援AIです。貼り付けられたヒアリングメモ・事前情報を整理しJSONのみ返答。前置き不要。`, `以下は"${confirmedCompany}"に関する事前情報です。\n---\n${priorInfo}\n---\nJSONのみ: {"summary":"全体サマリー3〜4文","key_points":["重要ポイント1","重要ポイント2"],"known_challenges":"判明している課題・ニーズ","known_contacts":"判明している担当者情報","known_systems":"判明しているシステム環境","other":"その他有用情報"}`);
        const pp = extractJSON(parseRaw);
        if (pp) { setParsedPrior(pp); priorContext = "\n\n【事前情報】" + JSON.stringify(pp); }
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
        const base = item.id === "market_share" ? PROMPTS.market_share(confirmedCompany, industry) : (PROMPTS[item.id]?.(confirmedCompany) || `"${confirmedCompany}"について「${item.label}」を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`);
        const prompt = priorContext ? base + priorContext + "\n※事前情報に記載された内容はそのままsummaryに含めてよい。推測は書かない。" : base;
        const raw = await deepResearchPremium(confirmedCompany, item.id, prompt, industry);
        const parsed = extractJSON(raw);
        newResults[item.id] = parsed || { summary: raw.slice(0, 200), status: "partial", missing: "", source_url: "" };
      } catch { newResults[item.id] = { summary: "取得できませんでした。", status: "unconfirmed", missing: "デスクトップリサーチでは確認不可。訪問時に直接確認が必要です。", source_url: "" }; }
      setResults({ ...newResults });
      await new Promise(r => setTimeout(r, 15000));
    }
    setProgress(p => ({ ...p, done: p.total, current: "" }));
    setAppStatus("done");
  };

  const handleRetry = async (itemId, hint = "") => {
    setRetryModal(null);
    setRetryingIds(s => { const n = new Set(s); n.add(itemId); return n; });
    const item = ALL_ITEMS.find(i => i.id === itemId);
    const industry = getIndustry();
    try {
      const base = itemId === "market_share" ? PROMPTS.market_share(company, industry) : (PROMPTS[itemId]?.(company) || `"${company}"について「${item?.label}」を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`);
      const raw = await retryResearch(company, itemId, base, industry, hint);
      const parsed = extractJSON(raw);
      if (parsed) {
        setResults(prev => {
          const current = prev[itemId];
          const rank = { confirmed: 3, partial: 2, unconfirmed: 1 };
          const currentRank = rank[current?.status] || 0;
          const newRank = rank[parsed?.status] || 0;
          if (newRank >= currentRank) {
            return { ...prev, [itemId]: parsed };
          } else {
            return {
              ...prev,
              [itemId]: {
                ...current,
                summary: current?.summary && current.summary !== "取得できませんでした。" ? current.summary : parsed.summary,
                source_url: current?.source_url || parsed.source_url,
              }
            };
          }
        });
      }
    } catch (_) {}
    setRetryingIds(s => { const n = new Set(s); n.delete(itemId); return n; });
  };

  const handleBulkRetry = async () => {
    setBulkRetrying(true);
    for (const id of unconfirmedResults) {
      if (abortRef.current) break;
      await handleRetry(id, "");
      await new Promise(r => setTimeout(r, 10000));
    }
    setBulkRetrying(false);
  };

  const handleStart = async () => {
    if (!company.trim() || selected.size === 0) return;
    setAppStatus("checking");
    const candidates = await checkCompany(company.trim());
    if (candidates) { setModal({ candidates }); setAppStatus("idle"); }
    else await startResearch(company.trim());
  };

  const resetResearch = () => { abortRef.current = true; setAppStatus("idle"); setResults({}); setParsedPrior(null); setProgress({ done: 0, total: 0, current: "" }); };
  const hasResults = Object.keys(results).length > 0;
  const isLoading = appStatus === "loading" || appStatus === "checking";
  const ResultCard = pptStyle === "visual" ? ResultCardVisual : ResultCardText;

  const inp = (label, val, setter, ph) => (
    <div style={{ marginBottom: 7 }}>
      <div style={{ fontSize: 10, fontWeight: 600, color: C.gray600, marginBottom: 3 }}>{label}</div>
      <input type="text" value={val} onChange={e => setter(e.target.value)} placeholder={ph}
        style={{ width: "100%", padding: "7px 9px", fontSize: 12, border: `1px solid ${C.gray200}`, borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none" }}
        onFocus={e => e.target.style.borderColor = C.blue} onBlur={e => e.target.style.borderColor = C.gray200} />
    </div>
  );

  return (
    <div style={{ fontFamily: "'Helvetica Neue','Hiragino Kaku Gothic ProN',Meiryo,sans-serif", background: C.gray50, minHeight: "100vh" }}>
      {modal && <CompanyModal company={company} candidates={modal.candidates} onConfirm={startResearch} />}
      {retryModal && <RetryModal item={retryModal.item} onConfirm={(hint) => handleRetry(retryModal.itemId, hint)} onClose={() => setRetryModal(null)} />}

      <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "12px 22px", display: "flex", alignItems: "center", gap: 11 }}>
        <div style={{ width: 6, height: 32, background: C.blue, borderRadius: 3 }} />
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: C.black }}>SBR 顧客リサーチ &amp; 社内報告ツール</div>
          <div style={{ fontSize: 11, color: C.gray400 }}>新規顧客の事前調査 → 社内報告PPT 一気通貫</div>
        </div>
      </div>

      <div style={{ display: "flex", height: "calc(100vh - 57px)" }}>
        <div style={{ width: 292, background: C.white, borderRight: `1px solid ${C.gray200}`, overflowY: "auto", display: "flex", flexDirection: "column", flexShrink: 0 }}>
          <div style={{ padding: "14px 14px 0", flex: 1 }}>
            <SectionLabel label="基本情報" tooltip={TOOLTIPS.basicInfo} />
            <div style={{ marginBottom: 7 }}>
              <div style={{ fontSize: 10, fontWeight: 600, color: C.gray600, marginBottom: 3 }}>顧客企業名 *</div>
              <input type="text" value={company} onChange={e => setCompany(e.target.value)}
                onKeyDown={e => e.key === "Enter" && !isLoading && company.trim() && selected.size && handleStart()}
                placeholder="例：ニチレイロジグループ株式会社"
                style={{ width: "100%", padding: "7px 9px", fontSize: 12, border: `1px solid ${C.gray200}`, borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none" }}
                onFocus={e => e.target.style.borderColor = C.blue} onBlur={e => e.target.style.borderColor = C.gray200} />
            </div>
            {inp("担当者部署", dept, setDept, "例：物流部・SCM部")}

            <SectionLabel label="業界" tooltip={TOOLTIPS.industry} />
            <div style={{ display: "flex", gap: 12, marginBottom: 8 }}>
              {["auto", "specify"].map(mode => (
                <label key={mode} style={{ display: "flex", alignItems: "center", gap: 5, cursor: "pointer", fontSize: 12 }}>
                  <input type="radio" name="industryMode" value={mode} checked={industryMode === mode} onChange={() => setIndustryMode(mode)} style={{ accentColor: C.blue }} />
                  <span style={{ color: C.black }}>{mode === "auto" ? "自動判定" : "指定する"}</span>
                </label>
              ))}
            </div>
            {industryMode === "specify" && (
              <div style={{ marginBottom: 8 }}>
                <select value={industrySelect} onChange={e => setIndustrySelect(e.target.value)}
                  style={{ width: "100%", padding: "7px 9px", fontSize: 12, border: `1px solid ${C.gray200}`, borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none", marginBottom: 6 }}>
                  {INDUSTRY_OPTIONS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                </select>
                {industrySelect === "その他" && <input type="text" value={industryCustom} onChange={e => setIndustryCustom(e.target.value)} placeholder="業界名を入力"
                  style={{ width: "100%", padding: "7px 9px", fontSize: 12, border: `1px solid ${C.gray200}`, borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none" }}
                  onFocus={e => e.target.style.borderColor = C.blue} onBlur={e => e.target.style.borderColor = C.gray200} />}
              </div>
            )}

            <div style={{ height: 1, background: C.gray100, margin: "10px 0" }} />
            <SectionLabel label="事前情報" tooltip={TOOLTIPS.priorInfo} sub="任意" />
            <textarea value={priorInfo} onChange={e => setPriorInfo(e.target.value)}
              placeholder={"例：\n・担当の山田部長から「WMS刷新を検討中」と聞いた\n・現状はExcel管理でミスが多い"} rows={4}
              style={{ width: "100%", padding: "8px 9px", fontSize: 12, border: `1px solid ${C.gray200}`, borderRadius: 7, fontFamily: "inherit", boxSizing: "border-box", outline: "none", resize: "vertical", lineHeight: 1.6 }}
              onFocus={e => e.target.style.borderColor = C.blue} onBlur={e => e.target.style.borderColor = C.gray200} />
            {priorInfo.trim() && <div style={{ fontSize: 11, color: C.green, marginBottom: 4, marginTop: 3 }}>✓ 事前情報あり — リサーチ時に自動反映します</div>}

            <div style={{ height: 1, background: C.gray100, margin: "10px 0" }} />
            <SectionLabel label="レポートスタイル" tooltip={TOOLTIPS.pptStyle} />
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 10 }}>
              {PPT_STYLES.map(s => (
                <label key={s.id} style={{ display: "flex", alignItems: "flex-start", gap: 8, padding: "8px 10px", borderRadius: 8, border: `1.5px solid ${pptStyle === s.id ? C.blue : C.gray200}`, background: pptStyle === s.id ? C.blueLight : C.white, cursor: "pointer" }}>
                  <input type="radio" name="pptStyle" value={s.id} checked={pptStyle === s.id} onChange={() => setPptStyle(s.id)} style={{ accentColor: C.blue, marginTop: 2 }} />
                  <div>
                    <div style={{ fontSize: 12, fontWeight: 600, color: C.black }}>{s.icon} {s.label}</div>
                    <div style={{ fontSize: 10, color: C.gray600, marginTop: 2 }}>{s.desc}</div>
                  </div>
                </label>
              ))}
            </div>

            <div style={{ height: 1, background: C.gray100, margin: "10px 0" }} />
            <SectionLabel label="調査項目" tooltip={TOOLTIPS.items} sub={`${selected.size}件選択`} />
            <div style={{ display: "flex", gap: 4, marginBottom: 10, flexWrap: "wrap" }}>
              {PRESETS.map((p, idx) => { const isActive = activePreset === idx; return (
                <button key={p.label} onClick={() => applyPreset(p)}
                  style={{ fontSize: 10, padding: "4px 10px", borderRadius: 99, fontFamily: "inherit", cursor: "pointer", fontWeight: isActive ? 700 : 500, background: isActive ? C.blue : C.gray100, color: isActive ? C.white : C.gray800, border: isActive ? `1.5px solid ${C.blue}` : `1.5px solid ${C.gray200}` }}>
                  {isActive ? "✓ " : ""}{p.label}
                </button>
              ); })}
            </div>
            {CATEGORIES.map(cat => {
              const ids = cat.items.map(i => i.id); const allOn = ids.every(id => selected.has(id)); const partial = ids.some(id => selected.has(id)) && !allOn;
              return (
                <div key={cat.id} style={{ marginBottom: 9 }}>
                  <div onClick={() => toggleCat(cat)} style={{ display: "flex", alignItems: "center", gap: 7, padding: "4px 6px", borderRadius: 6, cursor: "pointer", marginBottom: 3 }}>
                    <Checkbox checked={allOn} indeterminate={partial} color={cat.color} />
                    <div style={{ width: 7, height: 7, borderRadius: "50%", background: cat.color, flexShrink: 0 }} />
                    <span style={{ fontSize: 11, fontWeight: 600, color: C.black }}>{cat.label}</span>
                  </div>
                  <div style={{ paddingLeft: 10 }}>
                    {cat.items.map(item => { const on = selected.has(item.id); const r = results[item.id]; return (
                      <label key={item.id} style={{ display: "flex", alignItems: "flex-start", gap: 7, padding: "4px 6px", borderRadius: 6, cursor: "pointer", marginBottom: 1 }}>
                        <Checkbox checked={on} color={cat.color} />
                        <input type="checkbox" checked={on} onChange={() => toggle(item.id)} style={{ display: "none" }} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 12, color: on ? C.black : C.gray400, fontWeight: on ? 500 : 400, display: "flex", alignItems: "center", gap: 5, flexWrap: "wrap", lineHeight: 1.4 }}>
                            {item.label}{r && <StatusBadge status={r.status} />}
                          </div>
                        </div>
                      </label>
                    ); })}
                  </div>
                </div>
              );
            })}
          </div>

          <div style={{ padding: "12px 14px", borderTop: `1px solid ${C.gray100}` }}>
            {appStatus === "checking" && <div style={{ fontSize: 11, color: C.gray600, textAlign: "center", marginBottom: 8 }}>🔍 企業名を確認中...</div>}
            {appStatus === "loading" && <WarehouseProgress done={progress.done} total={progress.total} current={progress.current} />}
            <div style={{ display: "flex", gap: 6 }}>
              {hasResults && <button onClick={resetResearch} style={{ flex: 1, padding: "10px", fontSize: 12, fontWeight: 600, fontFamily: "inherit", background: C.white, color: C.gray800, border: `1px solid ${C.gray200}`, borderRadius: 8, cursor: "pointer" }}>新しいリサーチ</button>}
              {isLoading ?
                <button onClick={() => { abortRef.current = true; setAppStatus("idle"); }} style={{ flex: 1, padding: "10px", fontSize: 13, fontWeight: 700, fontFamily: "inherit", background: C.red, color: C.white, border: "none", borderRadius: 8, cursor: "pointer" }}>⏹ 中断</button> :
                <button onClick={handleStart} disabled={!company.trim() || !selected.size}
                  style={{ flex: 1, padding: "10px", fontSize: 13, fontWeight: 700, fontFamily: "inherit", background: (!company.trim() || !selected.size) ? C.gray200 : C.blue, color: C.white, border: "none", borderRadius: 8, cursor: (!company.trim() || !selected.size) ? "not-allowed" : "pointer" }}>
                  🔍 リサーチ開始
                </button>}
            </div>
          </div>
        </div>

        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          {!hasResults ? (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: C.gray200 }}>
              <div style={{ fontSize: 48, marginBottom: 12 }}>🔍</div>
              <div style={{ fontSize: 15, fontWeight: 600, color: C.gray400 }}>顧客名を入力してリサーチ開始</div>
              <div style={{ fontSize: 12, marginTop: 6, color: C.gray400 }}>1項目 = 1スライドで出力されます</div>
            </div>
          ) : (
            <>
              <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "0 20px", display: "flex", alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 16, flex: 1 }}>
                  {[{ id: "data", label: "📋 調査結果" }, { id: "slide", label: "📊 スライド" }].map(tab => (
                    <button key={tab.id} onClick={() => setActiveTab(tab.id)}
                      style={{ padding: "12px 4px", fontSize: 13, fontWeight: activeTab === tab.id ? 700 : 400, color: activeTab === tab.id ? C.blue : C.gray600, background: "none", border: "none", borderBottom: activeTab === tab.id ? `2px solid ${C.blue}` : "2px solid transparent", cursor: "pointer", fontFamily: "inherit" }}>
                      {tab.label}
                    </button>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  {["confirmed", "partial", "unconfirmed"].map(s => statusCounts[s] ? (
                    <div key={s} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <StatusBadge status={s} size="lg" /><span style={{ fontSize: 12, color: C.gray600 }}>{statusCounts[s]}件</span>
                    </div>
                  ) : null)}
                </div>
              </div>

              {activeTab === "data" ? (
                <div style={{ flex: 1, overflowY: "auto", padding: 20 }}>
                  <div style={{ maxWidth: 820 }}>
                    <div style={{ marginBottom: 16 }}>
                      <div style={{ fontSize: 20, fontWeight: 700, color: C.black }}>{company}</div>
                      {dept && <div style={{ fontSize: 12, color: C.gray600, marginTop: 2 }}>担当部署：{dept}</div>}
                    </div>

                    {parsedPrior && (
                      <div style={{ background: C.greenBg, border: `1px solid ${C.greenBorder}`, borderRadius: 10, padding: "12px 16px", marginBottom: 16 }}>
                        <div style={{ fontSize: 13, fontWeight: 600, color: C.green, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                          <span>📋 事前情報（整理済み）</span>
                          <span style={{ fontSize: 10, fontWeight: 400, color: C.green, background: C.greenBg, border: `1px solid ${C.greenBorder}`, padding: "1px 7px", borderRadius: 99 }}>リサーチに反映済み</span>
                        </div>
                        <p style={{ margin: "0 0 8px", fontSize: 13, color: "#166534", lineHeight: 1.75 }}>{safeStr(parsedPrior.summary)}</p>
                        {parsedPrior.key_points?.length > 0 && (
                          <div>
                            <div style={{ fontSize: 11, fontWeight: 600, color: C.green, marginBottom: 4 }}>重要ポイント</div>
                            {parsedPrior.key_points.map((pt, i) => <div key={i} style={{ display: "flex", gap: 6, fontSize: 12, color: "#166534", marginBottom: 2 }}><span>•</span><span>{safeStr(pt)}</span></div>)}
                          </div>
                        )}
                      </div>
                    )}

                    {confirmedResults.length > 0 && (
                      <div style={{ marginBottom: 16 }}>
                        {confirmedResults.map(id => {
                          const item = ALL_ITEMS.find(i => i.id === id);
                          return item ? <ResultCard key={id} item={item} result={results[id]} onRetry={() => setRetryModal({ itemId: id, item })} retrying={retryingIds.has(id)} /> : null;
                        })}
                      </div>
                    )}

                    {unconfirmedResults.length > 0 && (
                      <div style={{ background: C.redBg, border: `1px solid ${C.redBorder}`, borderRadius: 10, padding: "12px 16px" }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                          <div style={{ fontSize: 13, fontWeight: 600, color: C.red }}>⚠️ 確認できなかった情報（{unconfirmedResults.length}件）</div>
                          <button onClick={handleBulkRetry} disabled={bulkRetrying || retryingIds.size > 0}
                            style={{ padding: "5px 12px", fontSize: 11, fontWeight: 600, fontFamily: "inherit", background: (bulkRetrying || retryingIds.size > 0) ? C.gray100 : C.white, color: (bulkRetrying || retryingIds.size > 0) ? C.gray400 : C.red, border: `1px solid ${(bulkRetrying || retryingIds.size > 0) ? C.gray200 : C.red}`, borderRadius: 6, cursor: (bulkRetrying || retryingIds.size > 0) ? "not-allowed" : "pointer", whiteSpace: "nowrap" }}>
                            {bulkRetrying ? "再調査中..." : "⟳ まとめて再調査"}
                          </button>
                        </div>
                        {unconfirmedResults.map(id => {
                          const item = ALL_ITEMS.find(i => i.id === id);
                          const r = results[id];
                          return item ? (
                            <div key={id} style={{ display: "flex", gap: 8, fontSize: 12, marginBottom: 6, alignItems: "flex-start" }}>
                              <span style={{ color: C.red, flexShrink: 0 }}>•</span>
                              <div style={{ flex: 1 }}>
                                <span style={{ fontWeight: 600, color: C.black }}>{item.label}</span>
                                {r?.missing && <span style={{ color: C.gray600, marginLeft: 8 }}>{safeStr(r.missing)}</span>}
                              </div>
                              <button onClick={() => setRetryModal({ itemId: id, item })} disabled={retryingIds.has(id) || bulkRetrying}
                                style={{ padding: "2px 8px", fontSize: 10, fontWeight: 600, fontFamily: "inherit", background: C.white, color: retryingIds.has(id) ? C.gray400 : C.red, border: `1px solid ${retryingIds.has(id) ? C.gray200 : C.red}`, borderRadius: 5, cursor: (retryingIds.has(id) || bulkRetrying) ? "not-allowed" : "pointer", whiteSpace: "nowrap", flexShrink: 0 }}>
                                {retryingIds.has(id) ? "調査中..." : "🔍 再調査"}
                              </button>
                            </div>
                          ) : null;
                        })}
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div style={{ flex: 1, overflow: "hidden" }}>
                  <SlidePreview company={company} dept={dept} results={results} pptStyle={pptStyle} selectedIds={selected} />
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
