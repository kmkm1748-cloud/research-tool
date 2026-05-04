import { useState, useRef, useEffect } from "react";
import PptxGenJS from "pptxgenjs";

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
      { id: "market_share", label: "業界シェア・市場ポジション", desc: "業界内順位・シェア推移" },
    ],
  },
  {
    id: "finance", label: "財務", color: "#b45309",
    items: [
      { id: "pl_summary",       label: "売上・利益サマリー",  desc: "直近3期の売上・営業利益・率" },
      { id: "growth_profit",    label: "成長性・収益性評価",  desc: "CAGR・利益率トレンド・評価" },
      { id: "financial_health", label: "財務健全性",          desc: "自己資本比率・有利子負債・格付" },
      { id: "capex",            label: "投資・設備動向",      desc: "CAPEX・設備投資方針・DX予算" },
    ],
  },
  {
    id: "strategy", label: "戦略・方針", color: "#7c3aed",
    items: [
      { id: "mid_term_plan",  label: "中期経営計画",         desc: "KPI・重点施策・投資計画" },
      { id: "dx_strategy",    label: "DX・自動化戦略",       desc: "IT投資方針・システム化動向" },
      { id: "sustainability", label: "サステナビリティ方針", desc: "ESG・CO2削減・物流環境対応" },
    ],
  },
  {
    id: "logistics", label: "物流特性", color: "#be185d",
    items: [
      { id: "logistics_flow",   label: "物流フロー概要",   desc: "入荷〜保管〜出荷の基本フロー" },
      { id: "product_features", label: "取扱品の特性",     desc: "温度帯・サイズ・危険物・重量" },
      { id: "existing_systems", label: "既存システム環境", desc: "WMS・TMS・ERP・連携状況" },
    ],
  },
  {
    id: "contact", label: "担当者・組織", color: "#0369a1",
    items: [
      { id: "org_structure", label: "部署・組織構成",           desc: "物流系部署の階層・組織図" },
      { id: "recent_news",   label: "直近のニュース・トピック", desc: "IR・プレスリリース・受賞歴" },
    ],
  },
];

const ALL_ITEMS = CATEGORIES.flatMap(c => c.items.map(i => ({ ...i, categoryColor: c.color, categoryLabel: c.label })));

const PRESETS = [
  { label: "初回訪問前", ids: new Set(["company_overview","business_products","pl_summary","mid_term_plan","logistics_flow","recent_news"]) },
  { label: "財務・IR重点", ids: new Set(["pl_summary","growth_profit","financial_health","capex","mid_term_plan"]) },
  { label: "物流特性重点", ids: new Set(["logistics_flow","product_features","existing_systems","org_structure"]) },
];

const PPT_STYLES = [
  {
    id: "executive",
    label: "概要把握",
    desc: "重要ポイントを凝縮・数値ファースト",
    icon: "⚡",
    slides: 5,
  },
  {
    id: "sales",
    label: "営業準備",
    desc: "ヒアリング項目・提案ポイントを前面に",
    icon: "🤝",
    slides: 8,
  },
  {
    id: "report",
    label: "社内共有",
    desc: "根拠データ・ソースURL込みで網羅的に",
    icon: "📋",
    slides: 12,
  },
];

const INDUSTRY_OPTIONS = ["食品・飲料・冷凍食品","小売・EC・通販","医薬品・医療機器","化学・素材・危険物","自動車・輸送機器","アパレル・雑貨","電子部品・精密機器","その他"];

const STATUS = {
  confirmed:   { label: "確認済み",   bg: "#dcfce7", color: "#15803d", border: "#bbf7d0" },
  partial:     { label: "一部のみ",   bg: "#fef9c3", color: "#a16207", border: "#fde68a" },
  unconfirmed: { label: "確認できず", bg: "#fee2e2", color: "#dc2626", border: "#fecaca" },
};

// SBRカラー
const SBR = {
  navy: "#0D1B2A",
  cyan: "#00C8FF",
  lightGray: "#F0F4F8",
  white: "#FFFFFF",
  darkNavy: "#1a2744",
};

const FACT_ONLY = "公開情報・報道・公式サイトで確認できた事実のみ記載。推測・類推・一般論は禁止。不明な場合はstatusをunconfirmedにして空白にする。";

function guessDomain(company) {
  const cleaned = company.replace(/株式会社|有限会社|合同会社|ホールディングス|グループ|HD|Holdings/gi, "").replace(/[　\s]/g, "").toLowerCase();
  const map = {
    "subaru": "subaru.co.jp", "スバル": "subaru.co.jp",
    "toyota": "toyota.co.jp", "トヨタ": "toyota.co.jp",
    "honda": "honda.co.jp", "ホンダ": "honda.co.jp",
    "nissan": "nissan.co.jp", "日産": "nissan.co.jp",
    "nichirei": "nichirei.co.jp", "ニチレイ": "nichirei.co.jp",
    "yamato": "yamato-hd.co.jp", "ヤマト": "yamato-hd.co.jp",
    "sagawa": "sagawa-exp.co.jp", "佐川": "sagawa-exp.co.jp",
    "nipponexpress": "nipponexpress.com", "日本通運": "nipponexpress.com",
    "hitachi": "hitachi.co.jp", "日立": "hitachi.co.jp",
    "fujitsu": "fujitsu.com", "富士通": "fujitsu.com",
    "sony": "sony.co.jp", "ソニー": "sony.co.jp",
    "panasonic": "panasonic.com", "パナソニック": "panasonic.com",
  };
  for (const [key, domain] of Object.entries(map)) {
    if (cleaned.includes(key.toLowerCase())) return domain;
  }
  return null;
}

const URL_PATTERNS = {
  company_overview:  (d) => [`https://${d}`, `https://${d}/company/`, `https://${d}/corporate/`],
  business_products: (d) => [`https://${d}/products/`, `https://${d}/services/`, `https://${d}/business/`],
  group_structure:   (d) => [`https://${d}/corporate/group/`, `https://${d}/company/group/`],
  bases_network:     (d) => [`https://${d}/corporate/network/`, `https://${d}/company/bases/`],
  market_share:      (d) => [`https://${d}/ir/`, `https://${d}/investor/`],
  pl_summary:        (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/result/`, `https://${d}/ir/`],
  growth_profit:     (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/result/`],
  financial_health:  (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/`],
  capex:             (d) => [`https://${d}/ir/finance/`, `https://${d}/ir/strategy/`],
  mid_term_plan:     (d) => [`https://${d}/ir/strategy/`, `https://${d}/ir/management/`],
  dx_strategy:       (d) => [`https://${d}/corporate/dx/`, `https://${d}/news/`],
  sustainability:    (d) => [`https://${d}/sustainability/`, `https://${d}/csr/`],
  logistics_flow:    (d) => [`https://${d}/corporate/logistics/`, `https://${d}/business/logistics/`],
  product_features:  (d) => [`https://${d}/products/`, `https://${d}/business/products/`],
  existing_systems:  (d) => [`https://${d}/corporate/dx/`, `https://${d}/ir/strategy/`],
  org_structure:     (d) => [`https://${d}/corporate/organization/`, `https://${d}/company/organization/`],
  recent_news:       (d) => [`https://${d}/news/`, `https://${d}/press/`, `https://${d}/ir/news/`],
};

const SEARCH_QUERIES = {
  company_overview:  (c) => [`${c} 会社概要 設立 従業員数 本社`, `${c} 企業情報`],
  business_products: (c) => [`${c} 事業内容 主力商品 サービス`, `${c} 事業領域`],
  group_structure:   (c) => [`${c} グループ会社 子会社 組織`, `${c} グループ構成`],
  bases_network:     (c) => [`${c} 拠点 物流センター 倉庫 所在地`, `${c} 配送網`],
  market_share:      (c, ind) => [`${c} ${ind||""}業界シェア 市場ポジション`, `${c} 業界順位`],
  pl_summary:        (c) => [`${c} 決算 売上高 営業利益 2024 2025`, `${c} 業績 財務ハイライト`],
  growth_profit:     (c) => [`${c} 業績推移 成長率 収益性`, `${c} CAGR 利益率`],
  financial_health:  (c) => [`${c} 自己資本比率 有利子負債 財務`, `${c} 財務健全性`],
  capex:             (c) => [`${c} 設備投資 CAPEX 投資計画 2024 2025`, `${c} DX投資`],
  mid_term_plan:     (c) => [`${c} 中期経営計画 2025 2026 KPI`, `${c} 経営戦略`],
  dx_strategy:       (c) => [`${c} DX デジタル化 IT投資 自動化 2024 2025`, `${c} デジタル戦略`],
  sustainability:    (c) => [`${c} ESG CO2削減 サステナビリティ`, `${c} カーボンニュートラル`],
  logistics_flow:    (c) => [`${c} 物流フロー 入荷 出荷 保管`, `${c} サプライチェーン 物流`],
  product_features:  (c) => [`${c} 取扱品 商品特性 温度管理`, `${c} 商品カテゴリ`],
  existing_systems:  (c) => [`${c} WMS TMS ERP システム`, `${c} 基幹システム IT`],
  org_structure:     (c) => [`${c} 組織図 物流部門 SCM`, `${c} 部署構成`],
  recent_news:       (c) => [`${c} プレスリリース ニュース 2025`, `${c} IR 新着 2024 2025`],
};

const PROMPTS = {
  company_overview:  (c) => `"${c}"の会社概要（設立・本社・従業員数・上場区分）。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  business_products: (c) => `"${c}"の事業領域・主力商材。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  group_structure:   (c) => `"${c}"のグループ構造。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  bases_network:     (c) => `"${c}"の拠点・物流ネットワーク。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  market_share:      (c, ind) => `"${c}"の${ind?`${ind}業界における`:""}業界シェア・市場ポジション。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  pl_summary:        (c) => `"${c}"の直近3期の売上・営業利益・率。${FACT_ONLY} JSONのみ: {"summary":"3文以内","revenue":"売上高","growth":"成長率","margin":"営業利益率","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  growth_profit:     (c) => `"${c}"の成長性・収益性トレンド。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  financial_health:  (c) => `"${c}"の財務健全性。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  capex:             (c) => `"${c}"の設備投資・CAPEX・DX予算動向。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  mid_term_plan:     (c) => `"${c}"の中期経営計画（KPI・重点施策）。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  dx_strategy:       (c) => `"${c}"のDX・自動化戦略。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  sustainability:    (c) => `"${c}"のESG・CO2削減・サステナビリティ方針。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  logistics_flow:    (c) => `"${c}"の物流フロー概要。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  product_features:  (c) => `"${c}"が扱う商品の特性。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  existing_systems:  (c) => `"${c}"が使用するWMS・TMS・ERP等の既存システム。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  org_structure:     (c) => `"${c}"の物流・DX系部署の組織構成。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
  recent_news:       (c) => `"${c}"の直近1年のニュース・IR・プレスリリース。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`,
};

async function callClaudeNoSearch(system, user) {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-haiku-4-5-20251001", max_tokens: 1500, system, messages: [{ role: "user", content: user }] }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.content?.map(b => b.text || "").filter(Boolean).join("\n") || "";
}

async function callClaude(system, user) {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-haiku-4-5-20251001", max_tokens: 1500, system, messages: [{ role: "user", content: user }], tools: [{ type: "web_search_20250305", name: "web_search" }] }),
  });
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
    const search1Raw = await callClaude(`あなたは企業調査AIです。検索して最も信頼性の高い公式情報ページのURLを特定し、その内容を要約してください。`, `「${queries[0]}」で検索し、最も信頼性の高い公式情報を取得してください。\nURL: [参照したURL]\n内容: [取得した情報の要約（500文字以内）]`);
    if (search1Raw.length > 50) {
      collectedTexts.push(`【Web検索1回目】\n${search1Raw.slice(0, 2000)}`);
      const urlMatch = search1Raw.match(/URL:\s*(https?:\/\/[^\s\n]+)/);
      if (urlMatch) { collectedUrls.push(urlMatch[1]); const { text: deepText } = await fetchUrl(urlMatch[1]); if (deepText.length > 300) collectedTexts.push(`【${urlMatch[1]} 詳細】\n${deepText.slice(0, 3000)}`); }
    }
  } catch (_) {}
  await new Promise(r => setTimeout(r, 5000));
  if (queries.length > 1) {
    try {
      const search2Raw = await callClaude(`あなたは企業調査AIです。検索して補完情報・最新データを取得してください。`, `「${queries[1]}」で検索し、最新の情報・数値データを取得してください。\nURL: [参照したURL]\n内容: [取得した情報の要約（500文字以内）]`);
      if (search2Raw.length > 50) { collectedTexts.push(`【Web検索2回目】\n${search2Raw.slice(0, 2000)}`); const urlMatch2 = search2Raw.match(/URL:\s*(https?:\/\/[^\s\n]+)/); if (urlMatch2) collectedUrls.push(urlMatch2[1]); }
    } catch (_) {}
  }
  await new Promise(r => setTimeout(r, 2000));
  const allInfo = collectedTexts.join("\n\n---\n\n");
  const sourceUrls = [...new Set(collectedUrls)].slice(0, 3).join(", ");
  const finalPrompt = allInfo.length > 200 ? `${prompt}\n\n以下は複数の公式情報源から収集した情報です。これらを総合して、事実のみを記載してください。source_urlには「${sourceUrls}」を記載してください：\n\n${allInfo.slice(0, 8000)}` : `${prompt}\n\nsource_urlには参照したURLを記載してください。情報が見つからない場合はstatusをunconfirmedにしてください。`;
  return await callClaudeNoSearch(`あなたはB2B営業支援AIです。収集した複数の公式情報源を総合分析し、JSONのみ返答。【厳守ルール】- 収集した公式情報に記載された事実のみ記載する - 推測・類推・一般論は一切書かない - 情報が見つからない場合はstatusをunconfirmedにし「公開情報なし」と記載 - missingには「訪問時に直接確認すべき具体的な質問」を書く - source_urlには実際に参照したURLを記載する - 前置き・マークダウン不要。JSONのみ返答。`, finalPrompt);
}

// PPT生成関数
async function generatePPT(company, dept, results, pptStyle, selectedIds) {
  const prs = new PptxGenJS();
  prs.layout = "LAYOUT_WIDE";
  prs.author = "SBR 顧客リサーチツール";

  const style = PPT_STYLES.find(s => s.id === pptStyle) || PPT_STYLES[0];
  const today = new Date().toLocaleDateString("ja-JP");

  const addHeader = (slide, title) => {
    slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 1.2, fill: { color: "0D1B2A" } });
    slide.addShape(prs.ShapeType.rect, { x: 0, y: 1.2, w: "100%", h: 0.06, fill: { color: "00C8FF" } });
    slide.addText("SBR 顧客リサーチ", { x: 0.3, y: 0.1, w: 3, h: 0.5, fontSize: 11, color: "00C8FF", bold: false });
    slide.addText(title, { x: 0.3, y: 0.5, w: 9, h: 0.6, fontSize: 22, color: "FFFFFF", bold: true });
    slide.addText(company, { x: 7, y: 0.1, w: 2.7, h: 0.5, fontSize: 11, color: "FFFFFF", align: "right" });
  };

  const addFooter = (slide, pageNum, total) => {
    slide.addShape(prs.ShapeType.rect, { x: 0, y: 6.8, w: "100%", h: 0.6, fill: { color: "0D1B2A" } });
    slide.addText(`${today}　${dept ? `担当部署：${dept}` : ""}`, { x: 0.3, y: 6.85, w: 6, h: 0.4, fontSize: 10, color: "94a3b8" });
    slide.addText(`${pageNum} / ${total}`, { x: 8.5, y: 6.85, w: 1.2, h: 0.4, fontSize: 10, color: "94a3b8", align: "right" });
  };

  const getResult = (id) => results[id];
  const statusColor = (st) => st === "confirmed" ? "15803d" : st === "partial" ? "a16207" : "dc2626";
  const statusLabel = (st) => st === "confirmed" ? "✓ 確認済み" : st === "partial" ? "△ 一部のみ" : "⚠ 確認できず";

  // スライド構成をスタイルで分岐
  const slideConfigs = {
    executive: ["cover", "summary", "finance", "strategy", "hearing"],
    sales: ["cover", "company", "finance", "strategy", "logistics", "org", "proposal", "hearing"],
    report: ["cover", "summary", "company", "market", "finance", "strategy", "logistics", "org", "news", "hearing", "sources"],
  };

  const slides = slideConfigs[pptStyle] || slideConfigs.executive;
  const total = slides.length;

  slides.forEach((slideType, idx) => {
    const slide = prs.addSlide();
    const pageNum = idx + 1;
    slide.background = { color: "F0F4F8" };

    if (slideType === "cover") {
      // 表紙
      slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: "100%", h: "100%", fill: { color: "0D1B2A" } });
      slide.addShape(prs.ShapeType.rect, { x: 0, y: 3.8, w: "100%", h: 0.08, fill: { color: "00C8FF" } });
      slide.addText("顧客リサーチレポート", { x: 0.8, y: 1.2, w: 8, h: 0.8, fontSize: 18, color: "00C8FF", bold: false });
      slide.addText(company, { x: 0.8, y: 2.0, w: 8, h: 1.4, fontSize: 40, color: "FFFFFF", bold: true });
      if (dept) slide.addText(`担当部署：${dept}`, { x: 0.8, y: 3.5, w: 6, h: 0.4, fontSize: 14, color: "94a3b8" });
      slide.addText(`調査日：${today}`, { x: 0.8, y: 4.2, w: 4, h: 0.4, fontSize: 13, color: "94a3b8" });
      slide.addText(`スタイル：${style.label}（${style.slides}枚構成）`, { x: 0.8, y: 4.7, w: 6, h: 0.4, fontSize: 12, color: "64748b" });
      slide.addText("SoftBank Robotics Corp.", { x: 0.8, y: 6.3, w: 6, h: 0.4, fontSize: 11, color: "475569" });

    } else if (slideType === "summary") {
      addHeader(slide, "エグゼクティブサマリー");
      addFooter(slide, pageNum, total);
      const summaryItems = ["company_overview", "pl_summary", "mid_term_plan", "recent_news"].filter(id => results[id]);
      summaryItems.forEach((id, i) => {
        const r = results[id];
        const item = ALL_ITEMS.find(it => it.id === id);
        const x = i % 2 === 0 ? 0.3 : 5.1;
        const y = 1.5 + Math.floor(i / 2) * 2.4;
        slide.addShape(prs.ShapeType.rect, { x, y, w: 4.5, h: 2.2, fill: { color: "FFFFFF" }, line: { color: "e2e8f0", width: 1 } });
        slide.addShape(prs.ShapeType.rect, { x, y, w: 4.5, h: 0.45, fill: { color: "1a2744" } });
        slide.addText(item?.label || id, { x: x+0.1, y: y+0.05, w: 3.5, h: 0.35, fontSize: 11, color: "FFFFFF", bold: true });
        slide.addText(statusLabel(r?.status), { x: x+3.1, y: y+0.05, w: 1.3, h: 0.35, fontSize: 9, color: statusColor(r?.status), align: "right" });
        slide.addText(r?.summary || "情報なし", { x: x+0.15, y: y+0.55, w: 4.2, h: 1.55, fontSize: 10, color: "374151", valign: "top", wrap: true });
      });

    } else if (slideType === "company") {
      addHeader(slide, "企業基本情報");
      addFooter(slide, pageNum, total);
      const r1 = getResult("company_overview"), r2 = getResult("business_products");
      slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 1.5, w: 9.1, h: 0.4, fill: { color: "1a2744" } });
      slide.addText("会社概要", { x: 0.4, y: 1.55, w: 4, h: 0.3, fontSize: 12, color: "00C8FF", bold: true });
      slide.addText(r1?.summary || "情報なし", { x: 0.4, y: 2.0, w: 9, h: 1.0, fontSize: 11, color: "1f2937", wrap: true });
      slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 3.2, w: 9.1, h: 0.4, fill: { color: "1a2744" } });
      slide.addText("事業領域・主力商材", { x: 0.4, y: 3.25, w: 4, h: 0.3, fontSize: 12, color: "00C8FF", bold: true });
      slide.addText(r2?.summary || "情報なし", { x: 0.4, y: 3.7, w: 9, h: 1.2, fontSize: 11, color: "1f2937", wrap: true });

    } else if (slideType === "finance") {
      addHeader(slide, "財務ハイライト");
      addFooter(slide, pageNum, total);
      const r = getResult("pl_summary");
      const rg = getResult("growth_profit");
      // 財務KPIカード
      const kpis = [
        { label: "売上高", value: r?.revenue || "—" },
        { label: "成長率", value: r?.growth || rg?.summary?.match(/[\d.]+%/)?.[0] || "—" },
        { label: "営業利益率", value: r?.margin || "—" },
      ];
      kpis.forEach((kpi, i) => {
        const x = 0.3 + i * 3.1;
        slide.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.8, h: 1.4, fill: { color: "0D1B2A" } });
        slide.addShape(prs.ShapeType.rect, { x, y: 1.5, w: 2.8, h: 0.06, fill: { color: "00C8FF" } });
        slide.addText(kpi.label, { x: x+0.1, y: 1.6, w: 2.6, h: 0.4, fontSize: 11, color: "94a3b8" });
        slide.addText(kpi.value, { x: x+0.1, y: 2.0, w: 2.6, h: 0.7, fontSize: 20, color: "00C8FF", bold: true });
      });
      slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 3.1, w: 9.1, h: 0.35, fill: { color: "1a2744" } });
      slide.addText("財務サマリー", { x: 0.4, y: 3.15, w: 4, h: 0.25, fontSize: 11, color: "00C8FF", bold: true });
      slide.addText(r?.summary || "情報なし", { x: 0.4, y: 3.55, w: 9, h: 1.2, fontSize: 10, color: "1f2937", wrap: true });
      if (rg?.summary) {
        slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 4.9, w: 9.1, h: 0.35, fill: { color: "1a2744" } });
        slide.addText("成長性・収益性", { x: 0.4, y: 4.95, w: 4, h: 0.25, fontSize: 11, color: "00C8FF", bold: true });
        slide.addText(rg.summary, { x: 0.4, y: 5.35, w: 9, h: 1.0, fontSize: 10, color: "1f2937", wrap: true });
      }

    } else if (slideType === "strategy") {
      addHeader(slide, "戦略・方針");
      addFooter(slide, pageNum, total);
      const rs = [
        { id: "mid_term_plan", label: "中期経営計画" },
        { id: "dx_strategy", label: "DX・自動化戦略" },
        { id: "sustainability", label: "サステナビリティ" },
      ];
      rs.forEach((item, i) => {
        const r = getResult(item.id);
        const y = 1.5 + i * 1.7;
        slide.addShape(prs.ShapeType.rect, { x: 0.3, y, w: 0.08, h: 1.4, fill: { color: "00C8FF" } });
        slide.addText(item.label, { x: 0.55, y: y+0.05, w: 3, h: 0.35, fontSize: 12, color: "0D1B2A", bold: true });
        slide.addText(statusLabel(r?.status || "unconfirmed"), { x: 7.5, y: y+0.05, w: 2, h: 0.35, fontSize: 10, color: statusColor(r?.status || "unconfirmed"), align: "right" });
        slide.addText(r?.summary || "情報なし", { x: 0.55, y: y+0.45, w: 9, h: 0.9, fontSize: 10, color: "374151", wrap: true });
      });

    } else if (slideType === "logistics") {
      addHeader(slide, "物流特性");
      addFooter(slide, pageNum, total);
      const items = [
        { id: "logistics_flow", label: "物流フロー" },
        { id: "product_features", label: "取扱品の特性" },
        { id: "existing_systems", label: "既存システム環境" },
        { id: "bases_network", label: "拠点・ネットワーク" },
      ];
      items.forEach((item, i) => {
        const r = getResult(item.id);
        const x = i % 2 === 0 ? 0.3 : 5.1;
        const y = 1.5 + Math.floor(i / 2) * 2.5;
        slide.addShape(prs.ShapeType.rect, { x, y, w: 4.5, h: 2.3, fill: { color: "FFFFFF" }, line: { color: "e2e8f0", width: 1 } });
        slide.addShape(prs.ShapeType.rect, { x, y, w: 4.5, h: 0.45, fill: { color: "be185d" } });
        slide.addText(item.label, { x: x+0.1, y: y+0.07, w: 4.2, h: 0.32, fontSize: 11, color: "FFFFFF", bold: true });
        slide.addText(r?.summary || "情報なし", { x: x+0.15, y: y+0.55, w: 4.2, h: 1.6, fontSize: 10, color: "374151", wrap: true });
      });

    } else if (slideType === "org") {
      addHeader(slide, "組織・担当者");
      addFooter(slide, pageNum, total);
      const r = getResult("org_structure");
      slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 1.5, w: 9.1, h: 0.4, fill: { color: "1a2744" } });
      slide.addText("部署・組織構成", { x: 0.4, y: 1.55, w: 4, h: 0.3, fontSize: 12, color: "00C8FF", bold: true });
      slide.addText(r?.summary || "情報なし", { x: 0.4, y: 2.0, w: 9, h: 1.5, fontSize: 11, color: "1f2937", wrap: true });

    } else if (slideType === "news") {
      addHeader(slide, "直近のニュース・トピック");
      addFooter(slide, pageNum, total);
      const r = getResult("recent_news");
      slide.addText(r?.summary || "情報なし", { x: 0.4, y: 1.6, w: 9, h: 2.5, fontSize: 12, color: "1f2937", wrap: true });
      if (r?.source_url) slide.addText(`参照：${r.source_url}`, { x: 0.4, y: 5.5, w: 9, h: 0.4, fontSize: 9, color: "6b7280" });

    } else if (slideType === "proposal") {
      addHeader(slide, "提案ポイント");
      addFooter(slide, pageNum, total);
      // 確認できた情報から提案ポイントを整理
      const confirmedItems = ALL_ITEMS.filter(i => selectedIds.has(i.id) && results[i.id]?.status === "confirmed");
      const unconfirmedItems = ALL_ITEMS.filter(i => selectedIds.has(i.id) && results[i.id]?.status === "unconfirmed");
      slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 1.5, w: 4.4, h: 0.4, fill: { color: "15803d" } });
      slide.addText("✓ 確認済み情報", { x: 0.4, y: 1.55, w: 4, h: 0.3, fontSize: 12, color: "FFFFFF", bold: true });
      confirmedItems.slice(0, 5).forEach((item, i) => {
        slide.addText(`• ${item.label}：${results[item.id]?.summary?.slice(0, 40) || ""}…`, { x: 0.4, y: 2.05 + i * 0.5, w: 4.2, h: 0.4, fontSize: 10, color: "1f2937", wrap: true });
      });
      slide.addShape(prs.ShapeType.rect, { x: 5.1, y: 1.5, w: 4.4, h: 0.4, fill: { color: "dc2626" } });
      slide.addText("⚠ ヒアリング必須項目", { x: 5.2, y: 1.55, w: 4, h: 0.3, fontSize: 12, color: "FFFFFF", bold: true });
      unconfirmedItems.slice(0, 5).forEach((item, i) => {
        slide.addText(`• ${item.label}：${results[item.id]?.missing?.slice(0, 40) || "訪問時に確認"}…`, { x: 5.2, y: 2.05 + i * 0.5, w: 4.2, h: 0.4, fontSize: 10, color: "1f2937", wrap: true });
      });

    } else if (slideType === "hearing") {
      addHeader(slide, "ヒアリングシート");
      addFooter(slide, pageNum, total);
      const hearingItems = ALL_ITEMS.filter(i => selectedIds.has(i.id) && results[i.id]?.status === "unconfirmed");
      const partialItems = ALL_ITEMS.filter(i => selectedIds.has(i.id) && results[i.id]?.status === "partial");
      slide.addText(`要確認事項（${hearingItems.length + partialItems.length}件）`, { x: 0.3, y: 1.4, w: 9, h: 0.4, fontSize: 14, color: "0D1B2A", bold: true });
      let yPos = 1.9;
      [...hearingItems, ...partialItems].slice(0, 8).forEach((item) => {
        const r = results[item.id];
        const isUnconfirmed = r?.status === "unconfirmed";
        slide.addShape(prs.ShapeType.rect, { x: 0.3, y: yPos, w: 0.06, h: 0.55, fill: { color: isUnconfirmed ? "dc2626" : "a16207" } });
        slide.addText(item.label, { x: 0.5, y: yPos+0.02, w: 2.5, h: 0.25, fontSize: 10, color: "0D1B2A", bold: true });
        slide.addText(r?.missing || "訪問時に確認", { x: 0.5, y: yPos+0.27, w: 9, h: 0.25, fontSize: 9, color: "4b5563" });
        yPos += 0.65;
      });

    } else if (slideType === "market") {
      addHeader(slide, "市場・競合");
      addFooter(slide, pageNum, total);
      const r = getResult("market_share");
      slide.addShape(prs.ShapeType.rect, { x: 0.3, y: 1.5, w: 9.1, h: 0.4, fill: { color: "0f766e" } });
      slide.addText("業界シェア・市場ポジション", { x: 0.4, y: 1.55, w: 6, h: 0.3, fontSize: 12, color: "FFFFFF", bold: true });
      slide.addText(r?.summary || "情報なし", { x: 0.4, y: 2.05, w: 9, h: 2.0, fontSize: 11, color: "1f2937", wrap: true });

    } else if (slideType === "sources") {
      addHeader(slide, "情報ソース一覧");
      addFooter(slide, pageNum, total);
      const sourceItems = ALL_ITEMS.filter(i => selectedIds.has(i.id) && results[i.id]?.source_url);
      slide.addText("本レポートの参照元URL", { x: 0.3, y: 1.4, w: 9, h: 0.4, fontSize: 13, color: "0D1B2A", bold: true });
      sourceItems.slice(0, 10).forEach((item, i) => {
        const r = results[item.id];
        slide.addText(`${item.label}：${r.source_url}`, { x: 0.4, y: 1.95 + i * 0.48, w: 9, h: 0.4, fontSize: 9, color: "374151", wrap: true });
      });
    }
  });

  await prs.writeFile({ fileName: `SBR_リサーチ_${company}_${today}.pptx` });
}

// スライドプレビューコンポーネント
function SlidePreview({ company, dept, results, pptStyle, selectedIds }) {
  const [currentSlide, setCurrentSlide] = useState(0);
  const [generating, setGenerating] = useState(false);

  const style = PPT_STYLES.find(s => s.id === pptStyle) || PPT_STYLES[0];
  const slideConfigs = {
    executive: ["cover", "summary", "finance", "strategy", "hearing"],
    sales: ["cover", "company", "finance", "strategy", "logistics", "org", "proposal", "hearing"],
    report: ["cover", "summary", "company", "market", "finance", "strategy", "logistics", "org", "news", "hearing", "sources"],
  };
  const slides = slideConfigs[pptStyle] || slideConfigs.executive;
  const total = slides.length;
  const slideType = slides[currentSlide];

  const getResult = (id) => results[id];
  const statusColor = (st) => st === "confirmed" ? "#15803d" : st === "partial" ? "#a16207" : "#dc2626";
  const statusLabel = (st) => st === "confirmed" ? "✓ 確認済み" : st === "partial" ? "△ 一部のみ" : "⚠ 要確認";

  const renderSlide = () => {
    const headerStyle = { background: SBR.navy, padding: "10px 16px", borderBottom: `3px solid ${SBR.cyan}` };
    const headerText = { fontSize: 16, fontWeight: 700, color: SBR.white };
    const subText = { fontSize: 10, color: SBR.cyan };
    const bodyStyle = { padding: "16px", flex: 1, overflow: "hidden" };

    if (slideType === "cover") return (
      <div style={{ background: SBR.navy, height: "100%", display: "flex", flexDirection: "column", justifyContent: "center", padding: "32px" }}>
        <div style={{ fontSize: 12, color: SBR.cyan, marginBottom: 8 }}>顧客リサーチレポート</div>
        <div style={{ fontSize: 32, fontWeight: 700, color: SBR.white, marginBottom: 8 }}>{company}</div>
        <div style={{ width: 60, height: 3, background: SBR.cyan, marginBottom: 16 }} />
        {dept && <div style={{ fontSize: 13, color: "#94a3b8", marginBottom: 6 }}>担当部署：{dept}</div>}
        <div style={{ fontSize: 12, color: "#64748b" }}>調査日：{new Date().toLocaleDateString("ja-JP")}</div>
        <div style={{ fontSize: 11, color: "#475569", marginTop: 8 }}>{style.label}（{slides.length}枚構成）</div>
      </div>
    );

    if (slideType === "summary") return (
      <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
        <div style={headerStyle}><div style={subText}>Executive Summary</div><div style={headerText}>エグゼクティブサマリー</div></div>
        <div style={{ ...bodyStyle, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          {["company_overview","pl_summary","mid_term_plan","recent_news"].filter(id => results[id]).map(id => {
            const r = results[id]; const item = ALL_ITEMS.find(i => i.id === id);
            return (
              <div key={id} style={{ background: "#fff", borderRadius: 6, overflow: "hidden", border: "1px solid #e2e8f0" }}>
                <div style={{ background: SBR.darkNavy, padding: "4px 8px", display: "flex", justifyContent: "space-between" }}>
                  <span style={{ fontSize: 10, color: SBR.white, fontWeight: 600 }}>{item?.label}</span>
                  <span style={{ fontSize: 9, color: statusColor(r?.status) }}>{statusLabel(r?.status)}</span>
                </div>
                <div style={{ padding: "6px 8px", fontSize: 10, color: "#374151", lineHeight: 1.5 }}>{r?.summary?.slice(0, 80)}...</div>
              </div>
            );
          })}
        </div>
      </div>
    );

    if (slideType === "finance") return (
      <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
        <div style={headerStyle}><div style={subText}>Financial Highlights</div><div style={headerText}>財務ハイライト</div></div>
        <div style={bodyStyle}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 12 }}>
            {[{ label: "売上高", value: getResult("pl_summary")?.revenue || "—" }, { label: "成長率", value: getResult("pl_summary")?.growth || "—" }, { label: "営業利益率", value: getResult("pl_summary")?.margin || "—" }].map((kpi, i) => (
              <div key={i} style={{ background: SBR.navy, borderRadius: 6, padding: "10px", borderTop: `3px solid ${SBR.cyan}` }}>
                <div style={{ fontSize: 10, color: "#94a3b8", marginBottom: 4 }}>{kpi.label}</div>
                <div style={{ fontSize: 18, fontWeight: 700, color: SBR.cyan }}>{kpi.value}</div>
              </div>
            ))}
          </div>
          <div style={{ background: "#fff", borderRadius: 6, padding: 10, border: "1px solid #e2e8f0" }}>
            <div style={{ fontSize: 10, fontWeight: 600, color: SBR.darkNavy, marginBottom: 4 }}>財務サマリー</div>
            <div style={{ fontSize: 10, color: "#374151", lineHeight: 1.6 }}>{getResult("pl_summary")?.summary || "情報なし"}</div>
          </div>
        </div>
      </div>
    );

    if (slideType === "strategy") return (
      <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
        <div style={headerStyle}><div style={subText}>Strategy & Policy</div><div style={headerText}>戦略・方針</div></div>
        <div style={bodyStyle}>
          {[{ id: "mid_term_plan", label: "中期経営計画" }, { id: "dx_strategy", label: "DX・自動化戦略" }, { id: "sustainability", label: "サステナビリティ" }].map((item, i) => {
            const r = getResult(item.id);
            return (
              <div key={i} style={{ display: "flex", gap: 8, marginBottom: 10, background: "#fff", borderRadius: 6, overflow: "hidden", border: "1px solid #e2e8f0" }}>
                <div style={{ width: 4, background: SBR.cyan, flexShrink: 0 }} />
                <div style={{ padding: "8px 8px 8px 0" }}>
                  <div style={{ fontSize: 11, fontWeight: 600, color: SBR.darkNavy, marginBottom: 3 }}>{item.label}</div>
                  <div style={{ fontSize: 10, color: "#374151", lineHeight: 1.5 }}>{r?.summary?.slice(0, 100) || "情報なし"}</div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );

    if (slideType === "logistics") return (
      <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
        <div style={headerStyle}><div style={subText}>Logistics Features</div><div style={headerText}>物流特性</div></div>
        <div style={{ ...bodyStyle, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          {[{ id: "logistics_flow", label: "物流フロー" }, { id: "product_features", label: "取扱品の特性" }, { id: "existing_systems", label: "既存システム" }, { id: "bases_network", label: "拠点ネットワーク" }].map((item, i) => {
            const r = getResult(item.id);
            return (
              <div key={i} style={{ background: "#fff", borderRadius: 6, overflow: "hidden", border: "1px solid #e2e8f0" }}>
                <div style={{ background: "#be185d", padding: "4px 8px" }}><span style={{ fontSize: 10, color: "#fff", fontWeight: 600 }}>{item.label}</span></div>
                <div style={{ padding: "6px 8px", fontSize: 10, color: "#374151", lineHeight: 1.5 }}>{r?.summary?.slice(0, 80) || "情報なし"}</div>
              </div>
            );
          })}
        </div>
      </div>
    );

    if (slideType === "hearing") return (
      <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
        <div style={headerStyle}><div style={subText}>Hearing Sheet</div><div style={headerText}>ヒアリングシート</div></div>
        <div style={bodyStyle}>
          {ALL_ITEMS.filter(i => selectedIds.has(i.id) && (results[i.id]?.status === "unconfirmed" || results[i.id]?.status === "partial")).slice(0, 7).map((item, i) => {
            const r = results[item.id];
            return (
              <div key={i} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-start" }}>
                <div style={{ width: 4, height: 36, background: r?.status === "unconfirmed" ? "#dc2626" : "#a16207", borderRadius: 2, flexShrink: 0, marginTop: 2 }} />
                <div>
                  <div style={{ fontSize: 10, fontWeight: 600, color: SBR.darkNavy }}>{item.label}</div>
                  <div style={{ fontSize: 9, color: "#6b7280" }}>{r?.missing || "訪問時に確認"}</div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );

    // デフォルト
    const r = getResult(slides[currentSlide]);
    return (
      <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
        <div style={headerStyle}><div style={headerText}>{slideType}</div></div>
        <div style={bodyStyle}><div style={{ fontSize: 11, color: "#374151" }}>{r?.summary || "情報なし"}</div></div>
      </div>
    );
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      {/* スライドプレビュー */}
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: "16px", background: "#e2e8f0" }}>
        <div style={{ width: "100%", maxWidth: 640, aspectRatio: "16/9", background: SBR.lightGray, borderRadius: 8, boxShadow: "0 4px 20px rgba(0,0,0,0.15)", overflow: "hidden", display: "flex", flexDirection: "column" }}>
          {renderSlide()}
        </div>
      </div>

      {/* ナビゲーション */}
      <div style={{ padding: "12px 20px", background: "#fff", borderTop: "1px solid #e5e7eb", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <button onClick={() => setCurrentSlide(s => Math.max(0, s-1))} disabled={currentSlide === 0}
          style={{ padding: "8px 16px", borderRadius: 6, border: "1px solid #e5e7eb", background: currentSlide === 0 ? "#f9fafb" : "#fff", cursor: currentSlide === 0 ? "not-allowed" : "pointer", fontSize: 13, color: currentSlide === 0 ? "#9ca3af" : "#374151" }}>
          ← 前へ
        </button>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {slides.map((_, i) => (
            <div key={i} onClick={() => setCurrentSlide(i)}
              style={{ width: i === currentSlide ? 20 : 8, height: 8, borderRadius: 99, background: i === currentSlide ? SBR.navy : "#d1d5db", cursor: "pointer", transition: "all 0.2s" }} />
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <span style={{ fontSize: 12, color: "#6b7280" }}>{currentSlide + 1} / {total}</span>
          <button onClick={async () => { setGenerating(true); try { await generatePPT(company, dept, results, pptStyle, selectedIds); } catch(e) { alert("PPT生成エラー: " + e.message); } setGenerating(false); }}
            disabled={generating}
            style={{ padding: "8px 20px", borderRadius: 6, border: "none", background: generating ? "#d1d5db" : SBR.navy, color: "#fff", cursor: generating ? "not-allowed" : "pointer", fontSize: 13, fontWeight: 700, fontFamily: "inherit" }}>
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
  const mins = Math.floor(remaining / 60);
  const secs = remaining % 60;
  const timeStr = mins > 0 ? `約${mins}分` : `約${secs}秒`;
  const [tick, setTick] = useState(0);
  useEffect(() => { const id = setInterval(() => setTick(t => t+1), 100); return () => clearInterval(id); }, []);
  const cycle = tick % 200, CONV_END = 80, ARM_END = 140;
  const boxConvX = cycle < CONV_END ? 15 + (cycle/CONV_END)*100 : 115;
  const boxOnConv = cycle < CONV_END;
  const armPhase = cycle >= CONV_END && cycle < ARM_END ? (cycle-CONV_END)/(ARM_END-CONV_END) : 0;
  const armDown = Math.sin(armPhase*Math.PI);
  const armGripping = armPhase > 0.3 && armPhase < 0.85;
  const boxOnArm = armPhase > 0.35 && armPhase < 0.9;
  const forkX = 30 + pct*155;
  const beltOffset = (tick*1.5)%20;
  return (
    <div style={{ background:"#f8fafc", border:"1px solid #e2e8f0", borderRadius:10, padding:"10px 14px", marginBottom:8 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 }}>
        <span style={{ fontSize:11, color:"#475569", fontWeight:500 }}>{current?`「${current}」調査中...`:done===total&&total>0?"完了！":"準備中"}</span>
        <span style={{ fontSize:11, color:"#94a3b8" }}>{done}/{total}</span>
      </div>
      <svg viewBox="0 0 300 82" width="100%" style={{ display:"block", background:"#f1f5f9", borderRadius:8 }}>
        <rect x="0" y="65" width="300" height="17" fill="#e2e8f0"/><rect x="0" y="63" width="300" height="3" fill="#cbd5e1"/>
        <rect x="8" y="50" width="120" height="14" rx="3" fill="#334155"/>
        <clipPath id="wbc5"><rect x="8" y="50" width="120" height="14"/></clipPath>
        <g clipPath="url(#wbc5)">{Array.from({length:9}).map((_,i)=><line key={i} x1={8+((i*16-beltOffset+120)%120)} y1="50" x2={8+((i*16-beltOffset+120)%120)} y2="64" stroke="#475569" strokeWidth="1.5"/>)}</g>
        <rect x="8" y="50" width="120" height="14" rx="3" fill="none" stroke="#475569" strokeWidth="1"/>
        <circle cx="11" cy="57" r="5" fill="#1e293b" stroke="#475569" strokeWidth="1"/><circle cx="125" cy="57" r="5" fill="#1e293b" stroke="#475569" strokeWidth="1"/>
        {boxOnConv&&<g><rect x={boxConvX} y="40" width="16" height="12" rx="2" fill="#f59e0b" stroke="#d97706" strokeWidth="0.8"/><line x1={boxConvX+2} y1="44" x2={boxConvX+14} y2="44" stroke="#92400e" strokeWidth="0.8"/><line x1={boxConvX+8} y1="40" x2={boxConvX+8} y2="52" stroke="#92400e" strokeWidth="0.8"/></g>}
        <g transform="translate(145,8)">
          <rect x="-10" y="0" width="20" height="5" rx="2" fill="#475569"/><rect x="-4" y="-3" width="8" height="5" rx="1" fill="#64748b"/>
          <rect x="-3" y="4" width="6" height={20+armDown*16} rx="2" fill="#334155"/>
          <g transform={`translate(0,${24+armDown*16})`}>
            <rect x="-6" y="0" width="12" height="4" rx="1" fill="#1e293b"/>
            <rect x={armGripping?-7:-9} y="3" width="5" height="8" rx="1" fill="#94a3b8"/>
            <rect x={armGripping?2:4} y="3" width="5" height="8" rx="1" fill="#94a3b8"/>
            {boxOnArm&&<g transform="translate(-7,9)"><rect x="0" y="0" width="14" height="11" rx="2" fill="#f59e0b" stroke="#d97706" strokeWidth="0.6"/><line x1="2" y1="4" x2="12" y2="4" stroke="#92400e" strokeWidth="0.8"/><line x1="7" y1="0" x2="7" y2="11" stroke="#92400e" strokeWidth="0.8"/></g>}
          </g>
          <rect x="-10" y="3" width="7" height="9" rx="2" fill="#1e3a8a"/><circle cx="-6" cy="7" r="2.5" fill="#3b82f6"/>
        </g>
        <rect x="175" y="58" width="50" height="6" rx="1" fill="#92400e"/><rect x="179" y="54" width="42" height="5" rx="1" fill="#78350f"/>
        {[180,193,206,219].map(x=><rect key={x} x={x} y="63" width="4" height="2" fill="#78350f"/>)}
        {Array.from({length:Math.min(done,8)}).map((_,i)=>{const col=i%4,row=Math.floor(i/4);return(<g key={i}><rect x={180+col*10} y={46-row*11} width="9" height="9" rx="1" fill="#f59e0b" stroke="#d97706" strokeWidth="0.5"/><line x1={181+col*10} y1={49-row*11} x2={188+col*10} y2={49-row*11} stroke="#92400e" strokeWidth="0.6"/></g>);})}
        <g transform={`translate(${forkX},28)`}>
          <rect x="14" y="-2" width="4" height="30" rx="1" fill="#1e40af"/><rect x="17" y="20" width="22" height="2.5" rx="1" fill="#93c5fd"/><rect x="17" y="24" width="22" height="2.5" rx="1" fill="#93c5fd"/>
          <rect x="-13" y="10" width="29" height="18" rx="3" fill="#1d4ed8"/><rect x="-9" y="4" width="15" height="10" rx="2" fill="#2563eb"/><rect x="-7" y="5" width="11" height="7" rx="1" fill="#bfdbfe" opacity="0.85"/>
          <ellipse cx="14" cy="14" rx="2.5" ry="2" fill="#fef08a"/><rect x="-15" y="16" width="6" height="10" rx="2" fill="#1e3a8a"/>
          <ellipse cx="8" cy="29" rx="5" ry="4" fill="#1e293b"/><ellipse cx="8" cy="29" rx="3" ry="2.5" fill="#334155"/><circle cx="8" cy="29" r="1" fill="#64748b"/>
          <ellipse cx="-8" cy="29" rx="4" ry="3.5" fill="#1e293b"/><ellipse cx="-8" cy="29" rx="2.5" ry="2" fill="#334155"/><ellipse cx="-1" cy="7" rx="3" ry="3" fill="#1e3a8a"/>
        </g>
        <rect x="258" y="40" width="38" height="25" rx="2" fill="#f0fdf4" stroke="#86efac" strokeWidth="1" strokeDasharray="3,2"/>
        <text x="277" y="53" fontSize="7" fill="#16a34a" textAnchor="middle" fontWeight="600">GOAL</text>
        <text x="277" y="62" fontSize="6" fill="#22c55e" textAnchor="middle">出荷済</text>
      </svg>
      <div style={{ background:"#e2e8f0", borderRadius:99, height:5, overflow:"hidden", marginTop:8 }}>
        <div style={{ height:"100%", width:`${Math.round(pct*100)}%`, background:"linear-gradient(90deg,#2563eb,#22c55e)", transition:"width .5s", borderRadius:99 }}/>
      </div>
      <div style={{ display:"flex", justifyContent:"space-between", marginTop:4 }}>
        <span style={{ fontSize:10, color:"#64748b" }}>残り{timeStr}</span>
        <span style={{ fontSize:10, color:"#475569", fontWeight:600 }}>{Math.round(pct*100)}%</span>
      </div>
    </div>
  );
}

function CompanyModal({ company, candidates, onConfirm }) {
  const [selected, setSelected] = useState(candidates[0]||"");
  const [custom, setCustom] = useState("");
  const isOther = selected==="__other__";
  const finalName = isOther?custom:selected;
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 }}>
      <div style={{ background:"#fff", borderRadius:14, padding:"24px", width:380, maxWidth:"90vw", boxShadow:"0 20px 60px rgba(0,0,0,0.3)" }}>
        <div style={{ fontSize:15, fontWeight:700, color:"#111827", marginBottom:6 }}>企業名を確認してください</div>
        <div style={{ fontSize:12, color:"#6b7280", marginBottom:16 }}>「{company}」に該当する企業候補：</div>
        <div style={{ display:"flex", flexDirection:"column", gap:8, marginBottom:16 }}>
          {candidates.map(c=>(
            <label key={c} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", borderRadius:8, border:`1.5px solid ${selected===c?"#1e3a8a":"#e5e7eb"}`, background:selected===c?"#eff6ff":"#fff", cursor:"pointer" }}>
              <input type="radio" name="company" value={c} checked={selected===c} onChange={()=>setSelected(c)} style={{ accentColor:"#1e3a8a" }}/>
              <span style={{ fontSize:13, color:"#111827", fontWeight:selected===c?600:400 }}>{c}</span>
            </label>
          ))}
          <label style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", borderRadius:8, border:`1.5px solid ${isOther?"#1e3a8a":"#e5e7eb"}`, background:isOther?"#eff6ff":"#fff", cursor:"pointer" }}>
            <input type="radio" name="company" value="__other__" checked={isOther} onChange={()=>setSelected("__other__")} style={{ accentColor:"#1e3a8a" }}/>
            <span style={{ fontSize:13, color:"#111827" }}>その他</span>
          </label>
          {isOther&&<input autoFocus type="text" value={custom} onChange={e=>setCustom(e.target.value)} placeholder="企業名を入力" style={{ padding:"9px 12px", fontSize:13, border:"1.5px solid #1e3a8a", borderRadius:8, outline:"none", fontFamily:"inherit" }}/>}
        </div>
        <button onClick={()=>finalName.trim()&&onConfirm(finalName.trim())} disabled={!finalName.trim()}
          style={{ width:"100%", padding:"11px", fontSize:13, fontWeight:700, background:finalName.trim()?"#1e3a8a":"#d1d5db", color:"#fff", border:"none", borderRadius:8, cursor:finalName.trim()?"pointer":"not-allowed", fontFamily:"inherit" }}>
          確定してリサーチ開始
        </button>
      </div>
    </div>
  );
}

function StatusBadge({ status, size="sm" }) {
  if (!status) return null;
  const cfg = STATUS[status]||STATUS.unconfirmed;
  return <span style={{ fontSize:size==="lg"?12:10, fontWeight:600, color:cfg.color, background:cfg.bg, border:`1px solid ${cfg.border}`, padding:size==="lg"?"3px 9px":"1px 6px", borderRadius:99, whiteSpace:"nowrap", flexShrink:0 }}>{cfg.label}</span>;
}

function Checkbox({ checked, color, indeterminate }) {
  return (
    <div style={{ width:14, height:14, borderRadius:3, flexShrink:0, marginTop:1, border:`1.5px solid ${checked||indeterminate?color:"#d1d5db"}`, background:checked?color:"transparent", display:"flex", alignItems:"center", justifyContent:"center" }}>
      {checked&&<span style={{ fontSize:9, color:"#fff", fontWeight:700, lineHeight:1 }}>✓</span>}
      {indeterminate&&!checked&&<div style={{ width:6, height:2, background:color, borderRadius:1 }}/>}
    </div>
  );
}

function ResultCard({ item, result }) {
  const [open, setOpen] = useState(true);
  const st = result?.status||"unconfirmed";
  const cfg = STATUS[st];
  const borderColor = st==="unconfirmed"?"#fecaca":st==="partial"?"#fde68a":"#e5e7eb";
  const headerBg = st==="unconfirmed"?"#fff5f5":st==="partial"?"#fffdf0":"#fafafa";
  const sourceUrls = result?.source_url?result.source_url.split(",").map(u=>u.trim()).filter(u=>u.startsWith("http")):[];
  return (
    <div style={{ border:`1px solid ${borderColor}`, borderRadius:10, marginBottom:8, overflow:"hidden", background:"#fff" }}>
      <div onClick={()=>setOpen(o=>!o)} style={{ display:"flex", alignItems:"center", gap:8, padding:"9px 14px", cursor:"pointer", background:headerBg }}>
        <div style={{ width:3, height:14, background:cfg.color, borderRadius:2, flexShrink:0 }}/>
        <span style={{ fontSize:13, fontWeight:600, color:"#111827", flex:1 }}>{item.label}</span>
        <StatusBadge status={st}/>
        <span style={{ fontSize:11, color:"#9ca3af", marginLeft:4 }}>{open?"▲":"▼"}</span>
      </div>
      {open&&(
        <div style={{ padding:"10px 14px 12px", borderTop:`1px solid ${cfg.border}` }}>
          <p style={{ margin:"0 0 8px", fontSize:13, color:"#374151", lineHeight:1.85 }}>{result?.summary||"情報を取得できませんでした。"}</p>
          {result?.missing&&st!=="confirmed"&&(
            <div style={{ display:"flex", alignItems:"flex-start", gap:7, background:cfg.bg, border:`1px solid ${cfg.border}`, borderRadius:7, padding:"7px 11px", marginTop:6 }}>
              <span style={{ fontSize:11, fontWeight:700, color:cfg.color, whiteSpace:"nowrap", marginTop:1 }}>{st==="unconfirmed"?"⚠ 要ヒアリング":"△ 要確認"}</span>
              <span style={{ fontSize:12, color:cfg.color, lineHeight:1.6 }}>{result.missing}</span>
            </div>
          )}
          {sourceUrls.length>0&&(
            <div style={{ marginTop:8, display:"flex", alignItems:"flex-start", gap:5, flexWrap:"wrap" }}>
              <span style={{ fontSize:10, color:"#9ca3af", flexShrink:0, marginTop:2 }}>📎 参照：</span>
              <div style={{ display:"flex", flexDirection:"column", gap:2 }}>
                {sourceUrls.map((url,i)=>{let host=url;try{host=new URL(url).hostname;}catch{}return(<a key={i} href={url} target="_blank" rel="noopener noreferrer" style={{ fontSize:10, color:"#3b82f6", textDecoration:"none" }}>{host}</a>);})}
              </div>
            </div>
          )}
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
  const [pptStyle, setPptStyle] = useState("sales");
  const [appStatus, setAppStatus] = useState("idle");
  const [progress, setProgress] = useState({ done:0, total:0, current:"" });
  const [results, setResults] = useState({});
  const [activeCategory, setActiveCategory] = useState(null);
  const [activeTab, setActiveTab] = useState("data"); // "data" | "slide"
  const [modal, setModal] = useState(null);
  const abortRef = useRef(false);

  const getIndustry = () => { if (industryMode==="auto") return ""; if (industrySelect==="その他") return industryCustom; return industrySelect; };
  const activePreset = PRESETS.findIndex(p=>{ if(p.ids.size!==selected.size)return false; for(const id of p.ids)if(!selected.has(id))return false; return true; });
  const toggle = (id) => setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});
  const toggleCat = (cat) => { const ids=cat.items.map(i=>i.id); const allOn=ids.every(id=>selected.has(id)); setSelected(s=>{const n=new Set(s);ids.forEach(id=>allOn?n.delete(id):n.add(id));return n;}); };
  const applyPreset = (preset) => setSelected(new Set(preset.ids));

  const statusCounts = Object.values(results).reduce((acc,r)=>{ const k=r?.status||"unconfirmed"; acc[k]=(acc[k]||0)+1; return acc; },{});
  const hearingItems = ALL_ITEMS.filter(i=>selected.has(i.id)&&results[i.id]?.status==="unconfirmed");
  const partialItems = ALL_ITEMS.filter(i=>selected.has(i.id)&&results[i.id]?.status==="partial");

  const checkCompany = async (name) => {
    for (let attempt=0;attempt<3;attempt++) {
      try {
        if(attempt>0) await new Promise(r=>setTimeout(r,attempt*3000));
        const raw = await callClaude(`あなたは企業名の曖昧さを判定するAIです。企業名が曖昧・略称・グループ名の場合のみ候補リストを返してください。明確な企業名の場合はambiguous:falseを返してください。JSONのみ返答。`,`企業名「${name}」は曖昧ですか？曖昧な場合は日本の代表的な該当企業を最大5件リストアップ。JSONのみ: {"ambiguous":true|false,"candidates":["企業名1","企業名2"]}`);
        const parsed = extractJSON(raw);
        if(parsed?.ambiguous&&parsed?.candidates?.length>0) return parsed.candidates;
        return null;
      } catch(_){ if(attempt===2) return null; }
    }
    return null;
  };

  const startResearch = async (confirmedCompany) => {
    setModal(null); setCompany(confirmedCompany); abortRef.current=false;
    setAppStatus("loading"); setResults({}); setParsedPrior(null); setActiveCategory(null); setActiveTab("data");
    const items = ALL_ITEMS.filter(i=>selected.has(i.id));
    const hasPrior = priorInfo.trim().length>0;
    const industry = getIndustry();
    setProgress({ done:0, total:items.length+(hasPrior?1:0), current:"" });
    let priorContext = "";
    if(hasPrior) {
      setProgress({ done:0, total:items.length+1, current:"事前情報を整理中" });
      try {
        const parseRaw = await callClaudeNoSearch(`あなたはB2B営業支援AIです。貼り付けられたヒアリングメモ・事前情報を整理しJSONのみ返答。前置き不要。`,`以下は"${confirmedCompany}"に関する事前情報です。\n---\n${priorInfo}\n---\nJSONのみ: {"summary":"全体サマリー3〜4文","key_points":["重要ポイント1","重要ポイント2"],"known_challenges":"判明している課題・ニーズ","known_contacts":"判明している担当者情報","known_systems":"判明しているシステム環境","other":"その他有用情報"}`);
        const pp = extractJSON(parseRaw);
        if(pp){setParsedPrior(pp);priorContext="\n\n【事前情報】"+JSON.stringify(pp);}
      } catch(_){}
      setProgress({ done:1, total:items.length+1, current:"" });
    }
    const newResults = {};
    const offset = hasPrior?1:0;
    for(let i=0;i<items.length;i++) {
      if(abortRef.current) break;
      const item = items[i];
      setProgress({ done:i+offset, total:items.length+offset, current:item.label });
      try {
        const base = item.id==="market_share"?PROMPTS.market_share(confirmedCompany,industry):(PROMPTS[item.id]?.(confirmedCompany)||`"${confirmedCompany}"について「${item.label}」を調査。${FACT_ONLY} JSONのみ: {"summary":"3文以内","status":"confirmed|partial|unconfirmed","missing":"訪問時確認事項","source_url":"参照URL"}`);
        const prompt = priorContext?base+priorContext+"\n※事前情報に記載された内容はそのままsummaryに含めてよい。推測は書かない。":base;
        const raw = await deepResearchPremium(confirmedCompany,item.id,prompt,industry);
        const parsed = extractJSON(raw);
        newResults[item.id] = parsed||{summary:raw.slice(0,200),status:"partial",missing:"",source_url:""};
      } catch { newResults[item.id]={summary:"取得できませんでした。",status:"unconfirmed",missing:"デスクトップリサーチでは確認不可。訪問時に直接確認が必要です。",source_url:""}; }
      setResults({...newResults});
      await new Promise(r=>setTimeout(r,15000));
    }
    setProgress(p=>({...p,done:p.total,current:""}));
    setAppStatus("done"); setActiveCategory(CATEGORIES[0].id);
  };

  const handleStart = async () => {
    if(!company.trim()||selected.size===0) return;
    setAppStatus("checking");
    const candidates = await checkCompany(company.trim());
    if(candidates){setModal({candidates});setAppStatus("idle");}
    else await startResearch(company.trim());
  };

  const resetResearch = () => { abortRef.current=true; setAppStatus("idle"); setResults({}); setParsedPrior(null); setActiveCategory(null); setProgress({done:0,total:0,current:""}); };
  const hasResults = Object.keys(results).length>0;
  const displayCat = CATEGORIES.find(c=>c.id===activeCategory)||CATEGORIES[0];
  const isLoading = appStatus==="loading"||appStatus==="checking";

  const inp = (label, val, setter, ph) => (
    <div style={{ marginBottom:7 }}>
      <div style={{ fontSize:10, fontWeight:600, color:"#6b7280", marginBottom:3 }}>{label}</div>
      <input type="text" value={val} onChange={e=>setter(e.target.value)} placeholder={ph}
        style={{ width:"100%", padding:"7px 9px", fontSize:12, border:"1px solid #e5e7eb", borderRadius:7, fontFamily:"inherit", boxSizing:"border-box", outline:"none" }}
        onFocus={e=>e.target.style.borderColor="#1e3a8a"} onBlur={e=>e.target.style.borderColor="#e5e7eb"}/>
    </div>
  );

  return (
    <div style={{ fontFamily:"'Helvetica Neue','Hiragino Kaku Gothic ProN',Meiryo,sans-serif", background:"#f8f9fc", minHeight:"100vh" }}>
      {modal&&<CompanyModal company={company} candidates={modal.candidates} onConfirm={startResearch}/>}

      <div style={{ background:"#1e3a8a", padding:"12px 22px", display:"flex", alignItems:"center", gap:11 }}>
        <div style={{ width:28, height:28, background:"#F96167", borderRadius:6, display:"flex", alignItems:"center", justifyContent:"center", fontSize:15 }}>📦</div>
        <div>
          <div style={{ fontSize:14, fontWeight:700, color:"#fff" }}>SBR 顧客リサーチ &amp; 社内報告ツール</div>
          <div style={{ fontSize:11, color:"#93c5fd" }}>新規顧客の事前調査 → 社内報告PPT 一気通貫</div>
        </div>
      </div>

      <div style={{ display:"flex", height:"calc(100vh - 50px)" }}>
        {/* 左パネル */}
        <div style={{ width:292, background:"#fff", borderRight:"1px solid #e5e7eb", overflowY:"auto", display:"flex", flexDirection:"column", flexShrink:0 }}>
          <div style={{ padding:"14px 14px 0", flex:1 }}>
            <div style={{ fontSize:11, fontWeight:600, color:"#1e3a8a", marginBottom:8, display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:3, height:11, background:"#1e3a8a", borderRadius:2 }}/>基本情報
            </div>
            <div style={{ marginBottom:7 }}>
              <div style={{ fontSize:10, fontWeight:600, color:"#6b7280", marginBottom:3 }}>顧客企業名 *</div>
              <input type="text" value={company} onChange={e=>setCompany(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&!isLoading&&company.trim()&&selected.size&&handleStart()}
                placeholder="例：ニチレイロジグループ株式会社"
                style={{ width:"100%", padding:"7px 9px", fontSize:12, border:"1px solid #e5e7eb", borderRadius:7, fontFamily:"inherit", boxSizing:"border-box", outline:"none" }}
                onFocus={e=>e.target.style.borderColor="#1e3a8a"} onBlur={e=>e.target.style.borderColor="#e5e7eb"}/>
            </div>
            {inp("担当者部署", dept, setDept, "例：物流部・SCM部")}

            <div style={{ height:1, background:"#f3f4f6", margin:"10px 0" }}/>

            <div style={{ fontSize:11, fontWeight:600, color:"#1e3a8a", marginBottom:6, display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:3, height:11, background:"#1e3a8a", borderRadius:2 }}/>事前情報
              <span style={{ fontWeight:400, color:"#9ca3af", fontSize:10 }}>任意</span>
            </div>
            <textarea value={priorInfo} onChange={e=>setPriorInfo(e.target.value)}
              placeholder={"例：\n・担当の山田部長から「WMS刷新を検討中」と聞いた\n・現状はExcel管理でミスが多い"} rows={4}
              style={{ width:"100%", padding:"8px 9px", fontSize:12, border:"1px solid #e5e7eb", borderRadius:7, fontFamily:"inherit", boxSizing:"border-box", outline:"none", resize:"vertical", lineHeight:1.6 }}
              onFocus={e=>e.target.style.borderColor="#1e3a8a"} onBlur={e=>e.target.style.borderColor="#e5e7eb"}/>
            {priorInfo.trim()&&<div style={{ fontSize:11, color:"#15803d", marginBottom:4, marginTop:3 }}>✓ 事前情報あり — リサーチ時に自動反映します</div>}

            <div style={{ height:1, background:"#f3f4f6", margin:"10px 0" }}/>

            {/* PPTスタイル選択 */}
            <div style={{ fontSize:11, fontWeight:600, color:"#1e3a8a", marginBottom:8, display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:3, height:11, background:"#1e3a8a", borderRadius:2 }}/>レポートスタイル
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:6, marginBottom:10 }}>
              {PPT_STYLES.map(s=>(
                <label key={s.id} style={{ display:"flex", alignItems:"flex-start", gap:8, padding:"8px 10px", borderRadius:8, border:`1.5px solid ${pptStyle===s.id?"#1e3a8a":"#e5e7eb"}`, background:pptStyle===s.id?"#eff6ff":"#fff", cursor:"pointer" }}>
                  <input type="radio" name="pptStyle" value={s.id} checked={pptStyle===s.id} onChange={()=>setPptStyle(s.id)} style={{ accentColor:"#1e3a8a", marginTop:2 }}/>
                  <div>
                    <div style={{ fontSize:12, fontWeight:600, color:"#111827" }}>{s.icon} {s.label} <span style={{ fontSize:10, color:"#9ca3af", fontWeight:400 }}>（{s.slides}枚）</span></div>
                    <div style={{ fontSize:10, color:"#6b7280", marginTop:2 }}>{s.desc}</div>
                  </div>
                </label>
              ))}
            </div>

            <div style={{ height:1, background:"#f3f4f6", margin:"10px 0" }}/>

            <div style={{ fontSize:11, fontWeight:600, color:"#1e3a8a", marginBottom:8, display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:3, height:11, background:"#1e3a8a", borderRadius:2 }}/>業界シェアの調査軸
            </div>
            <div style={{ display:"flex", gap:12, marginBottom:8 }}>
              {["auto","specify"].map(mode=>(
                <label key={mode} style={{ display:"flex", alignItems:"center", gap:5, cursor:"pointer", fontSize:12 }}>
                  <input type="radio" name="industryMode" value={mode} checked={industryMode===mode} onChange={()=>setIndustryMode(mode)}/>
                  {mode==="auto"?"おまかせ":"指定する"}
                </label>
              ))}
            </div>
            {industryMode==="specify"&&(
              <div style={{ marginBottom:8 }}>
                <select value={industrySelect} onChange={e=>setIndustrySelect(e.target.value)}
                  style={{ width:"100%", padding:"7px 9px", fontSize:12, border:"1px solid #e5e7eb", borderRadius:7, fontFamily:"inherit", boxSizing:"border-box", outline:"none", marginBottom:6 }}>
                  {INDUSTRY_OPTIONS.map(opt=><option key={opt} value={opt}>{opt}</option>)}
                </select>
                {industrySelect==="その他"&&<input type="text" value={industryCustom} onChange={e=>setIndustryCustom(e.target.value)} placeholder="業界名を入力"
                  style={{ width:"100%", padding:"7px 9px", fontSize:12, border:"1px solid #e5e7eb", borderRadius:7, fontFamily:"inherit", boxSizing:"border-box", outline:"none" }}
                  onFocus={e=>e.target.style.borderColor="#1e3a8a"} onBlur={e=>e.target.style.borderColor="#e5e7eb"}/>}
              </div>
            )}

            <div style={{ height:1, background:"#f3f4f6", margin:"10px 0" }}/>

            <div style={{ fontSize:11, fontWeight:600, color:"#1e3a8a", marginBottom:6, display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:3, height:11, background:"#1e3a8a", borderRadius:2 }}/>
              調査項目　<span style={{ fontWeight:400, color:"#9ca3af", fontSize:10 }}>{selected.size}件選択</span>
            </div>
            <div style={{ display:"flex", gap:4, marginBottom:10, flexWrap:"wrap" }}>
              {PRESETS.map((p,idx)=>{ const isActive=activePreset===idx; return(
                <button key={p.label} onClick={()=>applyPreset(p)}
                  style={{ fontSize:10, padding:"4px 10px", borderRadius:99, fontFamily:"inherit", cursor:"pointer", fontWeight:isActive?700:500, background:isActive?"#1e3a8a":"#f1f5f9", color:isActive?"#fff":"#374151", border:isActive?"1.5px solid #1e3a8a":"1.5px solid #e2e8f0" }}>
                  {isActive?"✓ ":""}{p.label}
                </button>
              );})}
            </div>

            {CATEGORIES.map(cat=>{
              const ids=cat.items.map(i=>i.id); const allOn=ids.every(id=>selected.has(id)); const partial=ids.some(id=>selected.has(id))&&!allOn;
              return(
                <div key={cat.id} style={{ marginBottom:9 }}>
                  <div onClick={()=>toggleCat(cat)} style={{ display:"flex", alignItems:"center", gap:7, padding:"4px 6px", borderRadius:6, cursor:"pointer", marginBottom:3 }}>
                    <Checkbox checked={allOn} indeterminate={partial} color={cat.color}/><div style={{ width:7, height:7, borderRadius:"50%", background:cat.color, flexShrink:0 }}/><span style={{ fontSize:11, fontWeight:600, color:"#374151" }}>{cat.label}</span>
                  </div>
                  <div style={{ paddingLeft:10 }}>
                    {cat.items.map(item=>{ const on=selected.has(item.id); const r=results[item.id]; return(
                      <label key={item.id} style={{ display:"flex", alignItems:"flex-start", gap:7, padding:"4px 6px", borderRadius:6, cursor:"pointer", marginBottom:1 }}>
                        <Checkbox checked={on} color={cat.color}/><input type="checkbox" checked={on} onChange={()=>toggle(item.id)} style={{ display:"none" }}/>
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:12, color:on?"#111827":"#9ca3af", fontWeight:on?500:400, display:"flex", alignItems:"center", gap:5, flexWrap:"wrap", lineHeight:1.4 }}>
                            {item.label}{r&&<StatusBadge status={r.status}/>}
                          </div>
                        </div>
                      </label>
                    );})}
                  </div>
                </div>
              );
            })}
          </div>

          <div style={{ padding:"12px 14px", borderTop:"1px solid #f3f4f6" }}>
            {appStatus==="checking"&&<div style={{ fontSize:11, color:"#6b7280", textAlign:"center", marginBottom:8 }}>🔍 企業名を確認中...</div>}
            {appStatus==="loading"&&<WarehouseProgress done={progress.done} total={progress.total} current={progress.current}/>}
            <div style={{ display:"flex", gap:6 }}>
              {hasResults&&<button onClick={resetResearch} style={{ flex:1, padding:"10px", fontSize:12, fontWeight:600, fontFamily:"inherit", background:"#fff", color:"#374151", border:"1px solid #e5e7eb", borderRadius:8, cursor:"pointer" }}>新しいリサーチ</button>}
              {isLoading?
                <button onClick={()=>{abortRef.current=true;setAppStatus("idle");}} style={{ flex:1, padding:"10px", fontSize:13, fontWeight:700, fontFamily:"inherit", background:"#dc2626", color:"#fff", border:"none", borderRadius:8, cursor:"pointer" }}>⏹ 中断</button>:
                <button onClick={handleStart} disabled={!company.trim()||!selected.size}
                  style={{ flex:1, padding:"10px", fontSize:13, fontWeight:700, fontFamily:"inherit", background:(!company.trim()||!selected.size)?"#d1d5db":"#1e3a8a", color:"#fff", border:"none", borderRadius:8, cursor:(!company.trim()||!selected.size)?"not-allowed":"pointer" }}>
                  🔍 リサーチ開始
                </button>
              }
            </div>
          </div>
        </div>

        {/* 右パネル */}
        <div style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>
          {!hasResults?(
            <div style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", color:"#d1d5db" }}>
              <div style={{ fontSize:48, marginBottom:12 }}>🔍</div>
              <div style={{ fontSize:15, fontWeight:600 }}>顧客名を入力してリサーチ開始</div>
              <div style={{ fontSize:12, marginTop:6 }}>調査項目を選んでAIが自動でWeb検索・分析します</div>
            </div>
          ):(
            <>
              {/* タブ */}
              <div style={{ background:"#fff", borderBottom:"1px solid #e5e7eb", padding:"0 20px", display:"flex", alignItems:"center", gap:0 }}>
                <div style={{ display:"flex", alignItems:"center", gap:16, flex:1 }}>
                  {[{id:"data",label:"📋 詳細データ"},{id:"slide",label:"📊 スライド"}].map(tab=>(
                    <button key={tab.id} onClick={()=>setActiveTab(tab.id)}
                      style={{ padding:"12px 4px", fontSize:13, fontWeight:activeTab===tab.id?700:400, color:activeTab===tab.id?"#1e3a8a":"#6b7280", background:"none", border:"none", borderBottom:activeTab===tab.id?"2px solid #1e3a8a":"2px solid transparent", cursor:"pointer", fontFamily:"inherit" }}>
                      {tab.label}
                    </button>
                  ))}
                </div>
                <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                  {["confirmed","partial","unconfirmed"].map(s=>statusCounts[s]?(
                    <div key={s} style={{ display:"flex", alignItems:"center", gap:4 }}>
                      <StatusBadge status={s} size="lg"/><span style={{ fontSize:12, color:"#6b7280" }}>{statusCounts[s]}件</span>
                    </div>
                  ):null)}
                </div>
              </div>

              {/* タブコンテンツ */}
              {activeTab==="data"?(
                <div style={{ flex:1, overflowY:"auto", padding:20 }}>
                  <div style={{ maxWidth:820 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14 }}>
                      <div>
                        <div style={{ fontSize:20, fontWeight:700, color:"#111827" }}>{company}</div>
                        {dept&&<div style={{ fontSize:12, color:"#6b7280", marginTop:2 }}>担当部署：{dept}</div>}
                      </div>
                    </div>

                    {parsedPrior&&(
                      <div style={{ background:"#f0fdf4", border:"1px solid #bbf7d0", borderRadius:10, padding:"12px 16px", marginBottom:12 }}>
                        <div style={{ fontSize:13, fontWeight:600, color:"#15803d", marginBottom:8, display:"flex", alignItems:"center", gap:6 }}>
                          <span>📋 事前情報（整理済み）</span>
                          <span style={{ fontSize:10, fontWeight:400, color:"#16a34a", background:"#dcfce7", border:"1px solid #bbf7d0", padding:"1px 7px", borderRadius:99 }}>リサーチに反映済み</span>
                        </div>
                        <p style={{ margin:"0 0 8px", fontSize:13, color:"#166534", lineHeight:1.75 }}>{parsedPrior.summary}</p>
                        {parsedPrior.key_points?.length>0&&(
                          <div style={{ marginBottom:6 }}>
                            <div style={{ fontSize:11, fontWeight:600, color:"#15803d", marginBottom:4 }}>重要ポイント</div>
                            {parsedPrior.key_points.map((pt,i)=><div key={i} style={{ display:"flex", gap:6, fontSize:12, color:"#166534", marginBottom:2 }}><span style={{ flexShrink:0 }}>•</span><span>{pt}</span></div>)}
                          </div>
                        )}
                      </div>
                    )}

                    {appStatus==="done"&&hearingItems.length>0&&(
                      <div style={{ background:"#fff5f5", border:"1px solid #fecaca", borderRadius:10, padding:"12px 16px", marginBottom:12 }}>
                        <div style={{ fontSize:13, fontWeight:600, color:"#dc2626", marginBottom:8 }}>⚠️ デスクトップリサーチで確認できなかった項目 ({hearingItems.length}件) — 初回訪問で確認</div>
                        {hearingItems.map(item=>(
                          <div key={item.id} style={{ display:"flex", gap:8, fontSize:12, marginBottom:4 }}>
                            <span style={{ color:"#dc2626", flexShrink:0 }}>•</span>
                            <span style={{ fontWeight:600, color:"#111827", flexShrink:0, minWidth:120 }}>{item.label}</span>
                            <span style={{ color:"#6b7280" }}>{results[item.id]?.missing||"訪問時に確認"}</span>
                          </div>
                        ))}
                      </div>
                    )}

                    {appStatus==="done"&&partialItems.length>0&&(
                      <div style={{ background:"#fffdf0", border:"1px solid #fde68a", borderRadius:10, padding:"10px 16px", marginBottom:12 }}>
                        <div style={{ fontSize:13, fontWeight:600, color:"#a16207", marginBottom:6 }}>△ 情報が一部のみ取得できた項目 ({partialItems.length}件) — 追加確認推奨</div>
                        {partialItems.map(item=>(
                          <div key={item.id} style={{ display:"flex", gap:8, fontSize:12, marginBottom:6 }}>
                            <span style={{ color:"#a16207", flexShrink:0 }}>•</span>
                            <span style={{ fontWeight:600, color:"#111827", flexShrink:0, minWidth:120 }}>{item.label}</span>
                            <div style={{ flex:1 }}>
                              {results[item.id]?.summary&&<div style={{ fontSize:12, color:"#78350f", marginBottom:3 }}>{results[item.id].summary}</div>}
                              {results[item.id]?.missing&&<div style={{ fontSize:11, color:"#92400e", opacity:0.8 }}>△ {results[item.id].missing}</div>}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}

                    <div style={{ display:"flex", gap:4, marginBottom:14, flexWrap:"wrap" }}>
                      {CATEGORIES.filter(cat=>cat.items.some(i=>results[i.id])).map(cat=>{
                        const catItems=cat.items.filter(i=>results[i.id]);
                        const hasU=catItems.some(i=>results[i.id]?.status==="unconfirmed");
                        const hasP=catItems.some(i=>results[i.id]?.status==="partial");
                        const active=activeCategory===cat.id;
                        return(
                          <button key={cat.id} onClick={()=>setActiveCategory(cat.id)} style={{ padding:"5px 12px", fontSize:12, fontWeight:600, borderRadius:20, cursor:"pointer", fontFamily:"inherit", background:active?cat.color:"#fff", color:active?"#fff":"#374151", border:`1.5px solid ${active?cat.color:"#e5e7eb"}`, display:"flex", alignItems:"center", gap:5 }}>
                            {cat.label}
                            {hasU&&<span style={{ width:7, height:7, borderRadius:"50%", background:active?"#fca5a5":"#ef4444", display:"inline-block", flexShrink:0 }}/>}
                            {!hasU&&hasP&&<span style={{ width:7, height:7, borderRadius:"50%", background:active?"#fde68a":"#f59e0b", display:"inline-block", flexShrink:0 }}/>}
                          </button>
                        );
                      })}
                    </div>

                    {displayCat&&(
                      <div>
                        <div style={{ fontSize:13, fontWeight:600, color:displayCat.color, marginBottom:10, display:"flex", alignItems:"center", gap:6 }}>
                          <div style={{ width:3, height:14, background:displayCat.color, borderRadius:2 }}/>{displayCat.label}
                        </div>
                        {displayCat.items.filter(i=>results[i.id]).map(item=>(
                          <ResultCard key={item.id} item={item} result={results[item.id]}/>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              ):(
                <div style={{ flex:1, overflow:"hidden" }}>
                  <SlidePreview company={company} dept={dept} results={results} pptStyle={pptStyle} selectedIds={selected}/>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
