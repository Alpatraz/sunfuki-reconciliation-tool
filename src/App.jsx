
import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./style.css";

const UNIQUE_SIZE_PRODUCTS = new Set(["sac", "patch"]);
const PRODUCT_LABELS = {
  tshirt: "T-shirt",
  hoodie: "Hoodie",
  kangourou: "Kangourou",
  veste: "Veste",
  kimono: "Kimono",
  gants: "Gants",
  protege_pieds: "Protège-pieds",
  protege_tibias: "Protège-tibias",
  casque: "Casque",
  sac: "Sac",
  ceinture: "Ceinture",
  patch: "Patch",
  engagement: "Engagement",
  unknown: "Produit inconnu",
};

const SIZE_WORDS = [
  "0000","000","00","0","1","2","3","4","5","6","7","8",
  "XXS","XS","S","M","L","XL","XXL","XXXL","2XL","3XL","4XL",
  "YOUTH MEDIUM","M ADULTE"
];

function stripAccents(value = "") {
  return String(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
function clean(value = "") {
  return String(value ?? "")
    .replace(/\u0000/g, "")
    .replace(/[<>]/g, "")
    .replace(/\r?\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function norm(value = "") {
  return stripAccents(clean(value))
    .toLowerCase()
    .replace(/[’']/g, "")
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}
function pick(row, names) {
  const entries = Object.entries(row || {});
  for (const wanted of names) {
    const found = entries.find(([k]) => norm(k) === norm(wanted));
    if (found) return clean(found[1]);
  }
  for (const wanted of names) {
    const found = entries.find(([k]) => norm(k).includes(norm(wanted)));
    if (found) return clean(found[1]);
  }
  return "";
}
function detectDelimiter(text) {
  const sample = text.slice(0, 5000);
  const candidates = [",", ";", "\t"];
  return candidates.sort((a, b) => sample.split(b).length - sample.split(a).length)[0];
}
function parseCsv(text) {
  const delimiter = detectDelimiter(text);
  const rows = [];
  let row = [], cell = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i], next = text[i + 1];
    if (ch === '"' && quoted && next === '"') { cell += '"'; i++; }
    else if (ch === '"') quoted = !quoted;
    else if (ch === delimiter && !quoted) { row.push(cell); cell = ""; }
    else if ((ch === "\n" || ch === "\r") && !quoted) {
      if (ch === "\r" && next === "\n") i++;
      row.push(cell);
      if (row.some(v => clean(v))) rows.push(row);
      row = []; cell = "";
    } else cell += ch;
  }
  row.push(cell);
  if (row.some(v => clean(v))) rows.push(row);
  if (!rows.length) return [];
  const headers = rows[0].map((h, i) => clean(h) || `col_${i}`);
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = clean(r[i] || ""));
    return obj;
  });
}
async function readFileRows(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const all = [];
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
      json.forEach(r => all.push({ ...r, __sheet: sheetName }));
    }
    return all;
  }
  return parseCsv(await file.text());
}
function exportCsv(filename, rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const esc = v => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const csv = [headers.join(";"), ...rows.map(r => headers.map(h => esc(r[h])).join(";"))].join("\n");
  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
function normalizeOrder(order = "") {
  const raw = clean(order);
  const matches = raw.match(/#?\d+/g);
  if (!matches || !matches.length) return raw ? (raw.startsWith("#") ? raw : `#${raw}`) : "";
  return matches.map(m => m.startsWith("#") ? m : `#${m}`).join(" | ");
}
function splitOrders(order = "") {
  const raw = clean(order);
  const matches = raw.match(/#?\d+/g);
  if (!matches || !matches.length) return raw ? [normalizeOrder(raw)] : [];
  return matches.map(m => m.startsWith("#") ? m : `#${m}`);
}
function competitorKey(name = "") {
  return norm(name);
}
function canonicalProduct(raw = "") {
  const n = norm(raw);
  if (!n) return "unknown";
  if (n.includes("engagement")) return "engagement";
  if (n.includes("kangourou")) return "kangourou";
  if (n.includes("t-shirt") || n.includes("t shirt") || n.includes("tee") || n.includes("tshirt")) return "tshirt";
  if (n.includes("hoodie") || n.includes("chandail")) return "hoodie";
  if (n.includes("veste") || n.includes("softshell")) return "veste";
  if (n.includes("kimono") || n.includes("arawaza") || n.includes("saiko")) return "kimono";
  if (n.includes("gant")) return "gants";
  if (n.includes("protege-pied") || n.includes("protege pied") || n.includes("protège-pied") || n.includes("protège pied") || n.includes("pied")) return "protege_pieds";
  if (n.includes("tibia")) return "protege_tibias";
  if (n.includes("casque")) return "casque";
  if (n.includes("sac")) return "sac";
  if (n.includes("ceinture")) return "ceinture";
  if (n.includes("patch")) return "patch";
  return "unknown";
}
function productLabel(key) {
  return PRODUCT_LABELS[key] || key || "";
}
function hasSize(productKey) {
  return !UNIQUE_SIZE_PRODUCTS.has(productKey);
}
function extractSize(value = "") {
  const v = norm(value);
  if (!v) return "";
  if (v.includes("acompte") || v.includes("engagement") || v.includes("commande equipement") || v.includes("commande équipement")) return "";
  if (/\d+\s*%/.test(v)) return "";
  if (/\d+[,.]\d+\s*\$/.test(v)) return "";
  if (v.includes("youth medium")) return "YOUTH MEDIUM";
  if (v.includes("m adulte") || v.includes("medium adulte")) return "M ADULTE";
  const exact = SIZE_WORDS.find(s => norm(s) === v);
  if (exact) return exact;
  const match = v.match(/\b(0000|000|00|xxxl|xxl|xl|xxs|xs|s|m|l|2xl|3xl|4xl|[0-8])\b/i);
  return match ? match[1].toUpperCase() : "";
}
function productKey(order, prod) {
  return `${normalizeOrder(order)}|${prod}`;
}
function isFitofanComment(name, email, dojo) {
  const n = norm(name);
  const e = norm(email);
  if (!n && !e && !dojo) return true;
  if (!e && (n.startsWith("merci") || n.includes("commande#") || n.includes("commande #") || n.includes("bonne journee") || n.includes("bonne journée") || n.includes("finalement") || n.length > 60)) return true;
  return false;
}

function parseFitofan(rows) {
  const competitors = [];
  const comments = [];
  for (const row of rows) {
    const name = pick(row, ["Noms", "Nom", "Compétiteur", "Competitor", "Athlete"]);
    const email = pick(row, ["couriel", "courriel", "Email", "Courriel"]);
    const dojo = pick(row, ["école", "ecole", "Dojo", "École"]);
    const team = pick(row, ["Equipe", "Équipe", "Team"]);
    if (isFitofanComment(name, email, dojo)) {
      const txt = clean(Object.values(row).filter(Boolean).join(" | "));
      if (txt) comments.push({ Commentaire: txt });
      continue;
    }
    if (!name) continue;
    competitors.push({ competitor: name, key: competitorKey(name), email, dojo, team });
  }
  return { competitors, comments };
}
function parseShopify(rows) {
  const out = [];
  for (const row of rows) {
    const order = normalizeOrder(pick(row, ["Commande", "Order", "Name", "orderName", "Order name", "Numéro"]));
    const statusRaw = pick(row, ["Statut commande", "Status", "Financial status", "Fulfillment status", "Statut"]);
    const client = pick(row, ["Client", "Customer", "Nom client"]);
    const competitor = pick(row, ["Compétiteur", "Competitor", "Athlete", "Nom compétiteur", "Nom du compétiteur"]) || client;
    const email = pick(row, ["Courriel compétiteur", "Email", "Courriel", "Customer email", "recipient_email"]);
    const dojo = pick(row, ["Dojo", "École", "Ecole", "Location"]);
    const team = pick(row, ["Équipe", "Equipe", "Team"]);
    const rawProduct = pick(row, ["Produit", "Product", "Line item name", "Title", "Item"]);
    const qty = Number((pick(row, ["Quantité", "Quantity", "Qty"]) || "1").replace(",", ".")) || 1;
    const refunded = Number((pick(row, ["Qté remboursée", "Qte remboursee", "Refunded quantity", "quantityRefunded"]) || "0").replace(",", ".")) || 0;
    const effectiveRaw = pick(row, ["Qté effective", "Qte effective", "Effective quantity"]);
    let effectiveQty = effectiveRaw === "" ? Math.max(0, qty - refunded) : Number(effectiveRaw.replace(",", "."));
    if (Number.isNaN(effectiveQty)) effectiveQty = Math.max(0, qty - refunded);
    const prod = canonicalProduct(rawProduct);
    if (!rawProduct) continue;
    const cancelled = norm(statusRaw).includes("annul") || norm(statusRaw).includes("rembours");
    const excluded = cancelled || effectiveQty <= 0;
    if (prod === "engagement") {
      out.push({ kind: "engagement", order, competitor, ckey: competitorKey(competitor), client, email, dojo, team, rawProduct, productKey: prod, product: productLabel(prod), shopifySize: "", quantity: qty, refunded, effectiveQty, excluded: true, statusRaw, status: "ENGAGEMENT", comments: "" });
      continue;
    }
    const explicitSize = pick(row, ["Taille", "Size", "Variant", "Option", "Correction taille"]);
    let shopifySize = hasSize(prod) ? extractSize(explicitSize) : "Taille unique";
    out.push({ kind: "product", order, competitor, ckey: competitorKey(competitor), client, email, dojo, team, rawProduct, productKey: prod, product: productLabel(prod), shopifySize, quantity: qty, refunded, effectiveQty, excluded, statusRaw, status: excluded ? "ANNULÉ / REMBOURSÉ - exclu fournisseur" : "", comments: "" });
  }
  return out;
}
function parseSupabase(rows) {
  const out = [];
  for (const row of rows) {
    const orderRaw = pick(row, ["order_number", "Commande", "Order", "orderName", "Numéro"]);
    const orders = splitOrders(orderRaw);
    const competitor = pick(row, ["competitor", "Compétiteur", "Client", "Nom", "customerName"]);
    const email = pick(row, ["recipient_email", "Email", "Courriel", "customer_email"]);
    const rawProduct = pick(row, ["product_name", "Produit", "product", "Article", "item", "Nom produit"]);
    const size = extractSize(pick(row, ["confirmed_size", "Taille", "size", "Nouvelle taille", "Réponse"]));
    const comments = pick(row, ["comments", "Commentaire", "comment", "Notes", "Message", "Réponse", "response"]);
    const prod = canonicalProduct(rawProduct);
    const kind = prod === "engagement" ? "engagement" : (prod === "unknown" ? "comment" : "product");
    const base = { kind, orderRaw: normalizeOrder(orderRaw), competitor, ckey: competitorKey(competitor), email, rawProduct, productKey: prod, product: productLabel(prod), size, comments };
    if (orders.length) orders.forEach(order => out.push({ ...base, order }));
    else out.push({ ...base, order: "" });
  }
  return out;
}

export default function App() {
  const [fitofanRaw, setFitofanRaw] = useState([]);
  const [shopifyRaw, setShopifyRaw] = useState([]);
  const [supabaseRaw, setSupabaseRaw] = useState([]);
  const [files, setFiles] = useState({});
  const [tab, setTab] = useState("audit");
  const [search, setSearch] = useState("");

  async function upload(type, file) {
    if (!file) return;
    const rows = await readFileRows(file);
    setFiles(p => ({ ...p, [type]: `${file.name} (${rows.length} lignes)` }));
    if (type === "fitofan") setFitofanRaw(rows);
    if (type === "shopify") setShopifyRaw(rows);
    if (type === "supabase") setSupabaseRaw(rows);
  }

  const fitofan = useMemo(() => parseFitofan(fitofanRaw), [fitofanRaw]);
  const shopify = useMemo(() => parseShopify(shopifyRaw), [shopifyRaw]);
  const supabase = useMemo(() => parseSupabase(supabaseRaw), [supabaseRaw]);

  const reconciled = useMemo(() => {
    const corrections = new Map();
    const commentsByOrder = new Map();
    const commentsByCompetitor = new Map();

    supabase.forEach(s => {
      const comment = s.comments || (s.kind !== "product" ? s.rawProduct : "");
      if (comment) {
        if (s.order) {
          if (!commentsByOrder.has(s.order)) commentsByOrder.set(s.order, []);
          commentsByOrder.get(s.order).push(comment);
        }
        if (s.ckey) {
          if (!commentsByCompetitor.has(s.ckey)) commentsByCompetitor.set(s.ckey, []);
          commentsByCompetitor.get(s.ckey).push(comment);
        }
      }
      if (s.kind === "product" && s.size && s.order) corrections.set(productKey(s.order, s.productKey), s);
    });

    return shopify.map(item => {
      if (item.kind !== "product") return item;
      const corr = corrections.get(productKey(item.order, item.productKey));
      let finalSize = corr?.size || item.shopifySize || "";
      let sourceSize = corr?.size ? "Supabase" : (item.shopifySize ? "Shopify" : "");
      if (!hasSize(item.productKey)) { finalSize = "Taille unique"; sourceSize = "Taille unique"; }
      const missing = !finalSize && hasSize(item.productKey) && !item.excluded;
      const comments = [...new Set([...(commentsByOrder.get(item.order) || []), ...(commentsByCompetitor.get(item.ckey) || [])])].join(" | ");
      let status = item.excluded ? item.status : "OK SHOPIFY";
      if (corr && !item.shopifySize) status = "TAILLE AJOUTÉE PAR SUPABASE";
      if (corr && item.shopifySize && item.shopifySize !== corr.size) status = "TAILLE MODIFIÉE PAR SUPABASE";
      if (corr && item.shopifySize && item.shopifySize === corr.size) status = "TAILLE CONFIRMÉE PAR SUPABASE";
      if (missing) status = "TAILLE MANQUANTE";
      if (comments) status += " + COMMENTAIRE";
      return { ...item, supabaseSize: corr?.size || "", finalSize, sourceSize, missing, comments, status };
    });
  }, [shopify, supabase]);

  const detailRows = reconciled.map(r => ({
    Commande: r.order,
    Compétiteur: r.competitor,
    Client: r.client,
    Email: r.email,
    Dojo: r.dojo,
    Équipe: r.team,
    Type: r.kind,
    "Produit Shopify brut": r.rawProduct,
    "Produit normalisé": r.productKey,
    Produit: r.product,
    "Taille Shopify": r.shopifySize || "",
    "Taille Supabase": r.supabaseSize || "",
    "Taille finale": r.finalSize || "",
    "Source taille": r.sourceSize || "",
    Quantité: r.quantity,
    "Qté remboursée": r.refunded,
    "Qté effective": r.effectiveQty,
    "Exclu fournisseur": r.excluded ? "Oui" : "Non",
    Statut: r.status,
    Commentaires: r.comments || "",
  }));

  const supplierRows = useMemo(() => {
    const map = new Map();
    reconciled.forEach(r => {
      if (r.kind !== "product" || r.excluded || !r.finalSize) return;
      const key = `${r.productKey}|${r.team}|${r.finalSize}`;
      const prev = map.get(key) || { Produit: r.product, "Produit normalisé": r.productKey, Équipe: r.team || "Non précisée", Taille: r.finalSize, Quantité: 0 };
      prev.Quantité += r.effectiveQty;
      map.set(key, prev);
    });
    return Array.from(map.values()).sort((a,b) => `${a.Produit}${a.Équipe}${a.Taille}`.localeCompare(`${b.Produit}${b.Équipe}${b.Taille}`));
  }, [reconciled]);

  const fitofanReport = useMemo(() => {
    return fitofan.competitors.map(f => {
      const rows = reconciled.filter(r => r.ckey === f.key);
      const products = rows.filter(r => r.kind === "product" && !r.excluded);
      const engagements = rows.filter(r => r.kind === "engagement");
      const orders = [...new Set(rows.map(r => r.order).filter(Boolean))];

      const productSummary = products.map(r => `${r.product} (${r.finalSize || "taille manquante"})`).join(" | ");
      const shopifySizes = products.map(r => `${r.order} - ${r.product}: ${r.shopifySize || "vide"}`).join(" | ");
      const supabaseSizes = products.map(r => `${r.order} - ${r.product}: ${r.supabaseSize || "aucune"}`).join(" | ");
      const finalSizes = products.map(r => `${r.order} - ${r.product}: ${r.finalSize || "taille manquante"} (${r.sourceSize || "aucune source"})`).join(" | ");

      return {
        Compétiteur: f.competitor,
        Email: f.email,
        Dojo: f.dojo,
        Équipe: f.team,
        "Inscrit Fitofan": "Oui",
        "Engagement signé": engagements.length ? "Oui" : "Non",
        "Commandes liées": orders.join(" | "),
        "A acheté équipement": products.length ? "Oui" : "Non",
        "Produits commandés": productSummary,
        "Tailles Shopify": shopifySizes,
        "Tailles Supabase": supabaseSizes,
        "Tailles finales": finalSizes,
        "Tailles manquantes": products.filter(r => r.missing).length,
        Commentaires: [...new Set(rows.map(r => r.comments).filter(Boolean))].join(" | "),
      };
    });
  }, [fitofan, reconciled]);

  const supabaseLinkedKeys = new Set(reconciled.filter(r => r.supabaseSize).map(r => productKey(r.order, r.productKey)));
  const supabaseAudit = supabase.map(s => ({
    Commande: s.order,
    Compétiteur: s.competitor,
    Produit: s.rawProduct,
    "Produit normalisé": s.productKey,
    Taille: s.size,
    Commentaire: s.comments,
    Type: s.kind,
    Statut: s.kind === "product" && s.size ? (supabaseLinkedKeys.has(productKey(s.order, s.productKey)) ? "Reliée" : "NON RELIÉE") : "Commentaire / engagement conservé",
  }));

  const auditRows = [
    { Indicateur: "Lignes Fitofan compétiteurs", Valeur: fitofan.competitors.length },
    { Indicateur: "Commentaires Fitofan exclus", Valeur: fitofan.comments.length },
    { Indicateur: "Lignes Shopify produits/engagements", Valeur: shopify.length },
    { Indicateur: "Articles actifs analysés", Valeur: reconciled.filter(r => r.kind === "product" && !r.excluded).reduce((s,r) => s + r.effectiveQty, 0) },
    { Indicateur: "Articles exclus annulés/remboursés", Valeur: reconciled.filter(r => r.kind === "product" && r.excluded).length },
    { Indicateur: "Tailles Supabase appliquées", Valeur: reconciled.filter(r => r.supabaseSize).length },
    { Indicateur: "Tailles manquantes", Valeur: reconciled.filter(r => r.missing).length },
    { Indicateur: "Réponses Supabase total", Valeur: supabase.length },
    { Indicateur: "Réponses Supabase produits non reliées", Valeur: supabaseAudit.filter(r => r.Statut === "NON RELIÉE").length },
    { Indicateur: "Total produits fournisseur", Valeur: supplierRows.reduce((s,r) => s + r.Quantité, 0) },
  ];

  const views = {
    audit: { title: "Contrôle qualité", rows: auditRows },
    detail: { title: "Détail global Shopify + Supabase", rows: detailRows },
    fitofan: { title: "Suivi par compétiteur Fitofan", rows: fitofanReport },
    supplier: { title: "Commande fournisseur", rows: supplierRows },
    supabase: { title: "Audit réponses Supabase", rows: supabaseAudit },
    fitofanComments: { title: "Commentaires Fitofan", rows: fitofan.comments },
  };

  const current = views[tab];
  const filteredRows = (current.rows || []).filter(row => !search || norm(Object.values(row).join(" ")).includes(norm(search)));

  return (
    <div className="app">
      <header>
        <h1>Réconciliation Sunfuki V13</h1>
        <p>Fitofan + Shopify + Supabase · Shopify commande, Supabase corrige, Fitofan contrôle.</p>
      </header>

      <section className="uploads">
        <UploadBox title="1. Fitofan" info={files.fitofan || "Importer Liste Cobra Inter et Coach (.xlsx ou .csv)"} onFile={f => upload("fitofan", f)} />
        <UploadBox title="2. Shopify" info={files.shopify || "Importer detail-commandes Shopify (.csv)"} onFile={f => upload("shopify", f)} />
        <UploadBox title="3. Supabase" info={files.supabase || "Importer responses_rows Supabase (.csv)"} onFile={f => upload("supabase", f)} />
      </section>

      <nav>
        {Object.entries(views).map(([id, v]) => <button key={id} className={tab === id ? "active" : ""} onClick={() => setTab(id)}>{v.title}</button>)}
      </nav>

      <section className="toolbar">
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Rechercher nom, commande, produit, taille..." />
        <button onClick={() => exportCsv(`${tab}.csv`, filteredRows)} disabled={!filteredRows.length}>Exporter CSV</button>
      </section>

      <section className="tableCard">
        <h2>{current.title}</h2>
        <DataTable rows={filteredRows} />
      </section>
    </div>
  );
}

function UploadBox({ title, info, onFile }) {
  return (
    <div className="uploadBox">
      <strong>{title}</strong>
      <span>{info}</span>
      <input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={e => onFile(e.target.files?.[0])} />
    </div>
  );
}
function DataTable({ rows }) {
  if (!rows || !rows.length) return <p className="empty">Aucune donnée à afficher.</p>;
  const headers = Object.keys(rows[0]);
  return (
    <div className="tableWrap">
      <table>
        <thead><tr>{headers.map(h => <th key={h}>{h}</th>)}</tr></thead>
        <tbody>
          {rows.map((row, i) => <tr key={i}>{headers.map(h => <td key={h} title={String(row[h] ?? "")}>{String(row[h] ?? "")}</td>)}</tr>)}
        </tbody>
      </table>
    </div>
  );
}
