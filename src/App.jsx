import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./style.css";

const STORAGE_FITOFAN = "sunfuki_fitofan_rows_v14";
const STORAGE_FITOFAN_FILE = "sunfuki_fitofan_file_v14";
const UNIQUE_SIZE_PRODUCTS = new Set(["sac", "patch"]);
const PRODUCT_LABELS = { tshirt:"T-shirt", hoodie:"Hoodie", kangourou:"Kangourou", veste:"Veste", kimono:"Kimono", gants:"Gants", protege_pieds:"Protège-pieds", protege_tibias:"Protège-tibias", casque:"Casque", sac:"Sac", ceinture:"Ceinture", patch:"Patch", engagement:"Engagement", unknown:"Produit inconnu" };
const SIZE_WORDS = [
  "0000","000","00","0","1","2","3","4","5","6","7","8",
  "XXS","XS","S","M","L","XL","XXL","XXXL","2XL","3XL","4XL",
  "2-3T","3-4T","4-5T","5-6T","7-8T","9-10T","11-12T","12-14T","12-14 ans",
  "YOUTH MEDIUM","M ADULTE","M (ADULTE)"
];

function stripAccents(v=""){return String(v).normalize("NFD").replace(/[\u0300-\u036f]/g,"");}
function clean(v=""){return String(v??"").replace(/\u0000/g,"").replace(/[<>]/g,"").replace(/\r?\n/g," ").replace(/\s+/g," ").trim();}
function norm(v=""){return stripAccents(clean(v)).toLowerCase().replace(/[’']/g,"").replace(/[–—−]/g,"-").replace(/\s+/g," ").trim();}
function pick(row,names){const e=Object.entries(row||{});for(const w of names){const f=e.find(([k])=>norm(k)===norm(w));if(f)return clean(f[1]);}for(const w of names){const f=e.find(([k])=>norm(k).includes(norm(w)));if(f)return clean(f[1]);}return "";}
function detectDelimiter(text){const s=text.slice(0,5000);return [",",";","\t"].sort((a,b)=>s.split(b).length-s.split(a).length)[0];}
function parseCsv(text){const d=detectDelimiter(text);const rows=[];let row=[],cell="",q=false;for(let i=0;i<text.length;i++){const ch=text[i],n=text[i+1];if(ch==='"'&&q&&n==='"'){cell+='"';i++;}else if(ch==='"')q=!q;else if(ch===d&&!q){row.push(cell);cell="";}else if((ch==="\n"||ch==="\r")&&!q){if(ch==="\r"&&n==="\n")i++;row.push(cell);if(row.some(v=>clean(v)))rows.push(row);row=[];cell="";}else cell+=ch;}row.push(cell);if(row.some(v=>clean(v)))rows.push(row);if(!rows.length)return[];const h=rows[0].map((x,i)=>clean(x)||`col_${i}`);return rows.slice(1).map(r=>{const o={};h.forEach((x,i)=>o[x]=clean(r[i]||""));return o;});}
async function readFileRows(file){const name=file.name.toLowerCase();if(name.endsWith(".xlsx")||name.endsWith(".xls")){const b=await file.arrayBuffer();const wb=XLSX.read(b,{type:"array"});const all=[];for(const sn of wb.SheetNames){const ws=wb.Sheets[sn];XLSX.utils.sheet_to_json(ws,{defval:"",raw:false}).forEach(r=>all.push({...r,__sheet:sn}));}return all;}return parseCsv(await file.text());}
function exportCsv(filename,rows){if(!rows.length)return;const h=Object.keys(rows[0]);const esc=v=>`"${String(v??"").replace(/"/g,'""')}"`;const csv=[h.join(";"),...rows.map(r=>h.map(x=>esc(r[x])).join(";"))].join("\n");const blob=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);}
function normalizeOrder(o=""){const raw=clean(o);const m=raw.match(/#?\d+/g);if(!m||!m.length)return raw?(raw.startsWith("#")?raw:`#${raw}`):"";return m.map(x=>x.startsWith("#")?x:`#${x}`).join(" | ");}
function splitOrders(o=""){const raw=clean(o);const m=raw.match(/#?\d+/g);if(!m||!m.length)return raw?[normalizeOrder(raw)]:[];return m.map(x=>x.startsWith("#")?x:`#${x}`);}
function competitorKey(n=""){return norm(n);}
function canonicalProduct(raw=""){const n=norm(raw);if(!n)return"unknown";if(n.includes("engagement"))return"engagement";if(n.includes("kangourou"))return"kangourou";if(n.includes("t-shirt")||n.includes("t shirt")||n.includes("tee")||n.includes("tshirt"))return"tshirt";if(n.includes("hoodie")||n.includes("chandail"))return"hoodie";if(n.includes("veste")||n.includes("softshell"))return"veste";if(n.includes("kimono")||n.includes("arawaza")||n.includes("saiko"))return"kimono";if(n.includes("gant"))return"gants";if(n.includes("protege-pied")||n.includes("protege pied")||n.includes("protège-pied")||n.includes("protège pied")||n.includes("pied"))return"protege_pieds";if(n.includes("tibia"))return"protege_tibias";if(n.includes("casque"))return"casque";if(n.includes("sac"))return"sac";if(n.includes("ceinture"))return"ceinture";if(n.includes("patch"))return"patch";return"unknown";}
function productLabel(k){return PRODUCT_LABELS[k]||k||"";}
function hasSize(k){return !UNIQUE_SIZE_PRODUCTS.has(k);}
function extractSize(value=""){
  const original=clean(value);
  const v=norm(original);
  if(!v)return "";
  if(v.includes("engagement")||v.includes("commande equipement")||v.includes("commande équipement"))return "";

  const tailleProp=original.match(/taille\s*[:：]\s*([^|•,;]+)/i);
  if(tailleProp){
    const found=extractSize(tailleProp[1]);
    if(found)return found;
  }

  const youth=v.match(/\b(2\s*-\s*3|3\s*-\s*4|4\s*-\s*5|5\s*-\s*6|7\s*-\s*8|9\s*-\s*10|11\s*-\s*12|12\s*-\s*14)\s*(t|ans?)?\b/i);
  if(youth)return youth[1].replace(/\s+/g,"")+"T";

  if(v.includes("youth medium"))return "YOUTH MEDIUM";
  if(v.includes("m adulte")||v.includes("medium adulte")||v.includes("m (adulte)"))return "M ADULTE";

  const exact=SIZE_WORDS.find(s=>norm(s)===v);
  if(exact)return exact.toUpperCase().replace(" ANS","T").replace(/\s+/g,"");

  const match=v.match(/\b(0000|000|00|xxxl|xxl|xl|xxs|xs|s|m|l|2xl|3xl|4xl|[0-8])\b/i);
  return match?match[1].toUpperCase():"";
}
function productKey(order,prod){return`${normalizeOrder(order)}|${prod}`;}
function isFitofanComment(name,email,dojo){const n=norm(name),e=norm(email);if(!n&&!e&&!dojo)return true;if(!e&&(n.startsWith("merci")||n.includes("commande#")||n.includes("commande #")||n.includes("bonne journee")||n.includes("bonne journée")||n.includes("finalement")||n.length>60))return true;return false;}
function uniqueSorted(values){return[...new Set(values.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),"fr"));}
function parseFitofan(rows){const competitors=[],comments=[];for(const row of rows){const competitor=pick(row,["Noms","Nom","Compétiteur","Competitor","Athlete"]);const email=pick(row,["couriel","courriel","Email","Courriel"]);const dojo=pick(row,["école","ecole","Dojo","École"]);const team=pick(row,["Equipe","Équipe","Team"]);if(isFitofanComment(competitor,email,dojo)){const txt=clean(Object.values(row).filter(Boolean).join(" | "));if(txt)comments.push({Commentaire:txt});continue;}if(!competitor)continue;competitors.push({competitor,key:competitorKey(competitor),email,dojo,team});}return{competitors,comments};}
function findShopifySize(row,rawProduct,productKey){
 if(!hasSize(productKey))return "Taille unique";
 const raw=findShopifySizeRaw(row,rawProduct,productKey);
 return normalizeLooseSize(raw);
}

function normalizeLooseSize(value=""){
 const raw=clean(value); if(!raw)return ""; let v=norm(raw);
 v=v.replace(/taille|grandeur|size|pointure/g," ").replace(/[:：]/g," ").replace(/\s+/g," ").trim();
 const half=v.match(/\b([0-9])\s*(?:1\/2|½)\b/); if(half)return `${half[1]}.5`;
 const dec=v.match(/\b([0-9])\s*[,\.]\s*([05])\b/); if(dec)return `${dec[1]}.${dec[2]}`;
 const youth=v.match(/\b(2\s*-\s*3|3\s*-\s*4|4\s*-\s*5|5\s*-\s*6|7\s*-\s*8|9\s*-\s*10|11\s*-\s*12|12\s*-\s*14)\s*(t|ans?)?\b/i); if(youth)return youth[1].replace(/\s+/g,"")+"T";
 const range=v.match(/\b(xs|s|m|l|xl|xxl|xxxl|2xl|3xl|4xl)\s*[-/]\s*(xs|s|m|l|xl|xxl|xxxl|2xl|3xl|4xl)\b/i); if(range)return `${range[1].toUpperCase()}-${range[2].toUpperCase()}`;
 if(v.includes("youth medium"))return "YOUTH MEDIUM"; if(v.includes("m adulte")||v.includes("medium adulte")||v.includes("m (adulte)"))return "M ADULTE";
 const kim=v.match(/\b([0-9](?:[,.]5)?)\s*\/\s*([0-9]{3})\b/); if(kim)return `${kim[1].replace(",",".")}/${kim[2]}`;
 const alpha=v.match(/\b(0000|000|00|xxxl|xxl|xl|xxs|xs|s|m|l|2xl|3xl|4xl|[0-9])\b/i); if(alpha)return alpha[1].toUpperCase();
 return raw;
}
function isKnownStandardSize(s=""){
 const n=norm(s); if(!n)return false;
 if(SIZE_WORDS.some(x=>norm(x)===n))return true;
 return /^(0000|000|00|0|1|2|3|4|5|6|7|8|xxs|xs|s|m|l|xl|xxl|xxxl|2xl|3xl|4xl)$/i.test(s)
 || /^(2-3|3-4|4-5|5-6|7-8|9-10|11-12|12-14)t$/i.test(s)
 || /^[0-9]\.5$/.test(s)
 || /^[0-9](?:\.5)?\/[0-9]{3}$/.test(s)
 || /^(XS|S|M|L|XL|XXL|XXXL|2XL|3XL|4XL)-(XS|S|M|L|XL|XXL|XXXL|2XL|3XL|4XL)$/i.test(s);
}
function analyzeSize(value="",source=""){
 const raw=clean(value); if(!raw)return {raw:"",normalized:"",status:"MANQUANTE",source};
 const normalized=normalizeLooseSize(raw); if(!normalized)return {raw,normalized:"",status:"MANQUANTE",source};
 return {raw,normalized,status:isKnownStandardSize(normalized)?"STANDARD":"INHABITUELLE",source};
}
function findShopifySizeRaw(row,rawProduct,productKey){
 if(!hasSize(productKey))return "Taille unique";
 const parts=clean(rawProduct).split(/[—–-]/).map(p=>clean(p)).filter(Boolean);
 for(const part of parts){
  const n=norm(part); if(n.includes("equipe")||n.includes("cobra")||n.includes("acompte"))continue;
  const a=analyzeSize(part,"Shopify"); if(a.normalized)return part;
 }
 for(const value of Object.values(row||{})){const txt=clean(value);const prop=txt.match(/taille\s*[:：]\s*([^|•,;]+)/i);if(prop)return clean(prop[1]);}
 return pick(row,["Taille","Size","Variant","Option","Correction taille"]);
}

function parseShopify(rows){const out=[];for(const row of rows){const order=normalizeOrder(pick(row,["Commande","Order","Name","orderName","Order name","Numéro"]));const statusRaw=pick(row,["Statut commande","Status","Financial status","Fulfillment status","Statut"]);const client=pick(row,["Client","Customer","Nom client"]);const competitor=pick(row,["Compétiteur","Competitor","Athlete","Nom compétiteur","Nom du compétiteur"])||client;const email=pick(row,["Courriel compétiteur","Email","Courriel","Customer email","recipient_email"]);const dojo=pick(row,["Dojo","École","Ecole","Location"]);const team=pick(row,["Équipe","Equipe","Team"]);const rawProduct=pick(row,["Produit","Product","Line item name","Title","Item"]);const qty=Number((pick(row,["Quantité","Quantity","Qty"])||"1").replace(",","."))||1;const refunded=Number((pick(row,["Qté remboursée","Qte remboursee","Refunded quantity","quantityRefunded"])||"0").replace(",","."))||0;const er=pick(row,["Qté effective","Qte effective","Effective quantity"]);let effectiveQty=er===""?Math.max(0,qty-refunded):Number(er.replace(",","."));if(Number.isNaN(effectiveQty))effectiveQty=Math.max(0,qty-refunded);const prod=canonicalProduct(rawProduct);if(!rawProduct)continue;const cancelled=norm(statusRaw).includes("annul")||norm(statusRaw).includes("rembours");const excluded=cancelled||effectiveQty<=0;if(prod==="engagement"){out.push({kind:"engagement",order,competitor,ckey:competitorKey(competitor),client,email,dojo,team,rawProduct,productKey:prod,product:productLabel(prod),shopifySize:"",quantity:qty,refunded,effectiveQty,excluded:true,statusRaw,status:"ENGAGEMENT",comments:""});continue;}const explicitSize=pick(row,["Taille","Size","Variant","Option","Correction taille"]);const shopifySize=hasSize(prod)?extractSize(explicitSize):"Taille unique";let shopifySizeRaw=findShopifySizeRaw(row,rawProduct,prod);const shopifySizeAnalysis=analyzeSize(shopifySizeRaw||shopifySize,"Shopify");out.push({kind:"product",order,competitor,ckey:competitorKey(competitor),client,email,dojo,team,rawProduct,productKey:prod,product:productLabel(prod),shopifySizeRaw,shopifySize,shopifySizeStatus:shopifySizeAnalysis.status,quantity:qty,refunded,effectiveQty,excluded,statusRaw,status:excluded?"ANNULÉ / REMBOURSÉ - exclu fournisseur":"",comments:""});}return out;}
function parseSupabase(rows){const out=[];for(const row of rows){const orderRaw=pick(row,["order_number","Commande","Order","orderName","Numéro"]);const orders=splitOrders(orderRaw);const competitor=pick(row,["competitor","Compétiteur","Client","Nom","customerName"]);const email=pick(row,["recipient_email","Email","Courriel","customer_email"]);const rawProduct=pick(row,["product_name","Produit","product","Article","item","Nom produit"]);const sizeRaw=pick(row,["confirmed_size","Taille","size","Nouvelle taille","Réponse"]);const size=normalizeLooseSize(sizeRaw);const sizeAnalysis=analyzeSize(sizeRaw,"Supabase");const comments=pick(row,["comments","Commentaire","comment","Notes","Message","Réponse","response"]);const prod=canonicalProduct(rawProduct);const kind=prod==="engagement"?"engagement":(prod==="unknown"?"comment":"product");const base={kind,orderRaw:normalizeOrder(orderRaw),competitor,ckey:competitorKey(competitor),email,rawProduct,productKey:prod,product:productLabel(prod),sizeRaw,size,sizeStatus:sizeAnalysis.status,comments};if(orders.length)orders.forEach(order=>out.push({...base,order}));else out.push({...base,order:""});}return out;}
async function loadSupabaseResponses(){const url=import.meta.env.VITE_SUPABASE_URL;const key=import.meta.env.VITE_SUPABASE_ANON_KEY;if(!url||!key)throw new Error("Variables Supabase manquantes : VITE_SUPABASE_URL et VITE_SUPABASE_ANON_KEY.");const endpoint=`${url.replace(/\/$/,"")}/rest/v1/responses?select=*&order=created_at.desc`;const res=await fetch(endpoint,{headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json"}});if(!res.ok)throw new Error(`Erreur Supabase ${res.status}: ${await res.text()}`);return await res.json();}


function statusClass(value = "") {
  const v = norm(value);
  if (v.includes("taille manquante") || v.includes("non reliee") || v.includes("non reliée")) return "status warning";
  if (v.includes("annule") || v.includes("annulé") || v.includes("rembours")) return "status muted";
  if (v.includes("supabase")) return "status info";
  if (v.includes("commentaire") || v.includes("engagement")) return "status note";
  if (v.includes("ok")) return "status success";
  return "status";
}
function teamClass(value = "") {
  const v = norm(value);
  if (v.includes("international")) return "teamBadge international";
  if (v.includes("coach")) return "teamBadge coach";
  if (v.includes("cobra")) return "teamBadge cobra";
  return "teamBadge";
}
function productClass(value = "") {
  const v = norm(value);
  if (v.includes("kimono")) return "productBadge kimono";
  if (v.includes("t-shirt") || v.includes("tshirt")) return "productBadge tshirt";
  if (v.includes("gants")) return "productBadge gants";
  if (v.includes("pied")) return "productBadge pieds";
  if (v.includes("hoodie") || v.includes("kangourou")) return "productBadge hoodie";
  return "productBadge";
}
function dataQualityScore(auditRows = []) {
  const get = (name) => Number((auditRows.find(r => r.Indicateur === name) || {}).Valeur || 0);
  const active = get("Articles actifs analysés");
  const missing = get("Tailles manquantes");
  const unlinked = get("Réponses Supabase produits non reliées");
  if (!active) return 0;
  const score = Math.max(0, Math.min(100, Math.round(((active - missing) / active) * 100 - Math.min(15, unlinked))));
  return score;
}

function sourceBoxClass(count, required = true) {
  if (count > 0) return "sourceBox ok";
  return required ? "sourceBox danger" : "sourceBox warning";
}



function sourceStateClass(count) {
  return count > 0 ? "sourceCard sourceOk" : "sourceCard sourceMissing";
}

function sourceStateLabel(count) {
  return count > 0 ? "Données chargées" : "Données manquantes";
}


async function saveFitofanRowsToSupabase(rows, sourceFile = "") {
  const url = import.meta.env.VITE_SUPABASE_URL;
  const key = import.meta.env.VITE_SUPABASE_ANON_KEY;
  if (!url || !key) throw new Error("Variables Supabase manquantes.");

  const parsed = parseFitofan(rows).competitors.map((r) => ({
    competitor: r.competitor,
    email: r.email,
    dojo: r.dojo,
    team: r.team,
    source_file: sourceFile || "Import Fitofan",
  }));

  if (!parsed.length) throw new Error("Aucun compétiteur Fitofan valide à sauvegarder.");

  const endpoint = `${url.replace(/\/$/, "")}/rest/v1/fitofan_competitors`;

  // Remplacement complet : on vide la table puis on réinsère.
  const del = await fetch(`${endpoint}?id=not.is.null`, {
    method: "DELETE",
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
  });
  if (!del.ok) {
    const text = await del.text();
    throw new Error(`Erreur suppression Fitofan Supabase ${del.status}: ${text}`);
  }

  const res = await fetch(endpoint, {
    method: "POST",
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
    body: JSON.stringify(parsed),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Erreur sauvegarde Fitofan Supabase ${res.status}: ${text}`);
  }

  return parsed.length;
}

async function loadFitofanRowsFromSupabase() {
  const url = import.meta.env.VITE_SUPABASE_URL;
  const key = import.meta.env.VITE_SUPABASE_ANON_KEY;
  if (!url || !key) throw new Error("Variables Supabase manquantes.");

  const endpoint = `${url.replace(/\/$/, "")}/rest/v1/fitofan_competitors?select=*&order=competitor.asc`;
  const res = await fetch(endpoint, {
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Erreur chargement Fitofan Supabase ${res.status}: ${text}`);
  }

  const data = await res.json();

  // On remet sous une forme compatible avec parseFitofan().
  return (data || []).map((r) => ({
    Noms: r.competitor || "",
    Email: r.email || "",
    Dojo: r.dojo || "",
    Equipe: r.team || "",
  }));
}


function simpleNameScore(a="",b=""){
 const A=norm(a).split(" ").filter(Boolean),B=norm(b).split(" ").filter(Boolean); if(!A.length||!B.length)return 0;
 let hits=0; for(const x of A){if(B.includes(x))hits++;} for(const x of B){if(A.includes(x))hits++;}
 return Math.min(100,Math.round((hits/(A.length+B.length))*120+(norm(a)===norm(b)?40:0)));
}
function bestFitofanSuggestion(name="",fitofanCompetitors=[]){let best=null;for(const f of fitofanCompetitors){const score=simpleNameScore(name,f.competitor);if(!best||score>best.score)best={...f,score};}return best||{competitor:"",dojo:"",team:"",score:0};}
async function loadManualLinksFromSupabase(){const url=import.meta.env.VITE_SUPABASE_URL,key=import.meta.env.VITE_SUPABASE_ANON_KEY;if(!url||!key)throw new Error("Variables Supabase manquantes.");const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_order_links?select=*&order=created_at.desc`;const res=await fetch(endpoint,{headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json"}});if(!res.ok){const text=await res.text();throw new Error(`Erreur chargement liens manuels ${res.status}: ${text}`);}return await res.json();}
async function saveManualLinkToSupabase(link){const url=import.meta.env.VITE_SUPABASE_URL,key=import.meta.env.VITE_SUPABASE_ANON_KEY;if(!url||!key)throw new Error("Variables Supabase manquantes.");const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_order_links`;const res=await fetch(endpoint,{method:"POST",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=representation"},body:JSON.stringify(link)});if(!res.ok){const text=await res.text();throw new Error(`Erreur sauvegarde lien manuel ${res.status}: ${text}`);}return await res.json();}
async function deleteManualLinkFromSupabase(id){const url=import.meta.env.VITE_SUPABASE_URL,key=import.meta.env.VITE_SUPABASE_ANON_KEY;if(!url||!key)throw new Error("Variables Supabase manquantes.");const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_order_links?id=eq.${encodeURIComponent(id)}`;const res=await fetch(endpoint,{method:"DELETE",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=minimal"}});if(!res.ok){const text=await res.text();throw new Error(`Erreur suppression lien manuel ${res.status}: ${text}`);}}

export default function App(){const [fitofanCloudStatus,setFitofanCloudStatus]=useState("");const[manualLinks,setManualLinks]=useState([]);const[manualLinkStatus,setManualLinkStatus]=useState("");const[manualSelections,setManualSelections]=useState({});const[fitofanRaw,setFitofanRaw]=useState([]);const[shopifyRaw,setShopifyRaw]=useState([]);const[supabaseRaw,setSupabaseRaw]=useState([]);const[files,setFiles]=useState({});const[tab,setTab]=useState("dashboard");const[search,setSearch]=useState("");const[filters,setFilters]=useState({dojo:"",team:"",product:"",competitor:""});const[supabaseStatus,setSupabaseStatus]=useState("");useEffect(()=>{try{const saved=localStorage.getItem(STORAGE_FITOFAN);const savedFile=localStorage.getItem(STORAGE_FITOFAN_FILE);if(saved)setFitofanRaw(JSON.parse(saved));if(savedFile)setFiles(p=>({...p,fitofan:savedFile}));}catch(e){console.warn(e)}},[]);useEffect(()=>{refreshManualLinks();},[]);
async function upload(type,file){if(!file)return;const rows=await readFileRows(file);setFiles(p=>({...p,[type]:`${file.name} (${rows.length} lignes)`}));if(type==="fitofan"){setFitofanRaw(rows);localStorage.setItem(STORAGE_FITOFAN,JSON.stringify(rows));localStorage.setItem(STORAGE_FITOFAN_FILE,`${file.name} (${rows.length} lignes)`);}if(type==="shopify")setShopifyRaw(rows);if(type==="supabase")setSupabaseRaw(rows);}function resetFitofan(){localStorage.removeItem(STORAGE_FITOFAN);localStorage.removeItem(STORAGE_FITOFAN_FILE);setFitofanRaw([]);setFiles(p=>({...p,fitofan:""}));}
  async function saveFitofanCloud() {
    setFitofanCloudStatus("Sauvegarde Fitofan dans Supabase...");
    try {
      const source = files.fitofan || "Import Fitofan";
      const count = await saveFitofanRowsToSupabase(fitofanRaw, source);
      setFitofanCloudStatus(`Fitofan sauvegardé dans Supabase : ${count} compétiteurs`);
    } catch (e) {
      setFitofanCloudStatus(e.message || "Erreur sauvegarde Fitofan");
    }
  }

  async function loadFitofanCloud() {
    setFitofanCloudStatus("Chargement Fitofan depuis Supabase...");
    try {
      const rows = await loadFitofanRowsFromSupabase();
      setFitofanRaw(rows);
      localStorage.setItem(STORAGE_FITOFAN, JSON.stringify(rows));
      localStorage.setItem(STORAGE_FITOFAN_FILE, `Supabase fitofan_competitors (${rows.length} lignes)`);
      setFiles((p) => ({ ...p, fitofan: `Supabase fitofan_competitors (${rows.length} lignes)` }));
      setFitofanCloudStatus(`Fitofan chargé depuis Supabase : ${rows.length} lignes`);
    } catch (e) {
      setFitofanCloudStatus(e.message || "Erreur chargement Fitofan");
    }
  }


async function refreshManualLinks(){setManualLinkStatus("Chargement des liens manuels...");try{const rows=await loadManualLinksFromSupabase();setManualLinks(rows||[]);setManualLinkStatus(`Liens manuels chargés : ${(rows||[]).length}`);}catch(e){setManualLinkStatus(e.message||"Erreur chargement liens manuels");}}
async function createManualLink(order,shopifyName,fitofanName){setManualLinkStatus("Sauvegarde du lien manuel...");try{const f=fitofan.competitors.find(x=>x.competitor===fitofanName);if(!f)throw new Error("Compétiteur Fitofan introuvable.");await saveManualLinkToSupabase({order_number:order,shopify_competitor:shopifyName||"",fitofan_competitor:f.competitor,fitofan_email:f.email||"",fitofan_dojo:f.dojo||"",fitofan_team:f.team||""});await refreshManualLinks();setManualSelections(p=>({...p,[order]:""}));setManualLinkStatus(`Commande ${order} liée à ${f.competitor}`);}catch(e){setManualLinkStatus(e.message||"Erreur sauvegarde lien manuel");}}
async function removeManualLink(id){setManualLinkStatus("Suppression du lien manuel...");try{await deleteManualLinkFromSupabase(id);await refreshManualLinks();setManualLinkStatus("Lien manuel supprimé");}catch(e){setManualLinkStatus(e.message||"Erreur suppression lien manuel");}}
async function refreshSupabase(){setSupabaseStatus("Chargement Supabase...");try{const rows=await loadSupabaseResponses();setSupabaseRaw(rows);setFiles(p=>({...p,supabase:`Supabase responses (${rows.length} lignes)`}));setSupabaseStatus(`Supabase chargé : ${rows.length} lignes`);}catch(e){setSupabaseStatus(e.message||"Erreur Supabase");}}
const fitofan=useMemo(()=>parseFitofan(fitofanRaw),[fitofanRaw]);const shopify=useMemo(()=>parseShopify(shopifyRaw),[shopifyRaw]);const supabase=useMemo(()=>parseSupabase(supabaseRaw),[supabaseRaw]);
const manualLinkByOrder=useMemo(()=>{const m=new Map();manualLinks.forEach(l=>{if(l.order_number)m.set(l.order_number,l);});return m;},[manualLinks]);const reconciled=useMemo(()=>{const corrections=new Map(),commentsByOrder=new Map(),commentsByCompetitor=new Map();supabase.forEach(s=>{const comment=s.comments||(s.kind!=="product"?s.rawProduct:"");if(comment){if(s.order){if(!commentsByOrder.has(s.order))commentsByOrder.set(s.order,[]);commentsByOrder.get(s.order).push(comment);}if(s.ckey){if(!commentsByCompetitor.has(s.ckey))commentsByCompetitor.set(s.ckey,[]);commentsByCompetitor.get(s.ckey).push(comment);}}if(s.kind==="product"&&s.size&&s.order)corrections.set(productKey(s.order,s.productKey),s);});return shopify.map(item=>{const manualLink=manualLinkByOrder.get(item.order);if(manualLink){item={...item,competitor:manualLink.fitofan_competitor||item.competitor,ckey:competitorKey(manualLink.fitofan_competitor||item.competitor),dojo:manualLink.fitofan_dojo||item.dojo,team:manualLink.fitofan_team||item.team,manualLinked:true,manualOriginalCompetitor:item.competitor};}if(item.kind!=="product")return item;const corr=corrections.get(productKey(item.order,item.productKey));let finalSize=corr?.size||item.shopifySize||"";let sourceSize=corr?.size?"Supabase":(item.shopifySize?"Shopify":"");if(!hasSize(item.productKey)){finalSize="Taille unique";sourceSize="Taille unique";}const missing=!finalSize&&hasSize(item.productKey)&&!item.excluded;const comments=[...new Set([...(commentsByOrder.get(item.order)||[]),...(commentsByCompetitor.get(item.ckey)||[])])].join(" | ");let status=item.excluded?item.status:"OK SHOPIFY";if(corr&&!item.shopifySize)status="TAILLE AJOUTÉE PAR SUPABASE";if(corr&&item.shopifySize&&item.shopifySize!==corr.size)status="TAILLE MODIFIÉE PAR SUPABASE";if(corr&&item.shopifySize&&item.shopifySize===corr.size)status="TAILLE CONFIRMÉE PAR SUPABASE";if(missing)status="TAILLE MANQUANTE";if(comments)status+=" + COMMENTAIRE";const finalSizeAnalysis=analyzeSize(finalSize,sourceSize);return{...item,supabaseSizeRaw:corr?.sizeRaw||"",supabaseSize:corr?.size||"",finalSize,finalSizeStatus:missing?"MANQUANTE":finalSizeAnalysis.status,sourceSize,missing,initialShopifySizeMissing:!item.shopifySize&&hasSize(item.productKey)&&!item.excluded,comments,status};});},[shopify,supabase,manualLinkByOrder]);
const detailRows=reconciled.map(r=>({Commande:r.order,Compétiteur:r.competitor,"Nom Shopify original":r.manualOriginalCompetitor||"","Lien manuel":r.manualLinked?"Oui":"Non",Client:r.client,Email:r.email,Dojo:r.dojo,Équipe:r.team,Type:r.kind,"Produit Shopify brut":r.rawProduct,"Produit normalisé":r.productKey,Produit:r.product,"Taille Shopify brute":r.shopifySizeRaw||"","Taille Shopify":r.shopifySize||"","Statut taille Shopify":r.shopifySizeStatus||"","Taille Shopify initialement manquante":r.initialShopifySizeMissing?"Oui":"Non","Taille Supabase brute":r.supabaseSizeRaw||"","Taille Supabase":r.supabaseSize||"","Taille finale":r.finalSize||"","Statut taille finale":r.finalSizeStatus||"","Source taille":r.sourceSize||"",Quantité:r.quantity,"Qté remboursée":r.refunded,"Qté effective":r.effectiveQty,"Exclu fournisseur":r.excluded?"Oui":"Non",Statut:r.status,Commentaires:r.comments||""}));
const unusualSizeRows=reconciled.filter(r=>r.kind==="product"&&!r.excluded&&r.finalSizeStatus==="INHABITUELLE").map(r=>({Compétiteur:r.competitor,Dojo:r.dojo,Équipe:r.team,Commande:r.order,Produit:r.product,"Produit Shopify brut":r.rawProduct,"Taille Shopify brute":r.shopifySizeRaw||"","Taille Supabase brute":r.supabaseSizeRaw||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"",Statut:"TAILLE INHABITUELLE À VÉRIFIER",Commentaires:r.comments||""}));const shopifyInitialMissingRows=reconciled.filter(r=>r.kind==="product"&&r.initialShopifySizeMissing).map(r=>({Compétiteur:r.competitor,Dojo:r.dojo,Équipe:r.team,Commande:r.order,Produit:r.product,"Produit Shopify brut":r.rawProduct,"Taille Shopify initiale":r.shopifySize||"MANQUANTE","Taille Supabase":r.supabaseSize||"Aucune","Taille finale":r.finalSize||"MANQUANTE","Statut régularisation":r.supabaseSize&&r.finalSize?"RÉGULARISÉ PAR SUPABASE":"TOUJOURS MANQUANT",Commentaires:r.comments||""}));const supplierRows=useMemo(()=>{const map=new Map();reconciled.forEach(r=>{if(r.kind!=="product"||r.excluded||!r.finalSize)return;const key=`${r.productKey}|${r.team}|${r.dojo}|${r.finalSize}`;const prev=map.get(key)||{Produit:r.product,"Produit normalisé":r.productKey,Équipe:r.team||"Non précisée",Dojo:r.dojo||"Non précisé",Taille:r.finalSize,Quantité:0};prev.Quantité+=r.effectiveQty;map.set(key,prev);});return Array.from(map.values()).sort((a,b)=>`${a.Dojo}${a.Équipe}${a.Produit}${a.Taille}`.localeCompare(`${b.Dojo}${b.Équipe}${b.Produit}${b.Taille}`));},[reconciled]);
const fitofanReport=useMemo(()=>fitofan.competitors.map(f=>{const rows=reconciled.filter(r=>r.ckey===f.key);const products=rows.filter(r=>r.kind==="product"&&!r.excluded);const engagements=rows.filter(r=>r.kind==="engagement");const orders=[...new Set(rows.map(r=>r.order).filter(Boolean))];return{Compétiteur:f.competitor,Email:f.email,Dojo:f.dojo,Équipe:f.team,"Inscrit Fitofan":"Oui","Engagement signé":engagements.length?"Oui":"Non","Commandes liées":orders.join(" | "),"A acheté équipement":products.length?"Oui":"Non","Produits commandés":products.map(r=>`${r.product} (${r.finalSize||"taille manquante"})`).join(" | "),"Tailles Shopify":products.map(r=>`${r.order} - ${r.product}: ${r.shopifySize||"vide"}`).join(" | "),"Tailles Supabase":products.map(r=>`${r.order} - ${r.product}: ${r.supabaseSize||"aucune"}`).join(" | "),"Tailles finales":products.map(r=>`${r.order} - ${r.product}: ${r.finalSize||"taille manquante"} (${r.sourceSize||"aucune source"})`).join(" | "),"Tailles manquantes":products.filter(r=>r.missing).length,Commentaires:[...new Set(rows.map(r=>r.comments).filter(Boolean))].join(" | ")};}),[fitofan,reconciled]);
const supabaseLinkedKeys=new Set(reconciled.filter(r=>r.supabaseSize).map(r=>productKey(r.order,r.productKey)));const supabaseAudit=supabase.map(s=>({Commande:s.order,Compétiteur:s.competitor,Produit:s.rawProduct,"Produit normalisé":s.productKey,Taille:s.size,Commentaire:s.comments,Type:s.kind,Statut:s.kind==="product"&&s.size?(supabaseLinkedKeys.has(productKey(s.order,s.productKey))?"Reliée":"NON RELIÉE"):"Commentaire / engagement conservé"}));
const auditRows=[{Indicateur:"Lignes Fitofan compétiteurs",Valeur:fitofan.competitors.length},{Indicateur:"Fitofan source",Valeur:files.fitofan||"Non chargé"},{Indicateur:"Commentaires Fitofan exclus",Valeur:fitofan.comments.length},{Indicateur:"Lignes Shopify produits/engagements",Valeur:shopify.length},{Indicateur:"Articles actifs analysés",Valeur:reconciled.filter(r=>r.kind==="product"&&!r.excluded).reduce((s,r)=>s+r.effectiveQty,0)},{Indicateur:"Articles exclus annulés/remboursés",Valeur:reconciled.filter(r=>r.kind==="product"&&r.excluded).length},{Indicateur:"Tailles Supabase appliquées",Valeur:reconciled.filter(r=>r.supabaseSize).length},{Indicateur:"Tailles manquantes",Valeur:reconciled.filter(r=>r.missing).length},{Indicateur:"Réponses Supabase total",Valeur:supabase.length},{Indicateur:"Réponses Supabase produits non reliées",Valeur:supabaseAudit.filter(r=>r.Statut==="NON RELIÉE").length},{Indicateur:"Total produits fournisseur",Valeur:supplierRows.reduce((s,r)=>s+r.Quantité,0)}];
const dashboardRows=detailRows.filter(r=>r.Type==="product"&&r["Exclu fournisseur"]!=="Oui");const filterOptions={dojo:uniqueSorted([...dashboardRows.map(r=>r.Dojo),...fitofanReport.map(r=>r.Dojo)]),team:uniqueSorted([...dashboardRows.map(r=>r.Équipe),...fitofanReport.map(r=>r.Équipe)]),product:uniqueSorted(dashboardRows.map(r=>r.Produit)),competitor:uniqueSorted([...dashboardRows.map(r=>r.Compétiteur),...fitofanReport.map(r=>r.Compétiteur)])};function applyFilters(rows){return rows.filter(row=>{const text=norm(Object.values(row).join(" "));if(search&&!text.includes(norm(search)))return false;if(filters.dojo&&norm(row.Dojo)!==norm(filters.dojo))return false;if(filters.team&&norm(row.Équipe)!==norm(filters.team))return false;if(filters.product&&!norm(Object.values(row).join(" ")).includes(norm(filters.product)))return false;if(filters.competitor&&norm(row.Compétiteur)!==norm(filters.competitor))return false;return true;});}
const dashboardFiltered=applyFilters(dashboardRows);const dashboardStats={competitors:uniqueSorted(dashboardFiltered.map(r=>r.Compétiteur)).length,quantity:dashboardFiltered.reduce((s,r)=>s+(Number(r["Qté effective"])||0),0),missing:dashboardFiltered.filter(r=>r.Statut.includes("TAILLE MANQUANTE")).length,supabase:dashboardFiltered.filter(r=>r["Source taille"]==="Supabase").length,initialMissing:shopifyInitialMissingRows.length,regularized:shopifyInitialMissingRows.filter(r=>r["Statut régularisation"]==="RÉGULARISÉ PAR SUPABASE").length,unusual:unusualSizeRows.length};const competitorsMissingSizes=useMemo(()=>{const map=new Map();dashboardRows.filter(r=>String(r.Statut||"").includes("TAILLE MANQUANTE")).forEach(r=>{const name=r.Compétiteur||"Compétiteur non précisé";const current=map.get(name)||{Compétiteur:name,Dojo:r.Dojo||"",Équipe:r.Équipe||"",Commandes:new Set(),Produits:[]};if(r.Commande)current.Commandes.add(r.Commande);current.Produits.push(`${r.Produit||""} (${r.Commande||""})`);map.set(name,current);});return Array.from(map.values()).map(r=>({Compétiteur:r.Compétiteur,Dojo:r.Dojo,Équipe:r.Équipe,Commandes:Array.from(r.Commandes).join(" | "),"Produits sans taille":r.Produits.join(" | ")}));},[dashboardRows]);const competitorsMissingEngagement=useMemo(()=>fitofanReport.filter(r=>r["Inscrit Fitofan"]==="Oui"&&r["Engagement signé"]!=="Oui").map(r=>({Compétiteur:r.Compétiteur,Dojo:r.Dojo,Équipe:r.Équipe,Email:r.Email,"Commandes liées":r["Commandes liées"],"A acheté équipement":r["A acheté équipement"]})),[fitofanReport]);const linkedOrdersSet=useMemo(()=>new Set(manualLinks.map(l=>l.order_number)),[manualLinks]);const manualReconciliationRows=useMemo(()=>{const knownFitofanKeys=new Set(fitofan.competitors.map(f=>f.key));const orders=new Map();reconciled.forEach(r=>{if(!r.order)return;if(!orders.has(r.order))orders.set(r.order,{Commande:r.order,"Nom Shopify":r.manualOriginalCompetitor||r.competitor||"",Dojo:r.dojo||"",Équipe:r.team||"",Produits:[],DéjàLiée:linkedOrdersSet.has(r.order)?"Oui":"Non"});const o=orders.get(r.order);if(r.kind==="product")o.Produits.push(`${r.product} (${r.finalSize||"taille manquante"})`);});const rows=[];orders.forEach(o=>{const key=competitorKey(o["Nom Shopify"]);if(!knownFitofanKeys.has(key)||o.DéjàLiée==="Oui"){const suggestion=bestFitofanSuggestion(o["Nom Shopify"],fitofan.competitors);rows.push({...o,Produits:o.Produits.join(" | "),"Suggestion Fitofan":suggestion.competitor||"","Score suggestion":suggestion.score||0,Action:o.DéjàLiée==="Oui"?"Déjà liée":"À lier"});}});return rows;},[reconciled,fitofan,linkedOrdersSet]);const manualLinksRows=manualLinks.map(l=>({Commande:l.order_number,"Nom Shopify":l.shopify_competitor,"Compétiteur Fitofan":l.fitofan_competitor,Dojo:l.fitofan_dojo,Équipe:l.fitofan_team,Créé:l.created_at||"",Action:"Supprimer"}));const views={dashboard:{title:"Tableau de bord visuel",rows:dashboardFiltered},manualReconciliation:{title:"Réconciliation manuelle",rows:manualReconciliationRows},manualLinks:{title:"Liens manuels sauvegardés",rows:manualLinksRows},unusualSizes:{title:"Tailles inhabituelles à vérifier",rows:unusualSizeRows},shopifyInitialMissing:{title:"Tailles Shopify manquantes initialement",rows:shopifyInitialMissingRows},missingSizes:{title:"Compétiteurs avec tailles manquantes",rows:competitorsMissingSizes},missingEngagements:{title:"Engagements manquants",rows:competitorsMissingEngagement},audit:{title:"Contrôle qualité",rows:auditRows},detail:{title:"Détail global Shopify + Supabase",rows:detailRows},fitofan:{title:"Suivi par compétiteur Fitofan",rows:fitofanReport},supplier:{title:"Commande fournisseur",rows:supplierRows},supabase:{title:"Audit réponses Supabase",rows:supabaseAudit},fitofanComments:{title:"Commentaires Fitofan",rows:fitofan.comments}};const current=views[tab];const filteredRows=tab==="dashboard"?dashboardFiltered:applyFilters(current.rows||[]);
return <div className="app"><header><h1>Réconciliation Sunfuki V22</h1><p>Sources colorées · alertes compétiteurs · engagements manquants · exports filtrés par club.</p></header><section className="sourceStatusGrid">
        <div className={sourceStateClass(fitofanRaw.length)}>
          <div className="sourceTop">
            <strong>1. Fitofan</strong>
            <span>{fitofanRaw.length > 0 ? "OK" : "MANQUANT"}</span>
          </div>
          <p>{files.fitofan || "Importer une fois. Les données sont ensuite gardées en mémoire locale."}</p>
          <div className="sourceCount">{fitofan.competitors.length} compétiteurs</div>
          <input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={e => upload("fitofan", e.target.files?.[0])} /><div className="cloudActions"><button onClick={saveFitofanCloud} disabled={!fitofanRaw.length}>Sauvegarder Fitofan dans Supabase</button><button onClick={loadFitofanCloud}>Charger Fitofan depuis Supabase</button></div>{fitofanCloudStatus&&<small>{fitofanCloudStatus}</small>}
        </div>

        <div className={sourceStateClass(supabaseRaw.length)}>
          <div className="sourceTop">
            <strong>2. Supabase responses</strong>
            <span>{supabaseRaw.length > 0 ? "OK" : "MANQUANT"}</span>
          </div>
          <p>{files.supabase || "Source normale : chargement automatique depuis la table responses."}</p>
          <div className="sourceCount">{supabaseRaw.length} réponses</div>
          <button onClick={refreshSupabase}>Charger / actualiser Supabase</button>
          {supabaseStatus && <small>{supabaseStatus}</small>}
          <details className="fallbackImport">
            <summary>Import CSV Supabase de secours</summary>
            <p>À utiliser seulement si Supabase est temporairement indisponible ou pour tester un ancien export.</p>
            <input type="file" accept=".csv,.txt" onChange={e => upload("supabase", e.target.files?.[0])} />
          </details>
        </div>

        <div className={sourceStateClass(shopifyRaw.length)}>
          <div className="sourceTop">
            <strong>3. Shopify</strong>
            <span>{shopifyRaw.length > 0 ? "OK" : "MANQUANT"}</span>
          </div>
          <p>{files.shopify || "Importer le CSV Shopify actualisé à chaque analyse."}</p>
          <div className="sourceCount">{shopifyRaw.length} lignes Shopify</div>
          <input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={e => upload("shopify", e.target.files?.[0])} />
        </div>
      </section><section className="actions"><button onClick={resetFitofan}>Réinitialiser Fitofan mémorisé</button></section><section className="filters"><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Recherche libre..."/><Select label="Dojo" value={filters.dojo} options={filterOptions.dojo} onChange={v=>setFilters(p=>({...p,dojo:v}))}/><Select label="Équipe" value={filters.team} options={filterOptions.team} onChange={v=>setFilters(p=>({...p,team:v}))}/><Select label="Produit" value={filters.product} options={filterOptions.product} onChange={v=>setFilters(p=>({...p,product:v}))}/><Select label="Compétiteur" value={filters.competitor} options={filterOptions.competitor} onChange={v=>setFilters(p=>({...p,competitor:v}))}/><button onClick={()=>setFilters({dojo:"",team:"",product:"",competitor:""})}>Effacer filtres</button></section>{tab==="dashboard"&&<>
          <section className="metrics"><Metric label="Compétiteurs filtrés" value={dashboardStats.competitors}/><Metric label="Produits filtrés" value={dashboardStats.quantity}/><Metric label="Tailles Supabase" value={dashboardStats.supabase}/><Metric label="Tailles manquantes" value={dashboardStats.missing} warning={dashboardStats.missing>0}/><Metric label="Shopify sans taille au départ" value={dashboardStats.initialMissing||0}/><Metric label="Régularisées Supabase" value={dashboardStats.regularized||0}/><Metric label="Tailles inhabituelles" value={dashboardStats.unusual||0} warning={(dashboardStats.unusual||0)>0}/></section>

          <section className="missingAlertGrid">
            <div className={competitorsMissingSizes.length > 0 ? "alertCard orange" : "alertCard green"}>
              <div>
                <span>Compétiteurs avec taille manquante</span>
                <strong>{competitorsMissingSizes.length}</strong>
              </div>
              <button onClick={() => setTab("missingSizes")}>Voir la liste</button>
            </div>

            <div className={competitorsMissingEngagement.length > 0 ? "alertCard red" : "alertCard green"}>
              <div>
                <span>Engagements manquants</span>
                <strong>{competitorsMissingEngagement.length}</strong>
              </div>
              <button onClick={() => setTab("missingEngagements")}>Voir la liste</button>
            </div>
          </section></>}<nav>{Object.entries(views).map(([id,v])=><button key={id} className={tab===id?"active":""} onClick={()=>setTab(id)}>{v.title}</button>)}</nav><section className="toolbar"><button onClick={()=>exportCsv(`${tab}-filtre.csv`,filteredRows)} disabled={!filteredRows.length}>Exporter les données filtrées</button><button onClick={()=>exportCsv(`club-${filters.dojo||"tous"}-${filters.team||"toutes-equipes"}.csv`,filteredRows)} disabled={!filteredRows.length}>Export pour club / dojo</button></section>{tab==="manualReconciliation"&&<section className="manualLinkPanel"><h2>Créer un lien manuel</h2><p>Associe une commande Shopify au bon compétiteur Fitofan. Le lien sera sauvegardé dans Supabase.</p>{manualReconciliationRows.filter(r=>r.Action==="À lier").slice(0,25).map(r=><div className="manualLinkRow" key={r.Commande}><div><strong>{r.Commande}</strong><span>{r["Nom Shopify"]} → suggestion : {r["Suggestion Fitofan"]} ({r["Score suggestion"]}%)</span></div><select value={manualSelections[r.Commande]||r["Suggestion Fitofan"]||""} onChange={e=>setManualSelections(p=>({...p,[r.Commande]:e.target.value}))}><option value="">Choisir un compétiteur</option>{fitofan.competitors.map(f=><option key={f.competitor} value={f.competitor}>{f.competitor} — {f.dojo} — {f.team}</option>)}</select><button onClick={()=>createManualLink(r.Commande,r["Nom Shopify"],manualSelections[r.Commande]||r["Suggestion Fitofan"])}>Lier</button></div>)}{manualLinkStatus&&<small>{manualLinkStatus}</small>}</section>}{tab==="manualLinks"&&<section className="manualLinkPanel"><h2>Liens sauvegardés</h2>{manualLinks.map(l=><div className="manualLinkRow" key={l.id}><div><strong>{l.order_number}</strong><span>{l.shopify_competitor} → {l.fitofan_competitor}</span></div><button onClick={()=>removeManualLink(l.id)}>Supprimer</button></div>)}</section>}<section className="tableCard"><h2>{current.title}</h2><DataTable rows={filteredRows}/></section></div>}
function Select({label,value,options,onChange}){return <label><span>{label}</span><select value={value} onChange={e=>onChange(e.target.value)}><option value="">Tous</option>{options.map(o=><option key={o} value={o}>{o}</option>)}</select></label>}
function Metric({label,value,warning}){return <div className={`metric ${warning?"warning":""}`}><span>{label}</span><strong>{value}</strong></div>}
function UploadBox({title,info,onFile}){return <div className="uploadBox"><strong>{title}</strong><span>{info}</span><input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={e=>onFile(e.target.files?.[0])}/></div>}
function DataTable({rows}){if(!rows||!rows.length)return <p className="empty">Aucune donnée à afficher.</p>;const headers=Object.keys(rows[0]);return <div className="tableWrap"><table><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.map((row,i)=><tr key={i}>{headers.map(h=><td key={h} title={String(row[h]??"")}>{String(row[h]??"")}</td>)}</tr>)}</tbody></table></div>}
