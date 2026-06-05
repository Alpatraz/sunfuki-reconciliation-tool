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
function parseFitofan(rows){const competitors=[],comments=[];for(const row of rows){const competitor=pick(row,["Noms","Nom","Compétiteur","Competitor","Athlete"]);const email=pick(row,["couriel","courriel","Email","Courriel"]);const dojo=pick(row,["école","ecole","Dojo","École"]);const team=pick(row,["Equipe","Équipe","Team"]);if(isFitofanComment(competitor,email,dojo)){const txt=clean(Object.values(row).filter(Boolean).join(" | "));if(txt)comments.push({Commentaire:txt});continue;}if(!competitor)continue;competitors.push({competitor,key:competitorKey(competitor),email,dojo,team,fitofanStatus:pickFitofanStatus(row)});}return{competitors,comments};}
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
 const normalized=normalizeManualSizeForSave(raw); if(!normalized)return {raw,normalized:"",status:"MANQUANTE",source};
 if(isKnownStandardSizeList(normalized))return {raw,normalized,status:"STANDARD",source};
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

function parseShopify(rows){const out=[];for(const row of rows){const order=normalizeOrder(pick(row,["Commande","Order","Name","orderName","Order name","Numéro"]));const statusRaw=pick(row,["Statut commande","Status","Financial status","Fulfillment status","Statut"]);const client=pick(row,["Client","Customer","Nom client"]);const competitor=pick(row,["Compétiteur","Competitor","Athlete","Nom compétiteur","Nom du compétiteur"])||client;const email=pick(row,["Courriel compétiteur","Email","Courriel","Customer email","recipient_email"]);const dojo=pick(row,["Dojo","École","Ecole","Location"]);const team=pick(row,["Équipe","Equipe","Team"]);const rawProduct=pick(row,["Produit","Product","Line item name","Title","Item"]);const qty=Number((pick(row,["Quantité","Quantity","Qty"])||"1").replace(",","."))||1;const refunded=Number((pick(row,["Qté remboursée","Qte remboursee","Refunded quantity","quantityRefunded"])||"0").replace(",","."))||0;const er=pick(row,["Qté effective","Qte effective","Effective quantity"]);let effectiveQty=er===""?Math.max(0,qty-refunded):Number(er.replace(",","."));if(Number.isNaN(effectiveQty))effectiveQty=Math.max(0,qty-refunded);const prod=supplierProductKey(rawProduct);if(!rawProduct)continue;const cancelled=norm(statusRaw).includes("annul")||norm(statusRaw).includes("rembours");const excluded=cancelled||effectiveQty<=0;if(prod==="engagement"){out.push({kind:"engagement",order,competitor,ckey:competitorKey(competitor),client,email,dojo,team,rawProduct,productKey:prod,product:supplierProductName(rawProduct,prod),shopifySize:"",quantity:qty,refunded,effectiveQty,excluded:true,statusRaw,status:"ENGAGEMENT",comments:""});continue;}const explicitSize=pick(row,["Taille","Size","Variant","Option","Correction taille"]);const shopifySize=hasSize(prod)?extractSize(explicitSize):"Taille unique";let shopifySizeRaw=findShopifySizeRaw(row,rawProduct,prod);const shopifySizeAnalysis=analyzeSize(shopifySizeRaw||shopifySize,"Shopify");out.push({kind:"product",order,competitor,ckey:competitorKey(competitor),client,email,dojo,team,rawProduct,productKey:prod,product:supplierProductName(rawProduct,prod),shopifySizeRaw,shopifySize,shopifySizeStatus:shopifySizeAnalysis.status,quantity:qty,refunded,effectiveQty,excluded,statusRaw,status:excluded?"ANNULÉ / REMBOURSÉ - exclu fournisseur":"",comments:""});}return out;}
function parseSupabase(rows){const out=[];for(const row of rows){const orderRaw=pick(row,["order_number","Commande","Order","orderName","Numéro"]);const orders=splitOrders(orderRaw);const competitor=pick(row,["competitor","Compétiteur","Client","Nom","customerName"]);const email=pick(row,["recipient_email","Email","Courriel","customer_email"]);const rawProduct=pick(row,["product_name","Produit","product","Article","item","Nom produit"]);const sizeRaw=pick(row,["confirmed_size","Taille","size","Nouvelle taille","Réponse"]);const size=normalizeLooseSize(sizeRaw);const sizeAnalysis=analyzeSize(sizeRaw,"Supabase");const comments=pick(row,["comments","Commentaire","comment","Notes","Message","Réponse","response"]);const prod=supplierProductKey(rawProduct);const kind=prod==="engagement"?"engagement":(prod==="unknown"?"comment":"product");const base={kind,orderRaw:normalizeOrder(orderRaw),competitor,ckey:competitorKey(competitor),email,rawProduct,productKey:prod,product:supplierProductName(rawProduct,prod),sizeRaw,size,sizeStatus:sizeAnalysis.status,comments};if(orders.length)orders.forEach(order=>out.push({...base,order}));else out.push({...base,order:""});}return out;}
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
    fitofan_status: r.fitofanStatus || ""}));

  if (!parsed.length) throw new Error("Aucun compétiteur Fitofan valide à sauvegarder.");

  const endpoint = `${url.replace(/\/$/, "")}/rest/v1/fitofan_competitors`;

  // Remplacement complet : on vide la table puis on réinsère.
  const del = await fetch(`${endpoint}?id=not.is.null`, {
    method: "DELETE",
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
      Prefer: "return=minimal"}});
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
      Prefer: "return=minimal"},
    body: JSON.stringify(parsed)});
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
      "Content-Type": "application/json"}});
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
    FITOFAN: r.fitofan_status || ""}));
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


function isEngagementOrderRow(r){
  const txt=norm(`${r.rawProduct||""} ${r.product||""} ${r.productKey||""} ${r.status||""} ${r.kind||""}`);
  return txt.includes("engagement");
}


async function loadManualCompetitorsFromSupabase(){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_competitors?select=*&order=competitor.asc`;
  const res=await fetch(endpoint,{headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json"}});
  if(!res.ok){const text=await res.text();throw new Error(`Erreur chargement participants manuels ${res.status}: ${text}`);}
  return await res.json();
}
async function saveManualCompetitorToSupabase(row){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_competitors`;
  const res=await fetch(endpoint,{method:"POST",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=representation"},body:JSON.stringify(row)});
  if(!res.ok){const text=await res.text();throw new Error(`Erreur sauvegarde participant manuel ${res.status}: ${text}`);}
  return await res.json();
}
async function loadManualSizesFromSupabase(){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_size_overrides?select=*&order=created_at.desc`;
  const res=await fetch(endpoint,{headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json"}});
  if(!res.ok){const text=await res.text();throw new Error(`Erreur chargement tailles manuelles ${res.status}: ${text}`);}
  return await res.json();
}
async function saveManualSizeToSupabase(row){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/manual_size_overrides`;
  const res=await fetch(endpoint,{method:"POST",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=representation"},body:JSON.stringify(row)});
  if(!res.ok){const text=await res.text();throw new Error(`Erreur sauvegarde taille manuelle ${res.status}: ${text}`);}
  return await res.json();
}

async function loadCommentCorrectionsFromSupabase(){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/comment_corrections?select=*&order=created_at.desc`;
  const res=await fetch(endpoint,{headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json"}});
  if(!res.ok){const text=await res.text();throw new Error(`Erreur chargement validations commentaires ${res.status}: ${text}`);}
  return await res.json();
}
async function saveCommentCorrectionToSupabase(row){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/comment_corrections`;
  const res=await fetch(endpoint,{method:"POST",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=representation"},body:JSON.stringify(row)});
  if(!res.ok){const text=await res.text();throw new Error(`Erreur sauvegarde validation commentaire ${res.status}: ${text}`);}
  return await res.json();
}

function sizeNeedsValidation(row){if(!row||row.kind!=="product"||row.excluded)return false;if(!hasSize(row.productKey))return false;if(row.missing)return true;if(row.finalSizeStatus==="INHABITUELLE")return true;return false;}
function sizeValidationStatus(row){if(!row||row.excluded)return "EXCLU";if(!hasSize(row.productKey))return "VALIDÉE";if(row.sourceSize==="Manuel")return "VALIDÉE";if(row.missing)return "MANQUANTE";if(row.finalSizeStatus==="INHABITUELLE")return "À VÉRIFIER";return "VALIDÉE";}
function isSupplierReady(row){return row&&row.kind==="product"&&!row.excluded&&row.finalSize&&sizeValidationStatus(row)==="VALIDÉE";}


async function loadEngagementLinkCorrectionsFromSupabase(){
 const url=import.meta.env.VITE_SUPABASE_URL; const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
 if(!url||!key)throw new Error("Variables Supabase manquantes.");
 const endpoint=`${url.replace(/\/$/,"")}/rest/v1/engagement_link_corrections?select=*&order=created_at.desc`;
 const res=await fetch(endpoint,{headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json"}});
 if(!res.ok){const text=await res.text();throw new Error(`Erreur chargement corrections engagement ${res.status}: ${text}`);}
 return await res.json();
}
async function saveEngagementLinkCorrectionToSupabase(row){
 const url=import.meta.env.VITE_SUPABASE_URL; const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
 if(!url||!key)throw new Error("Variables Supabase manquantes.");
 const endpoint=`${url.replace(/\/$/,"")}/rest/v1/engagement_link_corrections`;
 const res=await fetch(endpoint,{method:"POST",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=representation"},body:JSON.stringify(row)});
 if(!res.ok){const text=await res.text();throw new Error(`Erreur sauvegarde correction engagement ${res.status}: ${text}`);}
 return await res.json();
}
async function deleteEngagementLinkCorrectionFromSupabase(id){
 const url=import.meta.env.VITE_SUPABASE_URL; const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
 if(!url||!key)throw new Error("Variables Supabase manquantes.");
 const endpoint=`${url.replace(/\/$/,"")}/rest/v1/engagement_link_corrections?id=eq.${encodeURIComponent(id)}`;
 const res=await fetch(endpoint,{method:"DELETE",headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=minimal"}});
 if(!res.ok){const text=await res.text();throw new Error(`Erreur suppression correction engagement ${res.status}: ${text}`);}
}
function engagementCorrectionKey(order,competitor,team){return `${normalizeOrder(order)}|${competitorKey(competitor)}|${canonicalTeamKey(team)}`;}


async function saveShopifyRowsToSupabase(items, sourceFile=""){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");
  const endpoint=`${url.replace(/\/$/,"")}/rest/v1/shopify_order_items`;

  const rows=(items||[]).map((r,idx)=>({
    unique_key:`${normalizeOrder(r.order)}|${r.kind||""}|${r.productKey||r.product||""}|${r.rawProduct||""}|${r.shopifySize||""}|${idx}`,
    order_number:normalizeOrder(r.order),
    kind:r.kind||"",
    competitor:r.competitor||"",
    email:r.email||"",
    dojo:r.dojo||"",
    team:r.team||"",
    product_key:r.productKey||"",
    product:r.product||"",
    raw_product:r.rawProduct||"",
    shopify_size_raw:r.shopifySizeRaw||"",
    shopify_size:r.shopifySize||"",
    shopify_size_status:r.shopifySizeStatus||"",
    quantity:Number(r.quantity)||0,
    refunded:Number(r.refunded)||0,
    effective_qty:Number(r.effectiveQty)||0,
    excluded:!!r.excluded,
    status_raw:r.statusRaw||"",
    status:r.status||"",
    comments:r.comments||"",
    source_file:sourceFile||"Import Shopify",
    imported_at:new Date().toISOString()
  })).filter(r=>r.order_number);

  if(!rows.length)throw new Error("Aucune ligne Shopify valide à sauvegarder.");

  // V33.19 : remplacement complet du dernier import Shopify
  const del=await fetch(`${endpoint}?id=not.is.null`,{
    method:"DELETE",
    headers:{
      apikey:key,
      Authorization:`Bearer ${key}`,
      "Content-Type":"application/json",
      Prefer:"return=minimal"
    }
  });

  if(!del.ok){
    const text=await del.text();
    throw new Error(`Erreur purge Shopify Supabase ${del.status}: ${text}`);
  }

  for(let i=0;i<rows.length;i+=500){
    const batch=rows.slice(i,i+500);
    const res=await fetch(endpoint,{
      method:"POST",
      headers:{
        apikey:key,
        Authorization:`Bearer ${key}`,
        "Content-Type":"application/json",
        Prefer:"return=minimal"
      },
      body:JSON.stringify(batch)
    });

    if(!res.ok){
      const text=await res.text();
      throw new Error(`Erreur sauvegarde Shopify Supabase ${res.status}: ${text}`);
    }
  }

  return rows.length;
}

async function loadShopifyRowsFromSupabase(){
  const url=import.meta.env.VITE_SUPABASE_URL;
  const key=import.meta.env.VITE_SUPABASE_ANON_KEY;
  if(!url||!key)throw new Error("Variables Supabase manquantes.");

  const base=`${url.replace(/\/$/,"")}/rest/v1/shopify_order_items?select=*&order=id.asc`;
  const pageSize=1000;
  let from=0;
  let all=[];

  while(true){
    const to=from+pageSize-1;
    const res=await fetch(base,{
      headers:{
        apikey:key,
        Authorization:`Bearer ${key}`,
        "Content-Type":"application/json",
        Range:`${from}-${to}`,
        Prefer:"count=exact"
      }
    });
    if(!res.ok){
      const text=await res.text();
      throw new Error(`Erreur chargement Shopify Supabase ${res.status}: ${text}`);
    }
    const batch=await res.json();
    all=all.concat(batch||[]);
    if(!batch || batch.length<pageSize)break;
    from+=pageSize;
    if(from>50000)throw new Error("Sécurité : plus de 50 000 lignes Shopify à charger.");
  }

  return (all||[]).map(r=>({
    kind:r.kind||"",
    order:r.order_number||"",
    competitor:r.competitor||"",
    email:r.email||"",
    dojo:r.dojo||"",
    team:r.team||"",
    productTeam:r.product_team||productTeamFromTitle(r.raw_product||""),
    ckey:competitorKey(r.competitor||""),
    productKey:r.product_key||"",
    product:r.product||"",
    rawProduct:r.raw_product||"",
    shopifySizeRaw:r.shopify_size_raw||"",
    shopifySize:r.shopify_size||"",
    shopifySizeStatus:r.shopify_size_status||"",
    quantity:Number(r.quantity)||0,
    refunded:Number(r.refunded)||0,
    effectiveQty:Number(r.effective_qty)||0,
    excluded:!!r.excluded,
    statusRaw:r.status_raw||"",
    status:r.status||"",
    comments:r.comments||""
  }));
}


function buildResolvedEngagementState({ fitofan, reconciled, engagementCorrections }) {
  const corrections = engagementCorrections || [];
  const unlinks = new Set();
  const manualLinks = new Set();

  corrections.forEach((c) => {
    const key = engagementCorrectionKey(c.order_number, c.competitor, c.team);
    if (c.action === "unlink") unlinks.add(key);
    if (c.action === "link") manualLinks.add(key);
  });

  const engagementRows = (reconciled || []).filter((r) => isEngagementOrderRow(r));

  const engagementByCompetitorTeam = new Map();
  const fitofanRows = fitofan?.competitors || [];

  fitofanRows.forEach((f) => {
    const fKey = `${f.key}|${canonicalTeamKey(f.team)}`;
    const linked = [];

    engagementRows.forEach((e) => {
      const autoMatch =
        competitorKey(e.competitor) === f.key &&
        sameTeam(e.team,f.team);

      const manualMatch = manualLinks.has(
        engagementCorrectionKey(e.order, f.competitor, f.team)
      );

      const manuallyUnlinked = unlinks.has(
        engagementCorrectionKey(e.order, f.competitor, f.team)
      );

      if ((autoMatch || manualMatch) && !manuallyUnlinked) {
        linked.push(e);
      }
    });

    engagementByCompetitorTeam.set(fKey, linked);
  });

  const missing = fitofanRows
    .filter((f) => {
      const key = `${f.key}|${canonicalTeamKey(f.team)}`;
      return !(engagementByCompetitorTeam.get(key) || []).length;
    })
    .map((f) => ({
      Compétiteur: f.competitor,
      Dojo: f.dojo,
      Équipe: f.team,
      Email: f.email || "",
      Engagement: "MANQUANT",
      "Action suggérée": "Gestion liens engagements"}));

  const managementRows = [];
  fitofanRows.forEach((f) => {
    const key = `${f.key}|${canonicalTeamKey(f.team)}`;
    const linked = engagementByCompetitorTeam.get(key) || [];

    if (linked.length) {
      linked.forEach((e) => {
        managementRows.push({
          Compétiteur: f.competitor,
          Dojo: f.dojo,
          Équipe: f.team,
          Commande: e.order,
          "Nom Shopify": e.manualOriginalCompetitor || e.competitor || "",
          Statut: "LIÉ",
          Action: "Délier"});
      });
    } else {
      managementRows.push({
        Compétiteur: f.competitor,
        Dojo: f.dojo,
        Équipe: f.team,
        Commande: "",
        "Nom Shopify": "",
        Statut: "AUCUN ENGAGEMENT LIÉ",
        Action: "Lier"});
    }
  });

  return {
    missing,
    managementRows,
    engagementByCompetitorTeam};
}


function stripProductNoise(raw=""){
  let s=clean(raw);
  s=s.replace(/\s+—\s+acompte\s*50\s*%/gi,"");
  s=s.replace(/\s+-\s+acompte\s*50\s*%/gi,"");
  s=s.replace(/\s+acompte\s*50\s*%/gi,"");
  s=s.replace(/\s+50\s*%/gi,"");
  s=s.replace(/\s+[-–—]\s*$/g,"");
  s=s.replace(/\s{2}/g," ").trim();
  return s;
}
function supplierProductName(raw="", fallback=""){
  const r=clean(raw||fallback);
  const n=norm(r);
  if(!r)return fallback||"Produit";
  if(n.includes("engagement"))return "Engagement";
  if(n.includes("saiko")||n.includes("saïko")||n.includes("arawaza")) return "Kimono Saïko (type Arawaza)";
  if(n.includes("combat")) return "Kimono Combat";
  if(n.includes("noir") && (n.includes("croise")||n.includes("croisé"))) return "Kimono noir croisé";
  if(n.includes("blanc") && (n.includes("100")||n.includes("coton"))) return "Kimono blanc 100% coton";
  if(n.includes("blanc")) return "Kimono blanc";
  if(n.includes("kimono")) return stripProductNoise(r)||"Kimono";
  if(n.includes("t-shirt")||n.includes("t shirt")||n.includes("tee")||n.includes("tshirt")) return "T-shirt";
  if(n.includes("kangourou")||n.includes("hoodie")) return "Kangourou / Hoodie";
  if(n.includes("veste")||n.includes("softshell")) return "Veste";
  if(n.includes("protege-pied")||n.includes("protège-pied")||n.includes("protege pied")||n.includes("protège pied")) return "Protège-pieds";
  if(n.includes("gant")) return "Gants";
  if(n.includes("casque")) return "Casque";
  if(n.includes("tibia")) return "Protège-tibias";
  if(n.includes("sac")) return stripProductNoise(r)||"Sac";
  return stripProductNoise(r)||fallback||"Produit";
}
function supplierProductKey(raw="", fallback=""){
  return norm(supplierProductName(raw,fallback));
}


function canonicalTeamKey(value=""){
  const n=norm(value);
  if(!n)return "";
  if(n.includes("assistant"))return "assistant coach";
  if(n.includes("coach"))return "coach";
  if(n.includes("international") || n.includes("inter"))return "international";
  if(n.includes("cobra"))return "cobra";
  return n;
}
function sameTeam(a="",b=""){
  const aa=canonicalTeamKey(a);
  const bb=canonicalTeamKey(b);
  if(!aa || !bb)return true;
  return aa===bb;
}


function normalizeFitofanStatus(value=""){
  const n=norm(value);
  if(!n)return "";
  if(["oui","yes","y","1","true","inscrit","fitofan"].includes(n))return "oui";
  if(["non","no","n","0","false","pas inscrit","pas fitofan"].includes(n))return "non";
  if(n.includes("oui"))return "oui";
  if(n.includes("non"))return "non";
  return clean(value);
}
function isFitofanYes(value=""){return normalizeFitofanStatus(value)==="oui";}
function isFitofanNo(value=""){return normalizeFitofanStatus(value)==="non";}


function pickFitofanStatus(row){
  const entries=Object.entries(row||{});
  const exact=entries.find(([k])=>norm(k).replace(/\s+/g,"")==="fitofan");
  if(exact)return normalizeFitofanStatus(exact[1]);

  const loose=entries.find(([k])=>{
    const kk=norm(k).replace(/\s+/g,"");
    return kk.includes("fitofan") || kk.includes("inscritfitofan") || kk.includes("presencefitofan");
  });
  if(loose)return normalizeFitofanStatus(loose[1]);

  return "";
}


function productTeamFromTitle(raw=""){
  const n=norm(raw);
  if(!n)return "";
  if(n.includes("assistant coach"))return "Assistant Coach";
  if(n.includes("coach"))return "Coach";
  if(n.includes("international cobra") || n.includes("cobra international"))return "International";
  if(n.includes("equipe international") || n.includes("équipe international"))return "International";
  if(n.includes("equipe cobra") || n.includes("équipe cobra"))return "Cobra";
  return "";
}
function displayProductTeam(itemTeam="", competitorTeam="", kind="product"){
  if(kind!=="product")return competitorTeam||"";
  return itemTeam || "Sans équipe";
}


function supplierTeamForRow(r){
  return displayProductTeam(r?.productTeam||"", "", r?.kind||"product");
}


function supplierLineStatus(r){
  if(!r)return "INCONNU";
  if(r.kind!=="product")return "HORS FOURNISSEUR";
  if(r.excluded)return "EXCLU";
  if(!hasSize(r.productKey))return "OK";
  if(!r.finalSize)return "TAILLE MANQUANTE";
  if(r.finalSizeStatus==="INHABITUELLE")return "TAILLE INHABITUELLE";
  return "OK";
}
function supplierFinalSizeLabel(r){
  if(!r)return "";
  if(!hasSize(r.productKey))return "Taille unique";
  if(!r.finalSize)return "TAILLE MANQUANTE";
  if(r.finalSizeStatus==="INHABITUELLE")return `À VALIDER: ${r.finalSize}`;
  return r.finalSize;
}
function supplierIncludeRow(r){
  return r&&r.kind==="product"&&!r.excluded;
}


function orderSearchMatch(row, query){
  const qRaw=clean(query);
  const q=norm(qRaw);
  const qOrder=normalizeOrder(qRaw);
  if(!q && !qOrder)return true;

  const order=clean(row?.order||row?.Commande||"");
  const orderNorm=normalizeOrder(order);
  const text=norm([
    row?.order,
    row?.Commande,
    orderNorm,
    `#${orderNorm}`,
    row?.competitor,
    row?.Compétiteur,
    row?.email,
    row?.Email,
    row?.dojo,
    row?.Dojo,
    row?.team,
    row?.Équipe,
    row?.product,
    row?.Produit,
    row?.rawProduct,
    row?.["Produit Shopify brut"],
    row?.finalSize,
    row?.["Taille finale"],
    row?.shopifySize,
    row?.["Taille Shopify"]
  ].filter(Boolean).join(" "));

  return text.includes(q) || (!!qOrder && orderNorm.includes(qOrder));
}

function hasUsefulComment(value=""){
  const c=clean(value);
  if(!c)return false;
  const n=norm(c);
  if(["non","none","aucun","aucune","na","n/a","ras","ok",""].includes(n))return false;
  return true;
}


function displayTeamCategory(value=""){
  const n=norm(value);
  if(!n)return "";
  if(n.includes("assistant"))return "Assistant-Coach";
  if(n.includes("coach"))return "Coach";
  if(n.includes("proprietaire")||n.includes("propriétaire")||n.includes("owner"))return "Propriétaire";
  if(n.includes("karate sunfuki")||n.includes("karaté sunfuki")||n==="sunfuki")return "Karaté Sunfuki";
  if(n.includes("international")||n.includes("inter"))return "Cobra International";
  if(n.includes("cobra"))return "Cobra";
  return clean(value);
}
function teamCategoryOptions(){
  return ["Cobra","Cobra International","Coach","Assistant-Coach","Propriétaire","Karaté Sunfuki"];
}


function splitManualSizeList(value=""){
  const raw=clean(value);
  if(!raw)return [];
  // Séparateurs acceptés pour plusieurs tailles : M / L, M | L, M ; L, M + L, M et L.
  // On ne coupe pas sur "-" pour ne pas casser 11-12T.
  const prepared=raw
    .replace(/\s+\bet\b\s+/gi," | ")
    .replace(/[;,|+]/g," | ")
    .replace(/\s+\/\s+/g," | ");
  const parts=prepared.split("|").map(x=>clean(x)).filter(Boolean);
  return parts.length?parts:[raw];
}
function normalizeManualSizeForSave(value=""){
  const parts=splitManualSizeList(value);
  if(parts.length<=1)return normalizeLooseSize(value);
  return parts.map(p=>normalizeLooseSize(p)).filter(Boolean).join(" | ");
}
function isMultiFinalSize(value=""){
  return clean(value).includes("|");
}
function finalSizeParts(value=""){
  const raw=clean(value);
  if(!raw)return [];
  return raw.split("|").map(x=>clean(x)).filter(Boolean);
}
function isKnownStandardSizeList(value=""){
  const parts=finalSizeParts(value);
  return parts.length>1 && parts.every(p=>isKnownStandardSize(p));
}
function expandRowByFinalSizes(row){
  const parts=finalSizeParts(row?.finalSize||"");
  const qty=Number(row?.effectiveQty ?? row?.quantity ?? 1)||1;
  if(parts.length<=1)return [row];
  // Si plusieurs tailles sont saisies, chaque taille représente 1 article.
  // Exemple Qté 2 + "XL | 12-14T" => 2 lignes de Qté 1.
  // Si le nombre de tailles est différent de la quantité, on garde quand même
  // une ligne par taille pour éviter d’écraser les tailles distinctes.
  return parts.map((size,idx)=>({
    ...row,
    finalSize:size,
    effectiveQty:1,
    quantity:1,
    unitIndex:idx,
    unitCount:parts.length,
    _multiSizeSource:row.finalSize,
    _originalQty:qty
  }));
}
function expandRowsByFinalSizes(rows=[]){
  const out=[];
  (rows||[]).forEach(r=>expandRowByFinalSizes(r).forEach(x=>out.push(x)));
  return out;
}



function multiSizeHelpForQty(row={}){
  const qty=Number(row.effectiveQty ?? row.quantity ?? row.Quantité ?? 1)||1;
  return qty>1?`Qté ${qty} : tu peux saisir ${qty} tailles, ex. XL / 12-14T`:"";
}
function quantityDisplay(row={}){
  return Number(row.effectiveQty ?? row.quantity ?? row.Quantité ?? 1)||1;
}
function multiSizeCountWarning(sizeValue="",row={}){
  const qty=quantityDisplay(row);
  const parts=splitManualSizeList(sizeValue);
  if(qty>1 && parts.length>1 && parts.length!==qty){
    return `Attention : ${parts.length} taille(s) saisie(s) pour une quantité de ${qty}.`;
  }
  return "";
}
function commentCorrectionKey(row={}){
  const order=normalizeOrder(row.order||row.Commande||row.order_number||"");
  const productKey=clean(row.productKey||row.product_key||row.Produit||row.product||"");
  const raw=clean(row.rawProduct||row.raw_product||row["Produit Shopify brut"]||"");
  return `${order}|${productKey}|${raw}`;
}

export default function App(){const [shopifyCloudStatus,setShopifyCloudStatus]=useState("");const [shopifyCloudRows,setShopifyCloudRows]=useState([]);const [manualCompetitors,setManualCompetitors]=useState([]);const [manualSizes,setManualSizes]=useState([]);const [manualParticipantStatus,setManualParticipantStatus]=useState("");const [newParticipant,setNewParticipant]=useState({competitor:"",email:"",dojo:"",team:""});const [manualSizeInputs,setManualSizeInputs]=useState({});const [orderSizeEditSearch,setOrderSizeEditSearch]=useState("");const [commentSizeReviewSearch,setCommentSizeReviewSearch]=useState("");const [commentCorrections,setCommentCorrections]=useState([]);const [commentCorrectionInputs,setCommentCorrectionInputs]=useState({});const [commentCorrectionStatus,setCommentCorrectionStatus]=useState("");const [savedSizeKeys,setSavedSizeKeys]=useState(new Set());const [engagementCorrections,setEngagementCorrections]=useState([]);const [engagementCorrectionStatus,setEngagementCorrectionStatus]=useState("");const [engagementLinkInputs,setEngagementLinkInputs]=useState({});const [engagementSearch,setEngagementSearch]=useState("");const [engagementStatusFilter,setEngagementStatusFilter]=useState("all");const [fitofanCloudStatus,setFitofanCloudStatus]=useState("");const[manualLinks,setManualLinks]=useState([]);const[manualLinkStatus,setManualLinkStatus]=useState("");const[manualSelections,setManualSelections]=useState({});const[fitofanRaw,setFitofanRaw]=useState([]);const[shopifyRaw,setShopifyRaw]=useState([]);const[supabaseRaw,setSupabaseRaw]=useState([]);const[files,setFiles]=useState({});const[tab,setTab]=useState("dashboard");const[search,setSearch]=useState("");const[filters,setFilters]=useState({dojo:"",team:"",product:"",competitor:""});const[supabaseStatus,setSupabaseStatus]=useState("");useEffect(()=>{try{const saved=localStorage.getItem(STORAGE_FITOFAN);const savedFile=localStorage.getItem(STORAGE_FITOFAN_FILE);if(saved)setFitofanRaw(JSON.parse(saved));if(savedFile)setFiles(p=>({...p,fitofan:savedFile}));}catch(e){console.warn(e)}},[]);useEffect(()=>{refreshManualLinks();refreshEngagementCorrections();refreshManualCompetitors();refreshManualSizes();refreshCommentCorrections();},[]);
async function upload(type,file){if(!file)return;const rows=await readFileRows(file);setFiles(p=>({...p,[type]:`${file.name} (${rows.length} lignes)`}));if(type==="fitofan"){setFitofanRaw(rows);localStorage.setItem(STORAGE_FITOFAN,JSON.stringify(rows));localStorage.setItem(STORAGE_FITOFAN_FILE,`${file.name} (${rows.length} lignes)`);}if(type==="shopify")setSavedSizeKeys(new Set());setShopifyCloudRows([]);setShopifyRaw(rows);if(type==="supabase")setSupabaseRaw(rows);}function resetFitofan(){localStorage.removeItem(STORAGE_FITOFAN);localStorage.removeItem(STORAGE_FITOFAN_FILE);setFitofanRaw([]);setFiles(p=>({...p,fitofan:""}));}
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
async function createManualLink(order,shopifyName,fitofanName){setManualLinkStatus("Sauvegarde du lien manuel...");try{const f=fitofan.competitors.find(x=>x.competitor===fitofanName);if(!order)throw new Error("Aucune commande d’engagement sélectionnée.");if(!f)throw new Error("Compétiteur Fitofan introuvable.");await saveManualLinkToSupabase({order_number:order,shopify_competitor:shopifyName||"",fitofan_competitor:f.competitor,fitofan_email:f.email||"",fitofan_dojo:f.dojo||"",fitofan_team:f.team||""});await refreshManualLinks();setManualSelections(p=>({...p,[order]:""}));setManualLinkStatus(`Commande ${order} liée à ${f.competitor}`);}catch(e){setManualLinkStatus(e.message||"Erreur sauvegarde lien manuel");}}
async function removeManualLink(id){setManualLinkStatus("Suppression du lien manuel...");try{await deleteManualLinkFromSupabase(id);await refreshManualLinks();setManualLinkStatus("Lien manuel supprimé");}catch(e){setManualLinkStatus(e.message||"Erreur suppression lien manuel");}}

async function refreshManualCompetitors(){
  setManualParticipantStatus("Chargement des participants manuels...");
  try{
    const rows=await loadManualCompetitorsFromSupabase();
    setManualCompetitors(rows||[]);
    setManualParticipantStatus(`Participants manuels chargés : ${(rows||[]).length}`);
  }catch(e){setManualParticipantStatus(e.message||"Erreur chargement participants manuels");}
}
async function refreshManualSizes(){
  try{const rows=await loadManualSizesFromSupabase();setManualSizes(rows||[]);}catch(e){console.warn(e);}
}
async function refreshCommentCorrections(){
  try{
    const rows=await loadCommentCorrectionsFromSupabase();
    setCommentCorrections(rows||[]);
    setCommentCorrectionStatus(`Commentaires validés chargés : ${(rows||[]).length}`);
  }catch(e){
    setCommentCorrectionStatus(e.message||"Table comment_corrections absente");
    console.warn("Chargement validations commentaires",e);
  }
}
async function validateCommentRow(item){
  const key=commentCorrectionKey(item);
  const input=commentCorrectionInputs[key]||{};
  const sizeRaw=clean(input.size ?? item.finalSize ?? "");
  const normalizedSize=sizeRaw?normalizeManualSizeForSave(sizeRaw):"";
  const dojo=clean(input.dojo ?? item.dojo ?? "");
  const team=clean(input.team ?? item.team ?? "");
  const countWarning=multiSizeCountWarning(sizeRaw,item);setCommentCorrectionStatus(countWarning||"Sauvegarde validation commentaire...");
  try{
    const payload={
      order_number:normalizeOrder(item.order),
      product_key:item.productKey||"",
      raw_product:item.rawProduct||"",
      competitor:item.competitor||"",
      email:item.email||"",
      dojo,
      team,
      size_raw:sizeRaw,
      size_normalized:normalizedSize,
      comment_text:item.comments||"",
      quantity:quantityDisplay(item),
      correction_key:key,
      validated:true
    };
    const saved=await saveCommentCorrectionToSupabase(payload);
    const savedRow=Array.isArray(saved)&&saved[0]?saved[0]:payload;
    setCommentCorrections(prev=>[savedRow,...(prev||[]).filter(r=>(r.correction_key||"")!==key)]);
    setCommentCorrectionInputs(p=>{const n={...p};delete n[key];return n;});
    setCommentCorrectionStatus(countWarning?`Commentaire traité / validé. ${countWarning}`:"Commentaire traité / validé");
  }catch(e){
    setCommentCorrectionStatus(e.message||"Erreur validation commentaire");
    console.error("Erreur validation commentaire",e);
  }
}

async function addManualParticipant(){
  setManualParticipantStatus("Ajout du participant...");
  try{
    if(!clean(newParticipant.competitor))throw new Error("Nom du compétiteur requis.");
    await saveManualCompetitorToSupabase({competitor:newParticipant.competitor,email:newParticipant.email||"",dojo:newParticipant.dojo||"",team:newParticipant.team||"",source:"manuel"});
    setNewParticipant({competitor:"",email:"",dojo:"",team:""});
    await refreshManualCompetitors();
    setManualParticipantStatus("Participant ajouté.");
  }catch(e){setManualParticipantStatus(e.message||"Erreur ajout participant");}
}
async function saveManualSize(order,productKey,competitor,rawProduct){
  const normalizedOrder=normalizeOrder(order);
  const k=`${normalizedOrder}|${productKey}`;
  const value=manualSizeInputs[k]||"";
  const countWarning=multiSizeCountWarning(value,itemOverride||{});setManualParticipantStatus(countWarning||"Sauvegarde taille manuelle...");
  try{
    if(!clean(value))throw new Error("Taille requise.");
    const normalized=normalizeManualSizeForSave(value);
    const payload={
      order_number:normalizedOrder,
      product_key:productKey,
      competitor:competitor||"",
      raw_product:rawProduct||"",
      size_raw:value,
      size_normalized:normalized,
      source:"manuel"
    };
    const saved=await saveManualSizeToSupabase(payload);
    const savedRow=Array.isArray(saved)&&saved[0]?saved[0]:payload;
    setManualSizes(prev=>{
      const without=(prev||[]).filter(s=>`${normalizeOrder(s.order_number)}|${s.product_key}`!==k);
      return [savedRow,...without];
    });
    setSavedSizeKeys(prev=>{const next=new Set(prev);next.add(k);return next;});
    setManualSizeInputs(p=>({...p,[k]:""}));
    await refreshManualSizes();
    setManualParticipantStatus(countWarning?`Taille sauvegardée : ${normalized}. ${countWarning}`:`Taille sauvegardée : ${normalized}`);
  }catch(e){
    setManualParticipantStatus(e.message||"Erreur sauvegarde taille");
    console.error("Erreur sauvegarde taille",e);
  }
}

async function refreshEngagementCorrections(){
 setEngagementCorrectionStatus("Chargement des corrections engagement...");
 try{const rows=await loadEngagementLinkCorrectionsFromSupabase();setEngagementCorrections(rows||[]);setEngagementCorrectionStatus(`Corrections engagement chargées : ${(rows||[]).length}`);}
 catch(e){setEngagementCorrectionStatus(e.message||"Erreur chargement corrections engagement");}
}
async function unlinkEngagement(order,competitor,team){
 setEngagementCorrectionStatus("Déliage engagement...");
 try{await saveEngagementLinkCorrectionToSupabase({order_number:order,competitor,team,action:"unlink"});await refreshEngagementCorrections();setEngagementCorrectionStatus(`Engagement ${order} délié de ${competitor}`);}
 catch(e){setEngagementCorrectionStatus(e.message||"Erreur déliaison engagement");}
}
async function linkEngagementToCompetitor(order,competitor,team){
 setEngagementCorrectionStatus("Liaison engagement...");
 try{if(!order)throw new Error("Commande engagement requise.");await saveEngagementLinkCorrectionToSupabase({order_number:order,competitor,team,action:"link"});await refreshEngagementCorrections();setEngagementLinkInputs(p=>({...p,[`${competitor}|${team}`]:""}));setEngagementCorrectionStatus(`Engagement ${order} lié à ${competitor}`);}
 catch(e){setEngagementCorrectionStatus(e.message||"Erreur liaison engagement");}
}
async function deleteEngagementCorrection(id){
 setEngagementCorrectionStatus("Suppression correction...");
 try{await deleteEngagementLinkCorrectionFromSupabase(id);await refreshEngagementCorrections();setEngagementCorrectionStatus("Correction supprimée.");}
 catch(e){setEngagementCorrectionStatus(e.message||"Erreur suppression correction");}
}

async function saveShopifyCloud(){
  setShopifyCloudStatus("Sauvegarde Shopify dans Supabase...");
  try{
    const count=await saveShopifyRowsToSupabase(shopify,files.shopify||"Import Shopify");
    setShopifyCloudStatus(`Shopify sauvegardé dans Supabase : ancien import remplacé par ${count} lignes`);
  }catch(e){
    setShopifyCloudStatus(e.message||"Erreur sauvegarde Shopify");
  }
}
async function loadShopifyCloud(){
  setShopifyCloudStatus("Chargement Shopify depuis Supabase...");
  try{
    const rows=await loadShopifyRowsFromSupabase();
    setSavedSizeKeys(new Set());setShopifyCloudRows(rows);
    setFiles(p=>({...p,shopify:`Supabase shopify_order_items (${rows.length} lignes)`}));
    setShopifyCloudStatus(`Shopify chargé depuis Supabase : ${rows.length} lignes complètes`);
  }catch(e){
    setShopifyCloudStatus(e.message||"Erreur chargement Shopify");
  }
}
async function refreshSupabase(){setSupabaseStatus("Chargement Supabase...");try{const rows=await loadSupabaseResponses();setSupabaseRaw(rows);setFiles(p=>({...p,supabase:`Supabase responses (${rows.length} lignes)`}));setSupabaseStatus(`Supabase chargé : ${rows.length} lignes`);}catch(e){setSupabaseStatus(e.message||"Erreur Supabase");}}
const fitofanBase=useMemo(()=>parseFitofan(fitofanRaw),[fitofanRaw]);const fitofan=useMemo(()=>{const base=fitofanBase.competitors||[];const keys=new Set(base.map(f=>f.key));const manual=(manualCompetitors||[]).map(r=>({competitor:r.competitor,key:competitorKey(r.competitor),email:r.email||"",dojo:r.dojo||"",team:r.team||"",manual:true,fitofanStatus:r.fitofan_status||r.fitofanStatus||"manuel"})).filter(f=>!keys.has(f.key));return{competitors:[...base,...manual],comments:fitofanBase.comments||[]};},[fitofanBase,manualCompetitors]);const shopifyParsed=useMemo(()=>parseShopify(shopifyRaw),[shopifyRaw]);const shopify=shopifyCloudRows.length?shopifyCloudRows:shopifyParsed;const supabase=useMemo(()=>parseSupabase(supabaseRaw),[supabaseRaw]);
const commentCorrectionMap=useMemo(()=>{
  const m=new Map();
  (commentCorrections||[]).forEach(r=>{
    const k=r.correction_key||`${normalizeOrder(r.order_number)}|${r.product_key||""}|${r.raw_product||""}`;
    if(k&&!m.has(k))m.set(k,r);
  });
  return m;
},[commentCorrections]);
const manualLinkByOrder=useMemo(()=>{const m=new Map();manualLinks.forEach(l=>{if(l.order_number)m.set(l.order_number,l);});return m;},[manualLinks]);const manualSizeMap=useMemo(()=>{const m=new Map();manualSizes.forEach(s=>{if(s.order_number&&s.product_key){const k=`${normalizeOrder(s.order_number)}|${s.product_key}`;if(!m.has(k))m.set(k,s);}});return m;},[manualSizes]);const reconciled=useMemo(()=>{const corrections=new Map(),commentsByOrder=new Map(),commentsByCompetitor=new Map();supabase.forEach(s=>{const comment=s.comments||(s.kind!=="product"?s.rawProduct:"");if(comment){if(s.order){if(!commentsByOrder.has(s.order))commentsByOrder.set(s.order,[]);commentsByOrder.get(s.order).push(comment);}if(s.ckey){if(!commentsByCompetitor.has(s.ckey))commentsByCompetitor.set(s.ckey,[]);commentsByCompetitor.get(s.ckey).push(comment);}}if(s.kind==="product"&&s.size&&s.order)corrections.set(productKey(s.order,s.productKey),s);});return shopify.map(item=>{const manualLink=manualLinkByOrder.get(item.order);if(manualLink){item={...item,competitor:manualLink.fitofan_competitor||item.competitor,ckey:competitorKey(manualLink.fitofan_competitor||item.competitor),dojo:manualLink.fitofan_dojo||item.dojo,team:manualLink.fitofan_team||item.team,productTeam:item.productTeam||productTeamFromTitle(item.rawProduct),manualLinked:true,manualOriginalCompetitor:item.competitor};}if(item.kind!=="product")return item;const corr=corrections.get(productKey(item.order,item.productKey));const commentCorrection=commentCorrectionMap.get(commentCorrectionKey(item));if(commentCorrection){item={...item,dojo:commentCorrection.dojo||item.dojo,team:commentCorrection.team||item.team,commentValidated:true,commentValidationText:commentCorrection.comment_text||""};}const manualSize=manualSizeMap.get(productKey(item.order,item.productKey));let finalSize=manualSize?.size_normalized||commentCorrection?.size_normalized||corr?.size||item.shopifySize||"";let sourceSize=manualSize?.size_normalized?"Manuel":(commentCorrection?.size_normalized?"Commentaire validé":(corr?.size?"Supabase":(item.shopifySize?"Shopify":"")));if(!hasSize(item.productKey)){finalSize="Taille unique";sourceSize="Taille unique";}const missing=!finalSize&&hasSize(item.productKey)&&!item.excluded;const comments=[...new Set([...(commentsByOrder.get(item.order)||[]),...(commentsByCompetitor.get(item.ckey)||[])])].join(" | ");let status=item.excluded?item.status:"OK SHOPIFY";if(corr&&!item.shopifySize)status="TAILLE AJOUTÉE PAR SUPABASE";if(corr&&item.shopifySize&&item.shopifySize!==corr.size)status="TAILLE MODIFIÉE PAR SUPABASE";if(corr&&item.shopifySize&&item.shopifySize===corr.size)status="TAILLE CONFIRMÉE PAR SUPABASE";if(missing)status="TAILLE MANQUANTE";if(comments)status+=" + COMMENTAIRE";if(commentCorrection)status+=" + COMMENTAIRE VALIDÉ";const finalSizeAnalysis=analyzeSize(finalSize,sourceSize);return{...item,supabaseSizeRaw:corr?.sizeRaw||"",supabaseSize:corr?.size||"",manualSizeRaw:manualSize?.size_raw||"",finalSize,finalSizeStatus:sourceSize==="Manuel"?"STANDARD":(missing?"MANQUANTE":finalSizeAnalysis.status),sourceSize,missing,initialShopifySizeMissing:!item.shopifySize&&hasSize(item.productKey)&&!item.excluded,comments,commentValidated:!!commentCorrection,commentCorrectionDojo:commentCorrection?.dojo||"",commentCorrectionTeam:commentCorrection?.team||"",commentCorrectionSize:commentCorrection?.size_normalized||"",status};});},[shopify,supabase,manualLinkByOrder,commentCorrectionMap,manualSizeMap]);
const detailRows=reconciled.map(r=>({Commande:r.order,Compétiteur:r.competitor,"Nom Shopify original":r.manualOriginalCompetitor||"","Lien manuel":r.manualLinked?"Oui":"Non",Client:r.client,Email:r.email,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),Type:r.kind,"Produit Shopify brut":r.rawProduct,"Produit normalisé":r.productKey,Produit:r.product,"Taille Shopify brute":r.shopifySizeRaw||"","Taille Shopify":r.shopifySize||"","Statut taille Shopify":r.shopifySizeStatus||"","Taille Shopify initialement manquante":r.initialShopifySizeMissing?"Oui":"Non","Taille Supabase brute":r.supabaseSizeRaw||"","Taille Supabase":r.supabaseSize||"","Taille manuelle":r.manualSizeRaw||"","Taille finale":r.finalSize||"","Statut taille finale":r.finalSizeStatus||"","Source taille":r.sourceSize||"",Quantité:r.quantity,"Qté remboursée":r.refunded,"Qté effective":r.effectiveQty,"Exclu fournisseur":r.excluded?"Oui":"Non",Statut:r.status,"Commentaire traité":r.commentValidated?"Oui":"Non",Commentaires:r.comments||""}));
const unusualSizeRows=reconciled.filter(r=>r.kind==="product"&&!r.excluded&&r.finalSizeStatus==="INHABITUELLE").map(r=>({Compétiteur:r.competitor,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),Commande:r.order,Produit:r.product,"Produit Shopify brut":r.rawProduct,"Taille Shopify brute":r.shopifySizeRaw||"","Taille Supabase brute":r.supabaseSizeRaw||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"",Statut:"TAILLE INHABITUELLE À VÉRIFIER",Commentaires:r.comments||""}));const shopifyInitialMissingRows=reconciled.filter(r=>r.kind==="product"&&r.initialShopifySizeMissing).map(r=>({Compétiteur:r.competitor,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),Commande:r.order,Produit:r.product,"Produit Shopify brut":r.rawProduct,"Taille Shopify initiale":r.shopifySize||"MANQUANTE","Taille Supabase":r.supabaseSize||"Aucune","Taille finale":r.finalSize||"MANQUANTE","Statut régularisation":r.supabaseSize&&r.finalSize?"RÉGULARISÉ PAR SUPABASE":"TOUJOURS MANQUANT",Commentaires:r.comments||""}));const supplierRows=useMemo(()=>{const map=new Map();expandRowsByFinalSizes(reconciled).forEach(r=>{if(r.kind!=="product"||r.excluded||!r.finalSize)return;const key=`${r.productKey}|${r.team}|${r.dojo}|${r.finalSize}`;const prev=map.get(key)||{Produit:r.product,"Produit normalisé":r.productKey,Équipe:r.team||"Non précisée",Taille:r.finalSize,Quantité:0};prev.Quantité+=r.effectiveQty;map.set(key,prev);});return Array.from(map.values()).sort((a,b)=>`${a.Dojo}${a.Équipe}${a.Produit}${a.Taille}`.localeCompare(`${b.Dojo}${b.Équipe}${b.Produit}${b.Taille}`));},[reconciled]);
const engagementOrders=useMemo(()=>reconciled.filter(r=>isEngagementOrderRow(r)).map(r=>({Commande:r.order,"Nom Shopify":r.manualOriginalCompetitor||r.competitor||"",Compétiteur:r.competitor,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),"Produit brut":r.rawProduct||""})),[reconciled]);
const resolvedEngagementState=useMemo(()=>buildResolvedEngagementState({fitofan,reconciled,engagementCorrections}),[fitofan,reconciled,engagementCorrections]);
const engagementManagementRows=resolvedEngagementState.managementRows;
const correctedMissingEngagementRows=resolvedEngagementState.missing;
const engagementManagementFilteredRows=useMemo(()=>{const q=norm(engagementSearch);return engagementManagementRows.filter(r=>{const txt=norm(`${r.Compétiteur} ${r.Dojo} ${r.Équipe} ${r.Commande} ${r["Nom Shopify"]} ${r.Statut}`);const okSearch=!q||txt.includes(q);const okStatus=engagementStatusFilter==="all"||(engagementStatusFilter==="linked"&&r.Statut==="LIÉ")||(engagementStatusFilter==="missing"&&r.Statut!=="LIÉ");return okSearch&&okStatus;});},[engagementManagementRows,engagementSearch,engagementStatusFilter]);
const fitofanReport=useMemo(()=>fitofan.competitors.map(f=>{const rows=reconciled.filter(r=>r.ckey===f.key);const products=rows.filter(r=>r.kind==="product"&&!r.excluded);const engagementKey=`${f.key}|${canonicalTeamKey(f.team)}`;const engagements=resolvedEngagementState.engagementByCompetitorTeam.get(engagementKey)||[];const orders=[...new Set([...rows.map(r=>r.order).filter(Boolean),...engagements.map(e=>e.order).filter(Boolean)])];return{Compétiteur:f.competitor,Email:f.email,Dojo:f.dojo,Équipe:displayTeamCategory(f.team),FITOFAN:f.fitofanStatus||"","Engagement signé":engagements.length?"Oui":"Non","Commandes liées":orders.join(" | "),"A acheté équipement":products.length?"Oui":"Non","Produits commandés":products.map(r=>`${r.product} (${r.finalSize||"taille manquante"})`).join(" | "),"Tailles Shopify":products.map(r=>`${r.order} - ${r.product}: ${r.shopifySize||"vide"}`).join(" | "),"Tailles Supabase":products.map(r=>`${r.order} - ${r.product}: ${r.supabaseSize||"aucune"}`).join(" | "),"Tailles finales":products.map(r=>`${r.order} - ${r.product}: ${r.finalSize||"taille manquante"} (${r.sourceSize||"aucune source"})`).join(" | "),"Tailles manquantes":products.filter(r=>r.missing).length,Commentaires:[...new Set(rows.map(r=>r.comments).filter(Boolean))].join(" | ")};}),[fitofan,reconciled,resolvedEngagementState]);
const supabaseLinkedKeys=new Set(reconciled.filter(r=>r.supabaseSize).map(r=>productKey(r.order,r.productKey)));const supabaseAudit=supabase.map(s=>({Commande:s.order,Compétiteur:s.competitor,Produit:s.rawProduct,"Produit normalisé":s.productKey,Taille:s.size,Commentaire:s.comments,Type:s.kind,Statut:s.kind==="product"&&s.size?(supabaseLinkedKeys.has(productKey(s.order,s.productKey))?"Reliée":"NON RELIÉE"):"Commentaire / engagement conservé"}));
const auditRows=[{Indicateur:"Lignes Fitofan compétiteurs",Valeur:fitofan.competitors.length},{Indicateur:"Colonne Inscrit Fitofan supprimée",Valeur:"Oui - utiliser FITOFAN"},{Indicateur:"Fitofan source",Valeur:files.fitofan||"Non chargé"},{Indicateur:"Commentaires Fitofan exclus",Valeur:fitofan.comments.length},{Indicateur:"Lignes Shopify produits/engagements",Valeur:shopify.length},{Indicateur:"Articles actifs analysés",Valeur:reconciled.filter(r=>r.kind==="product"&&!r.excluded).reduce((s,r)=>s+r.effectiveQty,0)},{Indicateur:"Articles exclus annulés/remboursés",Valeur:reconciled.filter(r=>r.kind==="product"&&r.excluded).length},{Indicateur:"Tailles Supabase appliquées",Valeur:reconciled.filter(r=>r.supabaseSize).length},{Indicateur:"Tailles manquantes",Valeur:reconciled.filter(r=>r.missing).length},{Indicateur:"Réponses Supabase total",Valeur:supabase.length},{Indicateur:"Réponses Supabase produits non reliées",Valeur:supabaseAudit.filter(r=>r.Statut==="NON RELIÉE").length},{Indicateur:"Total produits fournisseur",Valeur:supplierRows.reduce((s,r)=>s+r.Quantité,0)}];
const dashboardRows=detailRows.filter(r=>r.Type==="product"&&r["Exclu fournisseur"]!=="Oui");const filterOptions={dojo:uniqueSorted([...dashboardRows.map(r=>r.Dojo),...fitofanReport.map(r=>r.Dojo)]),team:uniqueSorted([...dashboardRows.map(r=>r.Équipe),...fitofanReport.map(r=>r.Équipe)]),product:uniqueSorted(dashboardRows.map(r=>r.Produit)),competitor:uniqueSorted([...dashboardRows.map(r=>r.Compétiteur),...fitofanReport.map(r=>r.Compétiteur)])};function applyFilters(rows){return rows.filter(row=>{const text=norm(Object.values(row).join(" "));if(search&&!text.includes(norm(search)))return false;if(filters.dojo&&norm(row.Dojo)!==norm(filters.dojo))return false;if(filters.team&&norm(row.Équipe)!==norm(filters.team))return false;if(filters.product&&!norm(Object.values(row).join(" ")).includes(norm(filters.product)))return false;if(filters.competitor&&norm(row.Compétiteur)!==norm(filters.competitor))return false;return true;});}
const dashboardFiltered=applyFilters(dashboardRows);const dashboardStats={competitors:uniqueSorted(dashboardFiltered.map(r=>r.Compétiteur)).length,quantity:dashboardFiltered.reduce((s,r)=>s+(Number(r["Qté effective"])||0),0),missing:dashboardFiltered.filter(r=>r.Statut.includes("TAILLE MANQUANTE")).length,supabase:dashboardFiltered.filter(r=>r["Source taille"]==="Supabase").length,initialMissing:shopifyInitialMissingRows.length,regularized:shopifyInitialMissingRows.filter(r=>r["Statut régularisation"]==="RÉGULARISÉ PAR SUPABASE").length,unusual:unusualSizeRows.length};const competitorsMissingSizes=useMemo(()=>{const map=new Map();dashboardRows.filter(r=>String(r.Statut||"").includes("TAILLE MANQUANTE")).forEach(r=>{const name=r.Compétiteur||"Compétiteur non précisé";const current=map.get(name)||{Compétiteur:name,Dojo:r.Dojo||"",Équipe:r.Équipe||"",Commandes:new Set(),Produits:[]};if(r.Commande)current.Commandes.add(r.Commande);current.Produits.push(`${r.Produit||""} (${r.Commande||""})`);map.set(name,current);});return Array.from(map.values()).map(r=>({Compétiteur:r.Compétiteur,Dojo:r.Dojo,Équipe:r.Équipe,Commandes:Array.from(r.Commandes).join(" | "),"Produits sans taille":r.Produits.join(" | ")}));},[dashboardRows]);const competitorsMissingEngagement=useMemo(()=>fitofanReport.filter(r=>r["Inscrit Fitofan"]==="Oui"&&r["Engagement signé"]!=="Oui").map(r=>({Compétiteur:r.Compétiteur,Dojo:r.Dojo,Équipe:r.Équipe,Email:r.Email,"Commandes liées":r["Commandes liées"],"A acheté équipement":r["A acheté équipement"]})),[fitofanReport]);const linkedOrdersSet=useMemo(()=>new Set(manualLinks.map(l=>l.order_number)),[manualLinks]);
const manualReconciliationRows=useMemo(()=>{
  const engagementOrdersList=(engagementOrders||[]).filter(o=>o.Commande);
  return correctedMissingEngagementRows.map(f=>{
    const suggestions=engagementOrdersList
      .map(o=>({
        ...o,
        Score:simpleNameScore(f.Compétiteur,o["Nom Shopify"]||o.Compétiteur||"")
      }))
      .sort((a,b)=>b.Score-a.Score)
      .slice(0,5);

    return {
      Compétiteur:f.Compétiteur,
      Dojo:f.Dojo,
      Équipe:f.Équipe,
      Email:f.Email||"",
      Engagement:"MANQUANT",
      "Meilleure suggestion":suggestions[0]?.Commande||"",
      "Nom Shopify suggéré":suggestions[0]?.["Nom Shopify"]||"",
      "Score suggestion":suggestions[0]?.Score||0,
      "Commandes engagement possibles":suggestions.map(s=>`${s.Commande} - ${s["Nom Shopify"]||s.Compétiteur||""} (${s.Score}%)`).join(" | ")
    };
  });
},[correctedMissingEngagementRows,engagementOrders]);const manualLinksRows=manualLinks.map(l=>({Commande:l.order_number,"Nom Shopify":l.shopify_competitor,"Compétiteur Fitofan":l.fitofan_competitor,Dojo:l.fitofan_dojo,Équipe:l.fitofan_team,Créé:l.created_at||"",Action:"Supprimer"}));const finalParticipantsRows=useMemo(()=>fitofan.competitors.map(f=>{const rows=reconciled.filter(r=>r.ckey===f.key);const products=rows.filter(r=>r.kind==="product"&&!r.excluded);const engagements=rows.filter(r=>isEngagementOrderRow(r));return{Compétiteur:f.competitor,Email:f.email,Dojo:f.dojo,Équipe:displayTeamCategory(f.team),Source:f.manual?"Ajout manuel":"Fitofan","Engagement signé":engagements.length?"Oui":"Non","Commandes liées":[...new Set(rows.map(r=>r.order).filter(Boolean))].join(" | "),"Produits commandés":products.map(r=>`${r.product} (${r.finalSize||"taille manquante"})`).join(" | "),"Tailles Shopify":products.map(r=>`${r.order} - ${r.product}: ${r.shopifySize||"vide"}`).join(" | "),"Tailles Supabase":products.map(r=>`${r.order} - ${r.product}: ${r.supabaseSize||"aucune"}`).join(" | "),"Tailles manuelles":products.map(r=>`${r.order} - ${r.product}: ${r.manualSizeRaw||"aucune"}`).join(" | "),"Tailles finales":products.map(r=>`${r.order} - ${r.product}: ${r.finalSize||"taille manquante"} (${r.sourceSize||"aucune source"})`).join(" | "),"Tailles manquantes":products.filter(r=>r.missing).length,Commentaires:[...new Set(rows.map(r=>r.comments).filter(Boolean))].join(" | ")};}),[fitofan,reconciled]);const sizeValidationRows=useMemo(()=>reconciled.filter(r=>sizeNeedsValidation(r)).map(r=>({Compétiteur:r.competitor,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),Commande:r.order,Produit:r.product,"Produit normalisé":r.productKey,"Produit Shopify brut":r.rawProduct,"Taille Shopify":r.shopifySize||"","Taille Supabase":r.supabaseSize||"","Taille manuelle":r.manualSizeRaw||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"","Statut validation":sizeValidationStatus(r),Commentaires:r.comments||""})),[reconciled]);
const supplierFinalRows=useMemo(()=>{const map=new Map();expandRowsByFinalSizes(reconciled).filter(r=>isSupplierReady(r)).forEach(r=>{const key=`${r.product}|${r.team||"Non précisée"}|${r.finalSize}`;const prev=map.get(key)||{Produit:r.product,Équipe:r.team||"Non précisée",Taille:r.finalSize,Quantité:0};prev.Quantité+=Number(r.effectiveQty)||0;map.set(key,prev);});return Array.from(map.values()).sort((a,b)=>`${a.Produit}${a.Équipe}${a.Taille}`.localeCompare(`${b.Produit}${b.Équipe}${b.Taille}`));},[reconciled]);
const supplierInternalRows=useMemo(()=>reconciled.filter(r=>r.kind==="product"&&!r.excluded).map(r=>({Commande:r.order,Compétiteur:r.competitor,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),Produit:r.product,"Produit brut":r.rawProduct,Quantité:r.effectiveQty||r.quantity||1,"Taille Shopify":r.shopifySize||"","Taille Supabase":r.supabaseSize||"","Taille manuelle":r.manualSizeRaw||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"","Statut validation":sizeValidationStatus(r),Quantité:r.effectiveQty,Commentaires:r.comments||""})),[reconciled]);

const engagementCorrectionsMap=useMemo(()=>{const unlink=new Set();const link=new Set();engagementCorrections.forEach(c=>{const k=engagementCorrectionKey(c.order_number,c.competitor,c.team);if(c.action==="unlink")unlink.add(k);if(c.action==="link")link.add(k);});return{unlink,link};},[engagementCorrections]);
const engagementCorrectionsRows=engagementCorrections.map(c=>({ID:c.id,Commande:c.order_number,Compétiteur:c.competitor,Équipe:c.team,Action:c.action,Créé:c.created_at||""}));
const fitofanKnownKeys=useMemo(()=>new Set((fitofan?.competitors||[]).map(f=>f.key)),[fitofan]);
const shopifyUnknownCompetitorRows=useMemo(()=>{
  const map=new Map();
  (reconciled||[]).forEach(r=>{
    if(!r.competitor)return;
    const key=competitorKey(r.competitor);
    if(fitofanKnownKeys.has(key))return;
    const cur=map.get(key)||{
      Compétiteur:r.competitor,
      Email:r.email||"",
      Dojo:r.dojo||"",
      Équipe:r.team||"",
      Commandes:new Set(),
      Produits:[],
      Engagement:false,
      Équipement:false,
      Statut:"ABSENT DU FICHIER COMPÉTITEURS"
    };
    if(r.order)cur.Commandes.add(r.order);
    if(isEngagementOrderRow(r))cur.Engagement=true;
    if(r.kind==="product"&&!r.excluded){
      cur.Équipement=true;
      cur.Produits.push(`${r.product} (${r.finalSize||"taille manquante"})`);
    }
    map.set(key,cur);
  });
  return Array.from(map.values()).map(r=>({
    Compétiteur:r.Compétiteur,
    Email:r.Email,
    Dojo:r.Dojo,
    Équipe:r.Équipe,
    Commandes:Array.from(r.Commandes).join(" | "),
    "A signé engagement":r.Engagement?"Oui":"Non",
    "A acheté équipement":r.Équipement?"Oui":"Non",
    Produits:r.Produits.join(" | "),
    Statut:r.Statut
  }));
},[reconciled,fitofanKnownKeys]);

const notFitofanWithEngagementRows=useMemo(()=>{
  return (fitofan?.competitors||[])
    .filter(f=>isFitofanNo(f.fitofanStatus))
    .map(f=>{
      const rows=(reconciled||[]).filter(r=>r.ckey===f.key);
      const engagements=rows.filter(r=>isEngagementOrderRow(r));
      const products=rows.filter(r=>r.kind==="product"&&!r.excluded);
      return {
        Compétiteur:f.competitor,
        Email:f.email,
        Dojo:f.dojo,
        Équipe:displayTeamCategory(f.team),
        FITOFAN:f.fitofanStatus||"non",
        "Engagement signé":engagements.length?"Oui":"Non",
        "A acheté équipement":products.length?"Oui":"Non",
        "Commandes liées":[...new Set(rows.map(r=>r.order).filter(Boolean))].join(" | "),
        Produits:products.map(r=>`${r.product} (${r.finalSize||"taille manquante"})`).join(" | ")
      };
    })
    .filter(r=>r["Engagement signé"]==="Oui"||r["A acheté équipement"]==="Oui");
},[fitofan,reconciled]);

const fitofanStatusAuditRows=useMemo(()=>{
  return (fitofan?.competitors||[]).map(f=>({
    Compétiteur:f.competitor,
    Email:f.email,
    Dojo:f.dojo,
    Équipe:displayTeamCategory(f.team),
    FITOFAN:f.fitofanStatus||"",
    Source:f.manual?"Ajout manuel":"Fichier / Supabase"
  }));
},[fitofan]);

const visibleSizeValidationRows=useMemo(()=>sizeValidationRows.filter(r=>!savedSizeKeys.has(`${normalizeOrder(r.order)}|${r.productKey}`)),[sizeValidationRows,savedSizeKeys]);
const orderSizeEditRows=useMemo(()=>{
  return (reconciled||[])
    .filter(r=>r.kind==="product"&&!r.excluded&&hasSize(r.productKey))
    .filter(r=>orderSearchMatch(r, orderSizeEditSearch))
    .map(r=>({
      Commande:r.order,
      Compétiteur:r.competitor,
      Email:r.email||"",
      Dojo:r.dojo||"",
      Équipe:displayTeamCategory(r.team),
      Produit:r.product,
      "Produit brut":r.rawProduct||"",
      "Taille Shopify":r.shopifySize||"",
      "Taille Supabase":r.supabaseSize||"",
      "Taille finale":r.finalSize||"",
      "Source taille":r.sourceSize||"",
      "Statut taille":sizeValidationStatus(r),
      productKey:r.productKey,
      raw:r
    }));
},[reconciled,orderSizeEditSearch]);
const productsNoProductTeamRows=useMemo(()=>reconciled.filter(r=>r.kind==="product"&&!r.excluded&&!r.productTeam).map(r=>({Commande:r.order,Compétiteur:r.competitor,Email:r.email||"",Dojo:r.dojo||"","Équipe compétiteur":displayTeamCategory(r.team)||"","Équipe produit":supplierTeamForRow(r),Produit:r.product,"Produit Shopify brut":r.rawProduct||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"",Quantité:r.effectiveQty||r.quantity||1})),[reconciled]);
const productTeamAuditRows=useMemo(()=>reconciled.filter(r=>r.kind==="product"&&!r.excluded).map(r=>({Commande:r.order,Compétiteur:r.competitor,Email:r.email||"",Dojo:r.dojo||"","Équipe compétiteur":displayTeamCategory(r.team)||"","Équipe produit":supplierTeamForRow(r),Produit:r.product,"Produit Shopify brut":r.rawProduct||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"",Quantité:r.effectiveQty||r.quantity||1})),[reconciled]);

const supplierAuditRows=useMemo(()=>{
  return expandRowsByFinalSizes(reconciled||[])
    .filter(r=>r.kind==="product")
    .map(r=>({
      Commande:r.order,
      Compétiteur:r.competitor,
      Email:r.email||"",
      Dojo:r.dojo||"",
      "Équipe compétiteur":displayTeamCategory(r.team)||"",
      "Équipe produit":supplierTeamForRow(r),
      Produit:r.product,
      "Produit Shopify brut":r.rawProduct||"",
      "Taille Shopify":r.shopifySize||"",
      "Taille finale fournisseur":supplierFinalSizeLabel(r),
      "Source taille":r.sourceSize||"",
      "Statut fournisseur":supplierLineStatus(r),
      "Inclus fournisseur":supplierIncludeRow(r)?"Oui":"Non",
      Quantité:Number(r.effectiveQty||r.quantity||0)||0
    }));
},[reconciled]);

const supplierAuditSummaryRows=useMemo(()=>{
  const rows=supplierAuditRows;
  const sum=list=>list.reduce((s,r)=>s+(Number(r.Quantité)||0),0);
  const total=sum(rows);
  const included=sum(rows.filter(r=>r["Inclus fournisseur"]==="Oui"));
  const ok=sum(rows.filter(r=>r["Statut fournisseur"]==="OK"));
  const missing=sum(rows.filter(r=>r["Statut fournisseur"]==="TAILLE MANQUANTE"));
  const unusual=sum(rows.filter(r=>r["Statut fournisseur"]==="TAILLE INHABITUELLE"));
  const excluded=sum(rows.filter(r=>r["Inclus fournisseur"]!=="Oui"));
  return [
    {Indicateur:"Articles produits Shopify",Valeur:total},
    {Indicateur:"Articles inclus fournisseur",Valeur:included},
    {Indicateur:"OK",Valeur:ok},
    {Indicateur:"Tailles manquantes incluses",Valeur:missing},
    {Indicateur:"Tailles inhabituelles incluses",Valeur:unusual},
    {Indicateur:"Exclus fournisseur",Valeur:excluded},
    {Indicateur:"Écart inclus - total produits",Valeur:included-total}
  ];
},[supplierAuditRows]);

const supplierFinalWithExceptionsRows=useMemo(()=>{
  const map=new Map();
  expandRowsByFinalSizes(reconciled||[]).filter(supplierIncludeRow).forEach(r=>{
    const product=r.product||"Produit";
    const team=supplierTeamForRow(r);
    const size=supplierFinalSizeLabel(r);
    const status=supplierLineStatus(r);
    const key=`${product}|${team}|${size}|${status}`;
    const cur=map.get(key)||{Produit:product,Équipe:team,Taille:size,"Statut fournisseur":status,Quantité:0};
    cur.Quantité += Number(r.effectiveQty||r.quantity||0)||0;
    map.set(key,cur);
  });
  return Array.from(map.values()).sort((a,b)=>String(a.Produit).localeCompare(String(b.Produit))||String(a.Équipe).localeCompare(String(b.Équipe))||String(a.Taille).localeCompare(String(b.Taille)));
},[reconciled]);

const commentedOrderKeys=useMemo(()=>{
  const keys=new Set();
  (reconciled||[]).forEach(r=>{
    if(hasUsefulComment(r.comments)){
      if(r.order)keys.add(normalizeOrder(r.order));
      if(r.ckey)keys.add(`ckey:${r.ckey}`);
    }
  });
  return keys;
},[reconciled]);

const commentSizeReviewRows=useMemo(()=>{
  return (reconciled||[])
    .filter(r=>r.kind==="product"&&!r.excluded&&hasSize(r.productKey))
    .filter(r=>!r.commentValidated)
    .filter(r=>hasUsefulComment(r.comments)||commentedOrderKeys.has(normalizeOrder(r.order))||commentedOrderKeys.has(`ckey:${r.ckey}`))
    .filter(r=>orderSearchMatch(r,commentSizeReviewSearch))
    .map(r=>({Commande:r.order,Compétiteur:r.competitor,Email:r.email||"",Dojo:r.dojo||"",Équipe:displayTeamCategory(r.team)||r.team||"",Produit:r.product,"Produit Shopify brut":r.rawProduct||"",Quantité:r.effectiveQty||r.quantity||1,"Taille Shopify":r.shopifySize||"","Taille Supabase":r.supabaseSize||"","Taille finale":r.finalSize||"","Source taille":r.sourceSize||"",Commentaire:r.comments||"",Quantité:r.effectiveQty||r.quantity||1,"Commentaire traité":"Non","Statut taille":sizeValidationStatus(r),productKey:r.productKey,raw:r}));
},[reconciled,commentedOrderKeys,commentSizeReviewSearch]);


const teamCategoryAuditRows=useMemo(()=>{
  const all=[
    ...(fitofan?.competitors||[]).map(r=>({Source:"Fitofan",Nom:r.competitor,Équipe:r.team||"",Catégorie:displayTeamCategory(r.team)})),
    ...(reconciled||[]).map(r=>({Source:"Shopify",Nom:r.competitor,Équipe:r.team||"",Catégorie:displayTeamCategory(r.team)}))
  ];
  const map=new Map();
  all.forEach(r=>{
    const key=`${r.Source}|${r.Catégorie||"Vide"}|${r.Équipe||"Vide"}`;
    const cur=map.get(key)||{Source:r.Source,Catégorie:r.Catégorie||"Vide","Équipe originale":r.Équipe||"Vide",Nombre:0};
    cur.Nombre+=1;
    map.set(key,cur);
  });
  return Array.from(map.values()).sort((a,b)=>String(a.Source).localeCompare(String(b.Source))||String(a.Catégorie).localeCompare(String(b.Catégorie)));
},[fitofan,reconciled]);


const supplierMultiSizeWarningRows=useMemo(()=>reconciled.filter(r=>{
  if(r.kind!=="product"||r.excluded)return false;
  const qty=quantityDisplay(r);
  const parts=finalSizeParts(r.finalSize||"");
  return qty>1 && parts.length>1 && parts.length!==qty;
}).map(r=>({Commande:r.order,Compétiteur:r.competitor,Produit:r.product,"Produit brut":r.rawProduct,Quantité:quantityDisplay(r),"Tailles saisies":r.finalSize,"Nb tailles":finalSizeParts(r.finalSize||"").length,Alerte:`${finalSizeParts(r.finalSize||"").length} tailles pour quantité ${quantityDisplay(r)}`})),[reconciled]);
const views={dashboard:{title:"Tableau de bord visuel",rows:dashboardFiltered},teamCategoryAudit:{title:"Audit catégories équipes",rows:teamCategoryAuditRows},supplierFinalWithExceptions:{title:"Export fournisseur avec exceptions",rows:supplierFinalWithExceptionsRows},supplierAuditSummary:{title:"Résumé audit fournisseur",rows:supplierAuditSummaryRows},supplierAudit:{title:"Audit fournisseur",rows:supplierAuditRows},productsNoProductTeam:{title:"Produits Shopify sans équipe",rows:productsNoProductTeamRows},productTeamAudit:{title:"Audit produits / équipes Shopify",rows:productTeamAuditRows},shopifyUnknownCompetitors:{title:"Shopify absent liste compétiteurs",rows:shopifyUnknownCompetitorRows},notFitofanWithOrders:{title:"Commandes avec FITOFAN = non",rows:notFitofanWithEngagementRows},fitofanStatusAudit:{title:"Audit statut FITOFAN",rows:fitofanStatusAuditRows},engagementManagement:{title:"Gestion liens engagements",rows:engagementManagementFilteredRows},engagementCorrections:{title:"Corrections engagements sauvegardées",rows:engagementCorrectionsRows},sizeValidation:{title:"Gestion / validation des tailles",rows:visibleSizeValidationRows},orderSizeEdit:{title:"Modifier tailles commandes",rows:orderSizeEditRows},commentSizeReview:{title:"Commandes avec commentaires",rows:commentSizeReviewRows},commentCorrections:{title:"Commentaires traités / validés",rows:commentCorrections.map(r=>({Commande:r.order_number,Compétiteur:r.competitor,Email:r.email,Dojo:r.dojo,Équipe:displayTeamCategory(r.team)||r.team,Produit:r.product_key,"Produit brut":r.raw_product,Quantité:r.quantity||"", "Taille validée":r.size_normalized||r.size_raw,Commentaire:r.comment_text,"Commentaire traité":r.validated?"Oui":"Non"}))},supplierFinal:{title:"Export fournisseur final",rows:supplierFinalRows},supplierInternal:{title:"Export interne détaillé",rows:supplierInternalRows},participantsFinal:{title:"Export final participants",rows:finalParticipantsRows},manualParticipants:{title:"Ajouter / gérer participants",rows:manualCompetitors.map(r=>({Compétiteur:r.competitor,Email:r.email,Dojo:r.dojo,Équipe:displayTeamCategory(r.team),Source:r.source||"manuel"}))},manualReconciliation:{title:"Engagements manquants à réconcilier",rows:manualReconciliationRows},manualLinks:{title:"Liens manuels sauvegardés",rows:manualLinksRows},unusualSizes:{title:"Tailles inhabituelles à vérifier",rows:unusualSizeRows},shopifyInitialMissing:{title:"Tailles Shopify manquantes initialement",rows:shopifyInitialMissingRows},missingSizes:{title:"Compétiteurs avec tailles manquantes",rows:competitorsMissingSizes},missingEngagements:{title:"Engagements manquants",rows:correctedMissingEngagementRows},audit:{title:"Contrôle qualité",rows:auditRows},supplierMultiSizeWarning:{title:"Alertes quantités / tailles multiples",rows:supplierMultiSizeWarningRows},detail:{title:"Détail global Shopify + Supabase",rows:detailRows},fitofan:{title:"Suivi par compétiteur Fitofan",rows:fitofanReport},supplier:{title:"Commande fournisseur",rows:supplierRows},supabase:{title:"Audit réponses Supabase",rows:supabaseAudit},fitofanComments:{title:"Commentaires Fitofan",rows:fitofan.comments}};const tabGroups=[
{id:"dashboard",label:"Tableau de bord",tabs:["dashboard"]},
{id:"engagements",label:"Engagements",tabs:["engagementManagement","missingEngagements","manualReconciliation","engagementCorrections","manualLinks"]},
{id:"sizes",label:"Tailles",tabs:["sizeValidation","orderSizeEdit","commentSizeReview","commentCorrections","missingSizes"]},
{id:"supplier",label:"Fournisseur",tabs:["supplierFinal","supplierFinalWithExceptions","supplierAuditSummary","supplierAudit","supplier","supplierInternal"]},
{id:"participants",label:"Participants",tabs:["fitofan","participantsFinal","manualParticipants"]},
{id:"audit",label:"Audit / contrôle",tabs:["audit","supplierMultiSizeWarning","teamCategoryAudit","productTeamAudit","productsNoProductTeam","shopifyUnknownCompetitors","notFitofanWithOrders","fitofanStatusAudit","detail","supabase","fitofanComments"]}
];
const currentGroup=tabGroups.find(g=>g.tabs.includes(tab))||tabGroups[0];
const visibleTabKeys=currentGroup.tabs.filter(k=>views[k]);
function selectGroup(g){const first=g.tabs.find(k=>views[k]);if(first)setTab(first);}
const current=views[tab];const filteredRows=tab==="dashboard"?dashboardFiltered:applyFilters(current.rows||[]);
return <div className="app"><header><h1>Réconciliation Sunfuki V34</h1><p>Sources colorées · alertes compétiteurs · engagements manquants · exports filtrés par club.</p></header><section className="sourceStatusGrid">
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
          <input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={e => upload("shopify", e.target.files?.[0])} /><div className="cloudActions"><button onClick={saveShopifyCloud} disabled={!shopify.length}>Sauvegarder Shopify dans Supabase</button><button onClick={loadShopifyCloud}>Charger Shopify depuis Supabase</button></div>{shopifyCloudStatus&&<small>{shopifyCloudStatus}</small>}
        </div>
      </section><section className="actions"><button onClick={resetFitofan}>Réinitialiser Fitofan mémorisé</button></section><section className="filters"><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Recherche libre..."/><Select label="Dojo" value={filters.dojo} options={filterOptions.dojo} onChange={v=>setFilters(p=>({...p,dojo:v}))}/><Select label="Équipe" value={filters.team} options={filterOptions.team} onChange={v=>setFilters(p=>({...p,team:v}))}/><Select label="Produit" value={filters.product} options={filterOptions.product} onChange={v=>setFilters(p=>({...p,product:v}))}/><Select label="Compétiteur" value={filters.competitor} options={filterOptions.competitor} onChange={v=>setFilters(p=>({...p,competitor:v}))}/><button onClick={()=>setFilters({dojo:"",team:"",product:"",competitor:""})}>Effacer filtres</button></section>{tab==="dashboard"&&<>
          <section className="metrics"><Metric label="Compétiteurs filtrés" value={dashboardStats.competitors}/><Metric label="Produits filtrés" value={dashboardStats.quantity}/><Metric label="Tailles Supabase" value={dashboardStats.supabase}/><Metric label="Tailles manquantes" value={dashboardStats.missing} warning={dashboardStats.missing>0}/><Metric label="Shopify sans taille au départ" value={dashboardStats.initialMissing||0}/><Metric label="Régularisées Supabase" value={dashboardStats.regularized||0}/><Metric label="À valider" value={dashboardStats.toValidate||0} warning={(dashboardStats.toValidate||0)>0}/><Metric label="Produits fournisseur prêts" value={dashboardStats.supplierReady||0}/><Metric label="Engagements manquants" value={correctedMissingEngagementRows.length} warning={correctedMissingEngagementRows.length>0}/><Metric label="Tailles inhabituelles" value={dashboardStats.unusual||0} warning={(dashboardStats.unusual||0)>0}/></section>

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
                <strong>{correctedMissingEngagementRows.length}</strong>
              </div>
              <button onClick={() => setTab("missingEngagements")}>Voir la liste</button>
            </div>
          </section></>}<nav className="groupedNavPanel"><div className="navGroups">{tabGroups.map(g=><button key={g.id} className={currentGroup.id===g.id?"active":""} onClick={()=>selectGroup(g)}>{g.label}</button>)}</div><div className="cleanTabButtons">{visibleTabKeys.map(id=><button key={id} className={tab===id?"active":""} onClick={()=>setTab(id)}>{views[id].title}</button>)}</div></nav><section className="toolbar"><button onClick={()=>exportCsv(`${current.title||tab}.csv`,filteredRows)} disabled={!filteredRows.length}>Exporter ce rapport CSV</button><button onClick={()=>exportCsv(`club-${filters.dojo||"tous"}-${filters.team||"toutes-equipes"}-${current.title||tab}.csv`,filteredRows)} disabled={!filteredRows.length}>Export club / dojo du rapport affiché</button></section>{tab==="manualReconciliation"&&<section className="manualLinkPanel"><h2>Réconcilier les engagements manquants</h2><p>Choisis un compétiteur Fitofan sans engagement, puis associe uniquement une commande Shopify de type engagement.</p>{manualReconciliationRows.map(r=><div className="manualLinkRow" key={r.Compétiteur}><div><strong>{r.Compétiteur}</strong><span>{r.Dojo} — {r.Équipe} · Suggestion : {r["Meilleure suggestion"]||"aucune"} {r["Nom Shopify suggéré"]?`(${r["Nom Shopify suggéré"]})`:""}</span></div><select value={manualSelections[`${r.Compétiteur}|${r.Équipe}`]||r["Meilleure suggestion"]||""} onChange={e=>setManualSelections(p=>({...p,[`${r.Compétiteur}|${r.Équipe}`]:e.target.value}))}><option value="">Choisir une commande d’engagement</option>{engagementOrders.map(x=><option key={`${r.Compétiteur}-${x.Commande}`} value={x.Commande}>{x.Commande} — {x["Nom Shopify"]} — {x["Produit brut"]}</option>)}</select><button onClick={()=>linkEngagementToCompetitor(manualSelections[`${r.Compétiteur}|${r.Équipe}`]||r["Meilleure suggestion"],r.Compétiteur,r.Équipe)}>Lier engagement</button></div>)}{(manualLinkStatus||engagementCorrectionStatus)&&<small>{manualLinkStatus||engagementCorrectionStatus}</small>}</section>}{tab==="manualLinks"&&<section className="manualLinkPanel"><h2>Liens sauvegardés</h2>{manualLinks.map(l=><div className="manualLinkRow" key={l.id}><div><strong>{l.order_number}</strong><span>{l.shopify_competitor} → {l.fitofan_competitor}</span></div><button onClick={()=>removeManualLink(l.id)}>Supprimer</button></div>)}</section>}{tab==="manualParticipants"&&<section className="manualParticipantPanel"><h2>Ajouter un participant absent de Fitofan</h2><div className="manualForm"><input placeholder="Nom compétiteur" value={newParticipant.competitor} onChange={e=>setNewParticipant(p=>({...p,competitor:e.target.value}))}/><input placeholder="Email" value={newParticipant.email} onChange={e=>setNewParticipant(p=>({...p,email:e.target.value}))}/><input placeholder="Dojo" value={newParticipant.dojo} onChange={e=>setNewParticipant(p=>({...p,dojo:e.target.value}))}/><input placeholder="Équipe" value={newParticipant.team} onChange={e=>setNewParticipant(p=>({...p,team:e.target.value}))}/><button onClick={addManualParticipant}>Ajouter participant</button></div>{manualParticipantStatus&&<small>{manualParticipantStatus}</small>}</section>}{tab==="missingSizes"&&<section className="manualParticipantPanel"><h2>Ajouter une taille manuellement</h2><p>Pour chaque produit encore sans taille, tu peux ajouter une taille sauvegardée dans Supabase.</p>{reconciled.filter(r=>r.missing).slice(0,50).map(r=>{const k=`${r.order}|${r.productKey}`;return <div className="manualSizeRow" key={k}><div><strong>{r.competitor}</strong><span>{r.order} — {r.product} — {r.rawProduct}</span></div><input placeholder="Ex: M ou M / L" value={manualSizeInputs[k]||""} onChange={e=>setManualSizeInputs(p=>({...p,[k]:e.target.value}))}/><button onClick={()=>saveManualSize(r.order,r.productKey,r.competitor,r.rawProduct)}>Sauvegarder taille</button></div>})}</section>}{tab==="sizeValidation"&&<section className="sizeValidationPanel"><h2>Validation des tailles</h2><p>Corrige ou valide les tailles manquantes / inhabituelles. Si un même produit est en quantité 2 ou plus, écris plusieurs tailles dans le même champ, par exemple : M / L. L’export fournisseur les comptera séparément.</p>{reconciled.filter(r=>sizeNeedsValidation(r)).slice(0,150).map(r=>{const k=`${normalizeOrder(r.order)}|${r.productKey}`;const current=manualSizeInputs[k]??(r.finalSize||"");return <div className={`sizeValidationRow ${sizeValidationStatus(r)==="MANQUANTE"?"missing":"warning"}`} key={k}><div><strong>{r.competitor}</strong><span>{r.order} — {r.product} — {r.rawProduct}</span><small>Actuel : {r.finalSize||"MANQUANTE"} · Qté : {r.effectiveQty||r.quantity||1} · Source : {r.sourceSize||"aucune"} · Statut : {sizeValidationStatus(r)}</small></div><input placeholder="Ex: M ou M / L si quantité 2" value={current} onChange={e=>setManualSizeInputs(p=>({...p,[k]:e.target.value}))}/><button onClick={()=>saveManualSize(r.order,r.productKey,r.competitor,r.rawProduct)}>{r.finalSize?"Valider / modifier":"Ajouter taille"}</button></div>})}{manualParticipantStatus&&<small>{manualParticipantStatus}</small>}</section>}{tab==="commentSizeReview"&&<section className="sizeManagementPanel"><h2>Commandes avec commentaires</h2><p>Liste les commandes avec commentaire non traité. Tu peux corriger le dojo, l’équipe et la taille, puis valider le commentaire. Si la quantité est 2 ou plus, saisis les tailles dans le même champ, par exemple : XL / 12-14T. L’export fournisseur les comptera séparément.</p><div className="sizeEditTools"><input placeholder="Rechercher commande (#1180), nom, produit, commentaire..." value={commentSizeReviewSearch} onChange={e=>setCommentSizeReviewSearch(e.target.value)}/><button onClick={()=>setCommentSizeReviewSearch("")}>Effacer</button><span>{commentSizeReviewRows.length} ligne(s) à traiter</span></div>{commentSizeReviewRows.map(r=>{const item=r.raw;const k=commentCorrectionKey(item);const input=commentCorrectionInputs[k]||{};const sizeValue=input.size??(item.finalSize||"");const dojoValue=input.dojo??(item.dojo||"");const teamValue=input.team??(displayTeamCategory(item.team)||item.team||"");return <div className="sizeValidationRow commentReviewRow" key={`${k}|${item.rawProduct}`}><div><strong>{item.competitor}</strong><span>{item.order} — {item.product} — {item.rawProduct}</span><small>Dojo actuel : {item.dojo||"vide"} · Équipe actuelle : {displayTeamCategory(item.team)||item.team||"vide"} · Shopify : {item.shopifySize||"vide"} · Supabase : {item.supabaseSize||"vide"} · Finale : {item.finalSize||"vide"} · Qté : {quantityDisplay(item)} {multiSizeHelpForQty(item)?`· ${multiSizeHelpForQty(item)}`:""}</small>{hasUsefulComment(item.comments)&&<p className="commentBox">💬 {item.comments}</p>}</div><div className="commentCorrectionGrid"><input placeholder="Dojo" value={dojoValue} onChange={e=>setCommentCorrectionInputs(p=>({...p,[k]:{...(p[k]||{}),dojo:e.target.value}}))}/><select value={teamValue} onChange={e=>setCommentCorrectionInputs(p=>({...p,[k]:{...(p[k]||{}),team:e.target.value}}))}><option value="">Équipe</option>{teamCategoryOptions().map(t=><option key={t} value={t}>{t}</option>)}</select><input placeholder={quantityDisplay(item)>1?`Ex: XL / 12-14T pour Qté ${quantityDisplay(item)}`:"Ex: M"} value={sizeValue} onChange={e=>setCommentCorrectionInputs(p=>({...p,[k]:{...(p[k]||{}),size:e.target.value}}))}/><button onClick={()=>validateCommentRow(item)}>Valider commentaire</button></div></div>})}{commentSizeReviewRows.length===0&&<div className="emptyState">Aucun commentaire à traiter pour cette recherche.</div>}{commentCorrectionStatus&&<small>{commentCorrectionStatus}</small>}</section>}{tab==="orderSizeEdit"&&<section className="sizeManagementPanel"><h2>Modifier une taille manuellement sur une commande</h2><p>Recherche une commande, un compétiteur, un produit ou une taille, puis saisis la taille finale à appliquer. Cette taille sera prioritaire dans les tableaux et exports.</p><div className="sizeEditTools"><input placeholder="Rechercher commande (#1180 ou 1180), nom, produit, taille..." value={orderSizeEditSearch} onChange={e=>setOrderSizeEditSearch(e.target.value)}/><button onClick={()=>setOrderSizeEditSearch("")}>Effacer</button><span>{orderSizeEditRows.length} ligne(s) trouvée(s)</span></div>{orderSizeEditRows.map(r=>{const item=r.raw;const k=`${normalizeOrder(item.order)}|${item.productKey}`;const current=manualSizeInputs[k]??(item.finalSize||"");return <div className="sizeValidationRow" key={k}><div><strong>{item.competitor}</strong><span>{item.order} — {item.product} — {item.rawProduct}</span><small>Shopify : {item.shopifySize||"vide"} · Supabase : {item.supabaseSize||"vide"} · Finale : {item.finalSize||"vide"} · Qté : {item.effectiveQty||item.quantity||1} · Source : {item.sourceSize||"aucune"}</small></div><input placeholder="Ex: M ou M / L" value={current} onChange={e=>setManualSizeInputs(p=>({...p,[k]:e.target.value}))}/><button onClick={()=>saveManualSize(item.order,item.productKey,item.competitor,item.rawProduct)}>Enregistrer taille</button></div>})}{orderSizeEditRows.length===0&&<div className="emptyState">Aucune ligne trouvée pour cette recherche.</div>}{manualParticipantStatus&&<small>{manualParticipantStatus}</small>}</section>}{tab==="engagementManagement"&&<section className="engagementManagementPanel"><h2>Gestion liens Engagement ↔ Compétiteur</h2><p>Délie un engagement mal attribué ou associe une commande d’engagement à la bonne personne.</p><div className="engagementTools"><input placeholder="Rechercher nom, dojo, équipe, commande..." value={engagementSearch} onChange={e=>setEngagementSearch(e.target.value)}/><select value={engagementStatusFilter} onChange={e=>setEngagementStatusFilter(e.target.value)}><option value="all">Tous</option><option value="linked">Déjà liés</option><option value="missing">Sans engagement lié</option></select><button onClick={()=>{setEngagementSearch("");setEngagementStatusFilter("all");}}>Effacer</button><span>{engagementManagementFilteredRows.length} résultat(s)</span></div>{engagementManagementFilteredRows.map(r=>{const k=`${r.Compétiteur}|${r.Équipe}`;return <div className={`engagementLinkRow ${r.Statut==="LIÉ"?"linked":"missing"}`} key={`${r.Compétiteur}-${r.Équipe}-${r.Commande||"none"}`}><div><strong>{r.Compétiteur}</strong><span>{r.Dojo} — {r.Équipe}</span><small>{r.Statut==="LIÉ"?`Engagement lié : ${r.Commande} (${r["Nom Shopify"]})`:"Aucun engagement lié"}</small></div>{r.Statut==="LIÉ"?<button onClick={()=>unlinkEngagement(r.Commande,r.Compétiteur,r.Équipe)}>Délier</button>:<><select value={engagementLinkInputs[k]||""} onChange={e=>setEngagementLinkInputs(p=>({...p,[k]:e.target.value}))}><option value="">Choisir une commande d’engagement</option>{engagementOrders.map(e=><option key={`${k}-${e.Commande}`} value={e.Commande}>{e.Commande} — {e["Nom Shopify"]} — {e["Produit brut"]}</option>)}</select><button onClick={()=>linkEngagementToCompetitor(engagementLinkInputs[k],r.Compétiteur,r.Équipe)}>Lier</button></>}</div>})}{engagementCorrectionStatus&&<small>{engagementCorrectionStatus}</small>}</section>}{tab==="engagementCorrections"&&<section className="engagementManagementPanel"><h2>Corrections engagement sauvegardées</h2>{engagementCorrections.map(c=><div className="engagementLinkRow linked" key={c.id}><div><strong>{c.order_number}</strong><span>{c.action} — {c.competitor} — {c.team}</span></div><button onClick={()=>deleteEngagementCorrection(c.id)}>Supprimer correction</button></div>)}</section>}<section className="tableCard"><div className="tableHeader"><h2>{current.title}</h2><button onClick={()=>exportCsv(`${current.title||tab}.csv`,filteredRows)} disabled={!filteredRows.length}>Exporter ce tableau CSV</button></div><DataTable rows={filteredRows}/></section></div>}
function Select({label,value,options,onChange}){return <label><span>{label}</span><select value={value} onChange={e=>onChange(e.target.value)}><option value="">Tous</option>{options.map(o=><option key={o} value={o}>{o}</option>)}</select></label>}
function Metric({label,value,warning}){return <div className={`metric ${warning?"warning":""}`}><span>{label}</span><strong>{value}</strong></div>}
function UploadBox({title,info,onFile}){return <div className="uploadBox"><strong>{title}</strong><span>{info}</span><input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={e=>onFile(e.target.files?.[0])}/></div>}
function DataTable({rows}){if(!rows||!rows.length)return <p className="empty">Aucune donnée à afficher.</p>;const headers=Object.keys(rows[0]);return <div className="tableWrap"><table><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.map((row,i)=><tr key={i}>{headers.map(h=><td key={h} title={String(row[h]??"")}>{String(row[h]??"")}</td>)}</tr>)}</tbody></table></div>}
