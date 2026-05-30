import React, { useEffect, useMemo, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import "./style.css";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

const SIZE_WORDS = [
  "0000","000","00",
  "XXS","XS","S","M","L","XL","2XL","3XL","4XL",
  "0","1","2","3","4","5","6","7","8",
  "2-3T","3-4T","4-5T","5-6T","7-8T",
  "9-10T","11-12T","12-14T"
];

function clean(v) {
  return String(v || "").trim();
}

function norm(v) {
  return clean(v)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function hasSize(product) {
  return !norm(product).includes("sac");
}

function pick(row, keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && row[k] !== "") {
      return row[k];
    }
  }
  return "";
}

function normalizeProductName(name = "") {
  const n = norm(name);

  if (n.includes("t-shirt")) return "T-shirt";
  if (n.includes("hoodie")) return "Hoodie";
  if (n.includes("kangourou")) return "Kangourou";
  if (n.includes("veste")) return "Veste";
  if (n.includes("gant")) return "Gants";
  if (n.includes("protege-pied") || n.includes("protège-pied")) return "Protège-pied";
  if (n.includes("kimono")) return "Kimono";
  if (n.includes("sac")) return "Sac";
  if (n.includes("engagement")) return "Engagement";

  return clean(name);
}

function extractSize(value = "") {
  const original = clean(value);
  const v = norm(original);

  if (!v) return "";

  if (
    v.includes("engagement") ||
    v.includes("acompte 50") ||
    v.includes("commande equipement")
  ) {
    // continue quand même
  }

  const youth = v.match(
    /\b(2\s*-\s*3|3\s*-\s*4|4\s*-\s*5|5\s*-\s*6|7\s*-\s*8|9\s*-\s*10|11\s*-\s*12|12\s*-\s*14)\s*(t|ans?)?\b/i
  );

  if (youth) {
    return youth[1].replace(/\s+/g, "") + "T";
  }

  if (v.includes("youth medium")) return "YOUTH MEDIUM";

  const exact = SIZE_WORDS.find(
    (s) => norm(s) === v
  );

  if (exact) {
    return exact.toUpperCase();
  }

  const match = v.match(
    /\b(0000|000|00|xxxl|xxl|xl|xxs|xs|s|m|l|2xl|3xl|4xl|[0-8])\b/i
  );

  return match ? match[1].toUpperCase() : "";
}

function findShopifySize(row, rawProduct, productKey) {
  if (!hasSize(productKey)) {
    return "Taille unique";
  }

  // 1. TITRE SHOPIFY
  const fromTitle = extractSize(rawProduct);

  if (fromTitle) {
    return fromTitle;
  }

  // 2. PROPRIÉTÉS SHOPIFY
  for (const value of Object.values(row || {})) {
    const txt = clean(value);

    const prop = txt.match(/taille\s*[:：]\s*([^|•,;]+)/i);

    if (prop) {
      const s = extractSize(prop[1]);

      if (s) return s;
    }
  }

  // 3. FALLBACK
  const direct = extractSize(
    pick(row, [
      "Taille",
      "Size",
      "Variant",
      "Option",
      "Correction taille"
    ])
  );

  if (direct) return direct;

  return "";
}

function Metric({ label, value, warning }) {
  return (
    <div className={`metric ${warning ? "warning" : "ok"}`}>
      <div className="metricLabel">{label}</div>
      <div className="metricValue">{value}</div>
    </div>
  );
}

export default function App() {
  const [fitofanRows, setFitofanRows] = useState([]);
  const [shopifyRows, setShopifyRows] = useState([]);
  const [supabaseRows, setSupabaseRows] = useState([]);

  async function loadSupabase() {
    const { data, error } = await supabase
      .from("responses")
      .select("*");

    if (error) {
      console.error(error);
      return;
    }

    setSupabaseRows(data || []);
  }

  useEffect(() => {
    loadSupabase();
  }, []);

  function parseFitofan(file) {
    const reader = new FileReader();

    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, {
        type: "binary"
      });

      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const rows = XLSX.utils.sheet_to_json(sheet);

      setFitofanRows(rows);
      localStorage.setItem(
        "fitofanData",
        JSON.stringify(rows)
      );
    };

    reader.readAsBinaryString(file);
  }

  function parseShopify(file) {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        const parsed = [];

        results.data.forEach((row) => {
          const rawProduct =
            pick(row, [
              "Lineitem name",
              "Produit",
              "Nom du produit"
            ]) || "";

          const product = normalizeProductName(rawProduct);

          const competitor =
            pick(row, [
              "Competiteur",
              "Compétiteur",
              "Customer"
            ]) || "";

          const dojo =
            pick(row, ["Dojo"]) || "";

          const team =
            pick(row, ["Equipe", "Équipe"]) || "";

          const order =
            pick(row, [
              "Name",
              "Commande",
              "Order"
            ]) || "";

          let shopifySize = findShopifySize(
            row,
            rawProduct,
            product
          );

          const supabaseMatch = supabaseRows.find((r) => {
            return (
              norm(r.competitor) === norm(competitor) &&
              normalizeProductName(r.product_name) === product
            );
          });

          let finalSize = shopifySize;
          let source = "Shopify";

          if (
            supabaseMatch &&
            clean(supabaseMatch.confirmed_size)
          ) {
            finalSize = clean(
              supabaseMatch.confirmed_size
            );

            source = "Supabase";
          }

          parsed.push({
            competitor,
            dojo,
            team,
            order,
            rawProduct,
            product,
            shopifySize,
            finalSize,
            source
          });
        });

        setShopifyRows(parsed);
      }
    });
  }

  const missingSizes = useMemo(() => {
    return shopifyRows.filter(
      (r) =>
        hasSize(r.product) &&
        !clean(r.finalSize)
    );
  }, [shopifyRows]);

  return (
    <div className="app">
      <h1>Réconciliation Sunfuki V18</h1>

      <div className="topGrid">
        <section className="card">
          <h2>1. Fitofan</h2>

          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) =>
              parseFitofan(e.target.files[0])
            }
          />

          <p>{fitofanRows.length} lignes</p>
        </section>

        <section className="card">
          <h2>2. Supabase</h2>

          <button onClick={loadSupabase}>
            Charger Supabase
          </button>

          <p>{supabaseRows.length} lignes</p>
        </section>

        <section className="card">
          <h2>3. Shopify</h2>

          <input
            type="file"
            accept=".csv"
            onChange={(e) =>
              parseShopify(e.target.files[0])
            }
          />

          <p>{shopifyRows.length} lignes</p>
        </section>
      </div>

      <div className="metricsGrid">
        <Metric
          label="Produits Shopify"
          value={shopifyRows.length}
        />

        <Metric
          label="Tailles Supabase"
          value={
            shopifyRows.filter(
              (r) => r.source === "Supabase"
            ).length
          }
        />

        <Metric
          label="Tailles manquantes"
          value={missingSizes.length}
          warning={missingSizes.length > 0}
        />
      </div>

      <section className="card">
        <h2>
          Compétiteurs avec tailles manquantes
        </h2>

        <table>
          <thead>
            <tr>
              <th>Compétiteur</th>
              <th>Commande</th>
              <th>Produit</th>
            </tr>
          </thead>

          <tbody>
            {missingSizes.map((r, i) => (
              <tr key={i}>
                <td>{r.competitor}</td>
                <td>{r.order}</td>
                <td>{r.rawProduct}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>

      <section className="card">
        <h2>Détail Shopify + Supabase</h2>

        <table>
          <thead>
            <tr>
              <th>Compétiteur</th>
              <th>Produit Shopify brut</th>
              <th>Produit</th>
              <th>Taille Shopify</th>
              <th>Taille finale</th>
              <th>Source</th>
            </tr>
          </thead>

          <tbody>
            {shopifyRows.map((r, i) => (
              <tr key={i}>
                <td>{r.competitor}</td>
                <td>{r.rawProduct}</td>
                <td>{r.product}</td>
                <td>{r.shopifySize}</td>
                <td>{r.finalSize}</td>
                <td>{r.source}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>
    </div>
  );
}
