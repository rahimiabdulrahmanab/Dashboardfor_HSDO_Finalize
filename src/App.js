import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

/* Leaflet */
import "leaflet/dist/leaflet.css";
import { MapContainer, TileLayer, GeoJSON } from "react-leaflet";
import L from "leaflet";

/* Charts */
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
} from "recharts";

/* =========================
   CONFIG (RAW links only)
========================= */
var LOGO_URL =
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/Logo.png";

/* ✅ ADM1 provinces */
var ADM1_GEOJSON_URL =
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/geoBoundaries-AFG-ADM1.geojson";

/* MH */
var MH_XLSX_URL =
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/MH.xlsx";

/* HER (read all automatically) */
var HER_XLSX_URLS = [
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data.xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data%20(2).xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data%20(3).xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data%20(4).xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data%20(5).xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data%20(6).xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/data%20(7).xls",
  "https://raw.githubusercontent.com/rahimiabdulrahmanab/Data/main/BDK-HER2%20SAFE%20M%26E%20Framwork.xlsx",
];

/* =========================
   THEME (NGO blue)
========================= */
var THEME = {
  blue: "#0B5ED7",
  blue2: "#0A4FB6",
  blue3: "#083B82",
  bg: "#EAF2FF",
  panel: "#FFFFFF",
  border: "#B9D3FF",
  text: "#0B1B33",
  muted: "#4B5B73",
  headerGrad: "linear-gradient(90deg, #EAF2FF 0%, #F7FBFF 55%, #EAF2FF 100%)",
  barTop: "linear-gradient(90deg, #0B5ED7 0%, #2F80ED 55%, #0B5ED7 100%)",
  softShadow: "0 6px 18px rgba(11,94,215,0.12)",
};

var CHART_COLORS = [
  "#0B5ED7",
  "#2F80ED",
  "#00A3FF",
  "#6C63FF",
  "#2DBE7F",
  "#F59E0B",
  "#EF4444",
  "#A855F7",
  "#14B8A6",
];

/* =========================
   HELPERS
========================= */
function safeText(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function normName(s) {
  var x = safeText(s).toLowerCase();
  x = x.replace(/\./g, " ");
  x = x.replace(/-/g, " ");
  x = x.replace(/_/g, " ");
  x = x.replace(/\s+/g, " ").trim();
  x = x
    .replace(/\bprovince\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return x;
}

function toNumber(v) {
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  if (!s) return null;
  s = s.replace(/,/g, "").replace(/\s+/g, "");
  var n = Number(s);
  if (Number.isFinite(n)) return n;
  return null;
}

function fmtInt(v) {
  var n = toNumber(v);
  if (n === null) return "0";
  try {
    return new Intl.NumberFormat("en-CA").format(Math.round(n));
  } catch (e) {
    return String(Math.round(n));
  }
}

function monthNum(m) {
  var x = normName(m);
  if (x === "january") return 1;
  if (x === "february") return 2;
  if (x === "march") return 3;
  if (x === "april") return 4;
  if (x === "may") return 5;
  if (x === "june") return 6;
  if (x === "july") return 7;
  if (x === "august") return 8;
  if (x === "september") return 9;
  if (x === "october") return 10;
  if (x === "november") return 11;
  if (x === "december") return 12;
  return 99;
}

function parsePeriodName(periodStr) {
  var s = safeText(periodStr);
  if (!s) return { month: "", year: "" };
  var parts = s.split(/\s+/);
  if (parts.length >= 2) return { month: parts[0], year: parts[1] };
  return { month: s, year: "" };
}

function sortPeriods(periods) {
  var copy = (periods || []).slice();
  copy.sort(function (a, b) {
    var pa = parsePeriodName(a);
    var pb = parsePeriodName(b);
    var ya = Number(pa.year) || 0;
    var yb = Number(pb.year) || 0;
    if (ya !== yb) return ya - yb;
    return monthNum(pa.month) - monthNum(pb.month);
  });
  return copy;
}

/* =========================
   Robust .xls parsing
========================= */
function sheetTo2D(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
}

function findHeaderRowIndex(rows2D) {
  var r;
  for (r = 0; r < rows2D.length && r < 80; r++) {
    var v = rows2D[r] && rows2D[r][0] ? String(rows2D[r][0]) : "";
    v = v.trim().toLowerCase();
    if (v === "periodname" || v === "period" || v === "period_name") return r;
  }
  return -1;
}

function parseWeirdSheetToObjects(ws) {
  var arr = sheetTo2D(ws);
  if (!arr || !arr.length) return [];

  var headerRow = findHeaderRowIndex(arr);
  if (headerRow < 0) headerRow = 0;

  var headers = arr[headerRow] || [];
  var out = [];
  var r, c;

  for (r = headerRow + 1; r < arr.length; r++) {
    var row = arr[r];
    if (!row) continue;

    var empty = true;
    for (c = 0; c < headers.length; c++) {
      if (safeText(row[c])) {
        empty = false;
        break;
      }
    }
    if (empty) continue;

    var obj = {};
    for (c = 0; c < headers.length; c++) {
      var h = safeText(headers[c]);
      if (!h) continue;
      obj[h] = row[c];
    }
    out.push(obj);
  }

  return out;
}

/* Wide -> Long */
function normalizeWideToLong(sheetRows, datasetName) {
  if (!sheetRows || !sheetRows.length) return [];

  var first = sheetRows[0];
  var cols = Object.keys(first || {});
  var periodCol = "";

  var i;
  for (i = 0; i < cols.length; i++) {
    var c = String(cols[i]).toLowerCase().trim();
    if (c === "periodname" || c === "period" || c === "period_name") {
      periodCol = cols[i];
      break;
    }
  }
  if (!periodCol && cols.length > 0) periodCol = cols[0];

  var out = [];
  var r, j;

  for (r = 0; r < sheetRows.length; r++) {
    var row = sheetRows[r];
    var period = safeText(row[periodCol]);
    if (!period) continue;

    var p = parsePeriodName(period);

    for (j = 0; j < cols.length; j++) {
      var col = cols[j];
      if (col === periodCol) continue;

      var num = toNumber(row[col]);
      if (num === null) continue;

      var indicator = safeText(col);
      var family = datasetName;

      if (indicator.indexOf(" - ") !== -1) {
        var parts = indicator.split(" - ");
        family = safeText(parts[0]) || datasetName;
        indicator = safeText(parts.slice(1).join(" - ")) || indicator;
      }

      out.push({
        Dataset: datasetName,
        Period: period,
        Year: safeText(p.year),
        Month: safeText(p.month),
        Family: family,
        Indicator: indicator,
        Value: num,
        Province: "Badakhshan",
      });
    }
  }

  return out;
}

/* =========================
   APP
========================= */
export default function App() {
  /* Desktop-only lock */
  var [isSmallScreen, setIsSmallScreen] = useState(false);
  useEffect(function () {
    function onResize() {
      var w = window.innerWidth || 0;
      var h = window.innerHeight || 0;
      /* allow 14" laptops; still block phones */
      if (w < 1024 || h < 640) setIsSmallScreen(true);
      else setIsSmallScreen(false);
    }
    onResize();
    window.addEventListener("resize", onResize);
    return function () {
      window.removeEventListener("resize", onResize);
    };
  }, []);

  /* Data */
  var [rows, setRows] = useState([]);
  var [loading, setLoading] = useState(true);
  var [error, setError] = useState("");

  /* Geo */
  var [geo, setGeo] = useState(null);
  var [geoError, setGeoError] = useState("");
  var [afgBounds, setAfgBounds] = useState(null);
  var mapRef = useRef(null);

  /* Filters */
  var [dataset, setDataset] = useState("HER");
  var [year, setYear] = useState("All");
  var [month, setMonth] = useState("All");
  var [indicator, setIndicator] = useState("Auto (recommended)");

  /* =========================
     Styles (one-page, no scroll)
========================= */
  var styles = {
    page: {
      height: "100vh",
      overflow: "hidden",
      background: THEME.bg,
      color: THEME.text,
      fontFamily: "Arial, sans-serif",
      display: "flex",
      flexDirection: "column",
    },

    header: {
      flex: "0 0 auto",
      background: THEME.headerGrad,
      borderBottom: "5px solid " + THEME.blue,
      padding: "10px 16px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: 12,
    },
    brand: { display: "flex", alignItems: "center", gap: 12 },
    logo: { width: 46, height: 46, objectFit: "contain" },
    title: { fontSize: 18, fontWeight: 900, lineHeight: 1.1 },
    subtitle: {
      fontSize: 12,
      fontWeight: 800,
      color: THEME.muted,
      marginTop: 2,
    },

    headerRight: { display: "flex", alignItems: "center", gap: 10 },
    pill: {
      border: "1px solid " + THEME.border,
      background: "#F7FBFF",
      borderRadius: 999,
      padding: "8px 12px",
      fontWeight: 900,
    },
    btn: {
      background: THEME.blue,
      border: "1px solid " + THEME.blue2,
      color: "white",
      fontWeight: 900,
      borderRadius: 10,
      padding: "8px 12px",
      cursor: "pointer",
    },

    body: {
      flex: "1 1 auto",
      minHeight: 0,
      padding: 10,
      display: "grid",
      gridTemplateColumns: "280px 1fr",
      gap: 10,
    },

    panel: {
      background: "#FFFFFF",
      border: "1px solid " + THEME.border,
      borderRadius: 14,
      boxShadow: THEME.softShadow,
      overflow: "hidden",
      minHeight: 0,
    },

    panelHeader: {
      background: THEME.barTop,
      color: "white",
      fontWeight: 900,
      padding: "10px 12px",
      fontSize: 14,
    },

    panelBody: { padding: 12 },

    /* Filters */
    tabRow: { display: "flex", gap: 10, marginBottom: 12 },
    tab: function (active) {
      return {
        flex: 1,
        borderRadius: 999,
        border: "1px solid " + (active ? THEME.blue : THEME.border),
        background: active ? THEME.blue : "#FFFFFF",
        color: active ? "white" : THEME.text,
        fontWeight: 900,
        padding: "10px 10px",
        cursor: "pointer",
        textAlign: "center",
      };
    },
    label: {
      fontSize: 12,
      fontWeight: 900,
      color: THEME.muted,
      marginBottom: 6,
    },
    select: {
      width: "100%",
      height: 40,
      borderRadius: 10,
      border: "1px solid " + THEME.border,
      padding: "0 10px",
      fontWeight: 900,
      outline: "none",
      background: "white",
      marginBottom: 10,
    },

    /* Main */
    main: {
      minHeight: 0,
      display: "grid",
      gridTemplateRows: "auto 1fr",
      gap: 10,
    },

    kpiRow: {
      display: "grid",
      gridTemplateColumns: "repeat(6, 1fr)",
      gap: 10,
    },

    kpi: {
      background: "#FFFFFF",
      border: "1px solid " + THEME.border,
      borderRadius: 12,
      padding: 10,
      position: "relative",
      overflow: "hidden",
      boxShadow: "0 4px 14px rgba(11,94,215,0.08)",
      minHeight: 62,
    },
    kpiTop: {
      position: "absolute",
      top: 0,
      left: 0,
      right: 0,
      height: 4,
      background: THEME.blue,
    },
    kpiLabel: { fontSize: 11, fontWeight: 900, color: THEME.muted },
    kpiValue: { fontSize: 18, fontWeight: 900, marginTop: 4 },

    /* Layout: Map left, Charts right (more charts) */
    bottomGrid: {
      minHeight: 0,
      display: "grid",
      gridTemplateColumns: "0.56fr 0.44fr",
      gap: 10,
    },

    leftCol: {
      minHeight: 0,
      display: "grid",
      gridTemplateRows: "auto 1fr",
      gap: 10,
    },

    mapBox: {
      border: "1px solid " + THEME.border,
      borderRadius: 12,
      overflow: "hidden",
      background: "white",
    },

    mapHeader: {
      background: THEME.barTop,
      color: "white",
      fontWeight: 900,
      padding: "10px 12px",
      fontSize: 13,
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
    },

    /* ✅ Bigger map height (and pan only) */
    mapWrap: {
      height: 380,
      width: "100%",
      background: "white",
    },

    infoGrid: {
      display: "grid",
      gridTemplateColumns: "1fr 1fr",
      gap: 10,
      minHeight: 0,
      alignItems: "stretch",
    },

    infoCard: {
      border: "1px solid " + THEME.border,
      borderRadius: 12,
      background: "#FFFFFF",
      padding: 12,
      boxShadow: "0 4px 14px rgba(11,94,215,0.06)",
      minHeight: 0,
      overflow: "auto",
      wordBreak: "break-word",
    },

    infoTitle: { fontWeight: 900, color: THEME.blue3, marginBottom: 8 },

    /* Right charts: 2x2 grid */
    rightCol: {
      minHeight: 0,
      display: "grid",
      gridTemplateRows: "1fr 1fr",
      gap: 10,
    },

    chartGrid2x2: {
      minHeight: 0,
      display: "grid",
      gridTemplateColumns: "1fr 1fr",
      gridTemplateRows: "1fr 1fr",
      gap: 10,
    },

    chartBox: {
      border: "1px solid " + THEME.border,
      borderRadius: 12,
      background: "#FFFFFF",
      padding: 10,
      boxShadow: "0 4px 14px rgba(11,94,215,0.06)",
      minHeight: 0,
      overflow: "hidden",
    },
    chartTitle: {
      fontWeight: 900,
      marginBottom: 8,
      color: THEME.blue3,
      fontSize: 13,
    },

    smallOverlay: {
      position: "fixed",
      left: 0,
      right: 0,
      top: 0,
      bottom: 0,
      background: "rgba(10,35,80,0.88)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 99999,
      padding: 16,
    },
    smallCard: {
      width: "100%",
      maxWidth: 560,
      background: "#FFFFFF",
      borderRadius: 16,
      border: "2px solid " + THEME.blue,
      padding: 18,
      textAlign: "center",
    },
    smallTitle: { fontSize: 18, fontWeight: 900, color: THEME.blue3 },
    smallText: {
      marginTop: 10,
      fontSize: 13,
      fontWeight: 800,
      color: THEME.muted,
      lineHeight: 1.5,
    },
  };

  /* =========================
     Load GeoJSON + bounds
========================= */
  useEffect(function () {
    setGeoError("");
    fetch(ADM1_GEOJSON_URL)
      .then(function (res) {
        if (!res.ok)
          throw new Error("ADM1 GeoJSON fetch failed. Check RAW link.");
        return res.json();
      })
      .then(function (j) {
        setGeo(j);

        /* bounds for Afghanistan */
        try {
          var layer = L.geoJSON(j);
          var b = layer.getBounds();
          setAfgBounds(b);
        } catch (e) {}

        /* fit after load */
        setTimeout(function () {
          if (mapRef.current) {
            try {
              mapRef.current.invalidateSize(true);
              /* keep Afghanistan centered; pan allowed */
              mapRef.current.setView([34.5, 66.0], 6, { animate: false });
              if (afgBounds) {
                mapRef.current.setMaxBounds(afgBounds.pad(0.15));
                mapRef.current.options.maxBoundsViscosity = 1.0;
              }
            } catch (e2) {}
          }
        }, 350);
      })
      .catch(function (e) {
        setGeoError(String(e && e.message ? e.message : e));
      });
  }, [afgBounds]);

  /* =========================
     Load MH + all HER files
========================= */
  useEffect(function () {
    setLoading(true);
    setError("");

    function fetchExcel(url) {
      return fetch(url).then(function (res) {
        if (!res.ok) throw new Error("Fetch failed: " + url);
        return res.arrayBuffer();
      });
    }

    function readWorkbook(buf) {
      return XLSX.read(buf, { type: "array" });
    }

    function parseFirstSheetStandard(wb) {
      var sh = wb.SheetNames[0];
      var ws = wb.Sheets[sh];
      return XLSX.utils.sheet_to_json(ws, { defval: "" });
    }

    function parseWeirdFromAllSheets(wb) {
      var all = [];
      var i;
      for (i = 0; i < wb.SheetNames.length; i++) {
        var sh = wb.SheetNames[i];
        var ws = wb.Sheets[sh];
        var objs = parseWeirdSheetToObjects(ws);
        if (objs && objs.length) all = all.concat(objs);
      }
      return all;
    }

    fetchExcel(MH_XLSX_URL)
      .then(function (buf) {
        var wb = readWorkbook(buf);
        var mhRows = parseFirstSheetStandard(wb);
        var mhLong = normalizeWideToLong(mhRows, "MH");
        return { mhLong: mhLong };
      })
      .then(function (state1) {
        var promises = [];
        var i;

        for (i = 0; i < HER_XLSX_URLS.length; i++) {
          (function (u) {
            var p = fetchExcel(u)
              .then(function (buf) {
                var wb = readWorkbook(buf);

                var herRows = parseWeirdFromAllSheets(wb);
                if (!herRows || !herRows.length)
                  herRows = parseFirstSheetStandard(wb);

                var herLong = normalizeWideToLong(herRows, "HER");
                return { ok: true, rows: herLong };
              })
              .catch(function () {
                return { ok: false, rows: [] };
              });

            promises.push(p);
          })(HER_XLSX_URLS[i]);
        }

        return Promise.all(promises).then(function (arr) {
          var herAll = [];
          var k;
          for (k = 0; k < arr.length; k++) {
            if (arr[k] && arr[k].rows && arr[k].rows.length)
              herAll = herAll.concat(arr[k].rows);
          }

          var allRows = [].concat(state1.mhLong).concat(herAll);

          /* If HER empty, default to MH silently */
          if (!herAll.length) setDataset("MH");

          setRows(allRows);
          setLoading(false);
        });
      })
      .catch(function (e) {
        setLoading(false);
        setError(String(e && e.message ? e.message : e));
      });
  }, []);

  function resetAll() {
    setDataset("HER");
    setYear("All");
    setMonth("All");
    setIndicator("Auto (recommended)");
  }

  /* =========================
     Options (year/month/indicator)
========================= */
  var years = useMemo(
    function () {
      var m = {};
      var i;
      for (i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (r.Dataset !== dataset) continue;
        if (safeText(r.Year)) m[r.Year] = true;
      }
      return Object.keys(m).sort(function (a, b) {
        return (Number(a) || 0) - (Number(b) || 0);
      });
    },
    [rows, dataset]
  );

  var months = useMemo(
    function () {
      var m = {};
      var i;
      for (i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (r.Dataset !== dataset) continue;
        if (year !== "All" && safeText(r.Year) !== safeText(year)) continue;
        if (safeText(r.Month)) m[r.Month] = true;
      }
      var keys = Object.keys(m);
      keys.sort(function (a, b) {
        return monthNum(a) - monthNum(b);
      });
      return keys;
    },
    [rows, dataset, year]
  );

  var indicatorOptions = useMemo(
    function () {
      var totals = {};
      var i;
      for (i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (r.Dataset !== dataset) continue;
        var name = safeText(r.Indicator);
        if (!name) continue;
        if (!totals[name]) totals[name] = 0;
        totals[name] += toNumber(r.Value) || 0;
      }

      var arr = [];
      var k;
      for (k in totals) arr.push({ name: k, total: totals[k] });

      arr.sort(function (a, b) {
        return (b.total || 0) - (a.total || 0);
      });

      return arr.slice(0, 60);
    },
    [rows, dataset]
  );

  var recommendedIndicator = useMemo(
    function () {
      if (!indicatorOptions.length) return "";
      return indicatorOptions[0].name;
    },
    [indicatorOptions]
  );

  function shortLabel(s) {
    var t = safeText(s);
    if (t.length <= 55) return t;
    return t.slice(0, 52) + "…";
  }

  var chosenIndicator =
    indicator === "Auto (recommended)" ? recommendedIndicator : indicator;

  /* =========================
     FILTERED rows (selected indicator only)
========================= */
  var filteredIndicatorRows = useMemo(
    function () {
      var out = [];
      var i;

      for (i = 0; i < rows.length; i++) {
        var r = rows[i];

        if (r.Dataset !== dataset) continue;
        if (year !== "All" && safeText(r.Year) !== safeText(year)) continue;
        if (month !== "All" && safeText(r.Month) !== safeText(month)) continue;

        if (
          chosenIndicator &&
          safeText(r.Indicator) !== safeText(chosenIndicator)
        )
          continue;

        out.push(r);
      }
      return out;
    },
    [rows, dataset, year, month, chosenIndicator]
  );

  /* =========================
     FILTERED rows (all indicators) for extra stats
========================= */
  var filteredAllRows = useMemo(
    function () {
      var out = [];
      var i;

      for (i = 0; i < rows.length; i++) {
        var r = rows[i];

        if (r.Dataset !== dataset) continue;
        if (year !== "All" && safeText(r.Year) !== safeText(year)) continue;
        if (month !== "All" && safeText(r.Month) !== safeText(month)) continue;

        out.push(r);
      }
      return out;
    },
    [rows, dataset, year, month]
  );

  /* =========================
     Trend (selected indicator)
========================= */
  var trendSelected = useMemo(
    function () {
      var m = {};
      var i;
      for (i = 0; i < filteredIndicatorRows.length; i++) {
        var r = filteredIndicatorRows[i];
        var k = safeText(r.Period);
        if (!k) continue;
        if (!m[k]) m[k] = 0;
        m[k] += toNumber(r.Value) || 0;
      }
      var keys = sortPeriods(Object.keys(m));
      var out = [];
      for (i = 0; i < keys.length; i++)
        out.push({ name: keys[i], value: Math.round(m[keys[i]] || 0) });
      return out;
    },
    [filteredIndicatorRows]
  );

  /* =========================
     Trend (TOTAL all indicators) - more meaningful overall activity
========================= */
  var trendTotal = useMemo(
    function () {
      var m = {};
      var i;
      for (i = 0; i < filteredAllRows.length; i++) {
        var r = filteredAllRows[i];
        var k = safeText(r.Period);
        if (!k) continue;
        if (!m[k]) m[k] = 0;
        m[k] += toNumber(r.Value) || 0;
      }
      var keys = sortPeriods(Object.keys(m));
      var out = [];
      for (i = 0; i < keys.length; i++)
        out.push({ name: keys[i], value: Math.round(m[keys[i]] || 0) });
      return out;
    },
    [filteredAllRows]
  );

  /* =========================
     Top indicators (colorful) - all indicators
========================= */
  var topIndicators = useMemo(
    function () {
      var m = {};
      var i;
      for (i = 0; i < filteredAllRows.length; i++) {
        var r = filteredAllRows[i];
        var k = safeText(r.Indicator);
        if (!k) continue;

        if (!m[k]) m[k] = 0;
        m[k] += toNumber(r.Value) || 0;
      }

      var arr = [];
      var key;
      for (key in m) arr.push({ name: key, value: Math.round(m[key] || 0) });

      arr.sort(function (a, b) {
        return (b.value || 0) - (a.value || 0);
      });

      return arr.slice(0, 8);
    },
    [filteredAllRows]
  );

  /* =========================
     Donut: share by family
========================= */
  var familyShare = useMemo(
    function () {
      var m = {};
      var i;
      for (i = 0; i < filteredAllRows.length; i++) {
        var r = filteredAllRows[i];
        var k = safeText(r.Family) || dataset;
        if (!m[k]) m[k] = 0;
        m[k] += toNumber(r.Value) || 0;
      }
      var arr = [];
      var key;
      for (key in m) arr.push({ name: key, value: Math.round(m[key] || 0) });

      arr.sort(function (a, b) {
        return (b.value || 0) - (a.value || 0);
      });
      return arr;
    },
    [filteredAllRows, dataset]
  );

  function renderBarCells(data) {
    var out = [];
    var i;
    for (i = 0; i < (data || []).length; i++) {
      out.push(
        <Cell key={"bar-" + i} fill={CHART_COLORS[i % CHART_COLORS.length]} />
      );
    }
    return out;
  }

  function renderPieCells(data) {
    var out = [];
    var i;
    for (i = 0; i < (data || []).length; i++) {
      out.push(
        <Cell key={"pie-" + i} fill={CHART_COLORS[i % CHART_COLORS.length]} />
      );
    }
    return out;
  }

  /* =========================
     KPI (selected indicator)
========================= */
  var kpiSelected = useMemo(
    function () {
      var total = 0;
      var i;
      for (i = 0; i < trendSelected.length; i++)
        total += trendSelected[i].value || 0;

      var latest = trendSelected.length
        ? trendSelected[trendSelected.length - 1].value || 0
        : 0;
      var prev =
        trendSelected.length >= 2
          ? trendSelected[trendSelected.length - 2].value || 0
          : 0;

      var changePct = 0;
      if (prev > 0) changePct = ((latest - prev) / prev) * 100;

      var avg = trendSelected.length ? total / trendSelected.length : 0;

      return {
        latest: Math.round(latest),
        prev: Math.round(prev),
        changePct: changePct,
        total: Math.round(total),
        avg: Math.round(avg),
        periods: trendSelected.length,
      };
    },
    [trendSelected]
  );

  /* KPI (overall total all indicators) */
  var kpiTotal = useMemo(
    function () {
      var total = 0;
      var i;
      for (i = 0; i < trendTotal.length; i++) total += trendTotal[i].value || 0;

      var latest = trendTotal.length
        ? trendTotal[trendTotal.length - 1].value || 0
        : 0;
      var prev =
        trendTotal.length >= 2
          ? trendTotal[trendTotal.length - 2].value || 0
          : 0;

      var changePct = 0;
      if (prev > 0) changePct = ((latest - prev) / prev) * 100;

      var avg = trendTotal.length ? total / trendTotal.length : 0;

      return {
        latest: Math.round(latest),
        prev: Math.round(prev),
        changePct: changePct,
        total: Math.round(total),
        avg: Math.round(avg),
        periods: trendTotal.length,
      };
    },
    [trendTotal]
  );

  /* =========================
     MAP: highlight only Badakhshan
========================= */
  function getProvName(feature) {
    var p = (feature && feature.properties) || {};
    return (
      safeText(p.shapeName) ||
      safeText(p.ADM1_NAME) ||
      safeText(p.NAME_1) ||
      safeText(p.province) ||
      safeText(p.PROVINCE) ||
      safeText(p.name) ||
      ""
    );
  }

  function isBadakhshanProvince(feature) {
    return normName(getProvName(feature)) === "badakhshan";
  }

  function styleAllProvinces() {
    return {
      color: "#8AA7D6",
      weight: 1,
      fillColor: "#CFE0FF",
      fillOpacity: 0.06,
      interactive: false,
    };
  }

  function styleBadakhshan() {
    return {
      color: THEME.blue3,
      weight: 2.8,
      fillColor: THEME.blue,
      fillOpacity: 0.7,
      interactive: false,
    };
  }

  function badakhshanOnlyGeo(geojson) {
    if (!geojson || !geojson.features) return null;
    var f = [];
    var i;
    for (i = 0; i < geojson.features.length; i++) {
      if (isBadakhshanProvince(geojson.features[i]))
        f.push(geojson.features[i]);
    }
    return { type: "FeatureCollection", features: f };
  }

  /* =========================
     SMALL SCREEN BLOCK
========================= */
  if (isSmallScreen) {
    return (
      <div style={styles.smallOverlay}>
        <div style={styles.smallCard}>
          <div style={styles.smallTitle}>
            This dashboard needs a bigger screen
          </div>
          <div style={styles.smallText}>
            Please open on a laptop/desktop.
            <br />
            Minimum recommended size: <b>1024px width</b>.
          </div>
        </div>
      </div>
    );
  }

  /* =========================
     LOADING
========================= */
  if (loading) {
    return (
      <div style={styles.page}>
        <div style={{ padding: 16 }}>
          <div style={styles.panel}>
            <div style={styles.panelHeader}>Loading Badakhshan Dashboard…</div>
            <div style={styles.panelBody}>
              <div style={{ fontWeight: 900 }}>
                Reading HER + MH Excel files…
              </div>
              <div
                style={{ marginTop: 8, color: THEME.muted, fontWeight: 800 }}
              >
                Please wait a moment.
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  var badGeo = geo ? badakhshanOnlyGeo(geo) : null;
  var hasBad = badGeo && badGeo.features && badGeo.features.length;

  var hasTrendSelected = trendSelected && trendSelected.length >= 2;
  var hasTrendTotal = trendTotal && trendTotal.length >= 2;

  return (
    <div style={styles.page}>
      <style>{`.leaflet-container { height: 100%; width: 100%; }`}</style>

      {/* HEADER */}
      <div style={styles.header}>
        <div style={styles.brand}>
          <img src={LOGO_URL} alt="Logo" style={styles.logo} />
          <div>
            <div style={styles.title}>Badakhshan Public Services Dashboard</div>
            <div style={styles.subtitle}>
              Uses HSDO's Own Data • Afghanistan map (pan only) • More charts
              & statistics
            </div>
          </div>
        </div>

        <div style={styles.headerRight}>
          <div style={styles.pill}>
            Province: <span style={{ color: THEME.blue2 }}>Badakhshan</span>
          </div>
          <button type="button" style={styles.btn} onClick={resetAll}>
            Reset
          </button>
        </div>
      </div>

      {/* BODY */}
      <div style={styles.body}>
        {/* LEFT FILTERS */}
        <div style={styles.panel}>
          <div style={styles.panelHeader}>Filters</div>
          <div style={styles.panelBody}>
            <div style={styles.tabRow}>
              <div
                style={styles.tab(dataset === "HER")}
                onClick={function () {
                  setDataset("HER");
                  setYear("All");
                  setMonth("All");
                  setIndicator("Auto (recommended)");
                }}
              >
                HER
              </div>

              <div
                style={styles.tab(dataset === "MH")}
                onClick={function () {
                  setDataset("MH");
                  setYear("All");
                  setMonth("All");
                  setIndicator("Auto (recommended)");
                }}
              >
                MH
              </div>
            </div>

            <div style={styles.label}>Key Indicator</div>
            <select
              style={styles.select}
              value={indicator}
              onChange={function (e) {
                setIndicator(e.target.value);
              }}
            >
              <option value="Auto (recommended)">Auto (recommended)</option>
              {indicatorOptions.map(function (x) {
                return (
                  <option key={x.name} value={x.name}>
                    {shortLabel(x.name)}
                  </option>
                );
              })}
            </select>

            <div style={styles.label}>Year</div>
            <select
              style={styles.select}
              value={year}
              onChange={function (e) {
                setYear(e.target.value);
                setMonth("All");
              }}
            >
              <option value="All">All</option>
              {years.map(function (y) {
                return (
                  <option key={y} value={y}>
                    {y}
                  </option>
                );
              })}
            </select>

            <div style={styles.label}>Month</div>
            <select
              style={styles.select}
              value={month}
              onChange={function (e) {
                setMonth(e.target.value);
              }}
            >
              <option value="All">All</option>
              {months.map(function (m) {
                return (
                  <option key={m} value={m}>
                    {m}
                  </option>
                );
              })}
            </select>

            <div
              style={{
                marginTop: 10,
                fontSize: 12,
                fontWeight: 900,
                color: THEME.muted,
                lineHeight: 1.45,
              }}
            >
              Key indicator:
              <br />
              <span style={{ color: THEME.blue3 }}>
                {shortLabel(chosenIndicator || "")}
              </span>
              <br />
              Records loaded:{" "}
              <span style={{ color: THEME.blue3 }}>{fmtInt(rows.length)}</span>
              <br />
              In view (current filters):{" "}
              <span style={{ color: THEME.blue3 }}>
                {fmtInt(filteredAllRows.length)}
              </span>
              <br />
              Map highlight:{" "}
              <span style={{ color: THEME.blue3 }}>
                {hasBad ? "Badakhshan OK" : "Not found"}
              </span>
            </div>

            {error ? (
              <div style={{ marginTop: 12, color: "#B00020", fontWeight: 900 }}>
                Data warning: {error}
              </div>
            ) : null}
          </div>
        </div>

        {/* MAIN */}
        <div style={styles.main}>
          {/* KPI ROW (selected indicator + overall total) */}
          <div style={styles.kpiRow}>
            <div style={styles.kpi}>
              <div style={styles.kpiTop} />
              <div style={styles.kpiLabel}>Latest (key indicator)</div>
              <div style={styles.kpiValue}>{fmtInt(kpiSelected.latest)}</div>
            </div>

            <div style={styles.kpi}>
              <div style={styles.kpiTop} />
              <div style={styles.kpiLabel}>Previous (key indicator)</div>
              <div style={styles.kpiValue}>{fmtInt(kpiSelected.prev)}</div>
            </div>

            <div style={styles.kpi}>
              <div style={styles.kpiTop} />
              <div style={styles.kpiLabel}>Change % (key indicator)</div>
              <div style={styles.kpiValue}>
                {kpiSelected.prev > 0
                  ? (kpiSelected.changePct >= 0 ? "+" : "") +
                    kpiSelected.changePct.toFixed(1) +
                    "%"
                  : "0.0%"}
              </div>
            </div>

            <div style={styles.kpi}>
              <div style={styles.kpiTop} />
              <div style={styles.kpiLabel}>Latest (overall total)</div>
              <div style={styles.kpiValue}>{fmtInt(kpiTotal.latest)}</div>
            </div>

            <div style={styles.kpi}>
              <div style={styles.kpiTop} />
              <div style={styles.kpiLabel}>Total (overall all periods)</div>
              <div style={styles.kpiValue}>{fmtInt(kpiTotal.total)}</div>
            </div>

            <div style={styles.kpi}>
              <div style={styles.kpiTop} />
              <div style={styles.kpiLabel}>Avg / period (overall)</div>
              <div style={styles.kpiValue}>{fmtInt(kpiTotal.avg)}</div>
            </div>
          </div>

          {/* BOTTOM GRID */}
          <div style={styles.bottomGrid}>
            {/* LEFT: Map + info */}
            <div style={styles.leftCol}>
              <div style={styles.mapBox}>
                <div style={styles.mapHeader}>
                  <span>Afghanistan Map — Badakhshan Only</span>
                  <span
                    style={{ fontSize: 12, fontWeight: 900, opacity: 0.95 }}
                  >
                    Pan only • No zoom
                  </span>
                </div>

                {geoError ? (
                  <div
                    style={{ padding: 10, color: "#B00020", fontWeight: 900 }}
                  >
                    Map error: {geoError}
                  </div>
                ) : null}

                <div style={styles.mapWrap}>
                  <MapContainer
                    center={[34.5, 66.0]}
                    zoom={6}
                    scrollWheelZoom={false}
                    zoomControl={false}
                    dragging={true}
                    doubleClickZoom={false}
                    touchZoom={false}
                    boxZoom={false}
                    keyboard={false}
                    whenCreated={function (m) {
                      mapRef.current = m;

                      /* Lock zoom */
                      try {
                        m.setMinZoom(6);
                        m.setMaxZoom(6);
                      } catch (e) {}

                      /* Keep view inside Afghanistan bounds (pan only) */
                      setTimeout(function () {
                        try {
                          m.invalidateSize(true);
                          m.setView([34.5, 66.0], 6, { animate: false });
                          if (afgBounds) {
                            m.setMaxBounds(afgBounds.pad(0.15));
                            m.options.maxBoundsViscosity = 1.0;
                          }
                        } catch (e2) {}
                      }, 250);
                    }}
                  >
                    <TileLayer
                      attribution="&copy; OpenStreetMap contributors"
                      url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                    />

                    {/* Grey all provinces */}
                    {geo && geo.features ? (
                      <GeoJSON data={geo} style={styleAllProvinces} />
                    ) : null}

                    {/* Blue Badakhshan */}
                    {badGeo && badGeo.features && badGeo.features.length ? (
                      <GeoJSON data={badGeo} style={styleBadakhshan} />
                    ) : null}
                  </MapContainer>
                </div>

                <div
                  style={{
                    padding: "8px 12px",
                    fontSize: 12,
                    fontWeight: 900,
                    color: THEME.muted,
                  }}
                >
                  Full Afghanistan map is shown. Only Badakhshan is highlighted
                  (province-level). Pan is enabled; zoom is locked.
                </div>
              </div>

              {/* Info boxes under map */}
              <div style={styles.infoGrid}>
                <div style={styles.infoCard}>
                  <div style={styles.infoTitle}>Dashboard overview</div>
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 800,
                      color: THEME.muted,
                      lineHeight: 1.5,
                    }}
                  >
                    This dashboard shows program results for{" "}
                    <b>Badakhshan province</b> only.
                    <br />
                    <br />
                    It automatically reads <b>
                      all uploaded HER Excel files
                    </b>{" "}
                    and the <b>MH Excel</b> file, then generates the KPIs and
                    charts.
                  </div>
                </div>

                <div style={styles.infoCard}>
                  <div style={styles.infoTitle}>Coverage & status</div>
                  <div
                    style={{
                      fontSize: 12,
                      fontWeight: 900,
                      color: THEME.muted,
                      lineHeight: 1.55,
                    }}
                  >
                    Province:{" "}
                    <span style={{ color: THEME.blue3 }}>Badakhshan</span>
                    <br />
                    Dataset:{" "}
                    <span style={{ color: THEME.blue3 }}>{dataset}</span>
                    <br />
                    Filters:{" "}
                    <span style={{ color: THEME.blue3 }}>
                      {year === "All" ? "All years" : year},{" "}
                      {month === "All" ? "All months" : month}
                    </span>
                    <br />
                    Key indicator periods:{" "}
                    <span style={{ color: THEME.blue3 }}>
                      {fmtInt(kpiSelected.periods)}
                    </span>
                    <br />
                    Overall periods:{" "}
                    <span style={{ color: THEME.blue3 }}>
                      {fmtInt(kpiTotal.periods)}
                    </span>
                  </div>
                </div>
              </div>
            </div>

            {/* RIGHT: More charts (2x2) */}
            <div style={styles.chartGrid2x2}>
              <div style={styles.chartBox}>
                <div style={styles.chartTitle}>Trend — Key indicator</div>
                {hasTrendSelected ? (
                  <ResponsiveContainer width="100%" height={190}>
                    <LineChart data={trendSelected}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="name" hide={true} />
                      <YAxis />
                      <Tooltip />
                      <Line
                        type="monotone"
                        dataKey="value"
                        stroke={THEME.blue}
                        strokeWidth={3}
                        dot={true}
                      />
                    </LineChart>
                  </ResponsiveContainer>
                ) : (
                  <div
                    style={{
                      fontWeight: 900,
                      color: THEME.muted,
                      fontSize: 12,
                      lineHeight: 1.5,
                    }}
                  >
                    Not enough periods for a trend (current filters).
                    <br />
                    Tip: set Month = <b>All</b>.
                  </div>
                )}
              </div>

              <div style={styles.chartBox}>
                <div style={styles.chartTitle}>
                  Trend — Overall (all indicators)
                </div>
                {hasTrendTotal ? (
                  <ResponsiveContainer width="100%" height={190}>
                    <LineChart data={trendTotal}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="name" hide={true} />
                      <YAxis />
                      <Tooltip />
                      <Line
                        type="monotone"
                        dataKey="value"
                        stroke={THEME.blue2}
                        strokeWidth={3}
                        dot={false}
                      />
                    </LineChart>
                  </ResponsiveContainer>
                ) : (
                  <div
                    style={{
                      fontWeight: 900,
                      color: THEME.muted,
                      fontSize: 12,
                      lineHeight: 1.5,
                    }}
                  >
                    Not enough periods for overall trend.
                  </div>
                )}
              </div>

              <div style={styles.chartBox}>
                <div style={styles.chartTitle}>Top indicators (colorful)</div>
                <ResponsiveContainer width="100%" height={190}>
                  <BarChart data={topIndicators}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" hide={true} />
                    <YAxis />
                    <Tooltip />
                    <Bar dataKey="value">{renderBarCells(topIndicators)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div style={styles.chartBox}>
                <div style={styles.chartTitle}>Share by indicator family</div>
                <ResponsiveContainer width="100%" height={190}>
                  <PieChart>
                    <Pie
                      data={familyShare}
                      dataKey="value"
                      nameKey="name"
                      innerRadius={55}
                      outerRadius={80}
                    >
                      {renderPieCells(familyShare)}
                    </Pie>
                    <Tooltip />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
          {/* end bottom */}
        </div>
      </div>
    </div>
  );
}
