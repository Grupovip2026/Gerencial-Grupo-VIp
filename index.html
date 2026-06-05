import { useState } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell } from "recharts";

// Convert Excel serial dates to month names
const months = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
const monthsFull = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

const companies = {
  VIP: {
    label: "VIP MX",
    color: "#00C9A7",
    accent: "#00856E",
    months: [
      { mes: "Jan", receitaPrevista: 266500, receitaRealizada: 435355, despesasPrevistas: 395348, despesasRealizadas: 439765, resultadoPrevisto: -128849, resultadoRealizado: -4410 },
      { mes: "Fev", receitaPrevista: 266500, receitaRealizada: 363194, despesasPrevistas: 343181, despesasRealizadas: 389392, resultadoPrevisto: -76681, resultadoRealizado: -26198 },
      { mes: "Mar", receitaPrevista: 263000, receitaRealizada: 446258, despesasPrevistas: 301861, despesasRealizadas: 464461, resultadoPrevisto: -38861, resultadoRealizado: -18203 },
      { mes: "Abr", receitaPrevista: 293150, receitaRealizada: 545913, despesasPrevistas: 228300, despesasRealizadas: 409522, resultadoPrevisto: 64850, resultadoRealizado: 136392 },
    ],
    despesas: [
      { categoria: "Matéria Prima", previsto: 263154, realizado: 252773 },
      { categoria: "Taxas Plataforma", previsto: 69290, realizado: 100406 },
      { categoria: "Cancelamentos", previsto: 13325, realizado: 35768 },
      { categoria: "Impostos", previsto: 29900, realizado: 33357 },
      { categoria: "Pessoal", previsto: 6228, realizado: 7654 },
      { categoria: "Financiamentos", previsto: 6602, realizado: 6602 },
    ],
    canais: [
      { canal: "Mercado Livre", previsto: 260000, realizado: 413272 },
      { canal: "Shopee", previsto: 3000, realizado: 11861 },
      { canal: "Amazon", previsto: 3500, realizado: 10223 },
    ],
  },
  VIDAL: {
    label: "VIDAL",
    color: "#6C63FF",
    accent: "#4B44CC",
    months: [
      { mes: "Jan", receitaPrevista: 207000, receitaRealizada: 196571, despesasPrevistas: 182854, despesasRealizadas: 71161, resultadoPrevisto: 24147, resultadoRealizado: 125409 },
      { mes: "Fev", receitaPrevista: 207000, receitaRealizada: 167144, despesasPrevistas: 182854, despesasRealizadas: 86782, resultadoPrevisto: 24147, resultadoRealizado: 80362 },
      { mes: "Mar", receitaPrevista: 207000, receitaRealizada: 229018, despesasPrevistas: 182854, despesasRealizadas: 101714, resultadoPrevisto: 24147, resultadoRealizado: 127305 },
      { mes: "Abr", receitaPrevista: 227700, receitaRealizada: 251048, despesasPrevistas: 194014, despesasRealizadas: 135847, resultadoPrevisto: 33686, resultadoRealizado: 115201 },
    ],
    despesas: [
      { categoria: "Matéria Prima", previsto: 90000, realizado: 0 },
      { categoria: "Taxas Plataforma", previsto: 40820, realizado: 22835 },
      { categoria: "Cancelamentos", previsto: 7850, realizado: 16053 },
      { categoria: "Pessoal", previsto: 22734, realizado: 17178 },
      { categoria: "Impostos", previsto: 13000, realizado: 8843 },
      { categoria: "Despesas Gerais", previsto: 8250, realizado: 6053 },
    ],
    canais: [
      { canal: "Mercado Livre", previsto: 155000, realizado: 138853 },
      { canal: "Loja Física", previsto: 50000, realizado: 44022 },
      { canal: "Shopee", previsto: 2000, realizado: 13696 },
    ],
  },
  V3: {
    label: "V3",
    color: "#FF6B6B",
    accent: "#CC4444",
    months: [
      { mes: "Jan", receitaPrevista: 7500, receitaRealizada: 2706, despesasPrevistas: 2250, despesasRealizadas: 578, resultadoPrevisto: 5250, resultadoRealizado: 2128 },
      { mes: "Fev", receitaPrevista: 7500, receitaRealizada: 3774, despesasPrevistas: 2250, despesasRealizadas: 1002, resultadoPrevisto: 5250, resultadoRealizado: 2772 },
      { mes: "Mar", receitaPrevista: 7500, receitaRealizada: 3913, despesasPrevistas: 2250, despesasRealizadas: 1206, resultadoPrevisto: 5250, resultadoRealizado: 2707 },
      { mes: "Abr", receitaPrevista: 3913, receitaRealizada: 3918, despesasPrevistas: 1174, despesasRealizadas: 1405, resultadoPrevisto: 2739, resultadoRealizado: 2514 },
    ],
    despesas: [
      { categoria: "Taxas Plataforma", previsto: 1875, realizado: 578 },
      { categoria: "Cancelamentos", previsto: 375, realizado: 0 },
    ],
    canais: [
      { canal: "Mercado Livre", previsto: 7500, realizado: 2706 },
    ],
  },
};

const fmt = (v) =>
  new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 }).format(v);

const pct = (real, prev) => {
  if (!prev) return "—";
  const d = ((real - prev) / Math.abs(prev)) * 100;
  return (d >= 0 ? "+" : "") + d.toFixed(1) + "%";
};

const KPI = ({ label, value, sub, color, positive }) => (
  <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 12, padding: "18px 22px", minWidth: 160, flex: 1 }}>
    <div style={{ fontSize: 11, color: "#888", textTransform: "uppercase", letterSpacing: 1.2, marginBottom: 6 }}>{label}</div>
    <div style={{ fontSize: 22, fontWeight: 700, color: positive ? "#00C9A7" : value < 0 ? "#FF6B6B" : "#f0f0f0", fontFamily: "'DM Mono', monospace" }}>{fmt(value)}</div>
    {sub != null && (
      <div style={{ fontSize: 12, color: sub >= 0 ? "#00C9A7" : "#FF6B6B", marginTop: 4 }}>
        {sub >= 0 ? "▲" : "▼"} {Math.abs(sub).toFixed(1)}% vs previsto
      </div>
    )}
  </div>
);

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#1a1a2e", border: "1px solid rgba(255,255,255,0.15)", borderRadius: 8, padding: "10px 14px", fontSize: 12 }}>
      <div style={{ color: "#aaa", marginBottom: 6, fontWeight: 600 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, marginBottom: 2 }}>{p.name}: {fmt(p.value)}</div>
      ))}
    </div>
  );
};

export default function Dashboard() {
  const [active, setActive] = useState("VIP");
  const co = companies[active];

  const totalReceitaReal = co.months.reduce((s, m) => s + m.receitaRealizada, 0);
  const totalReceitaPrev = co.months.reduce((s, m) => s + m.receitaPrevista, 0);
  const totalDespReal = co.months.reduce((s, m) => s + m.despesasRealizadas, 0);
  const totalDespPrev = co.months.reduce((s, m) => s + m.despesasPrevistas, 0);
  const totalResReal = co.months.reduce((s, m) => s + m.resultadoRealizado, 0);
  const totalResPrev = co.months.reduce((s, m) => s + m.resultadoPrevisto, 0);
  const receitaDelta = ((totalReceitaReal - totalReceitaPrev) / Math.abs(totalReceitaPrev)) * 100;
  const despDelta = ((totalDespReal - totalDespPrev) / Math.abs(totalDespPrev)) * 100;
  const resDelta = totalResPrev !== 0 ? ((totalResReal - totalResPrev) / Math.abs(totalResPrev)) * 100 : null;

  return (
    <div style={{ background: "#0d0d1a", minHeight: "100vh", color: "#f0f0f0", fontFamily: "'Inter', sans-serif", padding: "28px 24px" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 4px; } 
        ::-webkit-scrollbar-track { background: #111; }
        ::-webkit-scrollbar-thumb { background: #333; border-radius: 4px; }
        .tab-btn { transition: all 0.2s; cursor: pointer; }
        .tab-btn:hover { opacity: 0.85; }
      `}</style>

      {/* Header */}
      <div style={{ marginBottom: 28 }}>
        <div style={{ fontSize: 11, color: "#666", letterSpacing: 2, textTransform: "uppercase", marginBottom: 4 }}>Grupo VIP — Resultado Gerencial 2026</div>
        <h1 style={{ fontSize: 26, fontWeight: 700, margin: 0, color: "#fff" }}>Dashboard Financeiro</h1>
        <div style={{ fontSize: 13, color: "#555", marginTop: 4 }}>Janeiro – Abril 2026 · Previsto vs Realizado</div>
      </div>

      {/* Company Tabs */}
      <div style={{ display: "flex", gap: 8, marginBottom: 28, flexWrap: "wrap" }}>
        {Object.entries(companies).map(([key, c]) => (
          <button
            key={key}
            className="tab-btn"
            onClick={() => setActive(key)}
            style={{
              padding: "10px 22px",
              borderRadius: 8,
              border: active === key ? `2px solid ${c.color}` : "2px solid rgba(255,255,255,0.1)",
              background: active === key ? `${c.color}22` : "transparent",
              color: active === key ? c.color : "#888",
              fontWeight: active === key ? 700 : 400,
              fontSize: 14,
              letterSpacing: 0.5,
            }}
          >
            {c.label}
          </button>
        ))}
      </div>

      {/* KPI Cards */}
      <div style={{ display: "flex", gap: 12, marginBottom: 28, flexWrap: "wrap" }}>
        <KPI label="Receita Realizada" value={totalReceitaReal} sub={receitaDelta} positive />
        <KPI label="Despesas Realizadas" value={totalDespReal} sub={despDelta} />
        <KPI label="Resultado Realizado" value={totalResReal} sub={resDelta} />
        <KPI label="Receita Prevista" value={totalReceitaPrev} />
        <KPI label="Resultado Previsto" value={totalResPrev} />
      </div>

      {/* Charts Row 1 */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>

        {/* Receita Previsto vs Realizado */}
        <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "20px 16px" }}>
          <div style={{ fontSize: 12, color: "#aaa", marginBottom: 14, fontWeight: 600, letterSpacing: 0.8, textTransform: "uppercase" }}>Receita Bruta · Mês a Mês</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={co.months} barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="mes" tick={{ fill: "#666", fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#666", fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={v => `${(v/1000).toFixed(0)}k`} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#888" }} />
              <Bar dataKey="receitaPrevista" name="Previsto" fill="rgba(255,255,255,0.1)" radius={[4,4,0,0]} />
              <Bar dataKey="receitaRealizada" name="Realizado" fill={co.color} radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Resultado mensal */}
        <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "20px 16px" }}>
          <div style={{ fontSize: 12, color: "#aaa", marginBottom: 14, fontWeight: 600, letterSpacing: 0.8, textTransform: "uppercase" }}>Resultado Líquido · Mês a Mês</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={co.months} barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="mes" tick={{ fill: "#666", fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#666", fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={v => `${(v/1000).toFixed(0)}k`} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#888" }} />
              <Bar dataKey="resultadoPrevisto" name="Previsto" fill="rgba(255,255,255,0.1)" radius={[4,4,0,0]} />
              <Bar dataKey="resultadoRealizado" name="Realizado" radius={[4,4,0,0]}>
                {co.months.map((m, i) => (
                  <Cell key={i} fill={m.resultadoRealizado >= 0 ? co.color : "#FF6B6B"} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Charts Row 2 */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>

        {/* Despesas por categoria */}
        <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "20px 16px" }}>
          <div style={{ fontSize: 12, color: "#aaa", marginBottom: 14, fontWeight: 600, letterSpacing: 0.8, textTransform: "uppercase" }}>Despesas por Categoria · Jan</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={co.despesas} layout="vertical" barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" horizontal={false} />
              <XAxis type="number" tick={{ fill: "#666", fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={v => `${(v/1000).toFixed(0)}k`} />
              <YAxis type="category" dataKey="categoria" tick={{ fill: "#888", fontSize: 10 }} axisLine={false} tickLine={false} width={100} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#888" }} />
              <Bar dataKey="previsto" name="Previsto" fill="rgba(255,255,255,0.1)" radius={[0,4,4,0]} />
              <Bar dataKey="realizado" name="Realizado" fill={co.accent} radius={[0,4,4,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Canais de venda */}
        <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "20px 16px" }}>
          <div style={{ fontSize: 12, color: "#aaa", marginBottom: 14, fontWeight: 600, letterSpacing: 0.8, textTransform: "uppercase" }}>Canais de Venda · Jan</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={co.canais} barGap={4}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
              <XAxis dataKey="canal" tick={{ fill: "#666", fontSize: 10 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#666", fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={v => `${(v/1000).toFixed(0)}k`} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#888" }} />
              <Bar dataKey="previsto" name="Previsto" fill="rgba(255,255,255,0.1)" radius={[4,4,0,0]} />
              <Bar dataKey="realizado" name="Realizado" fill={co.color} radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Table */}
      <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.07)", borderRadius: 12, padding: "20px", overflowX: "auto" }}>
        <div style={{ fontSize: 12, color: "#aaa", marginBottom: 14, fontWeight: 600, letterSpacing: 0.8, textTransform: "uppercase" }}>Resumo Mensal</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
              {["Mês", "Receita Prevista", "Receita Realizada", "Δ Receita", "Despesas Prev.", "Despesas Real.", "Δ Despesas", "Resultado Prev.", "Resultado Real."].map(h => (
                <th key={h} style={{ padding: "8px 12px", textAlign: "right", color: "#666", fontWeight: 500, fontSize: 11, letterSpacing: 0.5, textTransform: "uppercase", whiteSpace: "nowrap" }}>
                  {h === "Mês" ? <span style={{ textAlign: "left", display: "block" }}>{h}</span> : h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {co.months.map((m, i) => {
              const rDelta = ((m.receitaRealizada - m.receitaPrevista) / Math.abs(m.receitaPrevista)) * 100;
              const dDelta = ((m.despesasRealizadas - m.despesasPrevistas) / Math.abs(m.despesasPrevistas)) * 100;
              return (
                <tr key={i} style={{ borderBottom: "1px solid rgba(255,255,255,0.04)" }}>
                  <td style={{ padding: "10px 12px", fontWeight: 600, color: "#ccc" }}>{m.mes}/26</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: "#777", fontFamily: "'DM Mono', monospace" }}>{fmt(m.receitaPrevista)}</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: "#f0f0f0", fontFamily: "'DM Mono', monospace", fontWeight: 600 }}>{fmt(m.receitaRealizada)}</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: rDelta >= 0 ? "#00C9A7" : "#FF6B6B", fontFamily: "'DM Mono', monospace" }}>{rDelta >= 0 ? "+" : ""}{rDelta.toFixed(1)}%</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: "#777", fontFamily: "'DM Mono', monospace" }}>{fmt(m.despesasPrevistas)}</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: "#f0f0f0", fontFamily: "'DM Mono', monospace", fontWeight: 600 }}>{fmt(m.despesasRealizadas)}</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: dDelta <= 0 ? "#00C9A7" : "#FF6B6B", fontFamily: "'DM Mono', monospace" }}>{dDelta >= 0 ? "+" : ""}{dDelta.toFixed(1)}%</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", color: "#777", fontFamily: "'DM Mono', monospace" }}>{fmt(m.resultadoPrevisto)}</td>
                  <td style={{ padding: "10px 12px", textAlign: "right", fontFamily: "'DM Mono', monospace", fontWeight: 700, color: m.resultadoRealizado >= 0 ? "#00C9A7" : "#FF6B6B" }}>{fmt(m.resultadoRealizado)}</td>
                </tr>
              );
            })}
            {/* Totals row */}
            <tr style={{ borderTop: "1px solid rgba(255,255,255,0.15)", background: "rgba(255,255,255,0.03)" }}>
              <td style={{ padding: "12px 12px", fontWeight: 700, color: "#fff" }}>TOTAL</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: "#777", fontFamily: "'DM Mono', monospace" }}>{fmt(totalReceitaPrev)}</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: "#f0f0f0", fontFamily: "'DM Mono', monospace", fontWeight: 700 }}>{fmt(totalReceitaReal)}</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: receitaDelta >= 0 ? "#00C9A7" : "#FF6B6B", fontFamily: "'DM Mono', monospace", fontWeight: 700 }}>{receitaDelta >= 0 ? "+" : ""}{receitaDelta.toFixed(1)}%</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: "#777", fontFamily: "'DM Mono', monospace" }}>{fmt(totalDespPrev)}</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: "#f0f0f0", fontFamily: "'DM Mono', monospace", fontWeight: 700 }}>{fmt(totalDespReal)}</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: despDelta <= 0 ? "#00C9A7" : "#FF6B6B", fontFamily: "'DM Mono', monospace", fontWeight: 700 }}>{despDelta >= 0 ? "+" : ""}{despDelta.toFixed(1)}%</td>
              <td style={{ padding: "12px 12px", textAlign: "right", color: "#777", fontFamily: "'DM Mono', monospace" }}>{fmt(totalResPrev)}</td>
              <td style={{ padding: "12px 12px", textAlign: "right", fontFamily: "'DM Mono', monospace", fontWeight: 700, color: totalResReal >= 0 ? "#00C9A7" : "#FF6B6B" }}>{fmt(totalResReal)}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div style={{ marginTop: 16, textAlign: "right", fontSize: 11, color: "#444" }}>
        Dados: Jan–Abr 2026 · Fonte: Resultado_Gerencial_2026.xlsx
      </div>
    </div>
  );
}
