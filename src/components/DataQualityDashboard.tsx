'use client';

import { useMemo, useState } from 'react';
import type { MappingStore } from '../utils/mappingStore';
import { isMapped } from '../utils/mappingStore';
import { OUTPUT_COLS } from '../utils/converter';

// ── Constants ─────────────────────────────────────────────

const FIXED_COLS = new Set([
  'Code', 'Actief vanaf', 'Inkoop', 'Ordergestuurd', 'Verkoop', 'Voorraad',
  'Btw-code: Verkoop', 'Eenheidsfactor', 'KMS Synchronisatie', '2026 Controle JW',
  'Hoofdleverancier', 'Btw-code: Inkoop',
]);
const OPTIONAL_COLS = new Set(['Merk', 'Artikelcode', 'Veiligheidsclassificatie']);

const DUPE_PREVIEW_COLS = ['Code', 'Omschrijving', 'Barcode', 'Kostprijs', 'Verkoopprijs', 'Productsoort'];
const PS_PREVIEW_COLS   = ['Code', 'Omschrijving', 'Productsoort', 'Artikelgroep: Omschrijving', 'Eenheid'];

const MINI_PAGE_SIZE = 10;

// Columns shown when previewing rows missing a given column.
// Always includes Code + Omschrijving + Productsoort as context, plus the column itself.
function colPreviewCols(col: string): string[] {
  return [...new Set(['Code', 'Omschrijving', 'Productsoort', col])];
}

// ── Types ─────────────────────────────────────────────────

interface Props {
  data: Record<string, unknown>[];
  mappingStore: MappingStore;
  lang: 'nl' | 'en';
}

interface ColStat {
  col: string;
  filled: number;
  empty: number;
  pct: number;
  optional: boolean;
}

type SuspiciousType =
  | 'kostprijs_zero'
  | 'verkoopprijs_zero'
  | 'invalid_barcode';

interface SuspiciousGroup {
  type: SuspiciousType;
  rows: Record<string, unknown>[];
}

interface QualityReport {
  colStats: ColStat[];
  overallFillRate: number;
  rowsWithEmpty: number;
  duplicateBarcodes: [string, number][];
  unknownProductsoort: string[];
  truncatedCount: number;
  suspiciousGroups: SuspiciousGroup[];
  totalSuspicious: number;
  codeMin: number | null;
  codeMax: number | null;
  issueCount: number;
}

const SUSP_PREVIEW_COLS: Record<SuspiciousType, string[]> = {
  kostprijs_zero:    ['Code', 'Omschrijving', 'Barcode', 'Kostprijs', 'Verkoopprijs'],
  verkoopprijs_zero: ['Code', 'Omschrijving', 'Barcode', 'Kostprijs', 'Verkoopprijs'],
  invalid_barcode:   ['Code', 'Omschrijving', 'Barcode', 'Productsoort'],
};

const SUSP_LABELS: Record<SuspiciousType, [string, string]> = {
  kostprijs_zero:    ['Kostprijs leeg of nul',        'Kostprijs empty or zero'],
  verkoopprijs_zero: ['Verkoopprijs leeg of nul',     'Verkoopprijs empty or zero'],
  invalid_barcode:   ['Barcode geen 13-cijferig EAN', 'Barcode not a 13-digit EAN'],
};

// ── Sub-components ────────────────────────────────────────

function FillBar({ pct }: { pct: number }) {
  const color = pct === 100 ? 'var(--green-dark)' : pct >= 80 ? '#BA7517' : 'var(--red-text)';
  return (
    <div style={{ height: 6, width: 100, background: 'var(--border)', borderRadius: 3, overflow: 'hidden', flexShrink: 0 }}>
      <div style={{ height: '100%', width: `${pct}%`, background: color, borderRadius: 3 }} />
    </div>
  );
}

function StatChip({ label, value, color }: { label: string; value: string; color: string }) {
  return (
    <div style={{
      display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2,
      padding: '8px 16px', background: 'var(--bg-secondary)', border: '0.5px solid var(--border)',
      borderRadius: 'var(--radius-md)', minWidth: 90,
    }}>
      <span style={{ fontSize: 20, fontWeight: 700, color, lineHeight: 1.1 }}>{value}</span>
      <span style={{ fontSize: 10, color: 'var(--text-secondary)', textAlign: 'center', whiteSpace: 'nowrap' }}>{label}</span>
    </div>
  );
}

function SectionHeading({ title, warn }: { title: string; warn?: boolean }) {
  return (
    <div style={{
      fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.07em',
      color: warn ? '#BA7517' : 'var(--text-secondary)', marginBottom: 6, marginTop: 14,
    }}>
      {title}
    </div>
  );
}

// Paginated table — never grows beyond MINI_PAGE_SIZE rows.
// Pass a stable `key` at the call site to reset the page when content changes.
function MiniTable({ cols, rows, lang }: { cols: string[]; rows: Record<string, unknown>[]; lang: 'nl' | 'en' }) {
  const [page, setPage] = useState(0);
  if (!rows.length) return null;

  const totalPages = Math.max(1, Math.ceil(rows.length / MINI_PAGE_SIZE));
  const safePage   = Math.min(page, totalPages - 1);
  const pageRows   = rows.slice(safePage * MINI_PAGE_SIZE, (safePage + 1) * MINI_PAGE_SIZE);
  const nl = lang === 'nl';

  return (
    <div style={{ marginTop: 8 }}>
      <div style={{ overflowX: 'auto', borderRadius: 4, border: '0.5px solid var(--border)' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
          <thead>
            <tr style={{ background: 'var(--bg-secondary)' }}>
              {cols.map(c => (
                <th key={c} style={{ padding: '5px 10px', textAlign: 'left', fontWeight: 600, color: 'var(--text-secondary)', borderBottom: '0.5px solid var(--border)', whiteSpace: 'nowrap' }}>
                  {c}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((row, i) => (
              <tr key={i} style={{ background: i % 2 === 1 ? 'var(--bg-secondary)' : undefined }}>
                {cols.map(c => (
                  <td key={c} style={{ padding: '4px 10px', color: 'var(--text)', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', borderBottom: i < pageRows.length - 1 ? '0.5px solid var(--border)' : 'none' }}>
                    {String(row[c] ?? '')}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {totalPages > 1 && (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, marginTop: 6 }}>
          <button className="btn btn-sm" style={{ padding: '1px 8px' }} disabled={safePage === 0} onClick={() => setPage(p => p - 1)}>‹</button>
          <span style={{ fontSize: 11, color: 'var(--text-secondary)' }}>
            {safePage + 1} / {totalPages} · {rows.length} {nl ? 'rijen' : 'rows'}
          </span>
          <button className="btn btn-sm" style={{ padding: '1px 8px' }} disabled={safePage >= totalPages - 1} onClick={() => setPage(p => p + 1)}>›</button>
        </div>
      )}
    </div>
  );
}

// Scrollable badge list — caps height so many badges never push content down
function BadgeList({ children }: { children: React.ReactNode }) {
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, maxHeight: 96, overflowY: 'auto', padding: '2px 0' }}>
      {children}
    </div>
  );
}

// ── Data logic ────────────────────────────────────────────

function computeReport(data: Record<string, unknown>[], mappingStore: MappingStore): QualityReport {
  const measuredCols = OUTPUT_COLS.filter(c => !FIXED_COLS.has(c));

  const colStats: ColStat[] = measuredCols.map(col => {
    let filled = 0;
    for (const row of data) if (String(row[col] ?? '').trim()) filled++;
    const empty = data.length - filled;
    return { col, filled, empty, pct: data.length > 0 ? Math.round((filled / data.length) * 100) : 100, optional: OPTIONAL_COLS.has(col) };
  });

  const sortedColStats = [...colStats].sort((a, b) => {
    const ap = !a.optional && a.pct < 100, bp = !b.optional && b.pct < 100;
    if (ap && !bp) return -1;
    if (!ap && bp) return 1;
    return a.pct - b.pct;
  });

  const requiredStats    = colStats.filter(s => !s.optional);
  const totalCells       = requiredStats.length * data.length;
  const filledCells      = requiredStats.reduce((s, c) => s + c.filled, 0);
  const overallFillRate  = totalCells > 0 ? Math.round((filledCells / totalCells) * 100) : 100;
  const requiredColNames = requiredStats.map(s => s.col);

  let rowsWithEmpty = 0;
  for (const row of data) if (requiredColNames.some(col => !String(row[col] ?? '').trim())) rowsWithEmpty++;

  const bcCounts = new Map<string, number>();
  for (const row of data) { const bc = String(row['Barcode'] ?? '').trim(); if (bc) bcCounts.set(bc, (bcCounts.get(bc) ?? 0) + 1); }
  const duplicateBarcodes = [...bcCounts.entries()].filter(([, n]) => n > 1);

  const uniquePS = [...new Set(data.map(r => String(r['Productsoort'] ?? '').trim()).filter(Boolean))];
  const unknownProductsoort = uniquePS.filter(ps => !isMapped(ps, mappingStore));

  let truncatedCount = 0;
  for (const row of data) {
    if (String(row['Omschrijving'] ?? '').length === 60 && String(row['Extra omschrijving'] ?? '').length > 60) truncatedCount++;
  }

  const groupMap = new Map<SuspiciousType, Record<string, unknown>[]>();
  const addGroup = (type: SuspiciousType, row: Record<string, unknown>) => {
    if (!groupMap.has(type)) groupMap.set(type, []);
    groupMap.get(type)!.push(row);
  };
  for (const row of data) {
    const kp = String(row['Kostprijs']    ?? '').trim();
    const vp = String(row['Verkoopprijs'] ?? '').trim();
    const bc = String(row['Barcode']      ?? '').trim();
    if (!kp || kp === '0')          addGroup('kostprijs_zero',    row);
    if (!vp || vp === '0')          addGroup('verkoopprijs_zero', row);
    if (bc && !/^\d{13}$/.test(bc)) addGroup('invalid_barcode',   row);
  }
  const suspiciousGroups: SuspiciousGroup[] = [...groupMap.entries()].map(([type, rows]) => ({ type, rows }));
  const totalSuspicious = suspiciousGroups.reduce((s, g) => s + g.rows.length, 0);

  const codes   = data.map(r => Number(r['Code'])).filter(n => !isNaN(n));
  const codeMin = codes.length ? Math.min(...codes) : null;
  const codeMax = codes.length ? Math.max(...codes) : null;

  const issueCount =
    colStats.filter(s => !s.optional && s.pct < 100).length +
    (duplicateBarcodes.length > 0 ? 1 : 0) +
    unknownProductsoort.length +
    (suspiciousGroups.length > 0 ? 1 : 0);

  return { colStats: sortedColStats, overallFillRate, rowsWithEmpty, duplicateBarcodes, unknownProductsoort, truncatedCount, suspiciousGroups, totalSuspicious, codeMin, codeMax, issueCount };
}

// ── Main component ────────────────────────────────────────

export default function DataQualityDashboard({ data, mappingStore, lang }: Props) {
  const [open,              setOpen]              = useState(true);
  const [expandedCol,       setExpandedCol]       = useState<string | null>(null);
  const [expandedBarcode,   setExpandedBarcode]   = useState<string | null>(null);
  const [expandedPS,        setExpandedPS]        = useState<string | null>(null);
  const [expandedSuspGroup, setExpandedSuspGroup] = useState<SuspiciousType | null>(null);
  const nl = lang === 'nl';

  const report = useMemo(() => data.length ? computeReport(data, mappingStore) : null, [data, mappingStore]);
  if (!report) return null;

  const { colStats, overallFillRate, rowsWithEmpty, duplicateBarcodes, unknownProductsoort, truncatedCount, suspiciousGroups, totalSuspicious, codeMin, codeMax, issueCount } = report;

  const allGood     = issueCount === 0;
  const headerColor = allGood ? 'var(--green-dark)' : issueCount >= 5 ? 'var(--red-text)' : '#BA7517';
  const summaryText = allGood
    ? (nl ? 'Alles ziet er goed uit' : 'Everything looks good')
    : `${issueCount} ${nl ? `aandachtspunt${issueCount === 1 ? '' : 'en'}` : `issue${issueCount === 1 ? '' : 's'}`}`;

  const incompleteCols = colStats.filter(s => s.pct < 100);
  const dupeRows = expandedBarcode ? data.filter(r => String(r['Barcode']      ?? '').trim() === expandedBarcode) : [];
  const psRows   = expandedPS      ? data.filter(r => String(r['Productsoort'] ?? '').trim() === expandedPS)      : [];
  const colRows  = expandedCol     ? data.filter(r => !String(r[expandedCol]   ?? '').trim())                     : [];

  return (
    <div className="card" style={{ padding: 0, overflow: 'hidden' }}>

      {/* ── Collapsible header ── */}
      <button
        onClick={() => setOpen(o => !o)}
        style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, padding: '12px 16px', background: 'none', border: 'none', cursor: 'pointer', borderBottom: open ? '0.5px solid var(--border)' : 'none', textAlign: 'left' }}
      >
        <span style={{ fontSize: 14, fontWeight: 700, color: headerColor, flexShrink: 0 }}>{allGood ? '✓' : '⚠'}</span>
        <span style={{ flex: 1, fontSize: 13, fontWeight: 600, color: 'var(--text)' }}>{nl ? 'Datakwaliteit' : 'Data Quality'}</span>
        <span style={{ fontSize: 12, color: headerColor, fontWeight: 500 }}>{summaryText}</span>
        <span style={{ fontSize: 10, color: 'var(--text-secondary)', marginLeft: 6 }}>{open ? '▲' : '▼'}</span>
      </button>

      {open && (
        <div style={{ padding: '14px 16px 18px' }}>

          {/* ── Stat chips ── */}
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 6 }}>
            <StatChip label={nl ? 'Totaal rijen' : 'Total rows'} value={String(data.length)} color="var(--text)" />
            <StatChip
              label={nl ? 'Vulgraad' : 'Fill rate'} value={`${overallFillRate}%`}
              color={overallFillRate === 100 ? 'var(--green-dark)' : overallFillRate >= 80 ? '#BA7517' : 'var(--red-text)'}
            />
            <StatChip
              label={nl ? 'Rijen incompleet' : 'Rows incomplete'} value={String(rowsWithEmpty)}
              color={rowsWithEmpty === 0 ? 'var(--green-dark)' : '#BA7517'}
            />
            {duplicateBarcodes.length > 0 && (
              <StatChip label={nl ? 'Dubbele barcodes' : 'Duplicate barcodes'} value={String(duplicateBarcodes.length)} color="var(--red-text)" />
            )}
            {codeMin !== null && (
              <StatChip label={nl ? 'Codebereik' : 'Code range'} value={`${codeMin}–${codeMax}`} color="var(--text)" />
            )}
          </div>

          {/* ── Column completeness ── */}
          <SectionHeading title={nl ? 'Kolomvulgraad' : 'Column completeness'} />
          {incompleteCols.length === 0 ? (
            <p style={{ fontSize: 11, color: 'var(--green-dark)', marginBottom: 4 }}>
              {nl ? 'Alle kolommen 100% gevuld.' : 'All columns 100% filled.'}
            </p>
          ) : (
            <div style={{ border: '0.5px solid var(--border)', borderRadius: 'var(--radius-md)', overflow: 'hidden' }}>
              {incompleteCols.map((s, si) => {
                const isExpanded = expandedCol === s.col;
                const isLast     = si === incompleteCols.length - 1;
                return (
                  <div key={s.col} style={{ borderBottom: isLast ? 'none' : '0.5px solid var(--border)' }}>
                    <button
                      onClick={() => setExpandedCol(prev => prev === s.col ? null : s.col)}
                      style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 8, padding: '7px 10px', background: isExpanded ? 'var(--bg-secondary)' : 'none', border: 'none', cursor: 'pointer', textAlign: 'left', borderBottom: isExpanded ? '0.5px solid var(--border)' : 'none' }}
                    >
                      <span style={{ flex: 1, fontSize: 12, color: 'var(--text)', minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {s.col}
                        {s.optional && <span style={{ fontSize: 10, color: 'var(--text-secondary)', marginLeft: 5 }}>(opt)</span>}
                      </span>
                      <FillBar pct={s.pct} />
                      <span style={{ fontSize: 11, fontWeight: 600, width: 34, textAlign: 'right', flexShrink: 0, color: s.pct === 0 ? (s.optional ? 'var(--text-secondary)' : 'var(--red-text)') : '#BA7517' }}>
                        {s.pct}%
                      </span>
                      <span style={{ fontSize: 10, color: 'var(--text-secondary)', width: 68, textAlign: 'right', flexShrink: 0 }}>
                        {s.empty} {nl ? 'leeg' : 'empty'}
                      </span>
                      <span style={{ fontSize: 10, color: 'var(--text-secondary)', flexShrink: 0, marginLeft: 2 }}>
                        {isExpanded ? '▲' : '▼'}
                      </span>
                    </button>
                    {isExpanded && (
                      <div style={{ padding: '0 10px 10px' }}>
                        <MiniTable key={s.col} cols={colPreviewCols(s.col)} rows={colRows} lang={lang} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* ── Duplicate barcodes ── */}
          {duplicateBarcodes.length > 0 && (
            <>
              <SectionHeading title={nl ? `Dubbele barcodes (${duplicateBarcodes.length})` : `Duplicate barcodes (${duplicateBarcodes.length})`} warn />
              <p style={{ fontSize: 11, color: 'var(--text-secondary)', marginBottom: 6 }}>
                {nl
                  ? 'Dezelfde barcode in meerdere rijen — klik een badge om de betrokken producten te zien.'
                  : 'The same barcode appears in multiple rows — click a badge to see the affected products.'}
              </p>
              <BadgeList>
                {duplicateBarcodes.map(([bc, count]) => (
                  <button
                    key={bc}
                    onClick={() => setExpandedBarcode(prev => prev === bc ? null : bc)}
                    className="badge badge-amber"
                    style={{ fontSize: 11, cursor: 'pointer', border: 'none', fontFamily: 'inherit', outline: expandedBarcode === bc ? '2px solid #BA7517' : 'none', background: expandedBarcode === bc ? '#BA751722' : undefined }}
                  >
                    {bc} ×{count} {expandedBarcode === bc ? '▲' : '▼'}
                  </button>
                ))}
              </BadgeList>
              {expandedBarcode && dupeRows.length > 0 && (
                <MiniTable key={expandedBarcode} cols={DUPE_PREVIEW_COLS} rows={dupeRows} lang={lang} />
              )}
            </>
          )}

          {/* ── Unknown productsoort ── */}
          {unknownProductsoort.length > 0 && (
            <>
              <SectionHeading title={nl ? `Onbekende productsoort (${unknownProductsoort.length})` : `Unknown productsoort (${unknownProductsoort.length})`} warn />
              <p style={{ fontSize: 11, color: 'var(--text-secondary)', marginBottom: 6 }}>
                {nl
                  ? 'Geen koppeling gevonden — klik een badge om de betrokken rijen te zien. Voeg ze toe via "Productsoort koppelingen" in stap 2.'
                  : 'No mapping found — click a badge to see the affected rows. Add them via "Productsoort mappings" in step 2.'}
              </p>
              <BadgeList>
                {unknownProductsoort.map(ps => (
                  <button
                    key={ps}
                    onClick={() => setExpandedPS(prev => prev === ps ? null : ps)}
                    className="badge badge-amber"
                    style={{ fontSize: 11, cursor: 'pointer', border: 'none', fontFamily: 'inherit', outline: expandedPS === ps ? '2px solid #BA7517' : 'none', background: expandedPS === ps ? '#BA751722' : undefined }}
                  >
                    {ps} {expandedPS === ps ? '▲' : '▼'}
                  </button>
                ))}
              </BadgeList>
              {expandedPS && psRows.length > 0 && (
                <MiniTable key={expandedPS} cols={PS_PREVIEW_COLS} rows={psRows} lang={lang} />
              )}
            </>
          )}

          {/* ── Omschrijving truncation ── */}
          {truncatedCount > 0 && (
            <>
              <SectionHeading title={nl ? `Omschrijving ingekort (${truncatedCount} rijen)` : `Omschrijving truncated (${truncatedCount} rows)`} />
              <p style={{ fontSize: 11, color: 'var(--text-secondary)' }}>
                {nl
                  ? `${truncatedCount} omschrijvingen zijn afgekapt op 60 tekens. De volledige tekst is bewaard in 'Extra omschrijving'.`
                  : `${truncatedCount} descriptions were cut at 60 characters. The full text is preserved in 'Extra omschrijving'.`}
              </p>
            </>
          )}

          {/* ── Suspicious values — grouped accordion ── */}
          {suspiciousGroups.length > 0 && (
            <>
              <SectionHeading
                title={nl
                  ? `Verdachte waarden — ${totalSuspicious} gevallen in ${suspiciousGroups.length} categorie${suspiciousGroups.length === 1 ? '' : 'ën'}`
                  : `Suspicious values — ${totalSuspicious} occurrence${totalSuspicious === 1 ? '' : 's'} in ${suspiciousGroups.length} categor${suspiciousGroups.length === 1 ? 'y' : 'ies'}`}
                warn
              />
              <div style={{ border: '0.5px solid var(--border)', borderRadius: 'var(--radius-md)', overflow: 'hidden' }}>
                {suspiciousGroups.map((group, gi) => {
                  const isExpanded  = expandedSuspGroup === group.type;
                  const label       = SUSP_LABELS[group.type][nl ? 0 : 1];
                  const previewCols = SUSP_PREVIEW_COLS[group.type];
                  const isLast      = gi === suspiciousGroups.length - 1;
                  return (
                    <div key={group.type} style={{ borderBottom: isLast ? 'none' : '0.5px solid var(--border)' }}>
                      <button
                        onClick={() => setExpandedSuspGroup(prev => prev === group.type ? null : group.type)}
                        style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', background: isExpanded ? 'var(--bg-secondary)' : 'none', border: 'none', cursor: 'pointer', textAlign: 'left', borderBottom: isExpanded ? '0.5px solid var(--border)' : 'none' }}
                      >
                        <span style={{ fontSize: 11, color: '#BA7517', flexShrink: 0 }}>⚠</span>
                        <span style={{ flex: 1, fontSize: 12, color: 'var(--text)', fontWeight: 500 }}>{label}</span>
                        <span className="badge badge-amber" style={{ fontSize: 10, flexShrink: 0 }}>
                          {group.rows.length} {nl ? 'rijen' : 'rows'}
                        </span>
                        <span style={{ fontSize: 10, color: 'var(--text-secondary)', marginLeft: 4, flexShrink: 0 }}>
                          {isExpanded ? '▲' : '▼'}
                        </span>
                      </button>
                      {isExpanded && (
                        <div style={{ padding: '0 12px 12px' }}>
                          <MiniTable key={group.type} cols={previewCols} rows={group.rows} lang={lang} />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </>
          )}

          {/* ── All good ── */}
          {allGood && (
            <p style={{ fontSize: 12, color: 'var(--green-dark)', textAlign: 'center', marginTop: 16, fontWeight: 500 }}>
              {nl ? 'Geen problemen gevonden — u kunt veilig downloaden.' : 'No issues found — you can safely download.'}
            </p>
          )}

        </div>
      )}
    </div>
  );
}
