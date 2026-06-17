'use client';

import { useState, useRef, useMemo, useCallback, memo, useEffect } from 'react';
import Stepper from '../components/Stepper';
import FileUploadCard from '../components/FileUploadCard';
import PreviewTable from '../components/PreviewTable';
import Banner from '../components/Banner';
import SheetPicker from '../components/SheetPicker';
import Tooltip from '../components/Tooltip';
import { useLang } from '../context/LangContext';
import { findBestCol, EAN_HINTS, CODE_HINTS } from '../utils/columns';
import { buildOutputRowsAsync, OUTPUT_COLS } from '../utils/converter';
import { sheetTo2D, downloadXLSX } from '../utils/xlsx';
import { loadMappingStore, isMapped, resolveArtikelgroep, resolveEenheid, type MappingStore } from '../utils/mappingStore';
import MappingEditor from '../components/MappingEditor';
import ColumnPreview from '../components/ColumnPreview';
import { detectHeaderRow, parseFromHeaderRow } from '../utils/headers';
import PresetBar from '../components/PresetBar';
import DataQualityDashboard from '../components/DataQualityDashboard';
import type { ParsedFile, BannerInfo, ColRef, Preset, SupplierConverterMappings } from '../types';

const PRODUCTSOORT_HINTS  = [
  'productsoort', 'soort', 'type', 'categorie', 'category',
  'producttype', 'artikeltype', 'itemtype', 'productgroep',
  'artikelsoort', 'klasse', 'classificatie', 'classification',
  'segment', 'division', 'productcategory', 'productcategorie',
];
const KOSTPRIJS_HINTS     = ['kostprijs', 'inkoopprijs', 'cost', 'inkoop', 'purchase'];
const BESTELNR_HINTS      = ['bestelnummer', 'bestel', 'ordernum', 'ordernr', 'leveranciersnr', 'artikelnummer', 'refnr'];
const VERKOOPPRIJS_HINTS  = ['verkoopprijs', 'verkoop', 'salesprice', 'selling', 'adviesprijs'];
const VEILIGHEID_HINTS    = ['veiligheid', 'safety', 'classif', 'classificatie', 'veiligheidscode'];
const GESLACHT_HINTS      = ['geslacht', 'gender', 'sex'];
const MAAT_HINTS          = ['maat', 'size', 'maten'];
const MERK_HINTS          = ['merk', 'brand', 'merknaam'];
const KLEUR_HINTS         = ['kleur', 'color', 'colour'];
const OMSCHR_HINTS        = ['omschrijving', 'omschr', 'description', 'descri', 'naam', 'name'];
const PRODUCTNAAM_HINTS   = ['productnaam', 'productname', 'titel', 'title'];

const KLEUR_PAGE_SIZE = 50;

const INPUT_STYLE: React.CSSProperties = {
  width: '100%', padding: '7px 10px',
  border: '0.5px solid var(--border-md)',
  borderRadius: 'var(--radius-md)' as string,
  background: 'var(--bg)', color: 'var(--text)',
  fontSize: 13, fontFamily: 'inherit',
};

// ── KleurRow — memoized so only the affected row re-renders ──────────────────
//
// Local `localValue` state decouples keystrokes from the parent kleurMapping
// update: the parent is only updated on blur or Enter, not on every keystroke.
// Combined with memo + stable useCallback handlers, this means typing in one
// color input does NOT re-render the other 499 rows.

interface KleurRowProps {
  original: string;
  mapped: string;
  isSelected: boolean;
  isModified: boolean;
  onSelect: (original: string, checked: boolean) => void;
  onValueChange: (original: string, value: string) => void;
  onReset: (original: string) => void;
}

const KleurRow = memo(function KleurRow({
  original, mapped, isSelected, isModified, onSelect, onValueChange, onReset,
}: KleurRowProps) {
  const [localValue, setLocalValue] = useState(mapped);

  // Sync when bulk rename or reset changes the committed value externally
  useEffect(() => { setLocalValue(mapped); }, [mapped]);

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <input
        type="checkbox"
        checked={isSelected}
        onChange={e => onSelect(original, e.target.checked)}
        style={{ flexShrink: 0 }}
      />
      <span style={{ flex: '0 0 200px', fontSize: 12, color: 'var(--text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
        {original}
      </span>
      <span style={{ fontSize: 12, color: 'var(--text-secondary)', width: 16, textAlign: 'center', flexShrink: 0 }}>→</span>
      <input
        type="text"
        value={localValue}
        onChange={e => setLocalValue(e.target.value)}
        onBlur={() => { if (localValue !== mapped) onValueChange(original, localValue); }}
        onKeyDown={e => { if (e.key === 'Enter') { onValueChange(original, localValue); (e.target as HTMLInputElement).blur(); } }}
        style={{
          flex: 1, fontSize: 12, padding: '3px 8px',
          border: `0.5px solid ${isModified ? 'var(--green-dark)' : 'var(--border-md)'}`,
          borderRadius: 4, background: 'var(--bg)', color: 'var(--text)', fontFamily: 'inherit',
        }}
      />
      {isModified && (
        <button
          className="btn btn-sm"
          style={{ padding: '1px 7px', fontSize: 11, flexShrink: 0 }}
          title="Reset"
          onClick={() => onReset(original)}
        >
          ↺
        </button>
      )}
    </div>
  );
});

// ── Multi-column combiner ────────────────────────────────

interface MultiColCombinerProps {
  sheetNames: string[];
  primarySheet: string;
  getSheetCols: (sheet: string) => string[];
  selected: ColRef[];
  onChange: (v: ColRef[]) => void;
  label: string;
  hint: string;
  required?: boolean;
  tooltip?: string;
  hasError?: boolean;
}

const COMBINE_PAGE_SIZE = 5;

function MultiColCombiner({ sheetNames, primarySheet, getSheetCols, selected, onChange, label, hint, required, tooltip, hasError }: MultiColCombinerProps) {
  const [activeSheet, setActiveSheet] = useState(primarySheet);
  const [search, setSearch] = useState('');
  const [combinePage, setCombinePage] = useState(0);
  const cols = getSheetCols(activeSheet);
  const isMultiSheet = sheetNames.length > 1;

  const filteredCols = search.trim()
    ? cols.filter(c => c.toLowerCase().includes(search.toLowerCase()))
    : cols;

  function isChecked(col: string) { return selected.some(r => r.sheet === activeSheet && r.col === col); }
  function toggle(col: string) {
    if (isChecked(col)) onChange(selected.filter(r => !(r.sheet === activeSheet && r.col === col)));
    else onChange([...selected, { sheet: activeSheet, col }]);
  }
  function move(idx: number, dir: -1 | 1) {
    const next = [...selected];
    [next[idx], next[idx + dir]] = [next[idx + dir], next[idx]];
    onChange(next);
  }

  return (
    <div>
      <label className="field-label">{label}{required && ' *'}{tooltip && <Tooltip text={tooltip} />}</label>
      {hint && <p className="field-hint" style={{ marginBottom: 6 }}>{hint}</p>}
      {isMultiSheet && <SheetPicker sheetNames={sheetNames} value={activeSheet} onChange={s => { setActiveSheet(s); setSearch(''); }} />}
      <div style={{ display: 'flex', gap: 6, marginBottom: 4, alignItems: 'center' }}>
        <input
          type="text"
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search columns…"
          style={{ flex: 1, fontSize: 12, padding: '4px 8px', border: '0.5px solid var(--border-md)', borderRadius: 4, background: 'var(--bg)', color: 'var(--text)', fontFamily: 'inherit' }}
        />
        {filteredCols.length > 0 && (
          <label style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11, color: 'var(--text-secondary)', cursor: 'pointer', flexShrink: 0, userSelect: 'none' }}>
            <input
              type="checkbox"
              checked={filteredCols.every(c => isChecked(c))}
              onChange={() => {
                const allChecked = filteredCols.every(c => isChecked(c));
                if (allChecked) {
                  onChange(selected.filter(r => !(r.sheet === activeSheet && filteredCols.includes(r.col))));
                } else {
                  const toAdd = filteredCols.filter(c => !isChecked(c)).map(c => ({ sheet: activeSheet, col: c }));
                  onChange([...selected, ...toAdd]);
                }
              }}
            />
            All
          </label>
        )}
      </div>
      <div className="ean-check-list" style={hasError && selected.length === 0 ? { border: '1.5px solid var(--red-text)', borderRadius: 'var(--radius-md)' } : undefined}>
        {filteredCols.length === 0
          ? <span style={{ fontSize: 11, color: 'var(--text-secondary)', padding: '4px 0' }}>No columns match "{search}"</span>
          : filteredCols.map(col => (
          <label key={`${activeSheet}:${col}`}>
            <input type="checkbox" checked={isChecked(col)} onChange={() => { toggle(col); }} />
            <span>{col}</span>
          </label>
        ))}
      </div>
      {selected.length > 0 && (() => {
        const totalPages = Math.ceil(selected.length / COMBINE_PAGE_SIZE);
        const safePage = Math.min(combinePage, totalPages - 1);
        const pageStart = safePage * COMBINE_PAGE_SIZE;
        const pageItems = selected.slice(pageStart, pageStart + COMBINE_PAGE_SIZE);
        return (
          <div style={{ marginTop: 8 }}>
            <div style={{ fontSize: 11, color: 'var(--text-secondary)', marginBottom: 4 }}>
              Combine order ({selected.length}):
            </div>
            {pageItems.map((ref, pageIdx) => {
              const idx = pageStart + pageIdx;
              return (
                <div key={`${ref.sheet}:${ref.col}:${idx}`} style={{ display: 'flex', alignItems: 'center', gap: 4, marginBottom: 3 }}>
                  <span style={{ flex: 1, fontSize: 11, background: 'var(--bg-secondary)', borderRadius: 4, padding: '3px 8px', border: '0.5px solid var(--border)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {idx + 1}. {isMultiSheet ? `[${ref.sheet}] ` : ''}{ref.col}
                  </span>
                  <button className="btn btn-sm" style={{ padding: '1px 7px' }} onClick={() => move(idx, -1)} disabled={idx === 0}>↑</button>
                  <button className="btn btn-sm" style={{ padding: '1px 7px' }} onClick={() => move(idx, 1)} disabled={idx === selected.length - 1}>↓</button>
                  <button className="btn btn-sm" style={{ padding: '1px 7px', color: 'var(--red-text)' }} onClick={() => { onChange(selected.filter((_, i) => i !== idx)); setCombinePage(p => Math.min(p, Math.ceil((selected.length - 1) / COMBINE_PAGE_SIZE) - 1)); }}>×</button>
                </div>
              );
            })}
            {totalPages > 1 && (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, marginTop: 4 }}>
                <button className="btn btn-sm" style={{ padding: '1px 7px' }} disabled={safePage === 0} onClick={() => setCombinePage(p => p - 1)}>‹</button>
                <span style={{ fontSize: 11, color: 'var(--text-secondary)' }}>{safePage + 1} / {totalPages}</span>
                <button className="btn btn-sm" style={{ padding: '1px 7px' }} disabled={safePage >= totalPages - 1} onClick={() => setCombinePage(p => p + 1)}>›</button>
              </div>
            )}
          </div>
        );
      })()}
    </div>
  );
}

// ── Main component ───────────────────────────────────────

const EMPTY_REF: ColRef = { sheet: '', col: '' };

export default function SupplierConverter() {
  const { lang, t } = useLang();

  const [step, setStep] = useState(1);
  const [supplFile, setSupplFile] = useState<ParsedFile | null>(null);
  const [banner, setBanner] = useState<BannerInfo | null>(null);

  const [hoofdleverancier, setHoofdleverancier] = useState('');
  const [btwInkoop, setBtwInkoop] = useState('3');
  const [startingCode, setStartingCode] = useState('');

  const [barcodeRef,          setBarcodeRef]          = useState<ColRef>(EMPTY_REF);
  const [productsoortRef,     setProductsoortRef]     = useState<ColRef>(EMPTY_REF);
  const [kostprijsRef,        setKostprijsRef]        = useState<ColRef>(EMPTY_REF);
  const [bestelnummerRef,     setBestelnummerRef]     = useState<ColRef>(EMPTY_REF);
  const [verkoopprijsRef,     setVerkoopprijsRef]     = useState<ColRef>(EMPTY_REF);
  const [maatRef,             setMaatRef]             = useState<ColRef>(EMPTY_REF);
  const [kleurRef,            setKleurRef]            = useState<ColRef>(EMPTY_REF);
  const [kleurMapping,        setKleurMapping]        = useState<Record<string, string>>({});
  const [kleurSearch,         setKleurSearch]         = useState('');
  const [kleurSelected,       setKleurSelected]       = useState<Record<string, boolean>>({});
  const [kleurPage,           setKleurPage]           = useState(0);
  const [kleurBulkValue,      setKleurBulkValue]      = useState('');
  const [veiligheidsclassRef, setVeiligheidsclassRef] = useState<ColRef>(EMPTY_REF);
  const [geslachtRef,         setGeslachtRef]         = useState<ColRef>(EMPTY_REF);
  const [merkRef,             setMerkRef]             = useState<ColRef>(EMPTY_REF);
  const [artikelcodeRef,      setArtikelcodeRef]      = useState<ColRef>(EMPTY_REF);
  const [omschrijvingRefs,    setOmschrijvingRefs]    = useState<ColRef[]>([]);
  const [productnaamRefs,     setProductnaamRefs]     = useState<ColRef[]>([]);
  const [outputData,          setOutputData]          = useState<Record<string, unknown>[] | null>(null);
  const [processing,          setProcessing]          = useState(false);
  const [processingProgress,  setProcessingProgress]  = useState(0);
  const [downloading,         setDownloading]         = useState(false);
  const [fieldErrors,         setFieldErrors]         = useState<Record<string, boolean>>({});
  const [codeReminder,        setCodeReminder]        = useState(false);

  const [mappingStore, setMappingStore] = useState<MappingStore>(() => loadMappingStore());
  const [showMappingEditor, setShowMappingEditor] = useState(false);

  const autoDetRef      = useRef<Record<string, string | null>>({});
  const parsedSheetsRef = useRef<Record<string, { cols: string[]; data: Record<string, unknown>[] }>>({})

  // ── Sheet helpers ────────────────────────────────────────

  const sheetNames   = supplFile?.workbook.SheetNames ?? [];
  const primarySheet = supplFile?.sheetName ?? sheetNames[0] ?? '';

  function getSheetInfo(sheetName: string): { cols: string[]; data: Record<string, unknown>[] } {
    if (!supplFile) return { cols: [], data: [] };
    if (!sheetName || sheetName === primarySheet) return { cols: supplFile.cols, data: supplFile.data };
    if (parsedSheetsRef.current[sheetName]) return parsedSheetsRef.current[sheetName];
    const ws = supplFile.workbook.Sheets[sheetName];
    if (!ws) return { cols: [], data: [] };
    const rows = sheetTo2D(ws);
    const hIdx = detectHeaderRow(rows);
    const parsed = parseFromHeaderRow(rows, hIdx);
    parsedSheetsRef.current[sheetName] = parsed;
    return parsed;
  }

  function getSheetCols(sheetName: string) { return getSheetInfo(sheetName).cols; }

  // Mapping store is persisted explicitly via the Save button inside MappingEditor

  // Unique productsoort values present in the currently selected column
  const fileValues = useMemo((): string[] => {
    if (!productsoortRef.col) return [];
    const { data } = getSheetInfo(productsoortRef.sheet || primarySheet);
    return [...new Set(data.map(r => String(r[productsoortRef.col] ?? '').trim()).filter(Boolean))].sort();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [productsoortRef.col, productsoortRef.sheet, supplFile]);

  function buildKleurMapping(ref: ColRef, existing: Record<string, string>): Record<string, string> {
    if (!ref.col) return {};
    const { data } = getSheetInfo(ref.sheet || primarySheet);
    const unique = [...new Set(data.map(r => String(r[ref.col] ?? '').trim()).filter(Boolean))].sort();
    const mapping: Record<string, string> = {};
    unique.forEach(v => { mapping[v] = existing[v] ?? v; });
    return mapping;
  }

  function handleKleurRefChange(ref: ColRef) {
    setKleurRef(ref);
    setKleurMapping(buildKleurMapping(ref, kleurMapping));
    setKleurSearch(''); setKleurSelected({}); setKleurBulkValue(''); setKleurPage(0);
  }

  // ── Memoized kleur computations ──────────────────────────
  // These are only recomputed when their specific inputs change,
  // preventing the entire color list from being rebuilt on every keystroke.

  const kleurAllEntries = useMemo(
    () => Object.entries(kleurMapping),
    [kleurMapping],
  );

  const kleurTotalModified = useMemo(
    () => kleurAllEntries.filter(([k, v]) => k !== v).length,
    [kleurAllEntries],
  );

  const kleurFiltered = useMemo(() => {
    if (!kleurSearch) return kleurAllEntries;
    const q = kleurSearch.toLowerCase();
    return kleurAllEntries.filter(([k]) => k.toLowerCase().includes(q));
  }, [kleurAllEntries, kleurSearch]);

  const kleurPageCount = useMemo(
    () => Math.max(1, Math.ceil(kleurFiltered.length / KLEUR_PAGE_SIZE)),
    [kleurFiltered],
  );

  const kleurSafePage = Math.min(kleurPage, kleurPageCount - 1);

  const kleurPagedEntries = useMemo(
    () => kleurFiltered.slice(kleurSafePage * KLEUR_PAGE_SIZE, (kleurSafePage + 1) * KLEUR_PAGE_SIZE),
    [kleurFiltered, kleurSafePage],
  );

  const kleurPagedUnmod = useMemo(
    () => kleurPagedEntries.filter(([k, v]) => k === v),
    [kleurPagedEntries],
  );

  const kleurPagedMod = useMemo(
    () => kleurPagedEntries.filter(([k, v]) => k !== v),
    [kleurPagedEntries],
  );

  const kleurSelectedCount = useMemo(
    () => Object.keys(kleurSelected).length,
    [kleurSelected],
  );

  // O(1) membership check — kleurSelected is Record<string, boolean>
  const allFilteredSelected =
    kleurFiltered.length > 0 && kleurFiltered.every(([k]) => kleurSelected[k]);

  // ── Stable KleurRow handlers (useCallback → stable refs → memo skips re-renders) ──

  const handleKleurRowSelect = useCallback((original: string, checked: boolean) => {
    setKleurSelected(prev => {
      if (checked) return { ...prev, [original]: true };
      const next = { ...prev };
      delete next[original];
      return next;
    });
  }, []);

  const handleKleurRowValueChange = useCallback((original: string, value: string) => {
    setKleurMapping(prev => ({ ...prev, [original]: value }));
  }, []);

  const handleKleurRowReset = useCallback((original: string) => {
    setKleurMapping(prev => ({ ...prev, [original]: original }));
  }, []);

  function applyBulk() {
    if (!kleurBulkValue.trim()) return;
    const selectedKeys = Object.keys(kleurSelected);
    setKleurMapping(prev => {
      const next = { ...prev };
      selectedKeys.forEach(k => { next[k] = kleurBulkValue; });
      return next;
    });
    setKleurSelected({}); setKleurBulkValue('');
  }

  // ── Presets ──────────────────────────────────────────────

  function getMappings(): Record<string, unknown> {
    return {
      hoofdleverancier, btwInkoop,
      barcodeRef, productsoortRef, kostprijsRef, bestelnummerRef,
      verkoopprijsRef, maatRef, kleurRef, kleurMapping,
      veiligheidsclassRef, geslachtRef, merkRef, artikelcodeRef,
      omschrijvingRefs, productnaamRefs,
    };
  }

  function applyPreset(preset: Preset) {
    setFieldErrors({});
    setCodeReminder(true);
    const m = preset.mappings as unknown as SupplierConverterMappings;

    // Eagerly parse ALL sheets so column lookup is complete
    const allCols = new Set<string>(supplFile!.cols);
    sheetNames.forEach(sn => getSheetInfo(sn).cols.forEach(c => allCols.add(c)));

    const checked: string[] = [];
    const missing: string[] = [];

    // Validate sheet exists in current file before applying it
    function validSheet(sheet: string): string {
      return sheetNames.includes(sheet) ? sheet : primarySheet;
    }

    function tryRef(saved: ColRef | undefined, setter: (v: ColRef) => void) {
      if (!saved?.col) return;
      checked.push(saved.col);
      const targetSheet = validSheet(saved.sheet);
      const sheetCols = getSheetInfo(targetSheet).cols;
      if (sheetCols.includes(saved.col)) setter({ sheet: targetSheet, col: saved.col });
      else if (allCols.has(saved.col)) setter({ sheet: primarySheet, col: saved.col });
      else missing.push(saved.col);
    }

    function tryRefs(saved: ColRef[] | undefined, setter: (v: ColRef[]) => void) {
      if (!saved?.length) return;
      const valid: ColRef[] = [];
      saved.forEach(ref => {
        if (!ref.col) return;
        checked.push(ref.col);
        const targetSheet = validSheet(ref.sheet);
        const sheetCols = getSheetInfo(targetSheet).cols;
        if (sheetCols.includes(ref.col)) valid.push({ sheet: targetSheet, col: ref.col });
        else if (allCols.has(ref.col)) valid.push({ sheet: primarySheet, col: ref.col });
        else missing.push(ref.col);
      });
      if (valid.length) setter(valid);
    }

    tryRef(m.barcodeRef, setBarcodeRef);
    tryRef(m.productsoortRef, setProductsoortRef);
    tryRef(m.kostprijsRef, setKostprijsRef);
    tryRef(m.bestelnummerRef, setBestelnummerRef);
    tryRef(m.verkoopprijsRef, setVerkoopprijsRef);
    tryRef(m.maatRef, setMaatRef);
    tryRef(m.veiligheidsclassRef, setVeiligheidsclassRef);
    tryRef(m.geslachtRef, setGeslachtRef);
    tryRef(m.merkRef, setMerkRef);
    tryRef(m.artikelcodeRef, setArtikelcodeRef);
    tryRefs(m.omschrijvingRefs, setOmschrijvingRefs);
    tryRefs(m.productnaamRefs, setProductnaamRefs);

    if (m.kleurRef?.col) {
      checked.push(m.kleurRef.col);
      const targetSheet = validSheet(m.kleurRef.sheet);
      const sheetCols = getSheetInfo(targetSheet).cols;
      const newRef = sheetCols.includes(m.kleurRef.col)
        ? { sheet: targetSheet, col: m.kleurRef.col }
        : allCols.has(m.kleurRef.col) ? { sheet: primarySheet, col: m.kleurRef.col } : null;
      if (newRef) {
        setKleurRef(newRef);
        setKleurMapping(buildKleurMapping(newRef, m.kleurMapping ?? {}));
        setKleurSearch(''); setKleurSelected({}); setKleurBulkValue(''); setKleurPage(0);
      } else {
        missing.push(m.kleurRef.col);
      }
    }

    if (m.hoofdleverancier !== undefined) setHoofdleverancier(m.hoofdleverancier);
    if (m.btwInkoop !== undefined) setBtwInkoop(m.btwInkoop);

    const matchCount = checked.length - missing.length;
    const matchRate = checked.length > 0 ? matchCount / checked.length : 1;
    const safeName = preset.name.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const name = `<strong>${safeName}</strong>`;
    setBanner(
      missing.length === 0
        ? { type: 'success', icon: '✓', message: lang === 'nl' ? `Preset ${name} geladen — alle ${checked.length} kolommen gevonden.` : `Preset ${name} loaded — all ${checked.length} columns matched.` }
        : matchRate >= 0.5
          ? { type: 'warning', icon: '⚠', message: lang === 'nl' ? `Preset ${name} gedeeltelijk geladen — ${matchCount}/${checked.length} kolommen gevonden. Niet gevonden: ${missing.join(', ')}.` : `Preset ${name} partially loaded — ${matchCount}/${checked.length} columns matched. Missing: ${missing.join(', ')}.` }
          : { type: 'warning', icon: '✗', message: lang === 'nl' ? `Preset ${name} lijkt niet te passen bij dit bestand (${matchCount}/${checked.length} kolommen gevonden). Gebruikt u de juiste preset?` : `Preset ${name} doesn't seem to match this file (${matchCount}/${checked.length} columns found). Are you using the right preset?` }
    );
  }

  // ── Navigation ──────────────────────────────────────────

  function enterStep2() {
    const ps = supplFile!.sheetName;
    parsedSheetsRef.current = {};
    const { cols, data } = supplFile!;
    const det: Record<string, string | null> = {
      barcode:      findBestCol(cols, data, EAN_HINTS, null),
      productsoort: findBestCol(cols, data, PRODUCTSOORT_HINTS, null),
      kostprijs:    findBestCol(cols, data, KOSTPRIJS_HINTS, null),
      bestelnummer: findBestCol(cols, data, BESTELNR_HINTS, null),
      verkoopprijs: findBestCol(cols, data, VERKOOPPRIJS_HINTS, null),
      veiligheid:   findBestCol(cols, data, VEILIGHEID_HINTS, null),
      geslacht:     findBestCol(cols, data, GESLACHT_HINTS, null),
      maat:         findBestCol(cols, data, MAAT_HINTS, null),
      merk:         findBestCol(cols, data, MERK_HINTS, null),
      kleur:        findBestCol(cols, data, KLEUR_HINTS, null),
      artikelcode:  findBestCol(cols, data, CODE_HINTS, null),
      omschr:       findBestCol(cols, data, OMSCHR_HINTS, null),
      productnaam:  findBestCol(cols, data, PRODUCTNAAM_HINTS, null),
    };
    autoDetRef.current = det;
    const mk = (key: string): ColRef => ({ sheet: ps, col: det[key] || '' });

    setBarcodeRef(mk('barcode'));
    setProductsoortRef(mk('productsoort'));
    setKostprijsRef(mk('kostprijs'));
    setBestelnummerRef(mk('bestelnummer'));
    setVerkoopprijsRef(mk('verkoopprijs'));
    setMaatRef(mk('maat'));
    const kleurDet = mk('kleur');
    setKleurRef(kleurDet);
    setKleurMapping(buildKleurMapping(kleurDet, {}));
    setKleurSearch(''); setKleurSelected({}); setKleurBulkValue(''); setKleurPage(0);
    setVeiligheidsclassRef(mk('veiligheid'));
    setGeslachtRef(mk('geslacht'));
    setMerkRef(mk('merk'));
    setArtikelcodeRef(mk('artikelcode'));
    setOmschrijvingRefs(det.omschr ? [{ sheet: ps, col: det.omschr }] : []);
    setProductnaamRefs(det.productnaam ? [{ sheet: ps, col: det.productnaam }] : []);

    const required = [det.barcode, det.productsoort, det.kostprijs, det.bestelnummer, det.verkoopprijs, det.maat, det.kleur];
    const detected = required.filter(Boolean).length;
    const allOk = detected === required.length;
    setBanner({
      type: allOk ? 'success' : 'warning',
      icon: allOk ? '✓' : '⚠',
      message: allOk
        ? t('scAutoDetAll').replace('{n}', String(required.length))
        : t('scAutoDetPartial').replace('{det}', String(detected)).replace('{n}', String(required.length)),
    });
    setStep(2);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function processFiles() {
    const code = parseInt(startingCode, 10);
    const errs: Record<string, boolean> = {};
    if (!hoofdleverancier.trim())      errs.hoofdleverancier = true;
    if (isNaN(code) || code < 0)       errs.startingCode     = true;
    if (!barcodeRef.col)               errs.barcodeRef       = true;
    if (!productsoortRef.col)          errs.productsoortRef  = true;
    if (!maatRef.col)                  errs.maatRef          = true;
    if (!kleurRef.col)                 errs.kleurRef         = true;
    if (omschrijvingRefs.length === 0) errs.omschrijvingRefs = true;
    if (productnaamRefs.length === 0)  errs.productnaamRefs  = true;
    if (Object.keys(errs).length > 0) { setFieldErrors(errs); alert(lang === 'nl' ? 'Vul alle verplichte velden in (gemarkeerd met *).' : 'Please fill in all required fields (marked with *).'); return; }
    setFieldErrors({});

    const allRefs: ColRef[] = [
      barcodeRef, productsoortRef, kostprijsRef, bestelnummerRef,
      verkoopprijsRef, veiligheidsclassRef, geslachtRef,
      maatRef, merkRef, kleurRef, artikelcodeRef,
      ...omschrijvingRefs, ...productnaamRefs,
    ];
    const usedSheets = [...new Set(allRefs.map(r => r.sheet || primarySheet))];
    const sheetsData: Record<string, Record<string, unknown>[]> = {};
    usedSheets.forEach(sn => { sheetsData[sn] = getSheetInfo(sn).data; });

    const rowCount = sheetsData[primarySheet]?.length ?? supplFile!.data.length;
    const resolve = (ref: ColRef): ColRef => ({ sheet: ref.sheet || primarySheet, col: ref.col });

    const args = {
      sheetsData, rowCount, startingCode: code,
      hoofdleverancier: hoofdleverancier.trim(), btwInkoop,
      mappingStore,
      barcodeRef:               resolve(barcodeRef),
      productsoortRef:          resolve(productsoortRef),
      kostprijsRef:             resolve(kostprijsRef),
      bestelnummerRef:          resolve(bestelnummerRef),
      verkoopprijsRef:          resolve(verkoopprijsRef),
      veiligheidsclassificatieRef: resolve(veiligheidsclassRef),
      geslachtRef:              resolve(geslachtRef),
      maatRef:                  resolve(maatRef),
      merkRef:                  resolve(merkRef),
      kleurRef:                 resolve(kleurRef),
      kleurMapping,
      artikelcodeRef:           resolve(artikelcodeRef),
      omschrijvingRefs:         omschrijvingRefs.map(resolve),
      productnaamRefs:          productnaamRefs.map(resolve),
    };

    setProcessing(true);
    setProcessingProgress(0);
    // defer so the loading state paints before heavy work begins
    await new Promise<void>(r => setTimeout(r, 30));
    try {
      const data = await buildOutputRowsAsync(args, pct => setProcessingProgress(pct));
      setOutputData(data);
      setStep(3);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } finally {
      setProcessing(false);
      setProcessingProgress(0);
    }
  }

  function downloadFile() {
    const name = supplFile!.fileName.replace(/\.(xlsx?|xlsm|xlsb|ods|csv|tsv|txt)$/i, '_converted.xls');
    const data = outputData!;
    setDownloading(true);
    setTimeout(() => {
      try { downloadXLSX(data, [...OUTPUT_COLS], name, 'Exact Import'); }
      finally { setDownloading(false); }
    }, 30);
  }

  // ── Field rendering helpers ──────────────────────────────

  const ad = autoDetRef.current;

  function selClass(key: string, ref: ColRef): string {
    if (!ref.col) return 'needs-review';
    return ref.col === ad[key] ? 'auto-detected' : '';
  }

  function renderField(
    label: string,
    ref: ColRef,
    setRef: (r: ColRef) => void,
    opts: { required?: boolean; optional?: boolean; defaultLabel?: string; hint?: string; autoKey?: string; tooltip?: string; errorKey?: string; extra?: React.ReactNode } = {},
  ) {
    const { required, optional, defaultLabel, hint, autoKey, tooltip, errorKey, extra } = opts;
    const cols = getSheetCols(ref.sheet || primarySheet);
    const { data } = getSheetInfo(ref.sheet || primarySheet);
    const hasError = errorKey ? !!fieldErrors[errorKey] : false;
    const cls = autoKey ? selClass(autoKey, ref) : (required && !ref.col ? 'needs-review' : '');
    return (
      <div>
        <label className="field-label">{label}{required && ' *'}{tooltip && <Tooltip text={tooltip} />}</label>
        <select value={ref.col} className={cls} style={hasError ? { border: '1.5px solid var(--red-text)' } : undefined} onChange={e => { setRef({ ...ref, sheet: ref.sheet || primarySheet, col: e.target.value }); if (errorKey) setFieldErrors(p => ({ ...p, [errorKey]: false })); }}>
          <option value="">{defaultLabel ?? (optional ? '— none —' : '— select column —')}</option>
          {cols.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <ColumnPreview col={ref.col} data={data} />
        {extra}
        {hint && <p className="field-hint">{hint}</p>}
        <SheetPicker sheetNames={sheetNames} value={ref.sheet || primarySheet} onChange={s => setRef({ sheet: s, col: '' })} />
      </div>
    );
  }

  const activiefVanaf = (() => {
    const now = new Date();
    return `01-${String(now.getMonth() + 1).padStart(2, '0')}-${now.getFullYear()}`;
  })();

  const FIXED_FIELDS = [
    { label: t('scActiviefVanaf'),   value: activiefVanaf },
    { label: t('scInkoop'),          value: 'Ja' },
    { label: t('scOrdergestuurd'),   value: 'Ja' },
    { label: t('scVerkoop'),         value: 'Ja' },
    { label: t('scVoorraad'),        value: 'Ja' },
    { label: t('scBtwVerkoop'),      value: '2' },
    { label: t('scEenheidsfactor'),  value: '1' },
    { label: t('scKmsSynch'),        value: 'Ja' },
    { label: t('scControle'),        value: 'NG' },
    { label: t('scEenheidCombo'),    value: t('scDerivedFrom') },
    { label: t('scArtikelcodeZoek'), value: t('scCopyOf') },
  ];

  const STEPS = [
    { num: 1, label: t('scStep1') },
    { num: 2, label: t('scStep2') },
    { num: 3, label: t('scStep3') },
  ];

  // ── Render ───────────────────────────────────────────────

  return (
    <div className="container">
      <header>
        <h1>{t('scTitle')}</h1>
        <p>{t('scDesc')}</p>
      </header>

      <Stepper step={step} onStepClick={setStep} steps={STEPS} />

      {/* ── Step 1: Upload ── */}
      {step === 1 && (
        <>
          <FileUploadCard title={lang === 'nl' ? 'Leveranciersproductbestand' : 'Supplier product file'} icon="📋" onFileLoaded={setSupplFile} initialFile={supplFile} />
          <div className="actions">
            <button className="btn btn-primary" disabled={!supplFile} onClick={enterStep2}>{t('btnContinue')}</button>
          </div>
        </>
      )}

      {/* ── Step 2: Field mapping ── */}
      {step === 2 && supplFile && (
        <>
          <PresetBar tool="supplier_converter" getMappings={getMappings} onLoad={applyPreset} />
          {banner && <Banner {...banner} />}

          {/* General settings */}
          <div className="card">
            <div className="card-title">{t('scSettingsCardTitle')}</div>
            <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: '0.75rem', marginTop: '-0.25rem' }}>{t('scSettingsCardDesc')}</p>
            <div className="grid-3">
              <div>
                <label className="field-label">{t('scHoofdleverancier')} *<Tooltip text={t('ttHoofdlev')} /></label>
                <input type="text" style={{ ...INPUT_STYLE, ...(fieldErrors.hoofdleverancier ? { border: '1.5px solid var(--red-text)' } : {}) }} value={hoofdleverancier} placeholder={t('scHoofdlevPlaceholder')} onChange={e => { setHoofdleverancier(e.target.value); setFieldErrors(p => ({ ...p, hoofdleverancier: false })); }} />
                <p className="field-hint">{t('scHoofdlevHint')}</p>
              </div>
              <div>
                <label className="field-label">{t('scBtwInkoop')} *<Tooltip text={t('ttBtwInkoop')} /></label>
                <select value={btwInkoop} onChange={e => setBtwInkoop(e.target.value)}>
                  <option value="3">{t('scBtwHolland')}</option>
                  <option value="9">{t('scBtwEurope')}</option>
                  <option value="10">{t('scBtwOutside')}</option>
                </select>
              </div>
              <div>
                <label className="field-label">{t('scLastCode')} *<Tooltip text={t('ttLastCode')} /></label>
                <input type="number" style={{ ...INPUT_STYLE, ...(fieldErrors.startingCode ? { border: '1.5px solid var(--red-text)' } : (codeReminder && !startingCode ? { border: '1.5px solid #BA7517' } : {})) }} min="0" value={startingCode} placeholder={t('scLastCodePlaceholder')} onChange={e => { setStartingCode(e.target.value); setCodeReminder(false); setFieldErrors(p => ({ ...p, startingCode: false })); }} />
                <p className="field-hint">{t('scLastCodeHint')}</p>
                {codeReminder && !startingCode && (
                  <p style={{ fontSize: 11, color: '#BA7517', marginTop: 3 }}>
                    {lang === 'nl' ? '↑ Vul het laatste gebruikte codenummer in' : '↑ Enter the last used code number'}
                  </p>
                )}
              </div>
            </div>
          </div>

          {/* Column mappings */}
          <div className="card">
            <div className="card-title">{t('scMappingCardTitle')}</div>
            <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: '0.75rem', marginTop: '-0.25rem' }}>{t('scMappingCardDesc')}</p>
            {sheetNames.length > 1 && (
              <p style={{ fontSize: 11, color: 'var(--text-secondary)', marginBottom: '0.75rem' }}>{t('scMultiSheetNote')}</p>
            )}
            <div className="grid-3">
              {renderField(t('scBarcode'), barcodeRef, setBarcodeRef, { required: true, autoKey: 'barcode', tooltip: t('ttBarcode'), errorKey: 'barcodeRef' })}
              {renderField(t('scProductsoort'), productsoortRef, setProductsoortRef, { required: true, autoKey: 'productsoort', tooltip: t('ttProductsoort'), errorKey: 'productsoortRef' })}
              {renderField(t('scKostprijs'), kostprijsRef, setKostprijsRef, { required: true, autoKey: 'kostprijs', tooltip: t('ttKostprijs') })}
              {renderField(t('scBestelnummer'), bestelnummerRef, setBestelnummerRef, { required: true, autoKey: 'bestelnummer', tooltip: t('ttBestelnummer') })}
              {renderField(t('scVerkoopprijs'), verkoopprijsRef, setVerkoopprijsRef, { required: true, autoKey: 'verkoopprijs', tooltip: t('ttVerkoopprijs') })}
              {renderField(t('scMaat'), maatRef, setMaatRef, { required: true, autoKey: 'maat', tooltip: t('ttMaat'), errorKey: 'maatRef' })}
              {renderField(t('scKleur'), kleurRef, handleKleurRefChange, { required: true, autoKey: 'kleur', tooltip: t('ttKleur'), errorKey: 'kleurRef' })}
              {renderField(t('scVeiligheid'), veiligheidsclassRef, setVeiligheidsclassRef, { optional: true, autoKey: 'veiligheid', tooltip: t('ttVeiligheid') })}
            </div>

            <hr className="divider" />
            <div style={{ fontSize: 11, color: 'var(--text-secondary)', marginBottom: '0.75rem' }}>{t('scOptionalCols')}</div>
            <div className="grid-3">
              {renderField(t('scGeslacht'), geslachtRef, setGeslachtRef, { optional: true, defaultLabel: t('scGeslachtDefault'), autoKey: 'geslacht', tooltip: t('ttGeslacht') })}
              {renderField(t('scMerk'), merkRef, setMerkRef, { optional: true, autoKey: 'merk', tooltip: t('ttMerk') })}
              {renderField(t('scArtikelcode'), artikelcodeRef, setArtikelcodeRef, { optional: true, autoKey: 'artikelcode', tooltip: t('ttArtikelcode') })}
            </div>
          </div>

          {/* Color value mapping */}
          {kleurRef.col && kleurAllEntries.length > 0 && (
            <div className="card">
              <div className="card-title">{t('scKleurMapping')}</div>
              <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: '0.75rem' }}>
                {kleurSearch
                  ? (lang === 'nl'
                      ? `${kleurFiltered.length} van ${kleurAllEntries.length} waarden komen overeen met "${kleurSearch}" — `
                      : `${kleurFiltered.length} of ${kleurAllEntries.length} values matching "${kleurSearch}" — `)
                  : (lang === 'nl'
                      ? `${kleurAllEntries.length} unieke kleurwaarden — `
                      : `${kleurAllEntries.length} unique color values — `)
                }
                <span style={{ color: 'var(--green-dark)', fontWeight: 500 }}>{kleurTotalModified} {t('scKleurModified')}</span>
                {', '}
                <span style={{ color: 'var(--text-secondary)' }}>{kleurAllEntries.length - kleurTotalModified} {t('scKleurUnmodified')}</span>
                {lang === 'nl'
                  ? '. Bewerk de rechterkolom om te hernoemen, of selecteer meerdere voor bulkhernoemen.'
                  : '. Edit the right column to rename, or select multiple and bulk-rename.'}
              </p>

              {/* Search + select-all + reset */}
              <div style={{ display: 'flex', gap: 8, marginBottom: 8, alignItems: 'center' }}>
                <input
                  type="text"
                  placeholder={t('scKleurSearchPh')}
                  value={kleurSearch}
                  onChange={e => { setKleurSearch(e.target.value); setKleurPage(0); }}
                  style={{ flex: 1, fontSize: 12, padding: '4px 8px', border: '0.5px solid var(--border-md)', borderRadius: 4, background: 'var(--bg)', color: 'var(--text)', fontFamily: 'inherit' }}
                />
                <label style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11, color: 'var(--text-secondary)', cursor: 'pointer', flexShrink: 0, userSelect: 'none' }}>
                  <input
                    type="checkbox"
                    checked={allFilteredSelected}
                    onChange={() => {
                      if (allFilteredSelected) {
                        setKleurSelected(prev => {
                          const next = { ...prev };
                          kleurFiltered.forEach(([k]) => delete next[k]);
                          return next;
                        });
                      } else {
                        setKleurSelected(prev => {
                          const next = { ...prev };
                          kleurFiltered.forEach(([k]) => { next[k] = true; });
                          return next;
                        });
                      }
                    }}
                  />
                  {kleurSearch ? t('scKleurSelectFiltered') : t('scKleurSelectAll')} ({kleurFiltered.length})
                </label>
                <button
                  className="btn btn-sm"
                  style={{ fontSize: 11, flexShrink: 0 }}
                  onClick={() => {
                    setKleurMapping(Object.fromEntries(kleurAllEntries.map(([k]) => [k, k])));
                    setKleurSelected({}); setKleurBulkValue('');
                  }}
                >
                  {t('scKleurResetAll')}
                </button>
              </div>

              {/* Bulk rename bar */}
              {kleurSelectedCount > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 10px', background: 'var(--bg-secondary)', border: '0.5px solid var(--border)', borderRadius: 4, marginBottom: 8 }}>
                  <span style={{ fontSize: 11, color: 'var(--text-secondary)', flexShrink: 0 }}>
                    {kleurSelectedCount} {lang === 'nl' ? 'geselecteerd' : 'selected'}
                  </span>
                  <input
                    type="text"
                    value={kleurBulkValue}
                    onChange={e => setKleurBulkValue(e.target.value)}
                    onKeyDown={e => { if (e.key === 'Enter') applyBulk(); }}
                    placeholder={t('scKleurNewValuePh')}
                    style={{ flex: 1, fontSize: 12, padding: '3px 8px', border: '0.5px solid var(--border-md)', borderRadius: 4, background: 'var(--bg)', color: 'var(--text)', fontFamily: 'inherit' }}
                  />
                  <button className="btn btn-sm" style={{ flexShrink: 0 }} onClick={applyBulk}>{t('scKleurApply')}</button>
                  <button className="btn btn-sm" style={{ flexShrink: 0 }} onClick={() => { setKleurSelected({}); setKleurBulkValue(''); }}>{t('scKleurCancel')}</button>
                </div>
              )}

              {/* Column headers */}
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4, paddingLeft: 22 }}>
                <span style={{ flex: '0 0 200px', fontSize: 11, fontWeight: 500, color: 'var(--text-secondary)' }}>{t('scKleurOriginal')}</span>
                <span style={{ width: 16 }} />
                <span style={{ flex: 1, fontSize: 11, fontWeight: 500, color: 'var(--text-secondary)' }}>{t('scKleurOutput')}</span>
              </div>

              {/* Paginated rows — max KLEUR_PAGE_SIZE DOM nodes at a time */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 3, maxHeight: 300, overflowY: 'auto' }}>
                {kleurPagedEntries.length === 0 && (
                  <div style={{ fontSize: 12, color: 'var(--text-secondary)', padding: '8px 0' }}>
                    {t('scKleurNoMatch')} "{kleurSearch}".
                  </div>
                )}

                {kleurPagedUnmod.length > 0 && (
                  <>
                    <div style={{ fontSize: 10, fontWeight: 600, color: 'var(--text-secondary)', letterSpacing: '0.06em', textTransform: 'uppercase', padding: '4px 0 2px', opacity: 0.7 }}>
                      {lang === 'nl' ? `Niet aangepast (${kleurPagedUnmod.length})` : `Unmodified (${kleurPagedUnmod.length})`}
                    </div>
                    {kleurPagedUnmod.map(([original, mapped]) => (
                      <KleurRow
                        key={original}
                        original={original}
                        mapped={mapped}
                        isSelected={kleurSelected[original] ?? false}
                        isModified={false}
                        onSelect={handleKleurRowSelect}
                        onValueChange={handleKleurRowValueChange}
                        onReset={handleKleurRowReset}
                      />
                    ))}
                  </>
                )}

                {kleurPagedMod.length > 0 && (
                  <>
                    {kleurPagedUnmod.length > 0 && <div style={{ borderTop: '0.5px solid var(--border)', margin: '6px 0 2px' }} />}
                    <div style={{ fontSize: 10, fontWeight: 600, color: 'var(--green-dark)', letterSpacing: '0.06em', textTransform: 'uppercase', padding: '2px 0', opacity: 0.85 }}>
                      {lang === 'nl' ? `Aangepast (${kleurPagedMod.length})` : `Modified (${kleurPagedMod.length})`}
                    </div>
                    {kleurPagedMod.map(([original, mapped]) => (
                      <KleurRow
                        key={original}
                        original={original}
                        mapped={mapped}
                        isSelected={kleurSelected[original] ?? false}
                        isModified={true}
                        onSelect={handleKleurRowSelect}
                        onValueChange={handleKleurRowValueChange}
                        onReset={handleKleurRowReset}
                      />
                    ))}
                  </>
                )}
              </div>

              {/* Pager */}
              {kleurPageCount > 1 && (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, paddingTop: 8, borderTop: '0.5px solid var(--border)', marginTop: 6 }}>
                  <button className="btn btn-sm" disabled={kleurSafePage === 0} onClick={() => setKleurPage(p => p - 1)}>‹</button>
                  <span style={{ fontSize: 11, color: 'var(--text-secondary)' }}>
                    {kleurSafePage + 1} / {kleurPageCount}
                    &nbsp;·&nbsp;
                    {lang === 'nl'
                      ? `${kleurFiltered.length} waarden, ${KLEUR_PAGE_SIZE} per pagina`
                      : `${kleurFiltered.length} values, ${KLEUR_PAGE_SIZE} per page`}
                  </span>
                  <button className="btn btn-sm" disabled={kleurSafePage >= kleurPageCount - 1} onClick={() => setKleurPage(p => p + 1)}>›</button>
                </div>
              )}
            </div>
          )}

          {/* Productsoort mapping trigger */}
          {(() => {
            const unknownCount = fileValues.filter(v => !isMapped(v, mappingStore)).length;
            const customCount = Object.keys(mappingStore).length;
            return (
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px', background: 'var(--bg-secondary)', border: '0.5px solid var(--border)', borderRadius: 'var(--radius-md)', marginBottom: 8 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <span style={{ fontSize: 13, fontWeight: 500, color: 'var(--text)' }}>
                    {lang === 'nl' ? 'Productsoort koppelingen' : 'Productsoort mappings'}
                  </span>
                  {fileValues.length > 0 && (
                    <span style={{ fontSize: 11, color: 'var(--text-secondary)', marginLeft: 8 }}>
                      {fileValues.length} {lang === 'nl' ? 'types in bestand' : 'types in file'}
                    </span>
                  )}
                  {unknownCount > 0 && <span className="badge badge-amber" style={{ marginLeft: 6 }}>{unknownCount} {lang === 'nl' ? 'onbekend' : 'unknown'}</span>}
                  {customCount > 0 && <span className="badge badge-green" style={{ marginLeft: 6 }}>{customCount} {lang === 'nl' ? 'aangepast' : 'custom'}</span>}
                </div>
                <button className="btn btn-sm" onClick={() => setShowMappingEditor(true)}>
                  {lang === 'nl' ? 'Bewerken →' : 'Edit →'}
                </button>
              </div>
            );
          })()}
          {showMappingEditor && (
            <MappingEditor
              fileValues={fileValues}
              store={mappingStore}
              onApply={setMappingStore}
              onClose={() => setShowMappingEditor(false)}
            />
          )}

          {/* Mapping preview — combined Artikelgroep + Merk result */}
          {productsoortRef.col && (() => {
            const psData = getSheetInfo(productsoortRef.sheet || primarySheet).data;
            const mkData = merkRef.col ? getSheetInfo(merkRef.sheet || primarySheet).data : null;
            type S = { ps: string; mk: string; ag: string; eh: string };
            const seen = new Set<string>();
            const samples: S[] = [];
            for (let i = 0; i < psData.length && samples.length < 5; i++) {
              const ps = String(psData[i]?.[productsoortRef.col] ?? '').trim();
              if (!ps) continue;
              const mk = mkData && merkRef.col ? String(mkData[i]?.[merkRef.col] ?? '').trim() : '';
              const key = `${ps}||${mk}`;
              if (seen.has(key)) continue;
              seen.add(key);
              const baseAg = resolveArtikelgroep(ps, mappingStore);
              samples.push({ ps, mk, ag: mk ? `${baseAg} - ${mk}` : baseAg, eh: resolveEenheid(ps, mappingStore) });
            }
            if (!samples.length) return null;
            return (
              <div className="card" style={{ padding: '10px 14px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 8 }}>
                  <span style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', color: 'var(--text-secondary)' }}>
                    {lang === 'nl' ? 'Koppeling voorbeeld' : 'Mapping preview'}
                  </span>
                  <Tooltip text={lang === 'nl' ? 'Eerste unieke rijen — toont hoe Productsoort en Merk samen de Artikelgroep en Eenheid bepalen.' : 'First unique rows — shows how Productsoort and Merk combine to produce Artikelgroep and Eenheid.'} />
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                  {samples.map(s => (
                    <div key={`${s.ps}||${s.mk}`} style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
                      <span className="col-preview-val" style={{ fontStyle: 'italic', color: 'var(--text-secondary)' }}>
                        {s.ps}{s.mk ? ` + ${s.mk}` : ''}
                      </span>
                      <span style={{ fontSize: 10, color: 'var(--text-secondary)' }}>→</span>
                      <span className="col-preview-val">{s.ag}</span>
                      <span className="col-preview-val" style={{ opacity: 0.7 }}>{s.eh}</span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })()}

          {/* Description columns */}
          <div className="card">
            <div className="card-title">{t('scDescColsCardTitle')}</div>
            <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: '0.75rem', marginTop: '-0.25rem' }}>{t('scDescColsCardDesc')}</p>
            <div className="grid-2" style={{ alignItems: 'flex-start' }}>
              <MultiColCombiner sheetNames={sheetNames} primarySheet={primarySheet} getSheetCols={getSheetCols} selected={omschrijvingRefs} onChange={v => { setOmschrijvingRefs(v); setFieldErrors(p => ({ ...p, omschrijvingRefs: false })); }} label={t('scOmschrijving')} hint="" tooltip={t('ttOmschrijving')} required hasError={!!fieldErrors.omschrijvingRefs} />
              <MultiColCombiner sheetNames={sheetNames} primarySheet={primarySheet} getSheetCols={getSheetCols} selected={productnaamRefs} onChange={v => { setProductnaamRefs(v); setFieldErrors(p => ({ ...p, productnaamRefs: false })); }} label={t('scProductnaam')} hint="" tooltip={t('ttProductnaam')} required hasError={!!fieldErrors.productnaamRefs} />
            </div>
          </div>

          {/* Fixed values */}
          <div className="card">
            <div className="card-title">{t('scFixedCardTitle')}</div>
            <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: '0.75rem' }}>{t('scFixedCardDesc')}</p>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
              {FIXED_FIELDS.map(f => (
                <div key={f.label} style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '4px 10px', background: 'var(--bg-secondary)', border: '0.5px solid var(--border)', borderRadius: 20, fontSize: 11 }}>
                  <span style={{ color: 'var(--text-secondary)' }}>{f.label}</span>
                  <span style={{ fontWeight: 500, color: 'var(--green-dark)' }}>→ {f.value}</span>
                </div>
              ))}
            </div>
          </div>

          <div className="actions">
            <button className="btn" disabled={processing} onClick={() => { setStep(1); window.scrollTo({ top: 0, behavior: 'smooth' }); }}>{t('scBack')}</button>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4, alignItems: 'stretch', flex: 1, maxWidth: 220 }}>
              <button className="btn btn-primary" disabled={processing} onClick={processFiles}>
                {processing ? `${lang === 'nl' ? 'Verwerken' : 'Processing'}… ${processingProgress > 0 ? processingProgress + '%' : ''}` : t('scConvert')}
              </button>
              {processing && (
                <div style={{ height: 3, background: 'var(--border)', borderRadius: 2, overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${processingProgress}%`, background: 'var(--btn-primary-bg)', transition: 'width 0.15s ease' }} />
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {/* ── Step 3: Preview & Download ── */}
      {step === 3 && outputData && (
        <>
          {outputData.length === 0 && (
            <Banner type="warning" icon="⚠" message={lang === 'nl' ? 'Het bronbestand bevat geen rijen. Er is niets om te converteren.' : 'The source file contains no rows. There is nothing to convert.'} />
          )}
          <div className="card">
            <div className="card-title">{t('scConvCompleteTitle')}</div>
            <div style={{ marginBottom: '1rem' }}>
              <span className="summary-chip">{t('scChipSupplier')}: {hoofdleverancier}</span>
              <span className="summary-chip">{t('scChipBtw')}: {btwInkoop}</span>
              <span className="summary-chip">{t('scChipCodesFrom')}: {parseInt(startingCode, 10) + 1}</span>
            </div>
            <div className="metrics" style={{ gridTemplateColumns: 'repeat(2, minmax(0,1fr))' }}>
              <div className="metric green">
                <div className="metric-val">{outputData.length}</div>
                <div className="metric-lbl"><span>{t('scMetricConverted')}</span><Tooltip text={t('ttScMetricConverted')} /></div>
              </div>
              <div className="metric">
                <div className="metric-val">{OUTPUT_COLS.length}</div>
                <div className="metric-lbl"><span>{t('scMetricCols')}</span><Tooltip text={t('ttScMetricCols')} /></div>
              </div>
            </div>
          </div>

          <div className="card">
            <div className="card-title">{t('scPreviewCardTitle')}</div>
            <p style={{ fontSize: 12, color: 'var(--text-secondary)', marginBottom: 8 }}>{t('scPreviewCardDesc')}</p>
            <PreviewTable cols={[...OUTPUT_COLS]} data={outputData} />
          </div>

          <DataQualityDashboard data={outputData} mappingStore={mappingStore} lang={lang} />

          <div className="actions">
            <button className="btn" onClick={() => { setStep(2); window.scrollTo({ top: 0, behavior: 'smooth' }); }}>{t('scBack')}</button>
            <button className="btn btn-download" disabled={downloading} onClick={downloadFile}>
              {downloading ? (lang === 'nl' ? 'Downloaden…' : 'Downloading…') : t('scDownload')}
            </button>
          </div>
        </>
      )}
    </div>
  );
}
