import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { Upload, Trash2, Users, X } from "lucide-react";
import { format, addDays, isToday, startOfMonth, endOfMonth, startOfWeek, differenceInCalendarWeeks, differenceInCalendarDays } from "date-fns";
import { it } from "date-fns/locale";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";


import {
  Select,
  SelectContent,
  SelectItem,
  SelectSeparator,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
} from "@/components/ui/dialog";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Calendario Turni — Orari e ore lavorate da Excel" },
      {
        name: "description",
        content:
          "Carica il file Excel dei turni, scegli la domenica di riferimento e consulta orari e ore lavorate in un calendario in stile iOS, anche offline.",
      },
      { property: "og:title", content: "Calendario Turni — Orari e ore lavorate da Excel" },
      {
        property: "og:description",
        content:
          "Carica il file Excel dei turni, scegli la domenica di riferimento e consulta orari e ore lavorate in un calendario in stile iOS, anche offline.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

const STORAGE_KEY = "turni.voci";
const FILE_CACHE_NAME = "turni-file";
// Sessione persistente: file caricato + selezioni, ripristinati alla riapertura
const ROWS_KEY = "turni.file.rows";
const FILENAME_KEY = "turni.file.name";
const SELECTED_KEY = "turni.selected";
const MONTH_KEY = "turni.month";
const FILE_SUNDAY_KEY = "turni.fileSunday";
const ACCENT_KEY = "turni.accent";

const DEFAULT_ACCENT = "#00B5CE";
/** Tavolozza colori per gradienti e badge totale ore. */
const PALETTE = [
  "#00A1D8", "#B383FE", "#EF3C2C", "#9B9823", "#76BB40",
  "#52C9C1", "#D5A4DB", "#FF8E1B", "#EAD346", "#C3D117",
];

/**
 * Svuota la cache offline dei dati del file precedente.
 * Non tocca MAI l'elenco voci (STORAGE_KEY): quello si cancella solo dall'utente.
 */
async function clearPreviousFileCache() {
  try {
    for (const storage of [localStorage, sessionStorage]) {
      const keys = Object.keys(storage).filter(
        (k) => k.startsWith("turni.file") && k !== STORAGE_KEY,
      );
      keys.forEach((k) => storage.removeItem(k));
    }
  } catch {
    /* ignore */
  }
  try {
    if (typeof caches !== "undefined") {
      await caches.delete(FILE_CACHE_NAME);
    }
  } catch {
    /* ignore */
  }
}

const DAY_LABELS = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì", "Sabato", "Domenica"];
// 1-based CSV column numbers per day (Mon..Sun)
const DAY_BLOCKS: number[][] = [
  [8, 9, 10, 11],
  [13, 14, 15, 16],
  [18, 19, 20, 21],
  [23, 24, 25, 26],
  [28, 29, 30, 31],
  [33, 34, 35, 36],
  [3, 4, 5, 6],
];

const TARGET_TABLE = "inserimento orari";

const MONTH_NAMES = [
  "gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno",
  "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre",
];

/** Estrae mese e anno di riferimento dal nome del file (es. "Turni Ottobre 2026.xlsx"). */
function monthYearFromFileName(name: string | null): Date | undefined {
  if (!name) return undefined;
  const lower = name.toLowerCase();
  const monthIdx = MONTH_NAMES.findIndex((m) => lower.includes(m));
  if (monthIdx === -1) return undefined;

  const yearMatch = lower.match(/(19|20)\d{2}/) ?? lower.match(/\b\d{2}\b/);
  let year = new Date().getFullYear();
  if (yearMatch) {
    const raw = yearMatch[0];
    year = raw.length === 2 ? 2000 + Number(raw) : Number(raw);
  }
  return new Date(year, monthIdx, 1);
}

function norm(v: string): string {
  return v.toLowerCase().replace(/\s+/g, " ").trim();
}

/** Estrae la data esatta dalla cella E7 del foglio "SETUP": è la domenica della prima settimana. */
function setupSunday(XLSX: typeof import("xlsx"), wb: import("xlsx").WorkBook): Date | undefined {
  const name = wb.SheetNames.find((n) => norm(n) === "setup");
  if (!name) return undefined;
  const cell = wb.Sheets[name]?.["E7"];
  if (!cell) return undefined;
  if (cell.v instanceof Date) return cell.v;
  if (typeof cell.v === "number") {
    const p = XLSX.SSF.parse_date_code(cell.v);
    if (p) return new Date(p.y, p.m - 1, p.d);
  }
  const s = String(cell.w ?? cell.v ?? "").trim();
  const m = s.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    const year = Number(m[3]!.length === 2 ? `20${m[3]}` : m[3]);
    return new Date(year, Number(m[2]) - 1, Number(m[1]));
  }
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

/** Legge la sola tabella "Inserimento Orari" dal file Excel. */
async function parseExcel(
  file: File,
): Promise<{ rows: string[][]; month: Date | undefined; fileSunday: Date | undefined }> {
  const XLSX = await import("xlsx");
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
  // Domenica esatta della prima settimana, letta da SETUP!E7 (es. 30/08 per un file di settembre).
  const fileSunday = setupSunday(XLSX, wb);
  // Priorità al mese/anno nel nome del file; fallback sul lunedì successivo a fileSunday.
  const month = monthYearFromFileName(file.name) ?? (fileSunday ? startOfMonth(addDays(fileSunday, 1)) : undefined);

  const toRows = (name: string): string[][] =>
    (
      XLSX.utils.sheet_to_json(wb.Sheets[name]!, {
        header: 1,
        raw: false,
        defval: "",
        blankrows: false,
      }) as unknown[][]
    ).map((r) => r.map((c) => String(c ?? "").trim()));

  // 1) foglio chiamato "Inserimento Orari"
  const sheetName = wb.SheetNames.find((n) => norm(n) === TARGET_TABLE);
  if (sheetName) return { rows: toRows(sheetName).filter((r) => r.some((c) => c !== "")), month, fileSunday };

  // 2) tabella identificata da una cella "Inserimento Orari": prendi le righe sotto
  for (const name of wb.SheetNames) {
    const rows = toRows(name);
    const idx = rows.findIndex((r) => r.some((c) => norm(c) === TARGET_TABLE));
    if (idx !== -1) {
      return { rows: rows.slice(idx + 1).filter((r) => r.some((c) => c !== "")), month, fileSunday };
    }
  }

  return { rows: [], month, fileSunday };
}


/** Restituisce nero o bianco per il massimo contrasto rispetto al colore esadecimale dato. */
function contrastTextColor(hex: string): string {
  const clean = hex.replace("#", "");
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  const lin = (c: number) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const luminance = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
  return luminance > 0.5 ? "#000000" : "#FFFFFF";
}

function toMinutes(value: string): number | null {
  const m = value.match(/^(\d{1,2})[:.,]?(\d{2})?$/);
  if (!m) return null;
  const h = Number(m[1]);
  const min = m[2] ? Number(m[2]) : 0;
  if (h > 24 || min > 59) return null;
  return h * 60 + min;
}

function formatHours(minutes: number): string {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return m === 0 ? `${h}h` : `${h}h ${String(m).padStart(2, "0")}m`;
}

type DayData = {
  start: string;
  end: string;
  total: string | null;
  late: boolean;
  shortShift: boolean;
} | null;
type WeekData = { monday: Date | undefined; days: DayData[] };

/** Ricostruisce la mappa offset -> turno (Sunday=0, Mon..Sat=1..6, poi +7 per ogni riga) per un nominativo. */
function offsetMapForName(rows: string[][], name: string): Map<number, DayData> {
  const matching = rows.filter((r) => (r[1] ?? "").trim() === name);
  const byOffset = new Map<number, DayData>();
  matching.forEach((row, i) => {
    const base = i * 7;
    byOffset.set(base, blockData(row, DAY_BLOCKS[6]!));
    for (let d = 0; d < 6; d++) {
      byOffset.set(base + 1 + d, blockData(row, DAY_BLOCKS[d]!));
    }
  });
  return byOffset;
}

function blockData(row: string[], cols: number[]): DayData {
  const values: string[] = [];
  for (const c of cols) {
    const v = (row[c - 1] ?? "").trim();
    if (v !== "") values.push(v);
  }
  if (values.length === 0) return null;
  const start = values[0] ?? "";
  const end = values[values.length - 1] ?? "";

  const s = toMinutes(start);
  const e = toMinutes(end);
  let total: string | null = null;
  let shortShift = false;
  if (s !== null && e !== null) {
    let diff = e - s;
    if (diff < 0) diff += 24 * 60;
    if (diff > 8 * 60) diff -= 60;
    const net = Math.max(diff, 0);
    shortShift = net < 7 * 60;
    total = formatHours(net);
  }
  const late = e !== null && e >= 20 * 60;
  return { start, end, total, late, shortShift };
}


function Index() {
  const [rows, setRows] = useState<string[][]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [month, setMonth] = useState<Date | undefined>(undefined);
  const [fileSunday, setFileSunday] = useState<Date | undefined>(undefined);
  const [selected, setSelected] = useState("");
  const [options, setOptions] = useState<string[]>([]);
  const sortedOptions = useMemo(
    () => [...options].sort((a, b) => a.localeCompare(b, "it", { sensitivity: "base" })),
    [options],
  );
  const [manageOpen, setManageOpen] = useState(false);
  const [colorOpen, setColorOpen] = useState(false);
  const [dayPopup, setDayPopup] = useState<{ date: Date; apre: string[]; chiude: string[] } | null>(null);
  const [newOption, setNewOption] = useState("");
  const [accent, setAccent] = useState(DEFAULT_ACCENT);
  const setAccentColor = (c: string) => {
    setAccent(c);
    try {
      localStorage.setItem(ACCENT_KEY, c);
    } catch {
      /* ignore */
    }
  };
  const fileRef = useRef<HTMLInputElement>(null);

  const [hydrated, setHydrated] = useState(false);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setOptions(JSON.parse(raw));
      const acc = localStorage.getItem(ACCENT_KEY);
      if (acc) setAccent(acc);
      const rawRows = localStorage.getItem(ROWS_KEY);
      if (rawRows) setRows(JSON.parse(rawRows));
      const name = localStorage.getItem(FILENAME_KEY);
      if (name) setFileName(name);
      const sel = localStorage.getItem(SELECTED_KEY);
      if (sel) setSelected(sel);
      const mon = localStorage.getItem(MONTH_KEY);
      if (mon) {
        const d = new Date(mon);
        if (!Number.isNaN(d.getTime())) setMonth(startOfMonth(d));
      }
      const fs = localStorage.getItem(FILE_SUNDAY_KEY);
      if (fs) {
        const d = new Date(fs);
        if (!Number.isNaN(d.getTime())) setFileSunday(d);
      }
    } catch {
      /* ignore */
    }
    setHydrated(true);
  }, []);

  // Salva stato per la prossima apertura della PWA
  useEffect(() => {
    if (!hydrated) return;
    try {
      if (rows.length > 0) localStorage.setItem(ROWS_KEY, JSON.stringify(rows));
      else localStorage.removeItem(ROWS_KEY);
      if (fileName) localStorage.setItem(FILENAME_KEY, fileName);
      else localStorage.removeItem(FILENAME_KEY);
      if (selected) localStorage.setItem(SELECTED_KEY, selected);
      else localStorage.removeItem(SELECTED_KEY);
      if (month) localStorage.setItem(MONTH_KEY, month.toISOString());
      else localStorage.removeItem(MONTH_KEY);
      if (fileSunday) localStorage.setItem(FILE_SUNDAY_KEY, fileSunday.toISOString());
      else localStorage.removeItem(FILE_SUNDAY_KEY);
    } catch {
      /* quota superata: ignora */
    }
  }, [hydrated, rows, fileName, selected, month, fileSunday]);

  const persist = (next: string[]) => {
    setOptions(next);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch {
      /* ignore */
    }
  };

  const handleFile = async (file: File) => {
    // svuota la cache offline del file precedente prima di prendere in carico il nuovo
    await clearPreviousFileCache();
    setRows([]);
    setFileName(null);
    const parsed = await parseExcel(file);
    setRows(parsed.rows);
    setMonth(parsed.month);
    setFileSunday(parsed.fileSunday);
    setFileName(file.name);
  };

  // Domenica di riferimento: l'ancora ESATTA letta da SETUP!E7 (anche se ricade nel mese
  // precedente, es. 30/08 per un file di settembre). Se il file non la fornisce, si ricade
  // sul calcolo basato sul mese (giorno prima del primo lunedì della griglia).
  const referenceSunday = useMemo(() => {
    if (fileSunday) return fileSunday;
    if (!month) return undefined;
    return addDays(startOfWeek(startOfMonth(month), { weekStartsOn: 1 }), -1);
  }, [fileSunday, month]);

  // Primo lunedì visualizzato: il giorno subito dopo la domenica di riferimento.
  const gridStart = useMemo(
    () => (referenceSunday ? addDays(referenceSunday, 1) : undefined),
    [referenceSunday],
  );

  const weeks: WeekData[] = useMemo(() => {
    if (!selected || rows.length === 0) return [];
    const matching = rows.filter((r) => (r[1] ?? "").trim() === selected);

    // Each CSV row covers Sunday -> Saturday starting from the reference Sunday.
    // key = day offset from the reference Sunday
    const byOffset = new Map<number, DayData>();
    matching.forEach((row, i) => {
      const base = i * 7;
      // Sunday block (cols 3-6) is the first day of the row
      byOffset.set(base, blockData(row, DAY_BLOCKS[6]!));
      // Monday..Saturday follow the next days
      for (let d = 0; d < 6; d++) {
        byOffset.set(base + 1 + d, blockData(row, DAY_BLOCKS[d]!));
      }
    });

    // Displayed weeks run Monday -> Sunday and cover the selected month.
    const weekCount =
      gridStart && month
        ? differenceInCalendarWeeks(endOfMonth(month), gridStart, { weekStartsOn: 1 }) + 1
        : matching.length + 1;

    // +1 settimana di testa (w=-1): mostra la Domenica di riferimento (offset0, letta da
    // SETUP!E7), che cade il giorno prima del primo lunedì visualizzato e altrimenti non
    // rientrerebbe in nessuna settimana della griglia.
    return Array.from({ length: weekCount + 1 }, (_, j) => {
      const w = j - 1;
      const mondayOffset = w * 7 + 1;
      const monday = gridStart ? addDays(gridStart, w * 7) : undefined;
      const days = Array.from(
        { length: 7 },
        (_, i) => byOffset.get(mondayOffset + i) ?? null,
      );
      return { monday, days };
    });
  }, [rows, selected, month, gridStart]);


  // Mappa offset -> turno per ogni nominativo presente nell'elenco voci (non solo il selezionato).
  const peopleByOffset = useMemo(() => {
    const map = new Map<string, Map<number, DayData>>();
    for (const name of options) {
      map.set(name, offsetMapForName(rows, name));
    }
    return map;
  }, [rows, options]);

  const openDayPopup = (date: Date | undefined) => {
    if (!date || !referenceSunday || rows.length === 0) return;
    const offset = differenceInCalendarDays(date, referenceSunday);
    const apre: string[] = [];
    const chiude: string[] = [];
    for (const name of options) {
      const data = peopleByOffset.get(name)?.get(offset);
      if (!data) continue;
      const s = toMinutes(data.start);
      const e = toMinutes(data.end);
      if (s !== null && s <= 10 * 60) apre.push(name);
      if (e !== null && e >= 20 * 60) chiude.push(name);
    }
    setDayPopup({ date, apre, chiude });
  };



  return (
    <main className="min-h-screen pb-16" style={{ backgroundColor: "#D8D8DE" }}>
      <div className="mx-auto max-w-6xl px-4 pt-6">
        {/* Titolo + toolbar + giorno in testata (sticky, tutto opaco) */}
        <section className="sticky top-0 z-20" style={{ backgroundColor: "#D8D8DE" }}>
          <header className="mb-4">
            <h1 className="text-3xl font-bold tracking-tight text-foreground">
              {month
                ? format(month, "MMMM yyyy", { locale: it }).replace(/^./, (c) => c.toUpperCase())
                : "Calendario Turni"}
            </h1>
          </header>

          <div className="rounded-2xl border border-border/60 bg-card p-3 shadow-sm">
            <div className="grid grid-cols-[auto_1fr] gap-2">
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) void handleFile(f);
                  e.target.value = "";
                }}
              />
              <Button
                className="w-full gap-1.5 whitespace-nowrap rounded-xl px-3 text-sm text-white hover:opacity-90"
                style={{ backgroundColor: "#003335" }}
                onClick={() => fileRef.current?.click()}
                title="Carica file Excel"
              >
                <Upload className="size-4 shrink-0" />
                <span>Carica file Excel</span>
              </Button>
              <div className="flex gap-2">
                <Select
                  value={selected}
                  onValueChange={(value) => {
                    if (value === "__add__") {
                      setManageOpen(true);
                      return;
                    }
                    setSelected(value);
                  }}
                >
                  <SelectTrigger
                    className="h-9 w-full min-w-0 flex-1 rounded-xl px-2 text-sm [&>span]:truncate"
                    aria-label="Filtra per voce"
                  >
                    <Users className="size-4 shrink-0 text-muted-foreground" />
                    <SelectValue placeholder="Nome" />
                  </SelectTrigger>
                  <SelectContent>
                    {sortedOptions.map((o) => (
                      <SelectItem key={o} value={o}>
                        {o}
                      </SelectItem>
                    ))}
                    <SelectSeparator />
                    <SelectItem value="__add__">Aggiungi nome</SelectItem>
                  </SelectContent>
                </Select>
                <Popover open={colorOpen} onOpenChange={setColorOpen}>
                  <PopoverTrigger asChild>
                    <button
                      type="button"
                      aria-label="Scegli colore"
                      className="size-9 shrink-0 rounded-xl border border-border shadow-sm"
                      style={{ backgroundColor: accent }}
                    />
                  </PopoverTrigger>
                  <PopoverContent align="end" className="w-auto rounded-2xl p-3">
                    <div className="grid grid-cols-5 gap-2">
                      {PALETTE.map((c) => (
                        <button
                          key={c}
                          type="button"
                          aria-label={`Colore ${c}`}
                          onClick={() => {
                            setAccentColor(c);
                            setColorOpen(false);
                          }}
                          className={cn(
                            "size-8 rounded-full border transition-transform hover:scale-110",
                            c === accent ? "border-foreground ring-2 ring-foreground/40" : "border-border",
                          )}
                          style={{ backgroundColor: c }}
                        />
                      ))}
                    </div>
                  </PopoverContent>
                </Popover>
              </div>
            </div>


            {fileName && (
              <p className="mt-2 truncate text-xs text-muted-foreground">File caricato: {fileName}</p>
            )}
          </div>
          {weeks.length > 0 && (
            <div
              className="mt-2 grid grid-cols-7 gap-1 px-1 py-2 sm:gap-2"
              style={{ backgroundColor: "#D8D8DE" }}
            >
              {DAY_LABELS.map((label) => (
                <div
                  key={label}
                  className="truncate text-center text-[13px] font-extrabold uppercase tracking-wide text-foreground sm:text-base"
                >
                  {label.slice(0, 3)}
                </div>
              ))}
            </div>
          )}
        </section>

        {/* Month: one week per row, ciascuna settimana inscritta in una cornice */}
        <section className="mt-3 space-y-2">
          {weeks.map((week, weekIdx) => (
            <div
              key={weekIdx}
              className="rounded-2xl border border-border/60 p-1.5 sm:rounded-3xl sm:p-2"
              style={{ backgroundColor: "#ECECF0" }}
            >
              <div className="grid grid-cols-7 gap-1 sm:gap-2">
                {DAY_LABELS.map((label, i) => {
                  const date = week.monday ? addDays(week.monday, i) : undefined;
                  const data = week.days[i];
                  const today = date ? isToday(date) : false;
                  return (
                    <div
                      key={`${weekIdx}-${label}`}
                      className="relative h-24 cursor-pointer sm:h-[148px]"
                      onClick={() => openDayPopup(date)}
                    >
                      {/* Data: posizione fissa, indipendente dall'altezza della tessera */}
                      <div className="absolute inset-x-0 top-2 z-10 pb-1 text-center">
                        <div
                          className={cn(
                            "block truncate rounded-full px-2 text-[12px] sm:text-[16px]",
                            today && !data && "border-2 border-foreground",
                          )}
                          style={{ color: "#5c5c5c" }}
                        >
                          {date ? format(date, "d", { locale: it }) : "—"}
                        </div>
                      </div>
                      {data && (
                        <>
                          <article
                            className={cn(
                              "absolute inset-x-0 bottom-0 flex flex-col items-center justify-end rounded-xl bg-card p-1 shadow-sm sm:rounded-2xl sm:p-2",
                              today ? "border-2 border-foreground" : "border border-border/60",
                            )}
                            style={{ height: data.shortShift ? "70%" : "100%" }}
                          >
                            {data.total && (
                              <span
                                className={cn(
                                  "w-full rounded-md px-1 py-[2px] text-center text-[10px] font-medium leading-none sm:rounded-xl sm:px-2 sm:py-[3px] sm:text-xs",
                                  data.late ? undefined : "bg-secondary text-foreground",
                                )}
                                style={data.late ? { backgroundColor: accent, color: contrastTextColor(accent) } : undefined}
                              >
                                {data.total}
                              </span>
                            )}
                          </article>
                          {/* Orari: posizione fissa, poco sotto al bordo superiore della tessera bassa */}
                          <div className="absolute inset-x-0 top-[36px] z-10 flex flex-col items-center gap-0.5 px-1 text-center sm:top-[55px] sm:gap-1 sm:px-2">
                            <span className="text-[12px] font-bold leading-tight text-foreground sm:text-base">
                              {data.start}
                            </span>
                            <span className="text-[12px] font-bold leading-tight text-foreground sm:text-base">
                              {data.end}
                            </span>
                          </div>
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </section>

        {rows.length === 0 && (
          <p className="mt-6 text-center text-sm text-muted-foreground">
            Carica il file Excel degli orari e seleziona un nome per iniziare.
          </p>
        )}
      </div>

      <Dialog open={!!dayPopup} onOpenChange={(o) => !o && setDayPopup(null)}>
        <DialogContent className="w-fit rounded-2xl" showCloseButton={false}>
          <DialogHeader>
            <DialogTitle className="whitespace-nowrap text-sm font-normal">
              {dayPopup
                ? format(dayPopup.date, "EEEE d MMMM", { locale: it }).replace(/^./, (c) => c.toUpperCase())
                : ""}
            </DialogTitle>
          </DialogHeader>

          <div>
            <h3 className="mb-1.5 text-lg font-semibold text-foreground">Chi apre</h3>
            {dayPopup && dayPopup.apre.length > 0 ? (
              <ul className="mb-4 space-y-1">
                {dayPopup.apre.map((n) => (
                  <li key={n} className="rounded-xl bg-secondary px-3 py-2 text-sm text-foreground">
                    {n}
                  </li>
                ))}
              </ul>
            ) : (
              <p className="mb-4 text-sm text-muted-foreground">Nessuno.</p>
            )}

            <h3 className="mb-1.5 text-lg font-semibold text-foreground">Chi chiude</h3>
            {dayPopup && dayPopup.chiude.length > 0 ? (
              <ul className="space-y-1">
                {dayPopup.chiude.map((n) => (
                  <li key={n} className="rounded-xl bg-secondary px-3 py-2 text-sm text-foreground">
                    {n}
                  </li>
                ))}
              </ul>
            ) : (
              <p className="text-sm text-muted-foreground">Nessuno.</p>
            )}
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={manageOpen} onOpenChange={setManageOpen}>
        <DialogContent className="rounded-2xl sm:max-w-md">
          <DialogHeader>
            <DialogTitle>Gestisci voci</DialogTitle>
            <DialogDescription>
              Aggiungi o elimina le voci disponibili nel menu a tendina.
            </DialogDescription>
          </DialogHeader>
          <form
            className="flex gap-2"
            onSubmit={(e) => {
              e.preventDefault();
              const v = newOption.trim();
              if (v && !options.includes(v)) {
                persist([...options, v]);
                setSelected(v);
                setManageOpen(false);
              }
              setNewOption("");
            }}
          >
            <Input
              value={newOption}
              onChange={(e) => setNewOption(e.target.value)}
              placeholder="Nuova voce"
              className="rounded-xl"
            />
            <Button type="submit" className="rounded-xl">
              Aggiungi
            </Button>
          </form>
          <ul className="max-h-64 space-y-1 overflow-y-auto">
            {sortedOptions.length === 0 && (
              <li className="py-4 text-center text-sm text-muted-foreground">Nessuna voce.</li>
            )}
            {sortedOptions.map((o) => (
              <li
                key={o}
                className="flex items-center justify-between rounded-xl bg-secondary px-3 py-2"
              >
                <span className="truncate text-sm text-foreground">{o}</span>
                <button
                  type="button"
                  aria-label={`Elimina ${o}`}
                  className="text-muted-foreground transition-colors hover:text-destructive"
                  onClick={() => {
                    persist(options.filter((x) => x !== o));
                    setSelected("");
                  }}
                >
                  <Trash2 className="size-4" />
                </button>
              </li>
            ))}
          </ul>
          <Button variant="outline" className="rounded-xl" onClick={() => setManageOpen(false)}>
            <X className="mr-1 size-4" />
            Chiudi
          </Button>
        </DialogContent>
      </Dialog>
    </main>
  );
}
