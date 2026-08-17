"use client";

import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { QRCodeSVG } from "qrcode.react";
import {
  ArrowRight,
  Check,
  Copy,
  Download,
  FileSpreadsheet,
  Heart,
  Layers3,
  MapPin,
  ShieldCheck,
  Sparkles,
  UploadCloud,
  X,
  Zap,
} from "lucide-react";

type Cell = string | number | boolean | Date | null | undefined;
type SheetRow = Cell[];

type SequenceInfo = {
  label: string;
  numeric: number | null;
  addNumber: number | null;
};

type GroupItem = {
  row: SheetRow;
  sourceIndex: number;
  sequence: SequenceInfo;
};

type AddressGroup = {
  key: string;
  street: string;
  number: string;
  zip: string;
  city: string;
  rows: GroupItem[];
  sequences: SequenceInfo[];
  firstSourceIndex: number;
  minNumeric?: number | null;
  minAdd?: number | null;
  finalComplement?: string;
};

function emvField(id: string, value: string) {
  return `${id}${String(value.length).padStart(2, "0")}${value}`;
}

function crc16Ccitt(value: string) {
  let crc = 0xffff;

  for (let index = 0; index < value.length; index += 1) {
    crc ^= value.charCodeAt(index) << 8;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 0x8000) !== 0 ? ((crc << 1) ^ 0x1021) & 0xffff : (crc << 1) & 0xffff;
    }
  }

  return crc.toString(16).toUpperCase().padStart(4, "0");
}

function createPixPayload(key: string) {
  const merchantAccount = emvField("00", "BR.GOV.BCB.PIX") + emvField("01", key);
  const additionalData = emvField("05", "***");
  const payload = [
    emvField("00", "01"),
    emvField("26", merchantAccount),
    emvField("52", "0000"),
    emvField("53", "986"),
    emvField("58", "BR"),
    emvField("59", "RAIMUNDO JUNIOR"),
    emvField("60", "FORTALEZA"),
    emvField("62", additionalData),
    "6304",
  ].join("");

  return `${payload}${crc16Ccitt(payload)}`;
}

type ProcessedFile = {
  originalName: string;
  outputName: string;
  headers: string[];
  rows: SheetRow[];
  sourcePackages: number;
  consolidatedStops: number;
  removedDuplicates: number;
  addPackages: number;
  mergedGroups: number;
  differentComplements: number;
  preview: Array<{ sequence: string; address: string; complement: string }>;
};

function plain(value: Cell) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function normalizeStreet(value: Cell) {
  const text = plain(value)
    .replace(/[.]/g, " ")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const expansions: Record<string, string> = {
    r: "rua",
    av: "avenida",
    avda: "avenida",
    prof: "professor",
    profa: "professora",
    des: "desembargador",
    cel: "coronel",
    gen: "general",
  };
  const tokens = text.split(" ");
  if (tokens[0] && expansions[tokens[0]]) tokens[0] = expansions[tokens[0]];
  return tokens.join(" ");
}

function parseAddress(address: Cell) {
  const parts = String(address ?? "")
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  const street = normalizeStreet(parts[0] ?? "");
  let number = "";
  const numberMatch = plain(parts[1] ?? "").match(/\b(\d+[a-z]?)\b/);
  if (numberMatch) number = numberMatch[1];
  if (!number) {
    const fallback = plain(address).match(/\b(\d+[a-z]?)\b/);
    if (fallback) number = fallback[1];
  }
  return { street, number };
}

function extractComplement(address: Cell) {
  const parts = String(address ?? "").split(",");
  return parts.length >= 3 ? parts[2].trim() : "";
}

function cleanDestinationAddress(address: Cell) {
  const parts = String(address ?? "").split(",").map((part) => part.trim());
  return parts.length >= 2 ? `${parts[0]}, ${parts[1]}` : String(address ?? "").trim();
}

function normalizeComplement(value: Cell) {
  return plain(value)
    .replace(/[.]/g, " ")
    .replace(/\b(apartamento|apto|apt)\b/g, "ap")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function levenshtein(a: string, b: string) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i += 1) {
    const current = [i];
    for (let j = 1; j <= b.length; j += 1) {
      current[j] = Math.min(
        current[j - 1] + 1,
        previous[j] + 1,
        previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[b.length];
}

function similarity(a: string, b: string) {
  const longest = Math.max(a.length, b.length);
  return longest === 0 ? 1 : 1 - levenshtein(a, b) / longest;
}

function sortSequences(a: SequenceInfo, b: SequenceInfo) {
  if (a.numeric !== null && b.numeric !== null) return a.numeric - b.numeric;
  if (a.numeric !== null) return -1;
  if (b.numeric !== null) return 1;
  if (a.addNumber !== null && b.addNumber !== null) return a.addNumber - b.addNumber;
  if (a.addNumber !== null) return 1;
  if (b.addNumber !== null) return -1;
  return a.label.localeCompare(b.label, "pt-BR");
}

function sequenceInfo(value: Cell, addNumber: number): SequenceInfo {
  const text = String(value ?? "").trim();
  if (text === "-" || text === "") {
    return { label: `ADD +${addNumber}`, numeric: null, addNumber };
  }
  const numeric = Number(text);
  return Number.isFinite(numeric)
    ? { label: String(numeric), numeric, addNumber: null }
    : { label: text, numeric: null, addNumber: null };
}

function normalizedHeader(value: Cell) {
  return plain(value).replace(/[^a-z0-9]/g, "");
}

function findHeader(headers: string[], aliases: string[]) {
  const normalized = headers.map(normalizedHeader);
  return normalized.findIndex((header) => aliases.some((alias) => header === alias || header.includes(alias)));
}

function processWorkbook(fileName: string, data: ArrayBuffer): ProcessedFile {
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) throw new Error("A planilha não possui nenhuma aba para processar.");
  const matrix = XLSX.utils.sheet_to_json<SheetRow>(workbook.Sheets[firstSheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
  if (matrix.length < 2) throw new Error("A planilha está vazia ou não possui linhas de entrega.");

  const headers = matrix[0].map((value) => String(value ?? "").trim());
  const sequenceIndex = findHeader(headers, ["sequence", "sequencia"]);
  const addressIndex = findHeader(headers, ["destinationaddress", "enderecodestino"]);
  const cityIndex = findHeader(headers, ["city", "cidade"]);
  const zipIndex = findHeader(headers, ["zipcodecep", "zipcodepostalcode", "zipcode", "cep"]);
  if (sequenceIndex < 0 || addressIndex < 0) {
    throw new Error('Não encontrei as colunas obrigatórias "Sequence" e "Destination Address".');
  }

  const sourceRows = matrix
    .slice(1)
    .filter((row) => row.some((value) => value !== null && value !== ""))
    .map((row) => Array.from({ length: headers.length }, (_, index) => row[index] ?? ""));
  if (!sourceRows.length) throw new Error("Não encontrei pacotes para processar.");

  const groups = new Map<string, AddressGroup>();
  let addCounter = 0;
  for (let sourceIndex = 0; sourceIndex < sourceRows.length; sourceIndex += 1) {
    const row = sourceRows[sourceIndex];
    const isAdd = ["", "-"].includes(String(row[sequenceIndex] ?? "").trim());
    if (isAdd) addCounter += 1;
    const sequence = sequenceInfo(row[sequenceIndex], addCounter);
    const parsed = parseAddress(row[addressIndex]);
    const zip = zipIndex >= 0 ? plain(row[zipIndex]).replace(/\D/g, "") : "";
    const city = cityIndex >= 0 ? plain(row[cityIndex]) : "";
    const exactKey = `${city}|${parsed.street}|${parsed.number}`;
    let key = exactKey;

    if (parsed.street && parsed.number && !groups.has(exactKey)) {
      let best: { key: string; score: number } | null = null;
      for (const [candidateKey, group] of groups) {
        if (group.city !== city || group.number !== parsed.number) continue;
        if (zip && group.zip && zip !== group.zip) continue;
        const score = similarity(parsed.street, group.street);
        if (score >= 0.88 && (!best || score > best.score)) best = { key: candidateKey, score };
      }
      if (best) key = best.key;
    }

    if (!groups.has(key)) {
      groups.set(key, {
        key,
        street: parsed.street,
        number: parsed.number,
        zip,
        city,
        rows: [],
        sequences: [],
        firstSourceIndex: sourceIndex,
      });
    }
    const group = groups.get(key)!;
    group.rows.push({ row, sourceIndex, sequence });
    group.sequences.push(sequence);
  }

  const consolidatedGroups = Array.from(groups.values());
  let differentComplements = 0;
  for (const group of consolidatedGroups) {
    group.sequences.sort(sortSequences);
    group.minNumeric = group.sequences.find((sequence) => sequence.numeric !== null)?.numeric ?? null;
    group.minAdd = group.sequences.find((sequence) => sequence.addNumber !== null)?.addNumber ?? null;
    const complements = new Map<string, { complement: string; sequences: SequenceInfo[]; first: number }>();
    for (const item of group.rows) {
      const complement = extractComplement(item.row[addressIndex]);
      const normalized = normalizeComplement(complement);
      if (!complements.has(normalized)) {
        complements.set(normalized, { complement, sequences: [], first: item.sourceIndex });
      }
      complements.get(normalized)!.sequences.push(item.sequence);
    }
    const complementGroups = Array.from(complements.values()).sort((a, b) => a.first - b.first);
    if (group.rows.length > 1 && complementGroups.length > 1) {
      differentComplements += 1;
      group.finalComplement = complementGroups
        .map((entry) => {
          const labels = entry.sequences.slice().sort(sortSequences).map((sequence) => sequence.label).join(",");
          return `${labels}: ${entry.complement || "Sem complemento"}`;
        })
        .join(" / ");
    } else {
      group.finalComplement = complementGroups[0]?.complement ?? "";
    }
  }

  consolidatedGroups.sort((a, b) => {
    const aOnlyAdd = a.minNumeric === null;
    const bOnlyAdd = b.minNumeric === null;
    if (aOnlyAdd !== bOnlyAdd) return aOnlyAdd ? -1 : 1;
    if (!aOnlyAdd) return (a.minNumeric ?? 0) - (b.minNumeric ?? 0);
    return (a.minAdd ?? Number.POSITIVE_INFINITY) - (b.minAdd ?? Number.POSITIVE_INFINITY);
  });

  const outputHeaders = [...headers.slice(0, addressIndex + 1), "Complemento", ...headers.slice(addressIndex + 1)];
  const outputRows = consolidatedGroups.map((group) => {
    const regular = group.rows.find((item) => item.sequence.numeric !== null);
    const representative = [...(regular ?? group.rows[0]).row];
    representative[sequenceIndex] = group.sequences.map((sequence) => sequence.label).join(", ");
    representative[addressIndex] = cleanDestinationAddress(representative[addressIndex]);
    return [
      ...representative.slice(0, addressIndex + 1),
      group.finalComplement ?? "",
      ...representative.slice(addressIndex + 1),
    ];
  });

  const baseName = fileName.replace(/\.(xlsx|xls)$/i, "");
  return {
    originalName: fileName,
    outputName: `${baseName} - Rota Organizada.xlsx`,
    headers: outputHeaders,
    rows: outputRows,
    sourcePackages: sourceRows.length,
    consolidatedStops: outputRows.length,
    removedDuplicates: sourceRows.length - outputRows.length,
    addPackages: addCounter,
    mergedGroups: consolidatedGroups.filter((group) => group.rows.length > 1).length,
    differentComplements,
    preview: consolidatedGroups.slice(0, 5).map((group, index) => ({
      sequence: String(outputRows[index][sequenceIndex] ?? ""),
      address: String(outputRows[index][addressIndex] ?? ""),
      complement: String(group.finalComplement ?? ""),
    })),
  };
}

function downloadProcessed(result: ProcessedFile) {
  const sheet = XLSX.utils.aoa_to_sheet([result.headers, ...result.rows]);
  sheet["!autofilter"] = { ref: XLSX.utils.encode_range(sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }) };
  sheet["!cols"] = result.headers.map((header) => {
    const normalized = normalizedHeader(header);
    if (normalized === "destinationaddress") return { wch: 38 };
    if (normalized === "complemento") return { wch: 54 };
    if (normalized === "sequence") return { wch: 24 };
    return { wch: Math.max(12, Math.min(24, header.length + 4)) };
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Entregas organizadas");
  XLSX.writeFile(workbook, result.outputName, { compression: true });
}

export default function Home() {
  const pixKey = "e0bb8e62-76d1-4a62-ba37-ac344c0a98ff";
  const pixPayload = createPixPayload(pixKey);
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [result, setResult] = useState<ProcessedFile | null>(null);
  const [error, setError] = useState("");
  const [pixCopied, setPixCopied] = useState(false);
  const [supportPrompt, setSupportPrompt] = useState<"processed" | null>(null);
  const hasSeenProcessedPrompt = useRef(false);

  useEffect(() => {
    if (!supportPrompt) return;

    function closeWithEscape(event: KeyboardEvent) {
      if (event.key === "Escape") setSupportPrompt(null);
    }

    window.addEventListener("keydown", closeWithEscape);
    return () => window.removeEventListener("keydown", closeWithEscape);
  }, [supportPrompt]);

  function goToSupport() {
    setSupportPrompt(null);
    window.requestAnimationFrame(() => {
      document.getElementById("apoie")?.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  }

  async function copyPixKey() {
    try {
      await navigator.clipboard.writeText(pixKey);
    } catch {
      const textArea = document.createElement("textarea");
      textArea.value = pixKey;
      textArea.style.position = "fixed";
      textArea.style.opacity = "0";
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      textArea.remove();
    }
    setPixCopied(true);
    window.setTimeout(() => setPixCopied(false), 2500);
  }

  async function handleFile(file?: File) {
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setError("Envie uma planilha no formato .xlsx ou .xls.");
      return;
    }
    setError("");
    setResult(null);
    setIsProcessing(true);
    try {
      const data = await file.arrayBuffer();
      const processedFile = processWorkbook(file.name, data);
      setResult(processedFile);
      window.setTimeout(() => {
        setSupportPrompt((current) => {
          if (current || hasSeenProcessedPrompt.current) return current;
          hasSeenProcessedPrompt.current = true;
          return "processed";
        });
      }, 900);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Não foi possível processar esta planilha.");
    } finally {
      setIsProcessing(false);
    }
  }

  return (
    <main>
      <header className="site-header">
        <a className="brand" href="#inicio" aria-label="Ir para o início">
          <span className="brand-mark"><MapPin size={20} strokeWidth={2.5} /></span>
          <span>Rota Fácil</span>
        </a>
        <nav aria-label="Navegação principal">
          <a className="nav-how" href="#como-funciona">Como funciona</a>
          <a className="nav-support" href="#apoie">
            <Heart size={15} fill="currentColor" />
            <span>Apoie</span>
          </a>
          <a className="nav-cta" href="#ferramenta">Organizar planilha</a>
        </nav>
      </header>

      <section className="hero" id="inicio">
        <div className="hero-copy">
          <div className="eyebrow"><Sparkles size={15} /> Feito para quem vive a rota</div>
          <h1>Menos planilha.<br /><span>Mais entregas.</span></h1>
          <p>
            Organize sua rota da Shopee em segundos. Endereços repetidos são agrupados,
            sequências são consolidadas e complementos ficam claros antes de importar no Spoke.
          </p>
          <div className="hero-actions">
            <a className="button button-primary" href="#ferramenta">
              Organizar minha planilha <ArrowRight size={18} />
            </a>
            <a className="text-link" href="#como-funciona">Entender como funciona</a>
          </div>
          <div className="trust-line">
            <span><ShieldCheck size={17} /> Processamento no seu dispositivo</span>
            <span><Zap size={17} /> Pronto em segundos</span>
          </div>
        </div>

        <div className="hero-visual" aria-label="Exemplo de consolidação de entregas">
          <div className="route-card route-card-back">
            <span>ANTES</span>
            <div><i>60</i><p>Rua Maria José, 123<br /><small>AP 106</small></p></div>
            <div><i>61</i><p>R Maria José, 123<br /><small>AP 313</small></p></div>
            <div><i>62</i><p>Rua Maria José, 123<br /><small>AP 313</small></p></div>
          </div>
          <div className="route-card route-card-front">
            <div className="card-label"><Check size={16} /> ROTA ORGANIZADA</div>
            <div className="route-pin"><MapPin size={25} fill="currentColor" /></div>
            <strong>Rua Maria José, 123</strong>
            <p className="sequence-pill">60, 61, 62</p>
            <p className="complement-preview"><b>60:</b> AP 106 <span>/</span> <b>61, 62:</b> AP 313</p>
          </div>
          <div className="saved-badge"><span>−2</span> paradas duplicadas</div>
        </div>
      </section>

      <section className="tool-section" id="ferramenta">
        <div className="section-heading centered">
          <div className="eyebrow"><FileSpreadsheet size={15} /> Sua rota começa aqui</div>
          <h2>Carregue. Organize. Entregue.</h2>
          <p>A planilha original não é alterada. Você baixa uma nova versão pronta para usar.</p>
        </div>

        <div className={`upload-shell ${result ? "has-result" : ""}`}>
          {!result ? (
            <div
              className={`dropzone ${isDragging ? "dragging" : ""}`}
              onDragEnter={(event) => { event.preventDefault(); setIsDragging(true); }}
              onDragOver={(event) => event.preventDefault()}
              onDragLeave={(event) => { event.preventDefault(); setIsDragging(false); }}
              onDrop={(event) => {
                event.preventDefault();
                setIsDragging(false);
                void handleFile(event.dataTransfer.files[0]);
              }}
            >
              <div className="upload-icon"><UploadCloud size={34} /></div>
              <h3>{isProcessing ? "Organizando sua rota..." : "Arraste sua planilha aqui"}</h3>
              <p>ou escolha o arquivo recebido da Shopee</p>
              <button className="button button-dark" onClick={() => inputRef.current?.click()} disabled={isProcessing}>
                <FileSpreadsheet size={18} /> Selecionar planilha
              </button>
              <input
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls"
                hidden
                onChange={(event) => void handleFile(event.target.files?.[0])}
              />
              <small>Formatos aceitos: XLSX ou XLS · Seus dados não saem do navegador</small>
            </div>
          ) : (
            <div className="result-panel">
              <div className="success-title">
                <div className="success-icon"><Check size={28} /></div>
                <div>
                  <span>Planilha organizada com sucesso</span>
                  <h3>{result.originalName}</h3>
                </div>
              </div>
              <div className="stats-grid">
                <div><strong>{result.sourcePackages}</strong><span>pacotes recebidos</span></div>
                <div><strong>{result.consolidatedStops}</strong><span>paradas finais</span></div>
                <div><strong>{result.removedDuplicates}</strong><span>duplicidades removidas</span></div>
                <div><strong>{result.addPackages}</strong><span>pacotes adicionados</span></div>
              </div>
              <div className="preview-table" role="region" aria-label="Prévia dos primeiros endereços" tabIndex={0}>
                <div className="preview-row preview-head"><span>Sequência</span><span>Endereço</span><span>Complemento</span></div>
                {result.preview.map((row, index) => (
                  <div className="preview-row" key={`${row.sequence}-${index}`}>
                    <span>{row.sequence}</span><span>{row.address}</span><span>{row.complement || "—"}</span>
                  </div>
                ))}
              </div>
              <div className="result-actions">
                <button className="button button-primary" onClick={() => downloadProcessed(result)}>
                  <Download size={18} /> Baixar planilha organizada
                </button>
                <button className="button button-ghost" onClick={() => { setResult(null); setError(""); if (inputRef.current) inputRef.current.value = ""; }}>
                  Processar outra
                </button>
              </div>
            </div>
          )}
          {error && <div className="error-message" role="alert">{error}</div>}
        </div>
      </section>

      <section className="how-section" id="como-funciona">
        <div className="section-heading">
          <div className="eyebrow"><Layers3 size={15} /> Simples de verdade</div>
          <h2>Como funciona?</h2>
          <p>Três passos entre a planilha original e uma rota muito mais limpa.</p>
        </div>
        <div className="steps-grid">
          <article>
            <span className="step-number">01</span>
            <div className="step-icon"><UploadCloud /></div>
            <h3>Carregue a planilha</h3>
            <p>Selecione o arquivo XLSX ou XLS recebido com as entregas da sua rota.</p>
          </article>
          <article>
            <span className="step-number">02</span>
            <div className="step-icon"><Sparkles /></div>
            <h3>Deixe o sistema organizar</h3>
            <p>Identificamos variações do mesmo endereço, agrupamos sequências, numeramos pacotes adicionados e separamos complementos.</p>
          </article>
          <article>
            <span className="step-number">03</span>
            <div className="step-icon"><Download /></div>
            <h3>Baixe e importe</h3>
            <p>Faça o download da nova planilha e importe no Spoke/Circuit para montar sua rota.</p>
          </article>
        </div>
        <div className="rules-card">
          <div>
            <span className="mini-icon"><MapPin /></span>
            <p><strong>Endereços equivalentes</strong>“Rua” e “R.” são reconhecidos como o mesmo local.</p>
          </div>
          <div>
            <span className="mini-icon"><Layers3 /></span>
            <p><strong>Apartamentos organizados</strong>Sequência e complemento aparecem juntos quando as unidades são diferentes.</p>
          </div>
          <div>
            <span className="mini-icon"><Zap /></span>
            <p><strong>Pacotes adicionados</strong>Registros sem sequência recebem ADD +1, ADD +2 e assim por diante.</p>
          </div>
        </div>
      </section>

      <section className="support-section" id="apoie">
        <div className="support-card">
          <div className="support-copy">
            <div className="eyebrow light"><Heart size={15} fill="currentColor" /> Apoie o projeto</div>
            <h2>Essa ferramenta poupou seu tempo?</h2>
            <p>
              O projeto nasceu de um problema real de quem trabalha na rua. Se ele ajudou sua rota,
              você pode contribuir com qualquer valor para manter e melhorar a ferramenta.
            </p>
          </div>
          <div className="pix-card">
            <span>PIX DO DESENVOLVEDOR</span>
            <div className="pix-desktop">
              <div className="pix-qr-shell">
                <QRCodeSVG
                  value={pixPayload}
                  size={176}
                  level="M"
                  marginSize={1}
                  title="QR Code Pix para apoiar o projeto"
                />
              </div>
              <div className="pix-qr-copy">
                <strong>Escaneie com o app do seu banco</strong>
                <p>Aponte a câmera do celular para o QR Code e escolha o valor da contribuição.</p>
              </div>
            </div>
            <div className="pix-mobile">
              <strong className="pix-key">{pixKey}</strong>
              <button type="button" onClick={copyPixKey} className={pixCopied ? "copied" : ""}>
                {pixCopied ? <Check size={17} /> : <Copy size={17} />}
                <span aria-live="polite">{pixCopied ? "Chave Pix copiada!" : "Copiar chave Pix"}</span>
              </button>
            </div>
            <small className="pix-note">Contribuição opcional. A ferramenta continua gratuita.</small>
          </div>
        </div>
      </section>

      <footer>
        <a className="brand footer-brand" href="#inicio">
          <span className="brand-mark"><MapPin size={20} strokeWidth={2.5} /></span>
          <span>Rota Fácil</span>
        </a>
        <p className="developer-credit">
          Desenvolvido por{" "}
          <a href="https://github.com/juniorod99" target="_blank" rel="noreferrer">
            Raimundo Junior
          </a>
        </p>
        <span>Projeto independente · Não afiliado à Shopee ou ao Spoke</span>
      </footer>

      {supportPrompt && (
        <div
          className="support-popup-backdrop"
          role="presentation"
          onMouseDown={() => setSupportPrompt(null)}
        >
          <section
            className="support-popup"
            role="dialog"
            aria-modal="true"
            aria-labelledby="support-popup-title"
            aria-describedby="support-popup-description"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <button
              className="support-popup-close"
              type="button"
              aria-label="Fechar convite de apoio"
              onClick={() => setSupportPrompt(null)}
            >
              <X size={20} />
            </button>
            <div className="support-popup-icon"><Heart size={25} fill="currentColor" /></div>
            <span className="support-popup-eyebrow">Sua planilha está pronta!</span>
            <h2 id="support-popup-title">Sua rota ficou mais simples?</h2>
            <p id="support-popup-description">
              Se a ferramenta poupou seu tempo, você pode apoiar a evolução do projeto com qualquer valor.
            </p>
            <div className="support-popup-actions">
              <button className="button button-primary" type="button" onClick={goToSupport} autoFocus>
                <Heart size={17} fill="currentColor" /> Apoiar com Pix
              </button>
              <button className="support-popup-later" type="button" onClick={() => setSupportPrompt(null)}>
                Agora não
              </button>
            </div>
            <small>É totalmente opcional. Qualquer valor já ajuda.</small>
          </section>
        </div>
      )}
    </main>
  );
}
