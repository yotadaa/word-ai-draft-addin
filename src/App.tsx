import {
  BadgeCheck,
  Check,
  ChevronLeft,
  ChevronRight,
  ClipboardCheck,
  FileSearch,
  Image as ImageIcon,
  Link2,
  Loader2,
  Minimize2,
  MousePointer2,
  PencilLine,
  RefreshCcw,
  ScanText,
  Sparkles,
  StretchHorizontal,
  Table2
} from "lucide-react";
import { useCallback, useEffect, useMemo, useState, type ReactNode } from "react";
import LiquidGlass from "liquid-glass-react";
import { requestRewrite } from "./api";
import {
  applyAllCaptionCandidates,
  applyAllPageSplitFixes,
  applyAllReferenceCandidates,
  applyCaptionCandidate,
  applyPageSplitFix,
  applyReferenceCandidate,
  CaptionCandidate,
  countWords,
  createPreviewSnapshot,
  DraftMode,
  PageSplitCandidate,
  ReferenceCandidate,
  readSelectionSnapshot,
  replaceSelectionWith,
  scanCaptionCandidates,
  scanPageSplitCandidates,
  scanReferenceCandidates,
  selectCaptionCandidate,
  selectPageSplitCandidate,
  selectReferenceCandidate,
  SelectionSnapshot,
  waitForOffice
} from "./office";

const previewSeed =
  "Algoritma genetika digunakan untuk mencari solusi penjadwalan praktikum dengan mempertimbangkan batasan ruang, waktu, dosen, dan kelompok praktikum.";

const modeOptions: Array<{
  id: DraftMode;
  label: string;
  icon: typeof StretchHorizontal;
}> = [
  { id: "extend", label: "Extend", icon: StretchHorizontal },
  { id: "shrink", label: "Shrink", icon: Minimize2 },
  { id: "rewrite", label: "Rewrite", icon: PencilLine }
];

type GlassPanelProps = {
  children: ReactNode;
  className?: string;
  radius?: number;
  padding?: string;
};

const glassDefaults = {
  displacementScale: 28,
  blurAmount: 0.09,
  saturation: 142,
  aberrationIntensity: 0.65,
  elasticity: 0.12,
  mode: "standard" as const,
  overLight: true
};

function GlassPanel({ children, className = "", radius = 12, padding = "0" }: GlassPanelProps) {
  if (typeof Office !== "undefined") {
    return (
      <div className={`native-glass ${className}`.trim()} style={{ borderRadius: radius, padding }}>
        {children}
      </div>
    );
  }

  return (
    <LiquidGlass
      {...glassDefaults}
      className={`glass-shell ${className}`.trim()}
      cornerRadius={radius}
      padding={padding}
    >
      {children}
    </LiquidGlass>
  );
}

export function App() {
  const [officeReady, setOfficeReady] = useState(false);
  const [checkingOffice, setCheckingOffice] = useState(true);
  const [mode, setMode] = useState<DraftMode>("extend");
  const [targetWords, setTargetWords] = useState(90);
  const [contextRadius, setContextRadius] = useState(2);
  const [snapshot, setSnapshot] = useState<SelectionSnapshot | null>(null);
  const [previewText, setPreviewText] = useState(previewSeed);
  const [result, setResult] = useState("");
  const [resultHtml, setResultHtml] = useState("");
  const [italicTerms, setItalicTerms] = useState<string[]>([]);
  const [model, setModel] = useState("");
  const [status, setStatus] = useState("Ready");
  const [error, setError] = useState("");
  const [isScanning, setIsScanning] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isReplacing, setIsReplacing] = useState(false);
  const [captionCandidates, setCaptionCandidates] = useState<CaptionCandidate[]>([]);
  const [captionIndex, setCaptionIndex] = useState(0);
  const [captionStatus, setCaptionStatus] = useState("");
  const [isCaptionScanning, setIsCaptionScanning] = useState(false);
  const [isCaptionApplying, setIsCaptionApplying] = useState(false);
  const [splitCandidates, setSplitCandidates] = useState<PageSplitCandidate[]>([]);
  const [splitIndex, setSplitIndex] = useState(0);
  const [splitStatus, setSplitStatus] = useState("");
  const [isSplitScanning, setIsSplitScanning] = useState(false);
  const [isSplitApplying, setIsSplitApplying] = useState(false);
  const [referenceCandidates, setReferenceCandidates] = useState<ReferenceCandidate[]>([]);
  const [referenceIndex, setReferenceIndex] = useState(0);
  const [referenceStatus, setReferenceStatus] = useState("");
  const [isReferenceScanning, setIsReferenceScanning] = useState(false);
  const [isReferenceApplying, setIsReferenceApplying] = useState(false);

  useEffect(() => {
    let alive = true;

    waitForOffice().then((ready) => {
      if (!alive) {
        return;
      }

      setOfficeReady(ready);
      setCheckingOffice(false);
      setStatus(ready ? "Word connected" : "Preview mode");
    });

    return () => {
      alive = false;
    };
  }, []);

  const activeWords = snapshot?.selectedWordCount ?? countWords(previewText);
  const resultWords = useMemo(() => countWords(result), [result]);
  const activeCaption = captionCandidates[captionIndex] ?? null;
  const captionStats = useMemo(() => summarizeCaptions(captionCandidates), [captionCandidates]);
  const activeSplit = splitCandidates[splitIndex] ?? null;
  const splitStats = useMemo(() => summarizePageSplits(splitCandidates), [splitCandidates]);
  const activeReference = referenceCandidates[referenceIndex] ?? null;
  const referenceStats = useMemo(() => summarizeReferences(referenceCandidates), [referenceCandidates]);

  const scanSelection = useCallback(
    async (options?: { silent?: boolean }) => {
      if (!options?.silent) {
        setError("");
        setResult("");
        setResultHtml("");
        setItalicTerms([]);
        setIsScanning(true);
      }

      try {
        const nextSnapshot = officeReady
          ? await readSelectionSnapshot(contextRadius)
          : createPreviewSnapshot(previewText);

        setSnapshot(nextSnapshot);
        setTargetWords(suggestTargetWords(mode, nextSnapshot.selectedWordCount));
        setStatus(nextSnapshot.source === "word" ? "Selection ready" : "Preview ready");
      } catch (scanError) {
        if (!options?.silent) {
          setError(scanError instanceof Error ? scanError.message : "Unable to read selection.");
        }
      } finally {
        if (!options?.silent) {
          setIsScanning(false);
        }
      }
    },
    [contextRadius, mode, officeReady, previewText]
  );

  async function generate() {
    setError("");
    setIsGenerating(true);
    setStatus(officeReady ? "Scanning selection" : "Generating");

    try {
      const activeSnapshot = officeReady ? await readSelectionSnapshot(contextRadius) : snapshot ?? createPreviewSnapshot(previewText);

      if (!activeSnapshot.selectedText) {
        throw new Error("Selection is empty.");
      }

      setSnapshot(activeSnapshot);
      setStatus("Generating");

      const response = await requestRewrite({
        mode,
        targetWords,
        snapshot: activeSnapshot
      });

      setResult(response.text);
      setResultHtml(response.html ?? "");
      setItalicTerms(response.italicTerms ?? []);
      setModel(response.model);
      setStatus("Suggestion ready");
    } catch (generateError) {
      setError(generateError instanceof Error ? generateError.message : "Unable to generate suggestion.");
      setStatus("Ready");
    } finally {
      setIsGenerating(false);
    }
  }

  async function replaceSelection() {
    if (!result) {
      return;
    }

    if (!officeReady) {
      setPreviewText(result);
      setSnapshot(createPreviewSnapshot(result));
      setResultHtml("");
      setItalicTerms([]);
      setStatus("Preview updated");
      return;
    }

    setIsReplacing(true);
    setError("");

    try {
      await replaceSelectionWith(result, resultHtml);
      setStatus("Selection replaced");
    } catch (replaceError) {
      setError(replaceError instanceof Error ? replaceError.message : "Unable to replace selection.");
    } finally {
      setIsReplacing(false);
    }
  }

  function updateMode(nextMode: DraftMode) {
    setMode(nextMode);

    const baselineWords = snapshot?.selectedWordCount ?? countWords(previewText);
    setTargetWords(suggestTargetWords(nextMode, baselineWords));
  }

  async function scanCaptions() {
    if (!officeReady) {
      setError("Open this add-in inside Word to scan tables and pictures.");
      return;
    }

    setError("");
    setCaptionStatus("");
    setIsCaptionScanning(true);

    try {
      const candidates = await scanCaptionCandidates();
      setCaptionCandidates(candidates);
      setCaptionIndex(0);
      setCaptionStatus(
        candidates.length
          ? `Found ${candidates.length} item${candidates.length === 1 ? "" : "s"}. ${summarizeCaptions(candidates).matched} matched.`
          : "No tables or inline pictures found."
      );
      setStatus("Caption scan ready");
    } catch (captionError) {
      setError(captionError instanceof Error ? captionError.message : "Unable to scan captions.");
      setStatus("Ready");
    } finally {
      setIsCaptionScanning(false);
    }
  }

  async function moveCaption(step: number) {
    if (!captionCandidates.length) {
      return;
    }

    const nextIndex = (captionIndex + step + captionCandidates.length) % captionCandidates.length;
    const nextCandidate = captionCandidates[nextIndex];

    setCaptionIndex(nextIndex);
    setCaptionStatus(`${nextCandidate.label} ${nextIndex + 1} of ${captionCandidates.length}. Click Select to jump.`);
  }

  async function selectActiveCaption() {
    if (!activeCaption) {
      return;
    }

    try {
      await selectCaptionCandidate(activeCaption);
      setCaptionStatus(`${activeCaption.label} target selected`);
    } catch (selectError) {
      setError(selectError instanceof Error ? selectError.message : "Unable to select caption target.");
    }
  }

  async function applyActiveCaption() {
    if (!activeCaption) {
      return;
    }

    setError("");
    setIsCaptionApplying(true);

    try {
      const updated = await applyCaptionCandidate(activeCaption);
      setCaptionCandidates((items) => items.map((item) => (item.id === activeCaption.id ? updated : item)));
      setCaptionStatus(`${activeCaption.label} caption applied; ${updated.syncedReferences ?? 0} refs synced.`);
      setStatus("Caption applied");
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Unable to apply caption.");
    } finally {
      setIsCaptionApplying(false);
    }
  }

  async function applyMatchedCaptions() {
    const matchedCandidates = captionCandidates.filter(isCaptionApplicable);

    if (!matchedCandidates.length) {
      setCaptionStatus("No matched captions to apply.");
      return;
    }

    setError("");
    setIsCaptionApplying(true);

    try {
      const summary = await applyAllCaptionCandidates(captionCandidates);
      const refreshed = await scanCaptionCandidates();
      setCaptionCandidates(refreshed);
      setCaptionIndex(0);
      setCaptionStatus(`Applied ${summary.applied}; skipped ${summary.skipped}; refs synced ${summary.referencesSynced}.`);
      setStatus("Captions applied");
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Unable to apply captions.");
    } finally {
      setIsCaptionApplying(false);
    }
  }

  async function scanReferences() {
    if (!officeReady) {
      setError("Open this add-in inside Word to scan table and figure mentions.");
      return;
    }

    setError("");
    setReferenceStatus("");
    setIsReferenceScanning(true);

    try {
      const candidates = await scanReferenceCandidates();
      setReferenceCandidates(candidates);
      setReferenceIndex(0);
      setReferenceStatus(
        candidates.length
          ? `Found ${candidates.length} mention${candidates.length === 1 ? "" : "s"}. ${summarizeReferences(candidates).pending} pending.`
          : "No table or figure mentions found."
      );
      setStatus("References scan ready");
    } catch (referenceError) {
      setError(referenceError instanceof Error ? referenceError.message : "Unable to scan references.");
      setStatus("Ready");
    } finally {
      setIsReferenceScanning(false);
    }
  }

  function moveReference(step: number) {
    if (!referenceCandidates.length) {
      return;
    }

    const nextIndex = (referenceIndex + step + referenceCandidates.length) % referenceCandidates.length;
    const nextCandidate = referenceCandidates[nextIndex];

    setReferenceIndex(nextIndex);
    setReferenceStatus(`${nextCandidate.label} mention ${nextIndex + 1} of ${referenceCandidates.length}.`);
  }

  async function selectActiveReference() {
    if (!activeReference) {
      return;
    }

    try {
      await selectReferenceCandidate(activeReference);
      setReferenceStatus(`${activeReference.mentionText} selected`);
    } catch (selectError) {
      setError(selectError instanceof Error ? selectError.message : "Unable to select reference mention.");
    }
  }

  async function applyActiveReference() {
    if (!activeReference) {
      return;
    }

    setError("");
    setIsReferenceApplying(true);

    try {
      const updated = await applyReferenceCandidate(activeReference);
      setReferenceCandidates((items) => items.map((item) => (item.id === activeReference.id ? updated : item)));
      setReferenceStatus(`${activeReference.mentionText} synced to ${activeReference.label} ${activeReference.newNumber}`);
      setStatus("Reference synced");
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Unable to apply reference.");
    } finally {
      setIsReferenceApplying(false);
    }
  }

  async function applyAllReferences() {
    const pendingCandidates = referenceCandidates.filter(isReferenceApplicable);

    if (!pendingCandidates.length) {
      setReferenceStatus("No matched references to apply.");
      return;
    }

    setError("");
    setIsReferenceApplying(true);

    try {
      const summary = await applyAllReferenceCandidates(referenceCandidates);
      const refreshed = await scanReferenceCandidates();
      setReferenceCandidates(refreshed);
      setReferenceIndex(0);
      setReferenceStatus(`Applied ${summary.applied}; skipped ${summary.skipped}.`);
      setStatus("References synced");
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Unable to apply references.");
    } finally {
      setIsReferenceApplying(false);
    }
  }

  async function scanPageSplits() {
    if (!officeReady) {
      setError("Open this add-in inside Word to scan page split risks.");
      return;
    }

    setError("");
    setSplitStatus("");
    setIsSplitScanning(true);

    try {
      const candidates = await scanPageSplitCandidates();
      const hasFallbackSplits = candidates.some((candidate) => candidate.source === "fallback");
      setSplitCandidates(candidates);
      setSplitIndex(0);
      setSplitStatus(
        candidates.length
          ? `Found ${candidates.length} ${hasFallbackSplits ? "probable " : ""}split risk${candidates.length === 1 ? "" : "s"}. ${summarizePageSplits(candidates).fixed} already guarded.`
          : "No split risks found."
      );
      setStatus("Page split scan ready");
    } catch (splitError) {
      setError(splitError instanceof Error ? splitError.message : "Unable to scan page split risks.");
      setStatus("Ready");
    } finally {
      setIsSplitScanning(false);
    }
  }

  function moveSplit(step: number) {
    if (!splitCandidates.length) {
      return;
    }

    const nextIndex = (splitIndex + step + splitCandidates.length) % splitCandidates.length;
    const nextCandidate = splitCandidates[nextIndex];

    setSplitIndex(nextIndex);
      setSplitStatus(`Split risk ${nextIndex + 1} of ${splitCandidates.length}. Sentence crosses page boundary.`);
  }

  async function selectActiveSplit() {
    if (!activeSplit) {
      return;
    }

    try {
      await selectPageSplitCandidate(activeSplit);
      setSplitStatus(`Paragraph ${activeSplit.index + 1} selected`);
    } catch (selectError) {
      setError(selectError instanceof Error ? selectError.message : "Unable to select paragraph.");
    }
  }

  async function applyActiveSplit() {
    if (!activeSplit) {
      return;
    }

    setError("");
    setIsSplitApplying(true);

    try {
      const updated = await applyPageSplitFix(activeSplit);
      setSplitCandidates((items) => items.map((item) => (item.id === activeSplit.id ? updated : item)));
      setSplitStatus("Page split guard applied");
      setStatus("Page split guarded");
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Unable to apply page split guard.");
    } finally {
      setIsSplitApplying(false);
    }
  }

  async function applyAllSplits() {
    const pendingCandidates = splitCandidates.filter((candidate) => !candidate.isFixed);

    if (!pendingCandidates.length) {
      setSplitStatus("No pending split risks to apply.");
      return;
    }

    setError("");
    setIsSplitApplying(true);

    try {
      const summary = await applyAllPageSplitFixes(splitCandidates);
      const refreshed = await scanPageSplitCandidates();
      setSplitCandidates(refreshed);
      setSplitIndex(0);
      setSplitStatus(`Applied ${summary.applied}; skipped ${summary.skipped}.`);
      setStatus("Page splits guarded");
    } catch (applyError) {
      setError(applyError instanceof Error ? applyError.message : "Unable to apply page split guards.");
    } finally {
      setIsSplitApplying(false);
    }
  }

  return (
    <main className="app-shell">
      <div className="background-accent accent-one" aria-hidden="true" />
      <div className="background-accent accent-two" aria-hidden="true" />
      <div className="background-accent accent-three" aria-hidden="true" />

      <header className="topbar">
        <div className="traffic-lights" aria-hidden="true">
          <span className="traffic-light red" />
          <span className="traffic-light yellow" />
          <span className="traffic-light green" />
        </div>
        <div>
          <h1>Draft Context AI</h1>
          <p>{checkingOffice ? "Connecting" : status}</p>
        </div>
        <div className="brand-mark">
          <Sparkles size={18} aria-hidden="true" />
        </div>
      </header>

      <GlassPanel className="panel compact" padding="14px">
        <div className="mode-tabs" role="tablist" aria-label="Rewrite mode">
          {modeOptions.map((option) => {
            const Icon = option.icon;
            const selected = mode === option.id;

            return (
              <button
                key={option.id}
                className={selected ? "mode-button selected" : "mode-button"}
                type="button"
                onClick={() => updateMode(option.id)}
                aria-pressed={selected}
                title={option.label}
              >
                <Icon size={16} aria-hidden="true" />
                <span>{option.label}</span>
              </button>
            );
          })}
        </div>

        <div className="controls-grid">
          <label className="field-label" htmlFor="targetWords">
            Final words
            <input
              id="targetWords"
              min={10}
              max={450}
              step={5}
              type="number"
              value={targetWords}
              onChange={(event) => setTargetWords(Number(event.target.value))}
            />
          </label>

          <label className="field-label" htmlFor="contextRadius">
            Context
            <select
              id="contextRadius"
              value={contextRadius}
              onChange={(event) => setContextRadius(Number(event.target.value))}
            >
              <option value={1}>1 para</option>
              <option value={2}>2 para</option>
              <option value={3}>3 para</option>
              <option value={4}>4 para</option>
            </select>
          </label>
        </div>

        {!officeReady && (
          <textarea
            className="preview-input"
            value={previewText}
            onChange={(event) => {
              setPreviewText(event.target.value);
              setSnapshot(null);
              setResult("");
              setResultHtml("");
              setItalicTerms([]);
            }}
            rows={5}
            aria-label="Preview text"
          />
        )}

        <div className="action-row">
          <button className="secondary-action" type="button" onClick={() => scanSelection()} disabled={isScanning}>
            {isScanning ? <Loader2 className="spin" size={16} /> : <ScanText size={16} />}
            <span>Scan</span>
          </button>
          <button className="primary-action" type="button" onClick={generate} disabled={isGenerating}>
            {isGenerating ? <Loader2 className="spin" size={16} /> : <Sparkles size={16} />}
            <span>Generate</span>
          </button>
        </div>
      </GlassPanel>

      <section className="stat-strip" aria-label="Word counts">
        <Metric label="Original" value={activeWords} />
        <Metric label="Target" value={targetWords} />
        <Metric label="Result" value={resultWords} />
      </section>

      <GlassPanel className="panel caption-panel" padding="14px">
        <div className="section-title">
          <FileSearch size={16} aria-hidden="true" />
          <h2>Captions</h2>
        </div>

        <div className="caption-summary" aria-label="Caption scan summary">
          <span>
            <Table2 size={14} aria-hidden="true" />
            {captionStats.tables}
          </span>
          <span>
            <ImageIcon size={14} aria-hidden="true" />
            {captionStats.figures}
          </span>
          <span>
            <BadgeCheck size={14} aria-hidden="true" />
            {captionStats.matched}
          </span>
        </div>

        <div className="action-row caption-action-row">
          <button
            className="secondary-action"
            type="button"
            onClick={scanCaptions}
            disabled={!officeReady || isCaptionScanning}
          >
            {isCaptionScanning ? <Loader2 className="spin" size={16} /> : <ScanText size={16} />}
            <span>Scan captions</span>
          </button>
          <button
            className="primary-action"
            type="button"
            title="Apply all matched caption structures"
            onClick={applyMatchedCaptions}
            disabled={!captionStats.matched || isCaptionApplying}
          >
            {isCaptionApplying ? <Loader2 className="spin" size={16} /> : <BadgeCheck size={16} />}
            <span>Apply all matched</span>
          </button>
        </div>

        {!officeReady && <p className="caption-hint">Open in Word to scan document objects.</p>}

        {activeCaption && (
          <div className="caption-card">
            <div className="caption-card-head">
              <span className={`caption-kind ${activeCaption.kind}`}>
                {activeCaption.kind === "table" ? <Table2 size={14} /> : <ImageIcon size={14} />}
                {activeCaption.label}
              </span>
              <span>
                {captionIndex + 1}/{captionCandidates.length}
              </span>
            </div>

            <div className="caption-copy">
              <span>{activeCaption.targetPosition === "above" ? "Text above" : "Text below"}</span>
              <p>{activeCaption.rawText || "No caption text found"}</p>
            </div>

            <div className="caption-copy title-copy">
              <span>Extracted title</span>
              <p>{activeCaption.title || "-"}</p>
            </div>

            <div className="caption-note">
              <span className={`confidence ${activeCaption.confidence}`}>{formatConfidence(activeCaption.confidence)}</span>
              <span>{activeCaption.isCaptionStyle ? "Already Caption style" : activeCaption.note}</span>
            </div>

            <div className="caption-nav">
              <button type="button" className="icon-action" onClick={() => moveCaption(-1)} title="Previous caption">
                <ChevronLeft size={16} aria-hidden="true" />
              </button>
              <button type="button" className="icon-action" onClick={() => moveCaption(1)} title="Next caption">
                <ChevronRight size={16} aria-hidden="true" />
              </button>
              <button type="button" className="secondary-action mini-action" onClick={selectActiveCaption}>
                <MousePointer2 size={15} aria-hidden="true" />
                <span>Select</span>
              </button>
              <button
                type="button"
                className="primary-action mini-action"
                onClick={applyActiveCaption}
                disabled={!isCaptionApplicable(activeCaption) || isCaptionApplying}
              >
                {isCaptionApplying ? <Loader2 className="spin" size={15} /> : <BadgeCheck size={15} />}
                <span>Apply</span>
              </button>
            </div>
          </div>
        )}

        {captionStatus && <p className="caption-status">{captionStatus}</p>}
      </GlassPanel>

      <GlassPanel className="panel caption-panel reference-panel" padding="14px">
        <div className="section-title">
          <Link2 size={16} aria-hidden="true" />
          <h2>References</h2>
        </div>

        <div className="caption-summary" aria-label="Reference scan summary">
          <span>
            <Table2 size={14} aria-hidden="true" />
            {referenceStats.tables}
          </span>
          <span>
            <ImageIcon size={14} aria-hidden="true" />
            {referenceStats.figures}
          </span>
          <span>
            <BadgeCheck size={14} aria-hidden="true" />
            {referenceStats.pending}
          </span>
        </div>

        <div className="action-row caption-action-row">
          <button
            className="secondary-action"
            type="button"
            onClick={scanReferences}
            disabled={!officeReady || isReferenceScanning}
          >
            {isReferenceScanning ? <Loader2 className="spin" size={16} /> : <ScanText size={16} />}
            <span>Scan refs</span>
          </button>
          <button
            className="primary-action"
            type="button"
            onClick={applyAllReferences}
            disabled={!referenceStats.pending || isReferenceApplying}
          >
            {isReferenceApplying ? <Loader2 className="spin" size={16} /> : <BadgeCheck size={16} />}
            <span>Apply refs</span>
          </button>
        </div>

        {!officeReady && <p className="caption-hint">Open in Word to sync table and figure mentions.</p>}

        {activeReference && (
          <div className="caption-card">
            <div className="caption-card-head">
              <span className={`caption-kind ${activeReference.kind}`}>
                {activeReference.kind === "table" ? <Table2 size={14} /> : <ImageIcon size={14} />}
                {activeReference.label}
              </span>
              <span>
                {referenceIndex + 1}/{referenceCandidates.length}
              </span>
            </div>

            <div className="caption-copy">
              <span>Mention</span>
              <p>
                {activeReference.mentionText} {"->"} {activeReference.label} {activeReference.newNumber}
              </p>
            </div>

            <div className="caption-copy title-copy">
              <span>Target caption</span>
              <p>{activeReference.targetTitle || `${activeReference.label} ${activeReference.newNumber}`}</p>
            </div>

            <div className="caption-note">
              <span className={`confidence ${activeReference.confidence}`}>{formatReferenceConfidence(activeReference)}</span>
              <span>{activeReference.note}</span>
            </div>

            <div className="caption-nav">
              <button type="button" className="icon-action" onClick={() => moveReference(-1)} title="Previous reference">
                <ChevronLeft size={16} aria-hidden="true" />
              </button>
              <button type="button" className="icon-action" onClick={() => moveReference(1)} title="Next reference">
                <ChevronRight size={16} aria-hidden="true" />
              </button>
              <button type="button" className="secondary-action mini-action" onClick={selectActiveReference}>
                <MousePointer2 size={15} aria-hidden="true" />
                <span>Select</span>
              </button>
              <button
                type="button"
                className="primary-action mini-action"
                onClick={applyActiveReference}
                disabled={!isReferenceApplicable(activeReference) || isReferenceApplying}
              >
                {isReferenceApplying ? <Loader2 className="spin" size={15} /> : <BadgeCheck size={15} />}
                <span>Apply</span>
              </button>
            </div>
          </div>
        )}

        {referenceStatus && <p className="caption-status">{referenceStatus}</p>}
      </GlassPanel>

      <GlassPanel className="panel caption-panel split-panel" padding="14px">
        <div className="section-title">
          <FileSearch size={16} aria-hidden="true" />
          <h2>Page Splits</h2>
        </div>

        <div className="caption-summary" aria-label="Page split scan summary">
          <span>
            <ScanText size={14} aria-hidden="true" />
            {splitStats.total}
          </span>
          <span>
            <BadgeCheck size={14} aria-hidden="true" />
            {splitStats.fixed}
          </span>
          <span>
            <FileSearch size={14} aria-hidden="true" />
            {splitStats.pending}
          </span>
        </div>

        <div className="action-row caption-action-row">
          <button
            className="secondary-action"
            type="button"
            onClick={scanPageSplits}
            disabled={!officeReady || isSplitScanning}
          >
            {isSplitScanning ? <Loader2 className="spin" size={16} /> : <ScanText size={16} />}
            <span>Scan splits</span>
          </button>
          <button
            className="primary-action"
            type="button"
            onClick={applyAllSplits}
            disabled={!splitStats.pending || isSplitApplying}
          >
            {isSplitApplying ? <Loader2 className="spin" size={16} /> : <BadgeCheck size={16} />}
            <span>Apply all</span>
          </button>
        </div>

        {!officeReady && <p className="caption-hint">Open in Word to scan page split risks.</p>}

        {activeSplit && (
          <div className="caption-card">
            <div className="caption-card-head">
              <span className={`caption-kind split-risk ${activeSplit.confidence}`}>
                <FileSearch size={14} />
                Split
              </span>
              <span>
                {splitIndex + 1}/{splitCandidates.length}
              </span>
            </div>

            <div className="caption-copy">
              <span>Split sentence</span>
              <p>{activeSplit.text}</p>
            </div>

            <div className="caption-copy title-copy">
              <span>Layout risk</span>
              <p>
                Ending punctuation appears after a rendered page break.
              </p>
            </div>

            <div className="caption-note">
              <span className={`confidence ${activeSplit.confidence}`}>{formatSplitConfidence(activeSplit)}</span>
              <span>{activeSplit.reason}</span>
            </div>

            <div className="caption-nav">
              <button type="button" className="icon-action" onClick={() => moveSplit(-1)} title="Previous split risk">
                <ChevronLeft size={16} aria-hidden="true" />
              </button>
              <button type="button" className="icon-action" onClick={() => moveSplit(1)} title="Next split risk">
                <ChevronRight size={16} aria-hidden="true" />
              </button>
              <button type="button" className="secondary-action mini-action" onClick={selectActiveSplit}>
                <MousePointer2 size={15} aria-hidden="true" />
                <span>Select</span>
              </button>
              <button
                type="button"
                className="primary-action mini-action"
                onClick={applyActiveSplit}
                disabled={activeSplit.isFixed || isSplitApplying}
              >
                {isSplitApplying ? <Loader2 className="spin" size={15} /> : <BadgeCheck size={15} />}
                <span>Apply</span>
              </button>
            </div>
          </div>
        )}

        {splitStatus && <p className="caption-status">{splitStatus}</p>}
      </GlassPanel>

      {snapshot && (
        <GlassPanel className="panel context-panel" padding="14px">
          <div className="section-title">
            <ClipboardCheck size={16} aria-hidden="true" />
            <h2>Context</h2>
          </div>
          {snapshot.context.heading && <p className="heading-chip">{snapshot.context.heading}</p>}
          <p className="selection-preview">{snapshot.selectedText}</p>
          {(snapshot.context.before.length > 0 || snapshot.context.after.length > 0) && (
            <div className="context-lines">
              {snapshot.context.before.map((item, index) => (
                <p key={`before-${index}`}>{item}</p>
              ))}
              {snapshot.context.after.map((item, index) => (
                <p key={`after-${index}`}>{item}</p>
              ))}
            </div>
          )}
        </GlassPanel>
      )}

      <GlassPanel className="panel result-panel" padding="14px">
        <div className="section-title">
          <RefreshCcw size={16} aria-hidden="true" />
          <h2>Suggestion</h2>
        </div>
        <textarea
          className="result-box"
          value={result}
          onChange={(event) => {
            setResult(event.target.value);
            setResultHtml("");
            setItalicTerms([]);
          }}
          placeholder="Result appears here"
          rows={8}
          aria-label="AI suggestion"
        />
        <div className="result-footer">
          <span>{model ? `Model: ${model}${italicTerms.length ? ` - Italic terms: ${italicTerms.length}` : ""}` : "No session stored"}</span>
          <button
            className="primary-action replace"
            type="button"
            onClick={replaceSelection}
            disabled={!result || isReplacing}
          >
            {isReplacing ? <Loader2 className="spin" size={16} /> : <Check size={16} />}
            <span>Replace</span>
          </button>
        </div>
      </GlassPanel>

      {error && <div className="error-banner">{error}</div>}
    </main>
  );
}

function Metric({ label, value }: { label: string; value: number }) {
  return (
    <GlassPanel className="metric" radius={12} padding="10px">
      <span>{label}</span>
      <strong>{value}</strong>
    </GlassPanel>
  );
}

function suggestTargetWords(mode: DraftMode, currentWords: number): number {
  if (mode === "extend") {
    return clamp(Math.ceil(currentWords * 1.8), 40, 450);
  }

  if (mode === "shrink") {
    return clamp(Math.max(Math.ceil(currentWords * 0.55), 15), 10, 300);
  }

  return clamp(currentWords, 15, 450);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function summarizeCaptions(candidates: CaptionCandidate[]) {
  return candidates.reduce(
    (summary, candidate) => {
      if (candidate.kind === "table") {
        summary.tables += 1;
      } else {
        summary.figures += 1;
      }

      if (isCaptionApplicable(candidate)) {
        summary.matched += 1;
      }

      return summary;
    },
    {
      tables: 0,
      figures: 0,
      matched: 0
    }
  );
}

function summarizePageSplits(candidates: PageSplitCandidate[]) {
  return candidates.reduce(
    (summary, candidate) => {
      summary.total += 1;

      if (candidate.isFixed) {
        summary.fixed += 1;
      } else {
        summary.pending += 1;
      }

      return summary;
    },
    {
      total: 0,
      fixed: 0,
      pending: 0
    }
  );
}

function summarizeReferences(candidates: ReferenceCandidate[]) {
  return candidates.reduce(
    (summary, candidate) => {
      if (candidate.kind === "table") {
        summary.tables += 1;
      } else {
        summary.figures += 1;
      }

      if (isReferenceApplicable(candidate)) {
        summary.pending += 1;
      }

      return summary;
    },
    {
      tables: 0,
      figures: 0,
      pending: 0
    }
  );
}

function isCaptionApplicable(candidate: CaptionCandidate): boolean {
  return candidate.confidence === "high" || candidate.confidence === "medium";
}

function isReferenceApplicable(candidate: ReferenceCandidate): boolean {
  return !candidate.isSynced && (candidate.confidence === "high" || candidate.confidence === "medium");
}

function formatConfidence(confidence: CaptionCandidate["confidence"]): string {
  const labels = {
    high: "Matched",
    medium: "Check",
    low: "Low",
    missing: "Missing"
  };

  return labels[confidence];
}

function formatSplitConfidence(candidate: PageSplitCandidate): string {
  if (candidate.isFixed) {
    return "Guarded";
  }

  const labels = {
    high: "High",
    medium: "Check",
    low: "Low"
  };

  return labels[candidate.confidence];
}

function formatReferenceConfidence(candidate: ReferenceCandidate): string {
  if (candidate.isSynced) {
    return "Synced";
  }

  const labels = {
    high: "Matched",
    medium: "Check",
    low: "Low"
  };

  return labels[candidate.confidence];
}
