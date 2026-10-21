"use client";

import { useState } from "react";
import { BilingualText } from "@/components/BilingualText";
import { ExcludedProductsModal } from "@/components/ExcludedProductsModal";
import { copy } from "@/lib/i18n";
import type { ConversionSummary, ExcludedProduct, OrderWorkbookData } from "@/lib/types";

interface DownloadResultsProps {
  summary: ConversionSummary;
  warningCount: number;
  excludedProducts: ExcludedProduct[];
  ps: OrderWorkbookData;
  li: OrderWorkbookData;
  onDownloadPS: () => void;
  onDownloadLI: () => void;
  onDownloadBoth: () => void;
  onPreviewPS: () => void;
  onPreviewLI: () => void;
  onReset: () => void;
}

function EyeIcon() {
  return (
    <svg viewBox="0 0 24 24" width="26" height="26" aria-hidden="true">
      <path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z" fill="none" stroke="currentColor" strokeWidth="1.8" />
      <circle cx="12" cy="12" r="2.7" fill="none" stroke="currentColor" strokeWidth="1.8" />
    </svg>
  );
}

export function DownloadResults(props: DownloadResultsProps) {
  const { summary } = props;
  const [showExcluded, setShowExcluded] = useState(false);

  return (
    <section className="result-panel success-panel" aria-labelledby="ready-heading" aria-live="polite">
      <div className="result-heading">
        <span className="result-symbol" aria-hidden="true">✓</span>
        <div>
          <p className="eyebrow">Check complete / जाँच पूरी हुई</p>
          <h2 id="ready-heading"><BilingualText text={copy.ready} /></h2>
          <p>Verified and ready to use. / जाँच पूरी हुई और उपयोग के लिए तैयार है।</p>
        </div>
      </div>

      <dl className="summary-grid">
        <div><dt>Employees processed<br /><span lang="hi">कुल कर्मचारी</span></dt><dd>{summary.employees}</dd></div>
        <div><dt>PS orders<br /><span lang="hi">PS ऑर्डर</span></dt><dd>{summary.psOrders}</dd></div>
        <div><dt>LI orders<br /><span lang="hi">LI ऑर्डर</span></dt><dd>{summary.liOrders}</dd></div>
        <div><dt>Total product lines<br /><span lang="hi">कुल प्रोडक्ट एंट्री</span></dt><dd>{summary.totalLines}</dd></div>
      </dl>

      {props.excludedProducts.length > 0 && (
        <div className="excluded-callout">
          <div>
            <strong>{props.excludedProducts.length} products excluded / {props.excludedProducts.length} प्रोडक्ट छोड़े गए</strong>
            <p>Missing product information did not stop file generation. / प्रोडक्ट की जानकारी कम होने पर फ़ाइल बनाना नहीं रोका गया।</p>
          </div>
          <button className="button button-secondary" type="button" onClick={() => setShowExcluded(true)}>View missing details / अधूरी जानकारी देखें</button>
        </div>
      )}

      {props.warningCount > props.excludedProducts.length && (
        <p className="success-warning">ℹ {props.warningCount} non-blocking warnings / {props.warningCount} सामान्य चेतावनियाँ</p>
      )}

      <div className="download-grid">
        <div className="file-action-group">
          <button className="button button-download" type="button" onClick={props.onDownloadPS}>
            <span aria-hidden="true">↓</span><BilingualText text={copy.ps} /><small>DO_UPLOAD_PS.xls</small>
          </button>
          <button className="button preview-icon-button" type="button" onClick={props.onPreviewPS} aria-label="Preview PS file / PS फ़ाइल का पूर्वावलोकन" title="Preview PS file / PS फ़ाइल देखें"><EyeIcon /></button>
        </div>
        <div className="file-action-group">
          <button className="button button-download" type="button" onClick={props.onDownloadLI}>
            <span aria-hidden="true">↓</span><BilingualText text={copy.li} /><small>DO_UPLOAD_LI.xls</small>
          </button>
          <button className="button preview-icon-button" type="button" onClick={props.onPreviewLI} aria-label="Preview LI file / LI फ़ाइल का पूर्वावलोकन" title="Preview LI file / LI फ़ाइल देखें"><EyeIcon /></button>
        </div>
      </div>
      <button className="button button-zip" type="button" onClick={props.onDownloadBoth}>
        <span aria-hidden="true">▣</span><BilingualText text={copy.both} /><small>DO_UPLOAD_FILES.zip</small>
      </button>
      <button className="button button-link" type="button" onClick={props.onReset}><BilingualText text={copy.another} /></button>

      {showExcluded && <ExcludedProductsModal products={props.excludedProducts} onClose={() => setShowExcluded(false)} />}
    </section>
  );
}
