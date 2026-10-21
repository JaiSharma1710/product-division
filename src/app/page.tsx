"use client";

import { useCallback, useState } from "react";
import { BilingualText } from "@/components/BilingualText";
import { DownloadResults } from "@/components/DownloadResults";
import { FileUploader } from "@/components/FileUploader";
import { ProcessingStatus } from "@/components/ProcessingStatus";
import { ValidationErrors } from "@/components/ValidationErrors";
import { WorkbookPreview } from "@/components/WorkbookPreview";
import { copy } from "@/lib/i18n";
import { convertWorkbook } from "@/lib/excel/convert";
import { workbookBlob, writeValidationReport } from "@/lib/excel/writeOutputs";
import type { ConversionResult, ValidationIssue } from "@/lib/types";

type Status = "idle" | "checking" | "error" | "success";

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
}

export default function Home() {
  const [status, setStatus] = useState<Status>("idle");
  const [filename, setFilename] = useState<string>();
  const [result, setResult] = useState<ConversionResult>();
  const [preview, setPreview] = useState<"PS" | "LI" | null>(null);

  const reset = useCallback(() => {
    setFilename(undefined);
    setResult(undefined);
    setStatus("idle");
    setPreview(null);
  }, []);

  const handleFile = useCallback(async (file: File) => {
    setFilename(file.name);
    setResult(undefined);
    setStatus("checking");
    await new Promise((resolve) => window.setTimeout(resolve, 40));

    if (!/\.(xlsx?|XLSX?)$/.test(file.name)) {
      const invalid: ValidationIssue = {
        id: "file-type",
        severity: "error",
        en: "This file type is not supported. Choose an Excel file ending in .xls or .xlsx.",
        hi: "यह फ़ाइल प्रकार समर्थित नहीं है। .xls या .xlsx नाम वाली Excel फ़ाइल चुनें।",
      };
      setResult({ issues: [invalid], excludedProducts: [] });
      setStatus("error");
      return;
    }

    try {
      const converted = await convertWorkbook(await file.arrayBuffer());
      setResult(converted);
      setStatus(converted.generated ? "success" : "error");
    } catch {
      const unexpected: ValidationIssue = {
        id: "unexpected",
        severity: "error",
        en: "The file could not be checked. Close and reopen it in Excel, save it, and try again.",
        hi: "फ़ाइल की जाँच नहीं हो सकी। इसे Excel में बंद करके दोबारा खोलें, सेव करें और फिर प्रयास करें।",
      };
      setResult({ issues: [unexpected], excludedProducts: [] });
      setStatus("error");
    }
  }, []);

  if (preview && result?.generated && result.ps && result.li) {
    const isPS = preview === "PS";
    return (
      <WorkbookPreview
        workbook={isPS ? result.ps : result.li}
        filename={isPS ? "DO_UPLOAD_PS.xls" : "DO_UPLOAD_LI.xls"}
        onBack={() => setPreview(null)}
        onDownload={() => download(workbookBlob(isPS ? result.generated!.ps : result.generated!.li), isPS ? "DO_UPLOAD_PS.xls" : "DO_UPLOAD_LI.xls")}
      />
    );
  }

  return (
    <main>
      <header className="hero">
        <div className="brand-mark" aria-hidden="true">XL</div>
        <div>
          <p className="kicker">ERP upload helper / ERP अपलोड सहायक</p>
          <h1><BilingualText text={copy.title} /></h1>
          <p className="hero-subtitle"><BilingualText text={copy.subtitle} /></p>
        </div>
      </header>

      <aside className="privacy-note">
        <span aria-hidden="true">🔒</span>
        <p><strong>Private and safe / निजी और सुरक्षित</strong><br /><BilingualText text={copy.privacy} /></p>
      </aside>

      <div className="workflow">
        <FileUploader filename={filename} disabled={status === "checking"} onSelect={handleFile} />
        <ProcessingStatus active={status === "checking"} />
        <section className={`step-card ${status === "success" ? "is-complete" : ""}`} aria-labelledby="download-heading">
          <div className="step-number" aria-hidden="true">3</div>
          <div className="step-content">
            <p className="eyebrow">Step 3 / चरण 3</p>
            <h2 id="download-heading">Download ERP files / ERP फ़ाइलें डाउनलोड करें</h2>
            <p className="supporting">Downloads appear only after every required check passes. / सभी आवश्यक जाँच पूरी होने के बाद ही डाउनलोड दिखाई देंगे।</p>
          </div>
        </section>
      </div>

      {status === "error" && result && (
        <ValidationErrors
          issues={result.issues}
          onReset={reset}
          onDownloadReport={() => download(
            new Blob([writeValidationReport(result.issues).slice().buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
            "Validation_Errors.xlsx",
          )}
        />
      )}

      {status === "success" && result?.generated && result.summary && result.ps && result.li && (
        <DownloadResults
          summary={result.summary}
          warningCount={result.issues.filter((issue) => issue.severity === "warning").length}
          excludedProducts={result.excludedProducts}
          ps={result.ps}
          li={result.li}
          onDownloadPS={() => download(workbookBlob(result.generated!.ps), "DO_UPLOAD_PS.xls")}
          onDownloadLI={() => download(workbookBlob(result.generated!.li), "DO_UPLOAD_LI.xls")}
          onDownloadBoth={() => download(result.generated!.zip, "DO_UPLOAD_FILES.zip")}
          onPreviewPS={() => setPreview("PS")}
          onPreviewLI={() => setPreview("LI")}
          onReset={reset}
        />
      )}

      <footer>
        <p>No sign-in · No server upload · No data storage</p>
        <p lang="hi">न साइन-इन · न सर्वर अपलोड · न डेटा संग्रह</p>
      </footer>
    </main>
  );
}
