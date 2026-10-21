import { BilingualText } from "@/components/BilingualText";
import { copy } from "@/lib/i18n";
import type { ValidationIssue } from "@/lib/types";

interface ValidationErrorsProps {
  issues: ValidationIssue[];
  onDownloadReport: () => void;
  onReset: () => void;
}

export function ValidationErrors({ issues, onDownloadReport, onReset }: ValidationErrorsProps) {
  const errors = issues.filter((issue) => issue.severity === "error");
  const warnings = issues.filter((issue) => issue.severity === "warning");

  return (
    <section className="result-panel error-panel" role="alert" aria-labelledby="error-heading">
      <div className="result-heading">
        <span className="result-symbol" aria-hidden="true">!</span>
        <div>
          <p className="eyebrow">Check stopped / जाँच रुक गई</p>
          <h2 id="error-heading"><BilingualText text={copy.errors} /></h2>
          <p>
            {errors.length} blocking {errors.length === 1 ? "issue" : "issues"} found. / {errors.length} आवश्यक {errors.length === 1 ? "त्रुटि मिली" : "त्रुटियाँ मिलीं"}।
          </p>
        </div>
      </div>

      <details className="issue-details" open={errors.length <= 3}>
        <summary><BilingualText text={copy.details} /></summary>
        <ol className="issue-list">
          {issues.map((issue) => (
            <li key={issue.id} className={issue.severity === "warning" ? "warning-issue" : ""}>
              <span className="issue-tag">{issue.severity === "error" ? "Error / त्रुटि" : "Warning / चेतावनी"}</span>
              <p lang="en">{issue.en}</p>
              <p lang="hi">{issue.hi}</p>
              {(issue.employeeCode || issue.sourceRow || issue.quantity !== undefined) && (
                <p className="issue-context">
                  Employee / कर्मचारी: {issue.employeeName || issue.employeeCode || "—"}
                  {issue.sourceRow ? ` · Row / पंक्ति: ${issue.sourceRow}` : ""}
                  {issue.quantity !== undefined ? ` · Quantity / मात्रा: ${String(issue.quantity)}` : ""}
                </p>
              )}
            </li>
          ))}
        </ol>
      </details>

      {warnings.length > 0 && <p className="warning-count">{warnings.length} warnings / {warnings.length} चेतावनियाँ</p>}
      <div className="button-row">
        <button className="button button-secondary" type="button" onClick={onDownloadReport}><BilingualText text={copy.errorReport} /></button>
        <button className="button button-primary" type="button" onClick={onReset}><BilingualText text={copy.another} /></button>
      </div>
    </section>
  );
}
