import { BilingualText } from "@/components/BilingualText";
import { copy } from "@/lib/i18n";

export function ProcessingStatus({ active }: { active: boolean }) {
  return (
    <section className={`step-card ${active ? "is-active" : ""}`} aria-labelledby="check-heading" aria-live="polite">
      <div className="step-number" aria-hidden="true">2</div>
      <div className="step-content">
        <p className="eyebrow">Step 2 / चरण 2</p>
        <h2 id="check-heading">Automatic file check / फ़ाइल की स्वतः जाँच</h2>
        {active ? (
          <div className="status-line">
            <span className="spinner" aria-hidden="true" />
            <p><BilingualText text={copy.checking} /></p>
          </div>
        ) : (
          <p className="supporting">The check starts when you choose a file. / फ़ाइल चुनते ही जाँच शुरू हो जाएगी।</p>
        )}
      </div>
    </section>
  );
}
