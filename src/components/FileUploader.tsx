import { useRef } from "react";
import { BilingualText } from "@/components/BilingualText";
import { copy } from "@/lib/i18n";

interface FileUploaderProps {
  filename?: string;
  disabled?: boolean;
  onSelect: (file: File) => void;
}

export function FileUploader({ filename, disabled, onSelect }: FileUploaderProps) {
  const inputRef = useRef<HTMLInputElement>(null);

  return (
    <section className="step-card" aria-labelledby="upload-heading">
      <div className="step-number" aria-hidden="true">1</div>
      <div className="step-content">
        <p className="eyebrow">Step 1 / चरण 1</p>
        <h2 id="upload-heading"><BilingualText text={copy.select} /></h2>
        <p className="supporting"><BilingualText text={copy.supported} /></p>
        <input
          ref={inputRef}
          className="sr-only"
          type="file"
          accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          disabled={disabled}
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) onSelect(file);
            event.currentTarget.value = "";
          }}
        />
        <button className="button button-primary upload-button" type="button" disabled={disabled} onClick={() => inputRef.current?.click()}>
          <span aria-hidden="true" className="button-icon">↑</span>
          <BilingualText text={copy.chooseFile} />
        </button>
        {filename && (
          <p className="selected-file" role="status">
            <span aria-hidden="true">📄</span>
            <span><strong>Selected file / चुनी गई फ़ाइल:</strong> {filename}</span>
          </p>
        )}
      </div>
    </section>
  );
}
