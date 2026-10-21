import type { BilingualText as BilingualTextValue } from "@/lib/types";

export function BilingualText({ text, className = "" }: { text: BilingualTextValue; className?: string }) {
  return (
    <span className={className}>
      <span lang="en">{text.en}</span>
      <span className="language-divider" aria-hidden="true"> / </span>
      <span lang="hi">{text.hi}</span>
    </span>
  );
}
