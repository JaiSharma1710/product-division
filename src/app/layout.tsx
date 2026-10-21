import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel to ERP Upload Generator",
  description: "Private browser-only converter for DO_UPLOAD_PS.xls and DO_UPLOAD_LI.xls",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
