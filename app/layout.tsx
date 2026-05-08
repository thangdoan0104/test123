import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Invitation Card Editor MVP",
  description: "Batch invitation card editor with custom fonts, Excel mapping, JPEG/PDF/ZIP export."
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}
