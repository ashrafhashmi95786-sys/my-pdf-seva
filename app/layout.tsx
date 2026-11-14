import type { Metadata } from "next";
import "./globals.css";
export const metadata: Metadata = {
  title: "My PDF Seva — Free Online PDF Tools",
  description: "Merge, split, rotate and organize PDF files for free. Simple, secure PDF tools that work in your browser.",
  applicationName: "My PDF Seva",
  keywords: ["free PDF tools", "merge PDF", "split PDF", "PDF India"]
};
export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body>{children}</body></html>;
}