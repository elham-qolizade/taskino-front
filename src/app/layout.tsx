import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Taskino",
  description: "Taskino project and task management dashboard",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="fa" dir="rtl">
      <body>{children}</body>
    </html>
  );
}
