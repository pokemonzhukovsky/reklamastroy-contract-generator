import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Генератор договоров РекламаСтрой",
  description: "Подготовка договоров и приложений для ООО «Мастера Рекламы» и ИП Гассиев.",
  icons: {
    icon: "/reklamastroy-logo.png",
    shortcut: "/reklamastroy-logo.png",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ru">
      <body>{children}</body>
    </html>
  );
}
