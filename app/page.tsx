import type { Metadata } from "next";
import { ContractGenerator } from "./contract-generator";

export const metadata: Metadata = {
  title: "Генератор договоров — Мастера Рекламы",
  description: "Подготовка договора и приложения для ООО или ИП.",
};

export default function Home() {
  return <ContractGenerator />;
}
