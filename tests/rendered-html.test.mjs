import assert from "node:assert/strict";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("renders the contract generator", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  const html = await response.text();
  assert.match(html, /Генератор договоров РекламаСтрой/);
  assert.match(html, /Основные данные/);
  assert.match(html, /Заказчик/);
  assert.match(html, /Услуги и расчёты/);
  assert.match(html, /Подпись и печать исполнителя/);
  assert.match(html, /Не загружено/);
  assert.match(html, /Word без печати/);
  assert.match(html, /PDF без печати/);
  assert.doesNotMatch(html, /src="\/signature\.png"|src="\/seal\.png"/);
  assert.doesNotMatch(html, /codex-preview|Building your site/);
});
