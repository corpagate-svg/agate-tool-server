const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");
const assert = require("node:assert/strict");
const { spawn } = require("node:child_process");

function inlineScripts(html) {
  return [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)]
    .map((match) => match[1])
    .filter((script) => script.trim());
}

async function waitForDashboard(port, child) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) throw new Error(`server exited: ${child.exitCode}`);
    try {
      const response = await fetch(`http://127.0.0.1:${port}/dashboard`);
      if (response.ok) return response.text();
    } catch (_) {}
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error("dashboard start timeout");
}

test("評価済みDASHBOARD_PAGEのクライアントJavaScriptに構文エラーがない", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "agate-dashboard-syntax-"));
  const port = 40000 + (process.pid % 1000);
  const child = spawn(process.execPath, [path.join(__dirname, "server.js")], {
    cwd: __dirname,
    env: { ...process.env, API_TOKEN: "dashboard-syntax-test", AGATE_DATA_DIR: dataDir, PORT: String(port) },
    stdio: ["ignore", "pipe", "pipe"],
    windowsHide: true,
  });
  try {
    const html = await waitForDashboard(port, child);
    const scripts = inlineScripts(html);
    assert.ok(scripts.length > 0, "インラインクライアントJavaScriptが取得できません");
    scripts.forEach((script, index) => assert.doesNotThrow(() => new vm.Script(script, { filename: `dashboard-inline-${index}.js` })));

    // 今回の不具合と同じ「ダブルクォート文字列内へ生改行が入る」状態を隔離文字列で再現し、
    // この構文検査が実際に失敗を検出できることも固定する。
    const broken = scripts.join("\n").replace("\\n現在US在庫", "\n現在US在庫");
    assert.notEqual(broken, scripts.join("\n"), "不具合再現用の置換対象が見つかりません");
    assert.throws(() => new vm.Script(broken, { filename: "dashboard-broken-regression.js" }), SyntaxError);
  } finally {
    child.kill();
    await new Promise((resolve) => child.once("exit", resolve));
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});
