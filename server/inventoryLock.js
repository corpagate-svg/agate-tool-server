// 在庫管理表.xlsx への「読み込み→変更→書き込み」をプロセス内で直列化するための共有ロック。
// server.js・realInventory.js・inventoryRebuild.js・ebay/inventorySync.js など、
// 在庫管理表.xlsxに書き込むすべての処理は、必ずこのモジュール経由でファイルI/Oを行うこと。
// Node のモジュールキャッシュにより require のたびに新しいインスタンスが作られることはないため、
// 全ファイルが同じキュー(tail)を共有できる。
const fs = require("fs");
const crypto = require("crypto");

let tail = Promise.resolve();

// fn を「直前に積まれた処理がすべて完了してから」実行する。fn自体の成功/失敗はそのまま
// 呼び出し元へ返すが、キュー自体は途切れさせない(前の処理が失敗していても次を実行する)。
function withInventoryLock(fn) {
  const next = tail.then(fn, fn);
  tail = next.then(() => undefined, () => undefined);
  return next;
}

// 同時書き込みが同じ一時ファイルを取り合わないよう、呼び出しごとに一意な一時ファイル名を作る。
function uniqueTmpPath(targetPath) {
  return `${targetPath}.${process.pid}-${Date.now()}-${crypto.randomBytes(4).toString("hex")}.tmp`;
}

// tmpファイルが存在すれば削除する。後片付け自体が失敗しても、呼び出し元が投げようとしている
// 本来のエラーを上書き・隠蔽しないよう、ここでは意図的に握りつぶす(元のxlsxには一切触れない)。
function cleanupTmp(tmpPath) {
  try {
    if (fs.existsSync(tmpPath)) fs.unlinkSync(tmpPath);
  } catch (cleanupErr) {
    // 無視する(本来のエラーはこの関数の外でそのまま呼び出し元へ伝播させる)
  }
}

// ExcelJSのワークブックを一意な一時ファイルへ書き出してから原子的にrenameする。
// 書き込み・rename のどちらが失敗しても一意tmpファイルを残さない(成功時はrenameで
// tmpパス自体が消えるため、片付けは失敗時のみ発生する)。
async function atomicWriteWorkbook(wb, targetPath) {
  const tmpPath = uniqueTmpPath(targetPath);
  try {
    await wb.xlsx.writeFile(tmpPath);
    fs.renameSync(tmpPath, targetPath);
  } catch (err) {
    cleanupTmp(tmpPath);
    throw err;
  }
}

// xlsxをExcelJS経由で構築しない場合(丸ごとアップロードされたバッファ等)向けの原子的書き込み。
function atomicWriteBuffer(buf, targetPath) {
  const tmpPath = uniqueTmpPath(targetPath);
  try {
    fs.writeFileSync(tmpPath, buf);
    fs.renameSync(tmpPath, targetPath);
  } catch (err) {
    cleanupTmp(tmpPath);
    throw err;
  }
}

module.exports = { withInventoryLock, atomicWriteWorkbook, atomicWriteBuffer, uniqueTmpPath, cleanupTmp };
