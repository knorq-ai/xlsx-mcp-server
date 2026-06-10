/**
 * Per-file write lock for serializing write operations.
 *
 * 2 層構成:
 *  1. プロセス内: Promise チェーンで同一ファイルへの書き込みを直列化する。
 *     チェーン登録は同期的に行う（await を挟むと登録順 = 呼び出し順の
 *     FIFO 保証が崩れ、並行発行された書き込みが逆順に適用され得る）。
 *  2. プロセス間: `<file>.mcplock` をアドバイザリロックとして O_EXCL で作成する。
 *     別の MCP サーバインスタンス（複数の Claude セッション等）が同じ
 *     ワークブックを同時に編集して更新を失わないようにする。
 *     保持中はハートビートで mtime を更新し続けるため、生きている所有者の
 *     ロックが「古い」と誤判定されて奪われることはない。
 *
 * プロセス間ロックのパスは realpath で正規化する（シンボリックリンク経由でも
 * 同一ファイルは同一ロックになる）。
 * 読み取り専用関数はロック不要。書き込み関数のみ withFileLock でラップする。
 */

import * as fs from "fs/promises";
import * as path from "path";
import { ErrorCode, EngineError } from "./xlsx-io.js";

const locks = new Map<string, Promise<void>>();

const LOCK_SUFFIX = ".mcplock";
/** プロセス間ロックの取得を諦めるまでの時間 */
const LOCK_TIMEOUT_MS = 10_000;
/** ロックファイルの mtime がこれより古ければ持ち主が死んだとみなして奪う */
const LOCK_STALE_MS = 60_000;
/** 保持中に mtime を更新する間隔（STALE 判定より十分短く） */
const HEARTBEAT_MS = 15_000;
const RETRY_INTERVAL_MS = 100;

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

function isProcessAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (e) {
    // EPERM = 存在するが権限がない（= 生きている）
    return (e as NodeJS.ErrnoException).code === "EPERM";
  }
}

/** シンボリックリンク等を解決した正規パスを返す。未作成ファイルは resolve のみ */
async function canonicalPath(filePath: string): Promise<string> {
  try {
    return await fs.realpath(filePath);
  } catch {
    return path.resolve(filePath);
  }
}

/**
 * プロセス間アドバイザリロックを取得する。返り値は解放関数。
 */
async function acquireCrossProcessLock(lockPath: string): Promise<() => Promise<void>> {
  const deadline = Date.now() + LOCK_TIMEOUT_MS;
  for (;;) {
    try {
      const handle = await fs.open(lockPath, "wx");
      try {
        await handle.writeFile(String(process.pid));
      } finally {
        await handle.close();
      }
      // ハートビート: 保持中は mtime を更新し続け、生きている所有者のロックが
      // stale 判定で奪われないようにする。unref でプロセス終了は妨げない。
      const heartbeat = setInterval(() => {
        const now = new Date();
        void fs.utimes(lockPath, now, now).catch(() => {});
      }, HEARTBEAT_MS);
      heartbeat.unref();
      return async () => {
        clearInterval(heartbeat);
        await fs.unlink(lockPath).catch(() => {});
      };
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code !== "EEXIST") throw e;
    }

    // 既存ロックの持ち主が死んでいる / 古すぎる場合は奪う。
    // unlink ではなく rename で奪う — rename はソースパスに対して原子的なので、
    // 複数の待機プロセスが同時に stale 判定しても勝者は 1 つに決まり、
    // 他プロセスが作り直した新しいロックを誤って消すことがない。
    try {
      const [stat, raw] = await Promise.all([
        fs.stat(lockPath),
        fs.readFile(lockPath, "utf8"),
      ]);
      const pid = Number.parseInt(raw, 10);
      const ownerDead = Number.isInteger(pid) && pid > 0 && !isProcessAlive(pid);
      const tooOld = Date.now() - stat.mtimeMs > LOCK_STALE_MS;
      if (ownerDead || tooOld) {
        const stolen = `${lockPath}.steal-${process.pid}-${Date.now()}`;
        try {
          await fs.rename(lockPath, stolen);
          await fs.unlink(stolen).catch(() => {});
        } catch {
          // 別の待機プロセスが先に奪った — そのまま再試行
        }
        continue;
      }
    } catch {
      // ロックファイルが消えた（解放された）— 即座に再試行
      continue;
    }

    if (Date.now() > deadline) {
      throw new EngineError(
        ErrorCode.FILE_LOCKED,
        `Could not acquire write lock for ${lockPath.slice(0, -LOCK_SUFFIX.length)} within ${LOCK_TIMEOUT_MS / 1000}s — ` +
          `another process is editing this workbook. Retry later, or remove ${lockPath} if no other editor is running.`,
      );
    }
    await sleep(RETRY_INTERVAL_MS);
  }
}

/**
 * 同一ファイルパスへの書き込み操作を直列化する。
 * 異なるファイルへの操作は並列に実行される。
 * 例外発生時もロックを正しく解放する。
 */
export async function withFileLock<T>(
  filePath: string,
  fn: () => Promise<T>,
): Promise<T> {
  // プロセス内チェーンのキーは同期計算する（FIFO 保証のため await 禁止）。
  // シンボリックリンク別名はプロセス内では別チェーンになるが、
  // realpath ベースの .mcplock が正しさを担保する。
  const key = path.resolve(filePath);
  const prev = locks.get(key) ?? Promise.resolve();

  let resolve!: () => void;
  const next = new Promise<void>((r) => {
    resolve = r;
  });
  locks.set(key, next);

  await prev;
  try {
    const canonical = await canonicalPath(filePath);
    const releaseCrossProcess = await acquireCrossProcessLock(canonical + LOCK_SUFFIX);
    try {
      return await fn();
    } finally {
      await releaseCrossProcess();
    }
  } finally {
    resolve();
    // GC: clean up if this is the last pending operation
    if (locks.get(key) === next) {
      locks.delete(key);
    }
  }
}
