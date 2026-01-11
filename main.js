// main.js

// HTML要素の参照を取得
const input = document.getElementById("imageInput");
const canvas = document.getElementById("canvas");
const ctx = canvas.getContext("2d");
const resultDiv = document.getElementById("result");

// ===== 範囲フィルタ（自由形ポリゴン）用の状態 =====

// 範囲フィルタのポリゴン。
// { points: [{x, y}, ...] } を想定。常に points.length >= 3 を維持する。
let rangeFilterPolygon = null;

// ドラッグ状態管理
let rfDraggingVertexIndex = -1;  // 節点ドラッグ中ならその index、していなければ -1
let rfDraggingWhole = false;     // ポリゴン全体をドラッグ中かどうか
let rfDragStartMouse = null;     // {x, y} マウスダウン時座標（キャンバス座標）
let rfDragStartPoints = null;    // ポリゴン全体ドラッグ開始時の points のコピー

// 節点・辺のヒットテスト用半径・許容距離
const RF_HANDLE_RADIUS = 8;      // 節点ヒット半径(px)
const RF_EDGE_HIT_DIST = 6;      // 辺ヒットとみなす距離(px)

// === キャンバス座標への変換 ===
function rfGetMousePosOnCanvas(evt) {
  const rect = canvas.getBoundingClientRect();
  const scaleX = canvas.width / rect.width;
  const scaleY = canvas.height / rect.height;
  return {
    x: (evt.clientX - rect.left) * scaleX,
    y: (evt.clientY - rect.top) * scaleY,
  };
}

// === 範囲フィルタ初期化：画像外枠から10px内側の矩形 ===
function initRangeFilterPolygonFromImage() {
  if (!originalImage) return;
  const w = canvas.width;
  const h = canvas.height;
  const margin = 10;

  rangeFilterPolygon = {
    points: [
      { x: margin,       y: margin },
      { x: w - margin,   y: margin },
      { x: w - margin,   y: h - margin },
      { x: margin,       y: h - margin },
    ],
  };

  rfDraggingVertexIndex = -1;
  rfDraggingWhole = false;
  rfDragStartMouse = null;
  rfDragStartPoints = null;

  rfRedrawOriginalWithPolygon();
}

// === ポリゴン＋元画像の再描画 ===
// 基本方針: 編集時は「元画像 only + 範囲フィルタの赤枠」を表示
function rfRedrawOriginalWithPolygon() {
  if (!originalImage) return;

  // OpenCV の Mat からキャンバスへ元画像を描画
  cv.imshow("canvas", originalImage);

  // その上に赤枠ポリゴンをオーバーレイ
  rfDrawPolygonOverlay();
}

// === 現在のキャンバス内容の上にポリゴンのみ描画（解析後の上書き用） ===
function rfDrawPolygonOverlay() {
  if (!rangeFilterPolygon || !rangeFilterPolygon.points || rangeFilterPolygon.points.length < 3) {
    return;
  }

  const pts = rangeFilterPolygon.points;
  ctx.save();

  // ポリゴン本体
  ctx.beginPath();
  ctx.moveTo(pts[0].x, pts[0].y);
  for (let i = 1; i < pts.length; i++) {
    ctx.lineTo(pts[i].x, pts[i].y);
  }
  ctx.closePath();

  // 枠線と薄い塗り
  ctx.lineWidth = 2;
  ctx.strokeStyle = "red";
  ctx.fillStyle = "rgba(255, 0, 0, 0.12)";
  ctx.stroke();
  ctx.fill();

  // 節点の丸
  ctx.fillStyle = "red";
  for (const p of pts) {
    ctx.beginPath();
    ctx.arc(p.x, p.y, 4, 0, Math.PI * 2);
    ctx.fill();
  }

  ctx.restore();
}

// === 節点ヒットテスト ===
function rfHitTestVertex(pos) {
  if (!rangeFilterPolygon) return -1;
  const pts = rangeFilterPolygon.points;
  for (let i = 0; i < pts.length; i++) {
    const dx = pts[i].x - pos.x;
    const dy = pts[i].y - pos.y;
    if (Math.hypot(dx, dy) <= RF_HANDLE_RADIUS) {
      return i;
    }
  }
  return -1;
}

// === 辺ヒットテスト (pos に最も近い辺を距離閾値付きで返す) ===
// 戻り値: { edgeIndex, t } または null
function rfHitTestEdge(pos) {
  if (!rangeFilterPolygon) return null;
  const pts = rangeFilterPolygon.points;
  const n = pts.length;
  let bestDist = Infinity;
  let bestEdge = -1;
  let bestT = 0;

  for (let i = 0; i < n; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % n];
    const vx = b.x - a.x;
    const vy = b.y - a.y;
    const wx = pos.x - a.x;
    const wy = pos.y - a.y;
    const vv = vx * vx + vy * vy;
    if (vv < 1e-6) continue;
    let t = (wx * vx + wy * vy) / vv;
    t = Math.max(0, Math.min(1, t));
    const px = a.x + vx * t;
    const py = a.y + vy * t;
    const dist = Math.hypot(pos.x - px, pos.y - py);
    if (dist < bestDist) {
      bestDist = dist;
      bestEdge = i;
      bestT = t;
    }
  }

  if (bestEdge >= 0 && bestDist <= RF_EDGE_HIT_DIST) {
    return { edgeIndex: bestEdge, t: bestT };
  }
  return null;
}

// === ポイントがポリゴン内部かどうか判定（odd-even rule） ===
function rfPointInPolygon(pos) {
  if (!rangeFilterPolygon || !rangeFilterPolygon.points || rangeFilterPolygon.points.length < 3) {
    return false;
  }
  const pts = rangeFilterPolygon.points;
  let inside = false;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const xi = pts[i].x, yi = pts[i].y;
    const xj = pts[j].x, yj = pts[j].y;
    const intersect =
      ((yi > pos.y) !== (yj > pos.y)) &&
      (pos.x < ((xj - xi) * (pos.y - yi)) / (yj - yi + 1e-9) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

// === 節点数が常に3以上になるように削除を制御 ===
function rfDeleteVertex(index) {
  if (!rangeFilterPolygon) return;
  const pts = rangeFilterPolygon.points;
  if (pts.length <= 3) {
    // C) 節数3未満にしない → 3個のときは削除しない
    return;
  }
  pts.splice(index, 1);
}

// === 線分とポリゴンの交差判定（Rule 1: segment intersects polygon） ===

function rfOrientation(p, q, r) {
  const val = (q.y - p.y) * (r.x - q.x) - (q.x - p.x) * (r.y - q.y);
  if (Math.abs(val) < 1e-9) return 0; // ほぼ一直線
  return val > 0 ? 1 : 2;             // 1:時計回り, 2:反時計回り
}

function rfOnSegment(p, q, r) {
  return (
    q.x >= Math.min(p.x, r.x) - 1e-9 &&
    q.x <= Math.max(p.x, r.x) + 1e-9 &&
    q.y >= Math.min(p.y, r.y) - 1e-9 &&
    q.y <= Math.max(p.y, r.y) + 1e-9
  );
}

// 2つの線分 (a–b) と (c–d) が交差または接しているか
function rfSegmentsIntersect(a, b, c, d) {
  const o1 = rfOrientation(a, b, c);
  const o2 = rfOrientation(a, b, d);
  const o3 = rfOrientation(c, d, a);
  const o4 = rfOrientation(c, d, b);

  // 一般位置での交差
  if (o1 !== o2 && o3 !== o4) return true;

  // コリニアな場合（端点を含む）
  if (o1 === 0 && rfOnSegment(a, c, b)) return true;
  if (o2 === 0 && rfOnSegment(a, d, b)) return true;
  if (o3 === 0 && rfOnSegment(c, a, d)) return true;
  if (o4 === 0 && rfOnSegment(c, b, d)) return true;

  return false;
}

// 1本の線分 seg がポリゴンと交差（or 内部に端点を持つ）しているか
// Rule 1: 「交差していれば丸ごと残す」仕様
function rfSegmentIntersectsPolygon(seg) {
  if (
    !rangeFilterPolygon ||
    !rangeFilterPolygon.points ||
    rangeFilterPolygon.points.length < 3
  ) {
    // ポリゴンが未定義ならフィルタ無し（全て通す）
    return true;
  }

  const a = { x: seg.x1, y: seg.y1 };
  const b = { x: seg.x2, y: seg.y2 };

  // 1) 端点が内側にある場合 → 交差ありとみなす
  if (rfPointInPolygon(a) || rfPointInPolygon(b)) {
    return true;
  }

  // 2) 線分がポリゴンの辺と交差しているかチェック
  const pts = rangeFilterPolygon.points;
  const n = pts.length;
  for (let i = 0; i < n; i++) {
    const c = pts[i];
    const d = pts[(i + 1) % n];
    if (rfSegmentsIntersect(a, b, c, d)) {
      return true;
    }
  }

  // 完全にポリゴンの外側（かつ交差なし）
  return false;
}

// 線分配列に範囲フィルタ（ポリゴン）を適用
function applyRangeFilterToSegments(segments) {
  if (
    !rangeFilterPolygon ||
    !rangeFilterPolygon.points ||
    rangeFilterPolygon.points.length < 3
  ) {
    // ポリゴン未設定ならフィルタしない
    return segments;
  }
  return segments.filter((seg) => rfSegmentIntersectsPolygon(seg));
}

// === 範囲フィルタ用マウスイベント ===
canvas.addEventListener("mousedown", (evt) => {
  if (!rangeFilterPolygon || !originalImage) return;

  const pos = rfGetMousePosOnCanvas(evt);

  // 1) 節点ヒット → 節移動 or 削除
  const vIdx = rfHitTestVertex(pos);
  if (vIdx >= 0) {
    // Ctrl または Meta (Cmd) + クリック → 即削除（A）※ただし3点未満禁止
    if (evt.ctrlKey || evt.metaKey) {
      rfDeleteVertex(vIdx);
      rfRedrawOriginalWithPolygon();
    } else {
      // 単純クリック → 節ドラッグ開始
      rfDraggingVertexIndex = vIdx;
    }
    evt.preventDefault();
    return;
  }

  // 2) 節点でない → 辺をクリックしていれば分割（節追加 B）
  const edgeHit = rfHitTestEdge(pos);
  if (edgeHit) {
    const pts = rangeFilterPolygon.points;
    const i = edgeHit.edgeIndex;
    const a = pts[i];
    const b = pts[(i + 1) % pts.length];

    // 分割点（実際にはクリック位置そのままでもよいが、辺上へ投影してもOK）
    const newPt = { x: pos.x, y: pos.y };
    pts.splice(i + 1, 0, newPt);

    // 追加した節をそのままドラッグ対象にする
    rfDraggingVertexIndex = i + 1;
    rfRedrawOriginalWithPolygon();
    evt.preventDefault();
    return;
  }

  // 3) ポリゴン内部 → 全体移動開始
  if (rfPointInPolygon(pos)) {
    rfDraggingWhole = true;
    rfDragStartMouse = pos;
    // ドラッグ開始時の位置をコピー
    rfDragStartPoints = rangeFilterPolygon.points.map((p) => ({ x: p.x, y: p.y }));
    evt.preventDefault();
    return;
  }

  // 4) それ以外（ポリゴン外部）は何もしない
});

canvas.addEventListener("mousemove", (evt) => {
  if (!rangeFilterPolygon || !originalImage) return;
  if (rfDraggingVertexIndex < 0 && !rfDraggingWhole) return;

  const pos = rfGetMousePosOnCanvas(evt);

  // キャンバス内にクランプ
  const clampX = (x) => Math.max(0, Math.min(canvas.width - 1, x));
  const clampY = (y) => Math.max(0, Math.min(canvas.height - 1, y));

  if (rfDraggingVertexIndex >= 0) {
    // 節点ドラッグ
    const pts = rangeFilterPolygon.points;
    pts[rfDraggingVertexIndex].x = clampX(pos.x);
    pts[rfDraggingVertexIndex].y = clampY(pos.y);
    rfRedrawOriginalWithPolygon();
  } else if (rfDraggingWhole && rfDragStartMouse && rfDragStartPoints) {
    // 全体ドラッグ
    const dx = pos.x - rfDragStartMouse.x;
    const dy = pos.y - rfDragStartMouse.y;

    // いったん平行移動した結果を作る
    const moved = rfDragStartPoints.map((p) => ({
      x: p.x + dx,
      y: p.y + dy,
    }));

    // 全点がキャンバス外に出過ぎないように簡易クランプ
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    for (const p of moved) {
      if (p.x < minX) minX = p.x;
      if (p.y < minY) minY = p.y;
      if (p.x > maxX) maxX = p.x;
      if (p.y > maxY) maxY = p.y;
    }

    let offX = 0, offY = 0;
    if (minX < 0) offX = -minX;
    if (minY < 0) offY = -minY;
    if (maxX > canvas.width - 1) offX = Math.min(offX, canvas.width - 1 - maxX);
    if (maxY > canvas.height - 1) offY = Math.min(offY, canvas.height - 1 - maxY);

    // オフセットを反映
    rangeFilterPolygon.points = moved.map((p) => ({
      x: p.x + offX,
      y: p.y + offY,
    }));

    rfRedrawOriginalWithPolygon();
  }

  evt.preventDefault();
});

function rfEndDrag() {
  rfDraggingVertexIndex = -1;
  rfDraggingWhole = false;
  rfDragStartMouse = null;
  rfDragStartPoints = null;
}

canvas.addEventListener("mouseup", () => {
  rfEndDrag();
});

canvas.addEventListener("mouseleave", () => {
  rfEndDrag();
});


// OpenCV.js の初期化フラグ
let cvReady = false;

// 元画像（RGBAのcv.Mat）を保持してサムネイル生成・再分析に使う
let originalImage = null;

// 短い線分を除外するための閾値（ピクセル単位） ※UIで変更するので let
let MIN_LINE_LENGTH = 20;

// === A/B/C 判定用の共通閾値（オレンジ線と黄色点線で共有する） ===
let EXT_A_MIN_PROJ = 10.0;   // 条件A: |proj| >= 10
let EXT_A_MAX_ORTH = 5.0;    // 条件A: orth <= 5

let EXT_B_MAX_ORTH = 5.0;    // 条件B: orthBoth <= 2.5

let EXT_C_MIN_PROJ = 10.0;   // 条件C: |proj_rev| >= 10
let EXT_C_MAX_ORTH = 5.0;    // 条件C: orth_rev <= 5

// === 平行一致用パラメータ ===
// 角度差しきい値（度）
let PARALLEL_ANGLE_THRESHOLD_DEG = 2.0;

// 評価線分に対して「長手方向」でどれだけ重なっていれば
// 平行一致とみなすか（0〜1 の割合）
let PARALLEL_MIN_INSTRIP_LENGTH = 0.01;

// === 同一間隔グループ用パラメータ ===
// 法線方向の間隔（px）が seed ± EPS なら同一間隔グループとみなす
let INTERVAL_EPS = 1.0;

// === 平行線比率用パラメータ ===
// 幾何級数上の生成点と実線の交点との許容距離（px）
let RATIO_POINT_EPS = 3.0;   // 

// 直近の解析結果を再分析用に保持（割合分析で使う）
let cachedLineSegments = [];
let cachedExtRelations = [];
let cachedParallelRelations = [];
let cachedEqualIntervalRelations = [];
let cachedRatioRelations = [];

// OpenCV.js 初期化コールバック
if (typeof cv !== "undefined") {
  cv["onRuntimeInitialized"] = () => {
    cvReady = true;
    resultDiv.textContent =
      "OpenCV.js の初期化が完了しました。画像を選択してください。 / OpenCV.js initialization completed. Please select an image.";
  };
} else {
  resultDiv.textContent =
    "エラー: OpenCV.js が読み込まれていません。パスを確認してください。 / Error: OpenCV.js is not loaded. Please check the script path.";
}

// ファイル選択時
input.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;

  if (!cvReady) {
    resultDiv.textContent =
      "OpenCV.js の初期化中です。少し待ってから再度お試しください。 / OpenCV.js is still initializing. Please wait and try again.";
    return;
  }

  resultDiv.textContent =
    `選択されたファイル: ${file.name}（キャンバスに表示しました。解析するには左の「分析 / 再分析を実行」ボタンを押してください。）`;

  const imgUrl = URL.createObjectURL(file);
  const img = new Image();

  img.onload = () => {
    // 元画像の実サイズ
    const origWidth = img.width;
    const origHeight = img.height;

    // 長辺
    const maxDim = Math.max(origWidth, origHeight);

    // 表示・解析に使う最終サイズ（デフォルトは元サイズ）
    let displayWidth = origWidth;
    let displayHeight = origHeight;

    // ★ n px を超える場合は、ポップアップを出してから長辺 n にスケーリング
    if (maxDim > 300) {
      alert(
        "高解像度画像のため、処理落ち防止として画像サイズを最大300ピクセルに調整します。\n" +
        "For this high-resolution image, the size will be scaled so that the longer side is 500 pixels to avoid processing slowdown."
      );
      const scale = 300 / maxDim;
      displayWidth = Math.round(origWidth * scale);
      displayHeight = Math.round(origHeight * scale);
    }

    // キャンバスサイズを（縮小後の）画像サイズに合わせる
    canvas.width = displayWidth;
    canvas.height = displayHeight;

    // 画像を描画（必要に応じて縮小済み）
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, displayWidth, displayHeight);

    // URL 開放
    URL.revokeObjectURL(imgUrl);

    // キャンバスから Mat を読み取り、originalImage に保持
    if (originalImage) {
      originalImage.delete();
      originalImage = null;
    }
    const src0 = cv.imread(canvas);   // RGBA（すでに縮小後の画像）
    originalImage = src0.clone();
    src0.delete();

    // UI の閾値を反映（必要なら）
    if (typeof applyThresholdFromUI === "function") {
      applyThresholdFromUI();
    }

    // ★ 画像サイズを使って「外枠から 10px 内側」の範囲フィルタポリゴンを初期化し、
    //    元画像＋赤ポリゴンをキャンバスに再描画
    initRangeFilterPolygonFromImage();
  };

  img.onerror = (e) => {
    console.error("画像の読み込みに失敗しました", e);
    resultDiv.textContent = "画像の読み込みに失敗しました。 / Failed to load the image.";
  };

  img.src = imgUrl;
});

// === Representative line (統合) 用パラメータ ===
let USE_REPRESENTATIVE_LINES = true;

// 初期グルーピング（Python の groups 作成に対応）
let REP_GROUP_ANGLE_THR_DEG = 3.0;     // 角度差
let REP_GROUP_NORMAL_THR_PX = 3.0;     // 法線距離
let REP_GROUP_LONG_GAP_PX = 5.0;       // 長手方向の“近さ”（少なくともどれかのメンバに近い）
let REP_MIN_GROUP_SIZE = 1;            // 1 にすると孤立線分も代表線化（=実質そのまま）

// グループ統合（Python の Union-Find マージに対応）
let REP_MERGE_ANGLE_THR_DEG = 2.5;
let REP_MERGE_NORMAL_THR_PX = 6.0;
let REP_MERGE_LONG_GAP_PX = 18.0;

// 代表線の最小長（Python の MIN_SEG_LEN に対応）
let REP_MIN_REP_LEN_PX = 1.0;

// “包含”の緩和条件（Python の is_candidate_contained_in_group の簡易版）
let REP_CONTAIN_OVERLAP_RATIO = 0.9;
let REP_CONTAIN_NORMAL_THR_PX = 3.0;

function buildRepresentativeLines(originalSegments, imgW, imgH) {
  const n = originalSegments.length;

  // ---- Python constants (合わせ込み) ----
  const PY_MIN_LEN_FOR_GROUP = 1.0;     // Python: if hypot >= 10 だけ lines[] に入れる
  const CLIP_BUFFER = 2.5;              // Python: CLIP_BUFFER
  const MIN_SEG_LEN = REP_MIN_REP_LEN_PX; // Python: MIN_SEG_LEN (=8)

  const dot2 = (a, b) => a[0] * b[0] + a[1] * b[1];
  const sub2 = (a, b) => [a[0] - b[0], a[1] - b[1]];
  const norm2 = (v) => Math.hypot(v[0], v[1]);

  function angleDiffDeg(a, b) {
    let d = Math.abs(a - b) % 180;
    if (d > 90) d = 180 - d;
    return d;
  }

  function meanAngle180(anglesDeg) {
    if (anglesDeg.length === 0) return 0;
    const base = anglesDeg[0];
    const adj = anglesDeg.map((a) => {
      let x = a;
      const d = x - base;
      if (d > 90) x -= 180;
      else if (d < -90) x += 180;
      return x;
    });
    let m = adj.reduce((s, v) => s + v, 0) / adj.length;
    m = ((m % 180) + 180) % 180;
    return m;
  }

  function clampToImage(x, y) {
    const xx = Math.max(0, Math.min(imgW - 1, x));
    const yy = Math.max(0, Math.min(imgH - 1, y));
    return [xx, yy];
  }

  function pointToSegmentDistance(px, py, ax, ay, bx, by) {
    const vx = bx - ax, vy = by - ay;
    const wx = px - ax, wy = py - ay;
    const vv = vx * vx + vy * vy;
    if (vv < 1e-12) return Math.hypot(px - ax, py - ay);
    let t = (wx * vx + wy * vy) / vv;
    if (t < 0) t = 0;
    else if (t > 1) t = 1;
    const cx = ax + t * vx;
    const cy = ay + t * vy;
    return Math.hypot(px - cx, py - cy);
  }

  // ---- segInfo (Python に合わせ：normal, mid, angle) ----
  const segInfo = [];
  for (let i = 0; i < n; i++) {
    const s = originalSegments[i];
    const dx = s.x2 - s.x1;
    const dy = s.y2 - s.y1;
    const L = Math.hypot(dx, dy);
    if (L < 1e-6) continue;

    const angle = ((Math.atan2(dy, dx) * 180) / Math.PI + 180) % 180;
    const u = [dx / L, dy / L];
    const nrm = [-u[1], u[0]];
    const mid = [(s.x1 + s.x2) / 2.0, (s.y1 + s.y2) / 2.0];

    segInfo.push({
      idx: i,
      angle,
      u,
      n: nrm,
      mid,
      p1: [s.x1, s.y1],
      p2: [s.x2, s.y2],
      length: L,
      eligible: L >= PY_MIN_LEN_FOR_GROUP, // Python: lines[] に入る条件
    });
  }

  const posByIdx = new Map();
  for (let k = 0; k < segInfo.length; k++) posByIdx.set(segInfo[k].idx, k);

  // ---- Python: is_candidate_contained_in_group (buffer(3) + intersection length >= 0.9*len) の近似 ----
  // 「candidate 線分上の点をサンプリングし、member までの距離 <= 3 の割合」を intersection-length の近似とする
  function containedByAnyMember(candidatePos, groupIdxList) {
    const cand = segInfo[candidatePos];
    const ax = cand.p1[0], ay = cand.p1[1], bx = cand.p2[0], by = cand.p2[1];

    const SAMPLES = 25; // 近似精度（必要なら増やす）
    for (const memberIdx of groupIdxList) {
      const mPos = posByIdx.get(memberIdx);
      if (mPos === undefined) continue;
      const mem = segInfo[mPos];

      let insideCount = 0;
      for (let t = 0; t < SAMPLES; t++) {
        const r = t / (SAMPLES - 1);
        const px = ax + (bx - ax) * r;
        const py = ay + (by - ay) * r;
        const d = pointToSegmentDistance(px, py, mem.p1[0], mem.p1[1], mem.p2[0], mem.p2[1]);
        if (d <= 3.0) insideCount++;
      }

      const ratio = insideCount / SAMPLES;
      if (ratio >= 0.9) return true; // Python: 0.9 * candidate length
    }
    return false;
  }

  // ---- 1) 初期グループ化（Python BFS と同等） ----
  const visited = new Set();
  const groups = [];

  for (let a = 0; a < segInfo.length; a++) {
    const seed = segInfo[a];
    if (!seed.eligible) continue;
    if (visited.has(seed.idx)) continue;

    const group = [seed.idx];
    visited.add(seed.idx);
    const queue = [seed.idx];

    while (queue.length > 0) {
      const curIdx = queue.pop();

      for (let b = 0; b < segInfo.length; b++) {
        const cand = segInfo[b];
        if (!cand.eligible) continue;
        if (visited.has(cand.idx)) continue;

        // angle_consistent: group 全員と角度差 <= thr
        const okAngle = group.every((mIdx) => {
          const mPos = posByIdx.get(mIdx);
          return angleDiffDeg(segInfo[mPos].angle, cand.angle) <= REP_GROUP_ANGLE_THR_DEG;
        });
        if (!okAngle) continue;

        // normal_dist: group 全員に対して <= thr
        let okNormalAll = true;
        for (const mIdx of group) {
          const mPos = posByIdx.get(mIdx);
          const mem = segInfo[mPos];
          const delta = sub2(cand.mid, mem.mid);
          const normalDist = Math.abs(dot2(delta, mem.n));
          if (normalDist > REP_GROUP_NORMAL_THR_PX) {
            okNormalAll = false;
            break;
          }
        }
        if (!okNormalAll) continue;

        // aligned_with_any: どれか 1 本に長手距離が近い
        let alignedWithAny = false;
        for (const mIdx of group) {
          const mPos = posByIdx.get(mIdx);
          const mem = segInfo[mPos];
          const delta = sub2(cand.mid, mem.mid);
          const longDist = Math.abs(dot2(delta, mem.u));
          if (longDist <= REP_GROUP_LONG_GAP_PX) {
            alignedWithAny = true;
            break;
          }
        }

        // Python: aligned false のとき contained 判定で救済
        if (!alignedWithAny) {
          if (!containedByAnyMember(b, group)) continue;
        }

        group.push(cand.idx);
        visited.add(cand.idx);
        queue.push(cand.idx);
      }
    }

    if (group.length >= REP_MIN_GROUP_SIZE) groups.push(group);
  }

  // ---- 2) グループ統合（Union-Find：Python と同様） ----
  const gStats = groups.map((g, gid) => {
    const angles = g.map((idx) => segInfo[posByIdx.get(idx)].angle);
    const avgAngle = meanAngle180(angles);

    const mids = g.map((idx) => segInfo[posByIdx.get(idx)].mid);
    const center = [
      mids.reduce((s, v) => s + v[0], 0) / mids.length,
      mids.reduce((s, v) => s + v[1], 0) / mids.length,
    ];

    const rad = (avgAngle * Math.PI) / 180;
    const u = [Math.cos(rad), Math.sin(rad)];
    const nrm = [-u[1], u[0]];
    return { gid, avgAngle, u, n: nrm, center };
  });

  const parent = Array.from({ length: groups.length }, (_, i) => i);
  function find(x) {
    while (parent[x] !== x) {
      parent[x] = parent[parent[x]];
      x = parent[x];
    }
    return x;
  }
  function union(a, b) {
    const ra = find(a), rb = find(b);
    if (ra !== rb) parent[rb] = ra;
  }

  function shouldMerge(gsI, gsJ) {
    const da = angleDiffDeg(gsI.avgAngle, gsJ.avgAngle);
    if (da > REP_MERGE_ANGLE_THR_DEG) return false;

    const nMean = [gsI.n[0] + gsJ.n[0], gsI.n[1] + gsJ.n[1]];
    const nm = norm2(nMean);
    const nUse = nm < 1e-9 ? gsI.n : [nMean[0] / nm, nMean[1] / nm];
    const delta = sub2(gsJ.center, gsI.center);
    const normalSep = Math.abs(dot2(delta, nUse));
    if (normalSep > REP_MERGE_NORMAL_THR_PX) return false;

    const uMean = [gsI.u[0] + gsJ.u[0], gsI.u[1] + gsJ.u[1]];
    const um = norm2(uMean);
    const uUse = um < 1e-9 ? gsI.u : [uMean[0] / um, uMean[1] / um];
    const longSep = Math.abs(dot2(delta, uUse));
    if (longSep > REP_MERGE_LONG_GAP_PX) return false;

    return true;
  }

  if (gStats.length <= 400) {
    for (let i = 0; i < gStats.length; i++) {
      for (let j = i + 1; j < gStats.length; j++) {
        if (shouldMerge(gStats[i], gStats[j])) union(i, j);
      }
    }
  }

  const mergedMap = new Map();
  for (let gi = 0; gi < groups.length; gi++) {
    const r = find(gi);
    if (!mergedMap.has(r)) mergedMap.set(r, []);
    mergedMap.get(r).push(...groups[gi]);
  }
  const mergedGroups = Array.from(mergedMap.values());

  // ---- 3) 代表線生成（Python: t0/t1 + offset + band で最終クリップ） ----
  // band(intersection) の近似：代表線上をサンプルし、member までの距離 <= CLIP_BUFFER の連続区間を採用
  function clipSegmentToBand(A0, B0, memberIdxList) {
    const ax = A0[0], ay = A0[1], bx = B0[0], by = B0[1];
    const L = Math.hypot(bx - ax, by - ay);
    if (L < 1e-6) return null;

    const S = Math.max(60, Math.min(400, Math.ceil(L * 0.8))); // 長いほど増やす（上限400）
    let bestLen = 0;
    let bestI0 = -1, bestI1 = -1;

    let curI0 = -1;
    for (let i = 0; i < S; i++) {
      const r = i / (S - 1);
      const px = ax + (bx - ax) * r;
      const py = ay + (by - ay) * r;

      let inside = false;
      for (const midx of memberIdxList) {
        const mp = posByIdx.get(midx);
        if (mp === undefined) continue;
        const mem = segInfo[mp];
        const d = pointToSegmentDistance(px, py, mem.p1[0], mem.p1[1], mem.p2[0], mem.p2[1]);
        if (d <= CLIP_BUFFER) { inside = true; break; }
      }

      if (inside) {
        if (curI0 < 0) curI0 = i;
      } else {
        if (curI0 >= 0) {
          const curLen = (i - 1 - curI0) * (L / (S - 1));
          if (curLen > bestLen) { bestLen = curLen; bestI0 = curI0; bestI1 = i - 1; }
          curI0 = -1;
        }
      }
    }
    if (curI0 >= 0) {
      const curLen = (S - 1 - curI0) * (L / (S - 1));
      if (curLen > bestLen) { bestLen = curLen; bestI0 = curI0; bestI1 = S - 1; }
    }
    if (bestI0 < 0) return null;

    const r0 = bestI0 / (S - 1);
    const r1 = bestI1 / (S - 1);
    const A = [ax + (bx - ax) * r0, ay + (by - ay) * r0];
    const B = [ax + (bx - ax) * r1, ay + (by - ay) * r1];
    if (Math.hypot(B[0] - A[0], B[1] - A[1]) < MIN_SEG_LEN) return null;
    return { A, B };
  }

  const repLines = [];
  const segmentGroupIds = new Array(n).fill(-1);

  for (let gid = 0; gid < mergedGroups.length; gid++) {
    const g = mergedGroups[gid];

    const angles = g.map((idx) => segInfo[posByIdx.get(idx)].angle);
    const avgAngle = meanAngle180(angles);
    const rad = (avgAngle * Math.PI) / 180;
    const u = [Math.cos(rad), Math.sin(rad)];
    const nrm = [-u[1], u[0]];

    const mids = g.map((idx) => segInfo[posByIdx.get(idx)].mid);
    const origin = [
      mids.reduce((s, v) => s + v[0], 0) / mids.length,
      mids.reduce((s, v) => s + v[1], 0) / mids.length,
    ];

    // endpoints の t 射影 min/max
    const ts = [];
    for (const idx of g) {
      const si = segInfo[posByIdx.get(idx)];
      ts.push(dot2(sub2(si.p1, origin), u));
      ts.push(dot2(sub2(si.p2, origin), u));
    }
    const t0 = Math.min(...ts);
    const t1 = Math.max(...ts);
    if ((t1 - t0) < MIN_SEG_LEN) continue;

    // offset = median(dot(mid-origin, n))
    const offs = mids.map((m) => dot2(sub2(m, origin), nrm)).sort((a, b) => a - b);
    const med = offs[Math.floor(offs.length / 2)] || 0;

    const basePt = [origin[0] + nrm[0] * med, origin[1] + nrm[1] * med];
    const A0 = [basePt[0] + u[0] * t0, basePt[1] + u[1] * t0];
    const B0 = [basePt[0] + u[0] * t1, basePt[1] + u[1] * t1];

    // Python の band intersection を近似
    const clipped = clipSegmentToBand(A0, B0, g);
    if (!clipped) continue;

    // 仕上げに画像内 clamp（念のため）
    let A = clampToImage(clipped.A[0], clipped.A[1]);
    let B = clampToImage(clipped.B[0], clipped.B[1]);

    const L = Math.hypot(B[0] - A[0], B[1] - A[1]);
    if (L < MIN_SEG_LEN) continue;

    const x1 = Math.round(A[0]);
    const y1 = Math.round(A[1]);
    const x2 = Math.round(B[0]);
    const y2 = Math.round(B[1]);
    const cx = (x1 + x2) / 2.0;
    const cy = (y1 + y2) / 2.0;

    repLines.push({
      x1, y1, x2, y2,
      length: L,
      angle: avgAngle,
      cx, cy,
      groupId: gid,
      members: g.slice(),
    });

    for (const idx of g) segmentGroupIds[idx] = gid;
  }

  return { repLines, segmentGroupIds };
}


// === 線分抽出方式（Python寄せ） ===
// "LSD" を優先。OpenCV.js ビルドに LSD が無い場合は自動で "HOUGHP" にフォールバック。
let LINE_EXTRACT_METHOD = "LSD"; // "LSD" | "HOUGHP"

// HoughP 用（必要に応じて調整）
let HOUGH_RHO = 1;
let HOUGH_THETA = Math.PI / 180;
let HOUGH_THRESHOLD = 100;
let HOUGH_MIN_LINE_LENGTH = 20; // ここは MIN_LINE_LENGTH と揃えてもOK
let HOUGH_MAX_LINE_GAP = 5;

// ===== HoughP を「補助」に抑えるための制御（過剰抽出対策） =====

// LSD が一定数以上拾えているなら HoughP をスキップ（補助のみ）
let HOUGH_SKIP_IF_LSD_AT_LEAST = 300;   // まずは 400 推奨（画像次第で 200〜800）

// LSD がある時に HoughP を回す場合でも、追加採用する本数に上限を設ける
let HOUGH_MAX_ADD_WHEN_LSD_PRESENT = 100; // まずは 250 推奨

// Canny 多パス（やりすぎ禁止：2パスまで）
let HOUGH_CANNY_PASSES = [
  { low: 50, high: 150, aperture: 3 },  // 標準
  { low: 30, high: 100, aperture: 3 },  // 細線の救済（控えめ）
];

// 途切れ接続はデフォルトOFF（ONにすると一気に増えやすい）
let HOUGH_USE_MORPH_CLOSE = false;      // 必要になったら true
let HOUGH_MORPH_ITER = 1;               // ON時の反復回数


// LSD 用（OpenCVの実装により引数の有無が異なるため、基本はデフォルトで使う）
let LSD_REFINE = 0; // 0: LSD_REFINE_NONE 相当（環境により定数が無いことがある）

function analyzeFromOriginalImage() {
  if (!cvReady) {
    resultDiv.textContent =
      "OpenCV.js の初期化が完了していません。 / OpenCV.js is not initialized yet.";
    return;
  }
  if (!originalImage) {
    resultDiv.textContent =
      "解析対象の画像がありません。 / No image is available for analysis.";
    return;
  }

  // ==== Python版に寄せるための定数（必要ならUI化してもOK）====
  const CANNY_LOW = 20;          // Python: cv2.Canny(base_img, 50, 110)
  const CANNY_HIGH = 110;
  const APPROX_EPS = 2.0;        // Python: approxPolyDP epsilon=2.0
  const OUTLIER_DIST_THR = 1.5;  // Python: threshold=1.5

  // UIの最短線分長をそのまま使う（ハード下限なし）
  const MIN_SRC_SEG_LEN = Number(MIN_LINE_LENGTH) || 0;

  const CLIP_BUFFER = 2.5;       // Python: CLIP_BUFFER=2.5（帯の厚み）
  const MIN_REP_LEN = 1.0;       // Python: MIN_SEG_LEN=8.0

  // OpenCV 定数が無いビルド向けフォールバック
  const DIST_L2 = typeof cv.DIST_L2 !== "undefined" ? cv.DIST_L2 : 2;

  // Mat.zeros が無い/挙動が怪しい環境向けフォールバック
  function matZeros(rows, cols, type) {
    if (cv.Mat && typeof cv.Mat.zeros === "function") {
      // 正：new は付けない
      return cv.Mat.zeros(rows, cols, type);
    }
    const m = new cv.Mat(rows, cols, type);
    m.setTo(new cv.Scalar(0));
    return m;
  }

  // ---- 幾何ユーティリティ ----
  function pointToSegmentDist(px, py, ax, ay, bx, by) {
    const vx = bx - ax;
    const vy = by - ay;
    const wx = px - ax;
    const wy = py - ay;

    const vv = vx * vx + vy * vy;
    if (vv < 1e-12) return Math.hypot(px - ax, py - ay);

    let t = (wx * vx + wy * vy) / vv;
    t = Math.max(0, Math.min(1, t));
    const cx = ax + t * vx;
    const cy = ay + t * vy;
    return Math.hypot(px - cx, py - cy);
  }

  function makeSegmentObj(x1, y1, x2, y2) {
    const dx = x2 - x1;
    const dy = y2 - y1;
    const length = Math.hypot(dx, dy);
    if (length < MIN_SRC_SEG_LEN) return null;

    let angleDeg = (Math.atan2(dy, dx) * 180) / Math.PI;
    angleDeg = (angleDeg + 180) % 180;

    const cx = (x1 + x2) / 2.0;
    const cy = (y1 + y2) / 2.0;

    return {
      x1: Math.round(x1),
      y1: Math.round(y1),
      x2: Math.round(x2),
      y2: Math.round(y2),
      length,
      angle: angleDeg,
      cx,
      cy,
      groupId: -1,
    };
  }

  // ---- Representative線を「帯（buffer）で最終クリップ」する簡易近似 ----
  function clipRepLineByMembers(repLine, allSegments) {
    if (!repLine || !repLine.members || repLine.members.length === 0) return repLine;

    const memSegs = [];
    for (const idx of repLine.members) {
      const s = allSegments[idx];
      if (!s) continue;
      memSegs.push(s);
    }
    if (memSegs.length === 0) return null;

    const Ax = repLine.x1, Ay = repLine.y1;
    const Bx = repLine.x2, By = repLine.y2;

    const vx = Bx - Ax;
    const vy = By - Ay;
    const L = Math.hypot(vx, vy);
    if (L < 1e-6) return null;

    let step = 2.0;
    if (L > 2000) step = 4.0;
    if (L > 4000) step = 6.0;

    const N = Math.max(2, Math.floor(L / step) + 1);
    const inside = new Array(N).fill(false);

    for (let i = 0; i < N; i++) {
      const t = i / (N - 1);
      const px = Ax + vx * t;
      const py = Ay + vy * t;

      let minD = Infinity;
      for (const ms of memSegs) {
        const d = pointToSegmentDist(px, py, ms.x1, ms.y1, ms.x2, ms.y2);
        if (d < minD) minD = d;
        if (minD <= CLIP_BUFFER) break;
      }
      inside[i] = minD <= CLIP_BUFFER;
    }

    let bestS = -1, bestE = -1, bestLen = 0;
    let curS = -1;

    for (let i = 0; i < N; i++) {
      if (inside[i]) {
        if (curS < 0) curS = i;
      } else {
        if (curS >= 0) {
          const curE = i - 1;
          const curLen = curE - curS + 1;
          if (curLen > bestLen) {
            bestLen = curLen;
            bestS = curS;
            bestE = curE;
          }
          curS = -1;
        }
      }
    }
    if (curS >= 0) {
      const curE = N - 1;
      const curLen = curE - curS + 1;
      if (curLen > bestLen) {
        bestLen = curLen;
        bestS = curS;
        bestE = curE;
      }
    }

    if (bestLen <= 1) return null;

    const tS = bestS / (N - 1);
    const tE = bestE / (N - 1);

    const nx1 = Ax + vx * tS;
    const ny1 = Ay + vy * tS;
    const nx2 = Ax + vx * tE;
    const ny2 = Ay + vy * tE;

    const newL = Math.hypot(nx2 - nx1, ny2 - ny1);
    if (newL < MIN_REP_LEN) return null;

    repLine.x1 = Math.round(nx1);
    repLine.y1 = Math.round(ny1);
    repLine.x2 = Math.round(nx2);
    repLine.y2 = Math.round(ny2);
    repLine.length = newL;
    repLine.cx = (repLine.x1 + repLine.x2) / 2.0;
    repLine.cy = (repLine.y1 + repLine.y2) / 2.0;

    return repLine;
  }

  // ==== 本処理 ====
  const src = originalImage.clone(); // RGBA（表示用にこのMatへ描画）
  const matsToDelete = [];

  try {
    // 1) Python相当: Canny
    const gray = new cv.Mat();
    cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
    matsToDelete.push(gray);

    const edges = new cv.Mat();
    cv.Canny(gray, edges, CANNY_LOW, CANNY_HIGH, 3, false);
    matsToDelete.push(edges);

    // 2) Python相当: findContours(RETR_EXTERNAL) → approxPolyDP → 連続点を線分化
    const contours = new cv.MatVector();
    const hierarchy = new cv.Mat();
    cv.findContours(edges, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    matsToDelete.push(contours, hierarchy);

    const originalLinesRaw = [];

    for (let i = 0; i < contours.size(); i++) {
      const cnt = contours.get(i);
      const approx = new cv.Mat();

      cv.approxPolyDP(cnt, approx, APPROX_EPS, false);

      const d = approx.data32S;
      if (d && d.length >= 4) {
        const pts = [];
        for (let k = 0; k + 1 < d.length; k += 2) {
          pts.push([d[k], d[k + 1]]);
        }

        for (let p = 0; p < pts.length - 1; p++) {
          const [x1, y1] = pts[p];
          const [x2, y2] = pts[p + 1];
          originalLinesRaw.push([x1, y1, x2, y2]);
        }
      }

      approx.delete();
      cnt.delete();
    }

    // 3) Python相当: Outliers抽出（distanceTransform 近似）
    const h = edges.rows;
    const w = edges.cols;

    const lineMask = matZeros(h, w, cv.CV_8UC1);
    matsToDelete.push(lineMask);

    for (const [x1, y1, x2, y2] of originalLinesRaw) {
      cv.line(
        lineMask,
        new cv.Point(x1, y1),
        new cv.Point(x2, y2),
        new cv.Scalar(255),
        1
      );
    }

    const inv = new cv.Mat();
    cv.bitwise_not(lineMask, inv);
    matsToDelete.push(inv);

    const dist = new cv.Mat();
    cv.distanceTransform(inv, dist, DIST_L2, 3);
    matsToDelete.push(dist);

    const outlierImg = matZeros(h, w, cv.CV_8UC1);
    matsToDelete.push(outlierImg);

    const eData = edges.data;
    const dData = dist.data32F;
    const oData = outlierImg.data;

    const total = eData.length;
    for (let idx = 0; idx < total; idx++) {
      if (eData[idx] === 0) continue;
      if (dData[idx] > OUTLIER_DIST_THR) {
        oData[idx] = 255;
      }
    }

    // 4) Python相当: Outliers画像から findContours→approxPolyDP→線分化
    const outContours = new cv.MatVector();
    const outHierarchy = new cv.Mat();
    cv.findContours(outlierImg, outContours, outHierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
    matsToDelete.push(outContours, outHierarchy);

    const outlierLinesRaw = [];

    for (let i = 0; i < outContours.size(); i++) {
      const cnt = outContours.get(i);
      const approx = new cv.Mat();

      cv.approxPolyDP(cnt, approx, APPROX_EPS, false);

      const d = approx.data32S;
      if (d && d.length >= 4) {
        const pts = [];
        for (let k = 0; k + 1 < d.length; k += 2) {
          pts.push([d[k], d[k + 1]]);
        }

        for (let p = 0; p < pts.length - 1; p++) {
          const [x1, y1] = pts[p];
          const [x2, y2] = pts[p + 1];
          outlierLinesRaw.push([x1, y1, x2, y2]);
        }
      }

      approx.delete();
      cnt.delete();
    }

    // 5) Python相当: CombinedLines（逆さ凹フィルタは不採用なので Filtered=Combined とみなす）
    const combinedRaw = originalLinesRaw.concat(outlierLinesRaw);

    // 6) 後段が扱える形へ変換
    const lineSegmentsRaw = [];
    for (const [x1, y1, x2, y2] of combinedRaw) {
      const s = makeSegmentObj(x1, y1, x2, y2);
      if (s) lineSegmentsRaw.push(s);
    }

    // ★ 範囲フィルタ（ポリゴン：Rule 1「交差していれば丸ごと残す」）を適用
    const lineSegments = applyRangeFilterToSegments(lineSegmentsRaw);

    // 7) RepresentativeLines（既存 buildRepresentativeLines を使用）
    const imgW = src.cols;
    const imgH = src.rows;

    const { repLines, segmentGroupIds } = buildRepresentativeLines(lineSegments, imgW, imgH);

    for (let i = 0; i < lineSegments.length; i++) {
      lineSegments[i].groupId = segmentGroupIds[i] ?? -1;
    }

    const clippedRepLines = [];
    for (const rl of repLines) {
      const clipped = clipRepLineByMembers({ ...rl }, lineSegments);
      if (!clipped) continue;
      if (clipped.length >= MIN_REP_LEN) clippedRepLines.push(clipped);
    }

    // 8) 後段解析に使う線分配列（代表線が作れたら置き換え）
    const analysisSegments =
      USE_REPRESENTATIVE_LINES && clippedRepLines.length > 0 ? clippedRepLines : lineSegments;

    // 9) 表示（緑線）
    for (const seg of analysisSegments) {
      cv.line(
        src,
        new cv.Point(seg.x1, seg.y1),
        new cv.Point(seg.x2, seg.y2),
        new cv.Scalar(0, 255, 0, 255),
        2
      );
    }
    cv.imshow("canvas", src);

    // 10) 後段解析（既存ロジック）
    const extRelations = analyzeExtensionRelations(analysisSegments);
    const parallelRelations = analyzeParallelRelations(analysisSegments);
    const equalIntervalRelations = analyzeEqualIntervalRelations(analysisSegments);
    const ratioRelations = analyzeParallelRatioRelations(analysisSegments);

    cachedLineSegments = analysisSegments;
    cachedExtRelations = extRelations;
    cachedParallelRelations = parallelRelations;
    cachedEqualIntervalRelations = equalIntervalRelations;
    cachedRatioRelations = ratioRelations;

    showLineInfo(
      analysisSegments,
      extRelations,
      parallelRelations,
      equalIntervalRelations,
      ratioRelations
    );

    // ★ 解析描画の上に範囲フィルタポリゴンを再度オーバーレイ
    rfDrawPolygonOverlay();
  } catch (err) {
    console.error(err);
    resultDiv.textContent =
      "解析中にエラーが発生しました。コンソールを確認してください。 / An error occurred during analysis. Please check the console.";
  } finally {
    for (const m of matsToDelete) {
      try {
        if (m && typeof m.delete === "function") m.delete();
      } catch (_) {}
    }
    try { src.delete(); } catch (_) {}
  }
}


/**
 * 延長一致に基づく A/B/C 関係を JS 版で簡易実装。
 * 代表線ではなく、検出線分配列 lineSegments[] をそのまま対象とする。
 */
function analyzeExtensionRelations(lineSegments) {
  const n = lineSegments.length;
  const result = [];

  function angleDiffDeg(a, b) {
    let d = Math.abs(a - b) % 180;
    if (d > 90) d = 180 - d;
    return d;
  }

  for (let i = 0; i < n; i++) {
    const segI = lineSegments[i];
    const pt1_i = [segI.x1, segI.y1];
    const pt2_i = [segI.x2, segI.y2];
    const base_vec = [pt2_i[0] - pt1_i[0], pt2_i[1] - pt1_i[1]];
    const norm = Math.hypot(base_vec[0], base_vec[1]);

    if (norm < 1e-6) {
      result.push({
        side1Match: false,
        side2Match: false,
        matchedA: [],
        matchedB: [],
        matchedC: [],
        summary: "Both side Non",
      });
      continue;
    }

    const unit_vec = [base_vec[0] / norm, base_vec[1] / norm];

    let side1Match = false;
    let side2Match = false;
    const matchedA = [];
    const matchedB = [];
    const matchedC = [];

    for (let j = 0; j < n; j++) {
      if (i === j) continue;

      const segJ = lineSegments[j];
      const pt1_j = [segJ.x1, segJ.y1];
      const pt2_j = [segJ.x2, segJ.y2];
      const vec_j = [pt2_j[0] - pt1_j[0], pt2_j[1] - pt1_j[1]];
      const norm_j = Math.hypot(vec_j[0], vec_j[1]);
      if (norm_j < 1e-6) continue;
      const unit_vec_j = [vec_j[0] / norm_j, vec_j[1] / norm_j];

      const endpoints_i = [pt1_i, pt2_i];
      const endpoints_j = [pt1_j, pt2_j];

      const cross2D = (u, v) => u[0] * v[1] - u[1] * v[0];
      const dot2D = (u, v) => u[0] * v[0] + u[1] * v[1];

      // --- 条件A ---
      let minDist = Infinity;
      let closestPair = null;
      for (const a of endpoints_i) {
        for (const b of endpoints_j) {
          const dx = a[0] - b[0];
          const dy = a[1] - b[1];
          const d = Math.hypot(dx, dy);
          if (d < minDist) {
            minDist = d;
            closestPair = { a, b };
          }
        }
      }
      if (closestPair) {
        const a = closestPair.a;
        const b = closestPair.b;
        const vec_ab = [b[0] - a[0], b[1] - a[1]];
        const proj = dot2D(vec_ab, unit_vec);
        const orth = Math.abs(cross2D(unit_vec, vec_ab));
        const angDiff = angleDiffDeg(segI.angle, segJ.angle);

        const other_a = a === pt1_i ? pt2_i : pt1_i;
        const other_b = b === pt1_j ? pt2_j : pt1_j;
        const orth_a_other = Math.abs(
          cross2D(unit_vec, [other_a[0] - pt1_i[0], other_a[1] - pt1_i[1]])
        );
        const orth_b_other = Math.abs(
          cross2D(unit_vec, [other_b[0] - pt1_i[0], other_b[1] - pt1_i[1]])
        );

        if (
          Math.abs(proj) >= EXT_A_MIN_PROJ &&
          orth <= EXT_A_MAX_ORTH &&
          angDiff <= 90 &&
          (orth_a_other > EXT_A_MAX_ORTH || orth_b_other > EXT_A_MAX_ORTH)
        ) {
          matchedA.push(j);
          // どちら側か
          const d1 = Math.hypot(a[0] - pt1_i[0], a[1] - pt1_i[1]);
          const d2 = Math.hypot(a[0] - pt2_i[0], a[1] - pt2_i[1]);
          if (d1 < d2) side1Match = true;
          else side2Match = true;
        }
      }

      // --- 条件B ---
      const orthBoth = endpoints_j.map((p) => {
        const v = [p[0] - pt1_i[0], p[1] - pt1_i[1]];
        return Math.abs(cross2D(unit_vec, v));
      });

      if (orthBoth.every((od) => od <= EXT_B_MAX_ORTH)) {
        matchedB.push(j);

        // 一番近い端点を見て側判定
        let minDistB = Infinity;
        let closestSelf = null;
        for (const se of endpoints_i) {
          for (const oe of endpoints_j) {
            const dx = se[0] - oe[0];
            const dy = se[1] - oe[1];
            const d = Math.hypot(dx, dy);
            if (d < minDistB) {
              minDistB = d;
              closestSelf = se;
            }
          }
        }
        if (closestSelf) {
          const d1 = Math.hypot(
            closestSelf[0] - pt1_i[0],
            closestSelf[1] - pt1_i[1]
          );
          const d2 = Math.hypot(
            closestSelf[0] - pt2_i[0],
            closestSelf[1] - pt2_i[1]
          );
          if (d1 < d2) side1Match = true;
          else side2Match = true;
        }
      }

      // --- 条件C ---
      let minDistRev = Infinity;
      let closestPairRev = null;
      for (const a of endpoints_j) {
        for (const b of endpoints_i) {
          const dx = a[0] - b[0];
          const dy = a[1] - b[1];
          const d = Math.hypot(dx, dy);
          if (d < minDistRev) {
            minDistRev = d;
            closestPairRev = { a, b };
          }
        }
      }
      if (closestPairRev) {
        const a_rev = closestPairRev.a;
        const b_rev = closestPairRev.b;
        const vec_ab_rev = [b_rev[0] - a_rev[0], b_rev[1] - a_rev[1]];
        const proj_rev = dot2D(vec_ab_rev, unit_vec_j);
        const orth_rev = Math.abs(cross2D(unit_vec_j, vec_ab_rev));
        const angDiffRev = angleDiffDeg(segJ.angle, segI.angle);

        const other_a_rev = a_rev === pt1_j ? pt2_j : pt1_j;
        const other_b_rev = b_rev === pt1_i ? pt2_i : pt1_i;
        const orth_a_other_rev = Math.abs(
          cross2D(unit_vec_j, [
            other_a_rev[0] - pt1_j[0],
            other_a_rev[1] - pt1_j[1],
          ])
        );
        const orth_b_other_rev = Math.abs(
          cross2D(unit_vec_j, [
            other_b_rev[0] - pt1_j[0],
            other_b_rev[1] - pt1_j[1],
          ])
        );

        if (
          Math.abs(proj_rev) >= EXT_C_MIN_PROJ &&
          orth_rev <= EXT_C_MAX_ORTH &&
          angDiffRev <= 90 &&
          (orth_a_other_rev > EXT_C_MAX_ORTH || orth_b_other_rev > EXT_C_MAX_ORTH)
        ) {
          matchedC.push(j);
        }
      }
    }

    const desc = [];
    if (!side1Match && !side2Match) {
      desc.push("Both side Non");
    } else if (!side1Match || !side2Match) {
      desc.push("One side Non");
    }
    if (matchedA.length > 0) {
      desc.push("Matched A: " + matchedA.join(","));
    }
    if (matchedB.length > 0) {
      desc.push("Matched B: " + matchedB.join(","));
    }
    if (matchedC.length > 0) {
      desc.push("Matched C: " + matchedC.join(","));
    }

    result.push({
      side1Match,
      side2Match,
      matchedA,
      matchedB,
      matchedC,
      summary: desc.join(" ; "),
    });
  }

  return result;
}

/**
 * 平行一致の解析
 */
function analyzeParallelRelations(lineSegments) {
  const n = lineSegments.length;
  const result = [];

  function angleSimilar(a, b, thresholdDeg) {
    let diff = Math.abs(a - b) % 180;
    return diff <= thresholdDeg || diff >= 180 - thresholdDeg;
  }

  // 評価線分 segI に対して、長手方向 unitVec でどれだけ重なっているか（0〜1）
  function overlapRatioOnLength(segI, segJ, unitVec) {
    const p1 = [segI.x1, segI.y1];
    const p2 = [segI.x2, segI.y2];
    const q1 = [segJ.x1, segJ.y1];
    const q2 = [segJ.x2, segJ.y2];

    const dot2D = (u, v) => u[0] * v[0] + u[1] * v[1];

    const proj_p1 = dot2D(p1, unitVec);
    const proj_p2 = dot2D(p2, unitVec);
    const p_min = Math.min(proj_p1, proj_p2);
    const p_max = Math.max(proj_p1, proj_p2);

    const proj_q1 = dot2D(q1, unitVec);
    const proj_q2 = dot2D(q2, unitVec);
    const q_min = Math.min(proj_q1, proj_q2);
    const q_max = Math.max(proj_q1, proj_q2);

    const overlapLen = Math.max(0, Math.min(p_max, q_max) - Math.max(p_min, q_min));
    const len_p = p_max - p_min;
    if (len_p <= 0) return 0;
    return overlapLen / len_p;
  }

  for (let i = 0; i < n; i++) {
    const segI = lineSegments[i];
    const baseVec = [segI.x2 - segI.x1, segI.y2 - segI.y1];
    const norm = Math.hypot(baseVec[0], baseVec[1]);
    if (norm < 1e-6) {
      result.push({ matchedParallel: [], summary: "length too short" });
      continue;
    }
    const unitVec = [baseVec[0] / norm, baseVec[1] / norm];

    const matchedParallel = [];

    for (let j = 0; j < n; j++) {
      if (i === j) continue;

      const segJ = lineSegments[j];

      // 角度が近くなければスキップ（UI の閾値を使用）
      if (!angleSimilar(segI.angle, segJ.angle, PARALLEL_ANGLE_THRESHOLD_DEG)) continue;

      // 長手方向オーバーラップ率（UI の閾値を使用）
      const ratio = overlapRatioOnLength(segI, segJ, unitVec);
      if (ratio < PARALLEL_MIN_INSTRIP_LENGTH) continue;

      matchedParallel.push(j);
    }

    let summary = "";
    if (matchedParallel.length === 0) {
      summary = "No parallel overlap";
    } else {
      summary = "Parallel with: " + matchedParallel.join(",");
    }

    result.push({
      matchedParallel,
      summary,
    });
  }

  return result;
}

/**
 * 同一間隔の解析
 */
function analyzeEqualIntervalRelations(lineSegments) {
  const n = lineSegments.length;

  // グローバルな「間隔グループ」リスト
  const groups = []; // { seed:number, pairs:[{a:number,b:number}], dists:number[] }

  function angleSimilar(a, b, thresholdDeg) {
    let diff = Math.abs(a - b) % 180;
    return diff <= thresholdDeg || diff >= 180 - thresholdDeg;
  }

  function overlapRatioOnLength(segI, segJ, unitVec) {
    const p1 = [segI.x1, segI.y1];
    const p2 = [segI.x2, segI.y2];
    const q1 = [segJ.x1, segJ.y1];
    const q2 = [segJ.x2, segJ.y2];

    const dot2D = (u, v) => u[0] * v[0] + u[1] * v[1];

    const proj_p1 = dot2D(p1, unitVec);
    const proj_p2 = dot2D(p2, unitVec);
    const p_min = Math.min(proj_p1, proj_p2);
    const p_max = Math.max(proj_p1, proj_p2);

    const proj_q1 = dot2D(q1, unitVec);
    const proj_q2 = dot2D(q2, unitVec);
    const q_min = Math.min(proj_q1, proj_q2);
    const q_max = Math.max(proj_q1, proj_q2);

    const overlapLen = Math.max(0, Math.min(p_max, q_max) - Math.max(p_min, q_min));
    const len_p = p_max - p_min;
    if (len_p <= 0) return 0;
    return overlapLen / len_p;
  }

  // --- 距離グループ構成 ---
  for (let i = 0; i < n; i++) {
    const segI = lineSegments[i];
    const baseVec = [segI.x2 - segI.x1, segI.y2 - segI.y1];
    const norm = Math.hypot(baseVec[0], baseVec[1]);
    if (norm < 1e-6) continue;

    const unitVec = [baseVec[0] / norm, baseVec[1] / norm];
    const normalUnit = [-unitVec[1], unitVec[0]];
    const midI = [(segI.x1 + segI.x2) / 2.0, (segI.y1 + segI.y2) / 2.0];

    const matchedLines = [];

    for (let j = 0; j < n; j++) {
      if (i === j) continue;

      const segJ = lineSegments[j];

      // 平行でなければスキップ
      if (!angleSimilar(segI.angle, segJ.angle, PARALLEL_ANGLE_THRESHOLD_DEG)) continue;

      // 長手方向オーバーラップ（固定 0.01）
      const ratio = overlapRatioOnLength(segI, segJ, unitVec);
      if (ratio < 0.01) continue;

      const midJ = [(segJ.x1 + segJ.x2) / 2.0, (segJ.y1 + segJ.y2) / 2.0];
      const offJ =
        (midJ[0] - midI[0]) * normalUnit[0] +
        (midJ[1] - midI[1]) * normalUnit[1];

      matchedLines.push({ idx: j, off: offJ });
    }

    // 自分自身も含める（オフセット 0）
    matchedLines.push({ idx: i, off: 0.0 });

    // 法線方向オフセットで並び替え
    matchedLines.sort((a, b) => a.off - b.off);

    // 隣接ペアの距離を INTERVAL_EPS でクラスタリング
    for (let k = 0; k < matchedLines.length - 1; k++) {
      const curr = matchedLines[k];
      const nxt  = matchedLines[k + 1];
      const dist = Math.abs(nxt.off - curr.off);
      if (dist <= 0) continue;

      let placed = false;
      for (const g of groups) {
        if (Math.abs(dist - g.seed) <= INTERVAL_EPS) {
          g.pairs.push({ a: curr.idx, b: nxt.idx });
          g.dists.push(dist);
          placed = true;
          break;
        }
      }
      if (!placed) {
        groups.push({
          seed: dist,
          pairs: [{ a: curr.idx, b: nxt.idx }],
          dists: [dist],
        });
      }
    }
  }

  // --- グループ情報から「線分ごとの関係」に落とし込む ---
  const neighborMapArr = Array.from({ length: n }, () => new Map());
  for (const g of groups) {
    if (!g.pairs || g.pairs.length === 0) continue;
    const meanDist =
      g.dists.reduce((s, v) => s + v, 0) / g.dists.length;

    for (const pair of g.pairs) {
      const a = pair.a;
      const b = pair.b;

      if (!neighborMapArr[a].has(b)) neighborMapArr[a].set(b, []);
      if (!neighborMapArr[b].has(a)) neighborMapArr[b].set(a, []);
      neighborMapArr[a].get(b).push(meanDist);
      neighborMapArr[b].get(a).push(meanDist);
    }
  }

  const equalRelations = [];
  for (let i = 0; i < n; i++) {
    const m = neighborMapArr[i];
    const neighbors = Array.from(m.keys()).sort((a, b) => a - b);

    let summary = "";
    if (neighbors.length === 0) {
      summary = "No equal-interval pattern";
    } else {
      const parts = neighbors.map((j) => {
        const dists = m.get(j);
        const meanD =
          dists.reduce((s, v) => s + v, 0) / dists.length;
        return `${j}(~${meanD.toFixed(2)})`;
      });
      summary = "Equal interval with: " + parts.join(", ");
    }

    equalRelations.push({
      matchedEqual: neighbors,
      summary,
    });
  }

  return equalRelations;
}

/**
 * 平行線比率の解析
 */
function analyzeParallelRatioRelations(lineSegments) {
  const n = lineSegments.length;
  if (!originalImage) {
    return Array.from({ length: n }, () => ({
      matchedRatio: [],
      summary: "No image",
    }));
  }

  const imgW = originalImage.cols;
  const imgH = originalImage.rows;

  function angleSimilar(a, b, thresholdDeg) {
    let diff = Math.abs(a - b) % 180;
    return diff <= thresholdDeg || diff >= 180 - thresholdDeg;
  }

  function overlapRatioOnLength(segI, segJ, unitVec) {
    const p1 = [segI.x1, segI.y1];
    const p2 = [segI.x2, segI.y2];
    const q1 = [segJ.x1, segJ.y1];
    const q2 = [segJ.x2, segJ.y2];

    const dot2D = (u, v) => u[0] * v[0] + u[1] * v[1];

    const proj_p1 = dot2D(p1, unitVec);
    const proj_p2 = dot2D(p2, unitVec);
    const p_min = Math.min(proj_p1, proj_p2);
    const p_max = Math.max(proj_p1, proj_p2);

    const proj_q1 = dot2D(q1, unitVec);
    const proj_q2 = dot2D(q2, unitVec);
    const q_min = Math.min(proj_q1, proj_q2);
    const q_max = Math.max(proj_q1, proj_q2);

    const overlapLen = Math.max(0, Math.min(p_max, q_max) - Math.max(p_min, q_min));
    const len_p = p_max - p_min;
    const len_q = q_max - q_min;
    const denom = Math.min(len_p, len_q);
    if (denom <= 0) return 0;
    return overlapLen / denom;
  }

  function inImage(pt) {
    return pt[0] >= 0 && pt[0] < imgW && pt[1] >= 0 && pt[1] < imgH;
  }

  function intersectLine(p1, p2, p3, p4) {
    const r = [p2[0] - p1[0], p2[1] - p1[1]];
    const s = [p4[0] - p3[0], p4[1] - p3[1]];
    const denom = r[0] * s[1] - r[1] * s[0];
    if (Math.abs(denom) < 1e-8) return null;
    const t =
      ((p3[0] - p1[0]) * s[1] - (p3[1] - p1[1]) * s[0]) / denom;
    return [p1[0] + t * r[0], p1[1] + t * r[1]];
  }

  // --- 1) 平行グループ作成 ---
  const groups = [];
  const used = new Array(n).fill(false);

  for (let i = 0; i < n; i++) {
    if (used[i]) continue;

    const segI = lineSegments[i];
    const dx = segI.x2 - segI.x1;
    const dy = segI.y2 - segI.y1;
    const len = Math.hypot(dx, dy);
    if (len < 1e-6) continue;

    const dirVec = [dx / len, dy / len];

    const grp = [i];
    used[i] = true;

    for (let j = i + 1; j < n; j++) {
      if (used[j]) continue;

      const segJ = lineSegments[j];
      if (!angleSimilar(segI.angle, segJ.angle, PARALLEL_ANGLE_THRESHOLD_DEG)) {
        continue;
      }

      const ratio = overlapRatioOnLength(segI, segJ, dirVec);
      if (ratio < PARALLEL_MIN_INSTRIP_LENGTH) continue;

      grp.push(j);
      used[j] = true;
    }

    if (grp.length > 0) {
      groups.push(grp);
    }
  }

  // --- 2) グループ内で 3 本組の幾何級数パターンを探索 ---
  const patterns = []; // { triplet:[i,j,k], center:number, ratio:number, extras:number[] }

  for (const grp of groups) {
    if (grp.length < 3) continue;

    for (let a = 0; a < grp.length - 2; a++) {
      for (let b = a + 1; b < grp.length - 1; b++) {
        for (let c = b + 1; c < grp.length; c++) {
          const idxs = [grp[a], grp[b], grp[c]];
          const lines = idxs.map((idx) => lineSegments[idx]);

          // A の方向ベクトルで法線軸を定義
          const baseSeg = lines[0];
          let base_p1 = [baseSeg.x1, baseSeg.y1];
          let base_p2 = [baseSeg.x2, baseSeg.y2];
          let v_dir = [base_p2[0] - base_p1[0], base_p2[1] - base_p1[1]];
          let v_norm = Math.hypot(v_dir[0], v_dir[1]);
          if (v_norm < 1e-6) continue;
          v_dir = [v_dir[0] / v_norm, v_dir[1] / v_norm];
          const normal_axis = [-v_dir[1], v_dir[0]];

          // 中点投影値で並び替え → B が中央
          function projVal(seg) {
            const mx = (seg.x1 + seg.x2) / 2.0;
            const my = (seg.y1 + seg.y2) / 2.0;
            return mx * normal_axis[0] + my * normal_axis[1];
          }

          const sortedIdxs = idxs.slice().sort((i1, i2) => {
            return projVal(lineSegments[i1]) - projVal(lineSegments[i2]);
          });

          const Aidx = sortedIdxs[0];
          const Bidx = sortedIdxs[1];
          const Cidx = sortedIdxs[2];

          const segA = lineSegments[Aidx];
          const segB = lineSegments[Bidx];
          const segC = lineSegments[Cidx];

          const pA1 = [segA.x1, segA.y1];
          const pA2 = [segA.x2, segA.y2];
          const pB1 = [segB.x1, segB.y1];
          const pB2 = [segB.x2, segB.y2];
          const pC1 = [segC.x1, segC.y1];
          const pC2 = [segC.x2, segC.y2];

          // point2: B の中点
          const point2 = [
            (pB1[0] + pB2[0]) / 2.0,
            (pB1[1] + pB2[1]) / 2.0,
          ];

          // 法線 unit
          let dirB = [pB2[0] - pB1[0], pB2[1] - pB1[1]];
          const n_len = Math.hypot(dirB[0], dirB[1]);
          if (n_len < 1e-6) continue;
          const u_n = [-dirB[1] / n_len, dirB[0] / n_len];

          // 無限法線 q1-q2
          const q1 = [point2[0] - u_n[0] * 20000.0, point2[1] - u_n[1] * 20000.0];
          const q2 = [point2[0] + u_n[0] * 20000.0, point2[1] + u_n[1] * 20000.0];

          // A, C との交点
          const intA = intersectLine(q1, q2, pA1, pA2);
          const intC = intersectLine(q1, q2, pC1, pC2);
          if (!intA || !intC) continue;

          const dA = Math.hypot(intA[0] - point2[0], intA[1] - point2[1]);
          const dC = Math.hypot(intC[0] - point2[0], intC[1] - point2[1]);
          if (dA < 1e-4 || dC < 1e-4) continue;

          let point1, point3, d21, d32;
          if (dA <= dC) {
            point1 = intA;
            point3 = intC;
            d21 = dA;
            d32 = dC;
          } else {
            point1 = intC;
            point3 = intA;
            d21 = dC;
            d32 = dA;
          }

          const rRatio = d32 / d21;
          if (rRatio <= 1.0 + 1e-6) continue;

          const d10 = d21 / rRatio;
          const u_dir = [
            (point2[0] - point1[0]) / d21,
            (point2[1] - point1[1]) / d21,
          ];
          const point0 = [
            point1[0] - u_dir[0] * d10,
            point1[1] - u_dir[1] * d10,
          ];

          // 幾何級数上の点列生成
          const pointsAlong = []; // { idx:number, pt:[x,y] }

          // 内側（負）
          let curPt = [point0[0], point0[1]];
          let prevStep = d10;
          let idxPos = 0;
          while (true) {
            if (!inImage(curPt)) break;
            pointsAlong.push({ idx: idxPos, pt: [curPt[0], curPt[1]] });

            const nextStep = prevStep / rRatio;
            if (nextStep < 0.5) break;

            curPt = [
              curPt[0] - u_dir[0] * nextStep,
              curPt[1] - u_dir[1] * nextStep,
            ];
            prevStep = nextStep;
            idxPos -= 1;
          }

          // 既知点 1,2,3
          pointsAlong.push({ idx: 1, pt: [point1[0], point1[1]] });
          pointsAlong.push({ idx: 2, pt: [point2[0], point2[1]] });
          pointsAlong.push({ idx: 3, pt: [point3[0], point3[1]] });

          // 外側（正）
          curPt = [point3[0], point3[1]];
          prevStep = d32;
          idxPos = 4;
          while (true) {
            const nextStep = prevStep * rRatio;
            const candPt = [
              curPt[0] + u_dir[0] * nextStep,
              curPt[1] + u_dir[1] * nextStep,
            ];
            if (!inImage(candPt)) break;
            pointsAlong.push({ idx: idxPos, pt: [candPt[0], candPt[1]] });
            curPt = candPt;
            prevStep = nextStep;
            idxPos += 1;
          }

          // グループ内の他線分との照合
          const extras = [];
          const baseSet = new Set([Aidx, Bidx, Cidx]);

          for (const k of grp) {
            if (baseSet.has(k)) continue;
            const segK = lineSegments[k];
            const pk1 = [segK.x1, segK.y1];
            const pk2 = [segK.x2, segK.y2];
            const inter = intersectLine(q1, q2, pk1, pk2);
            if (!inter) continue;

            let hit = false;
            for (const pa of pointsAlong) {
              const d = Math.hypot(inter[0] - pa.pt[0], inter[1] - pa.pt[1]);
              if (d <= RATIO_POINT_EPS) {
                hit = true;
                break;
              }
            }
            if (hit) extras.push(k);
          }

          if (extras.length === 0) continue; // Python と同様、追加線が無ければ捨てる

          patterns.push({
            triplet: [Aidx, Bidx, Cidx],
            center: Bidx,
            ratio: rRatio,
            extras: Array.from(new Set(extras)),
          });
        }
      }
    }
  }

  // --- 3) パターン情報を「線分ごとの関係」に落とし込む ---
  const neighborMapArr = Array.from({ length: n }, () => new Map());

  for (const pat of patterns) {
    const participants = new Set([
      ...pat.triplet,
      ...pat.extras,
    ]);
    const arr = Array.from(participants);

    for (let i = 0; i < arr.length; i++) {
      for (let j = 0; j < arr.length; j++) {
        if (i === j) continue;
        const a = arr[i];
        const b = arr[j];

        const mA = neighborMapArr[a];
        if (!mA.has(b)) mA.set(b, []);
        mA.get(b).push(pat.ratio);
      }
    }
  }

  const ratioRelations = [];
  for (let i = 0; i < n; i++) {
    const m = neighborMapArr[i];
    const neighbors = Array.from(m.keys()).sort((a, b) => a - b);

    let summary = "";
    if (neighbors.length === 0) {
      summary = "No ratio pattern";
    } else {
      const parts = neighbors.map((j) => {
        const rArr = m.get(j);
        const meanR =
          rArr.reduce((s, v) => s + v, 0) / rArr.length;
        return `${j}(r~${meanR.toFixed(2)})`;
      });
      summary = "Ratio with: " + parts.join(", ");
    }

    ratioRelations.push({
      matchedRatio: neighbors,
      summary,
    });
  }

  return ratioRelations;
}

/**
 * 結果を表形式 + サムネイルで表示
 *  - 延長一致
 *  - 平行一致
 *  - 同一間隔
 *  - 平行線比率
 */
function showLineInfo(
  lineSegments,
  extRelations,
  parallelRelations,
  equalIntervalRelations,
  ratioRelations
) {
  const filteredCount = lineSegments.length;

  let html = "";
  html += `<b>延長一致の分析結果 / Extension (A/B/C) Results</b><br>`;
  html += `線分数（フィルタ後） / Line segments (after filtering): ${filteredCount}<br>`;

  // 閾値の表示
  html += "<b>閾値（現在の設定） / Thresholds (current)</b><br>";
  html += `最短線分長 / Min line length: ${MIN_LINE_LENGTH} px<br>`;
  html += `A: |proj| ≥ ${EXT_A_MIN_PROJ}, orth ≤ ${EXT_A_MAX_ORTH}<br>`;
  html += `B: orth ≤ ${EXT_B_MAX_ORTH}<br>`;
  html += `C: |proj| ≥ ${EXT_C_MIN_PROJ}, orth ≤ ${EXT_C_MAX_ORTH}<br>`;
  html += `平行 / Parallel: angleDiff ≤ ${PARALLEL_ANGLE_THRESHOLD_DEG}°, overlap ≥ ${PARALLEL_MIN_INSTRIP_LENGTH}<br>`;
  html += `同一間隔EPS / Equal-interval EPS: ±${INTERVAL_EPS} px<br>`;
  html += `比率点距離EPS / Ratio point distance EPS: ±${RATIO_POINT_EPS} px<br><br>`;

  // 分析ボタン + Excel ダウンロードボタン（延長一致）
  html += `
    <button id="analyzeBtn" style="margin-bottom:4px;">
      条件A/B/C の抽出割合を分析 / Analyze extraction ratio of A/B/C
    </button>
    <button id="downloadExtExcel" style="margin-left:8px; margin-bottom:4px;">
      延長一致結果をExcelダウンロード / Download extension results (Excel)
    </button>
    <div id="analysisSummary" style="margin-bottom:8px;"></div>
  `;

  // ===== 延長一致テーブル =====
  html += `<div id="tableWrapper">`;
  html += `<table id="lineTable" border="1" cellspacing="0" cellpadding="4">`;
  html += "<thead><tr>";
  html += `<th class="col-small">表示 / Show</th>`;
  html += `<th class="col-small">#</th>`;
  html += `<th class="col-small">x1</th>`;
  html += `<th class="col-small">y1</th>`;
  html += `<th class="col-small">x2</th>`;
  html += `<th class="col-small">y2</th>`;
  html += `<th class="col-small">length</th>`;
  html += `<th class="col-small">angle(deg)</th>`;
  html += `<th class="col-canvas">Object</th>`;
  html += `<th class="col-canvas">AllPatterns(A+B+C)</th>`;
  html += `<th class="col-small">side1</th>`;
  html += `<th class="col-small">side2</th>`;
  html += `<th class="col-abc">A</th>`;
  html += `<th class="col-abc">B</th>`;
  html += `<th class="col-abc">C</th>`;
  html += `<th class="col-summary">summary</th>`;
  html += "</tr></thead><tbody>";

  lineSegments.forEach((seg, idx) => {
    const rel = extRelations[idx];
    html += `<tr data-idx="${idx}">`;
    html += `<td class="col-small keep-visible"><input type="checkbox" class="row-toggle" data-index="${idx}" checked></td>`;
    html += `<td class="col-small keep-visible">${idx}</td>`;
    html += `<td class="col-small">${seg.x1}</td>`;
    html += `<td class="col-small">${seg.y1}</td>`;
    html += `<td class="col-small">${seg.x2}</td>`;
    html += `<td class="col-small">${seg.y2}</td>`;
    html += `<td class="col-small">${seg.length.toFixed(1)}</td>`;
    html += `<td class="col-small">${seg.angle.toFixed(1)}</td>`;

    html += `<td class="col-canvas"><canvas id="obj_${idx}" style="border:1px solid #ccc;"></canvas></td>`;
    html += `<td class="col-canvas"><canvas id="all_${idx}" style="border:1px solid #ccc;"></canvas></td>`;

    if (rel) {
      html += `<td class="col-small">${rel.side1Match ? "○" : ""}</td>`;
      html += `<td class="col-small">${rel.side2Match ? "○" : ""}</td>`;
      html += `<td class="col-abc">${rel.matchedA.join(",")}</td>`;
      html += `<td class="col-abc">${rel.matchedB.join(",")}</td>`;
      html += `<td class="col-abc">${rel.matchedC.join(",")}</td>`;
      html += `<td class="col-summary">${rel.summary}</td>`;
    } else {
      html += `<td class="col-small"></td>`;
      html += `<td class="col-small"></td>`;
      html += `<td class="col-abc"></td>`;
      html += `<td class="col-abc"></td>`;
      html += `<td class="col-abc"></td>`;
      html += `<td class="col-summary"></td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table></div>";

  // ===== 平行一致ブロック =====
  html += `<hr style="margin:12px 0;">`;
  html += `<b>平行一致の分析結果 / Parallel Results</b><br>`;

  let parallelHasMatch = 0;
  for (const pr of parallelRelations) {
    if (pr && pr.matchedParallel && pr.matchedParallel.length > 0) {
      parallelHasMatch++;
    }
  }
  html += `平行かつ条件を満たす相手を1本以上持つ線分 / Segments with ≥1 valid parallel match: ${parallelHasMatch} / ${filteredCount}<br><br>`;

  html += `
    <button id="parallelAnalyzeBtn" style="margin-bottom:4px;">
      平行抽出の割合を分析 / Analyze parallel extraction ratio
    </button>
    <button id="downloadParallelExcel" style="margin-left:8px; margin-bottom:4px;">
      平行一致結果をExcelダウンロード / Download parallel results (Excel)
    </button>
    <div id="parallelAnalysisSummary" style="margin-bottom:8px;"></div>
  `;

  html += `<div id="parallelTableWrapper">`;
  html += `<table id="parallelTable" border="1" cellspacing="0" cellpadding="4">`;
  html += "<thead><tr>";
  html += `<th class="col-small">表示 / Show</th>`;
  html += `<th class="col-small">#</th>`;
  html += `<th class="col-small">x1</th>`;
  html += `<th class="col-small">y1</th>`;
  html += `<th class="col-small">x2</th>`;
  html += `<th class="col-small">y2</th>`;
  html += `<th class="col-small">angle(deg)</th>`;
  html += `<th class="col-canvas">Parallel Object</th>`;
  html += `<th class="col-canvas">Parallel All</th>`;
  html += `<th class="col-abc">Parallel idx</th>`;
  html += `<th class="col-summary">Parallel summary</th>`;
  html += "</tr></thead><tbody>";

  lineSegments.forEach((seg, idx) => {
    const pr = parallelRelations[idx];
    html += `<tr data-pidx="${idx}">`;
    html += `<td class="col-small keep-visible"><input type="checkbox" class="prow-toggle" data-index="${idx}" checked></td>`;
    html += `<td class="col-small keep-visible">${idx}</td>`;
    html += `<td class="col-small">${seg.x1}</td>`;
    html += `<td class="col-small">${seg.y1}</td>`;
    html += `<td class="col-small">${seg.x2}</td>`;
    html += `<td class="col-small">${seg.y2}</td>`;
    html += `<td class="col-small">${seg.angle.toFixed(1)}</td>`;

    html += `<td class="col-canvas"><canvas id="pobj_${idx}" style="border:1px solid #ccc;"></canvas></td>`;
    html += `<td class="col-canvas"><canvas id="pall_${idx}" style="border:1px solid #ccc;"></canvas></td>`;

    if (pr) {
      html += `<td class="col-abc">${(pr.matchedParallel || []).join(",")}</td>`;
      html += `<td class="col-summary">${pr.summary}</td>`;
    } else {
      html += `<td class="col-abc"></td>`;
      html += `<td class="col-summary"></td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table></div>";

  // ===== 同一間隔ブロック =====
  html += `<hr style="margin:12px 0;">`;
  html += `<b>同一間隔の分析結果 / Equal-Interval Results</b><br>`;

  let equalHasMatch = 0;
  for (const er of equalIntervalRelations) {
    if (er && er.matchedEqual && er.matchedEqual.length > 0) {
      equalHasMatch++;
    }
  }
  html += `同一間隔パターンを持つ線分 / Segments with equal-interval pattern: ${equalHasMatch} / ${filteredCount}<br><br>`;

  html += `
    <button id="intervalAnalyzeBtn" style="margin-bottom:4px;">
      同一間隔抽出の割合を分析 / Analyze equal-interval extraction ratio
    </button>
    <button id="downloadIntervalExcel" style="margin-left:8px; margin-bottom:4px;">
      同一間隔結果をExcelダウンロード / Download equal-interval results (Excel)
    </button>
    <div id="intervalAnalysisSummary" style="margin-bottom:8px;"></div>
  `;

  html += `<div id="intervalTableWrapper">`;
  html += `<table id="intervalTable" border="1" cellspacing="0" cellpadding="4">`;
  html += "<thead><tr>";
  html += `<th class="col-small">表示 / Show</th>`;
  html += `<th class="col-small">#</th>`;
  html += `<th class="col-small">x1</th>`;
  html += `<th class="col-small">y1</th>`;
  html += `<th class="col-small">x2</th>`;
  html += `<th class="col-small">y2</th>`;
  html += `<th class="col-small">angle(deg)</th>`;
  html += `<th class="col-canvas">Interval Object</th>`;
  html += `<th class="col-canvas">Interval All</th>`;
  html += `<th class="col-abc">Interval idx</th>`;
  html += `<th class="col-summary">Interval summary</th>`;
  html += "</tr></thead><tbody>";

  lineSegments.forEach((seg, idx) => {
    const er = equalIntervalRelations[idx];
    html += `<tr data-iidx="${idx}">`;

    html += `<td class="col-small keep-visible"><input type="checkbox" class="irow-toggle" data-index="${idx}" checked></td>`;
    html += `<td class="col-small keep-visible">${idx}</td>`;
    html += `<td class="col-small">${seg.x1}</td>`;
    html += `<td class="col-small">${seg.y1}</td>`;
    html += `<td class="col-small">${seg.x2}</td>`;
    html += `<td class="col-small">${seg.y2}</td>`;
    html += `<td class="col-small">${seg.angle.toFixed(1)}</td>`;

    html += `<td class="col-canvas"><canvas id="iobj_${idx}" style="border:1px solid #ccc;"></canvas></td>`;
    html += `<td class="col-canvas"><canvas id="iall_${idx}" style="border:1px solid #ccc;"></canvas></td>`;

    if (er) {
      html += `<td class="col-abc">${(er.matchedEqual || []).join(",")}</td>`;
      html += `<td class="col-summary">${er.summary}</td>`;
    } else {
      html += `<td class="col-abc"></td>`;
      html += `<td class="col-summary"></td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table></div>";

  // ===== 平行線比率ブロック =====
  html += `<hr style="margin:12px 0;">`;
  html += `<b>平行線比率の分析結果 / Parallel-Line Ratio Results</b><br>`;

  let ratioHasMatch = 0;
  for (const rr of ratioRelations) {
    if (rr && rr.matchedRatio && rr.matchedRatio.length > 0) {
      ratioHasMatch++;
    }
  }
  html += `比率パターンを持つ線分 / Segments with ratio pattern: ${ratioHasMatch} / ${filteredCount}<br><br>`;

  html += `
    <button id="ratioAnalyzeBtn" style="margin-bottom:4px;">
      平行線比率抽出の割合を分析 / Analyze parallel-line ratio extraction ratio
    </button>
    <button id="downloadRatioExcel" style="margin-left:8px; margin-bottom:4px;">
      平行線比率結果をExcelダウンロード / Download ratio results (Excel)
    </button>
    <div id="ratioAnalysisSummary" style="margin-bottom:8px;"></div>
  `;

  html += `<div id="ratioTableWrapper">`;
  html += `<table id="ratioTable" border="1" cellspacing="0" cellpadding="4">`;
  html += "<thead><tr>";
  html += `<th class="col-small">表示 / Show</th>`;
  html += `<th class="col-small">#</th>`;
  html += `<th class="col-small">x1</th>`;
  html += `<th class="col-small">y1</th>`;
  html += `<th class="col-small">x2</th>`;
  html += `<th class="col-small">y2</th>`;
  html += `<th class="col-small">angle(deg)</th>`;
  html += `<th class="col-canvas">Ratio Object</th>`;
  html += `<th class="col-canvas">Ratio All</th>`;
  html += `<th class="col-abc">Ratio idx</th>`;
  html += `<th class="col-summary">Ratio summary</th>`;
  html += "</tr></thead><tbody>";

  lineSegments.forEach((seg, idx) => {
    const rr = ratioRelations[idx];
    html += `<tr data-ridx="${idx}">`;

    html += `<td class="col-small keep-visible"><input type="checkbox" class="rrow-toggle" data-index="${idx}" checked></td>`;
    html += `<td class="col-small keep-visible">${idx}</td>`;
    html += `<td class="col-small">${seg.x1}</td>`;
    html += `<td class="col-small">${seg.y1}</td>`;
    html += `<td class="col-small">${seg.x2}</td>`;
    html += `<td class="col-small">${seg.y2}</td>`;
    html += `<td class="col-small">${seg.angle.toFixed(1)}</td>`;

    html += `<td class="col-canvas"><canvas id="robj_${idx}" style="border:1px solid #ccc;"></canvas></td>`;
    html += `<td class="col-canvas"><canvas id="rall_${idx}" style="border:1px solid #ccc;"></canvas></td>`;

    if (rr) {
      html += `<td class="col-abc">${(rr.matchedRatio || []).join(",")}</td>`;
      html += `<td class="col-summary">${rr.summary}</td>`;
    } else {
      html += `<td class="col-abc"></td>`;
      html += `<td class="col-summary"></td>`;
    }

    html += "</tr>";
  });

  html += "</tbody></table></div>";

  // ここまでの HTML を反映
  resultDiv.innerHTML = html;

  // ===== 延長側イベント・サムネイル =====
  const table = document.getElementById("lineTable");
  const tbody = table.querySelector("tbody");
  const analysisSummary = document.getElementById("analysisSummary");

  tbody.querySelectorAll("input.row-toggle").forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const tr = e.target.closest("tr");
      if (!tr) return;
      if (e.target.checked) tr.classList.remove("hidden-row");
      else tr.classList.add("hidden-row");
    });
  });

  const analyzeBtn = document.getElementById("analyzeBtn");
  analyzeBtn.addEventListener("click", () => {
    const checkedIdx = [];
    tbody.querySelectorAll("input.row-toggle").forEach((cb) => {
      if (cb.checked) {
        const idx = parseInt(cb.dataset.index, 10);
        if (!Number.isNaN(idx)) checkedIdx.push(idx);
      }
    });

    const total = checkedIdx.length;
    if (total === 0) {
      analysisSummary.textContent = "チェックされた線分がありません。 / No line segments are checked.";
      return;
    }

    const checkedSet = new Set(checkedIdx);
    let successA = 0;
    let successB = 0;
    let successC = 0;

    for (const i of checkedIdx) {
      const rel = extRelations[i];
      if (!rel) continue;

      if (rel.matchedA.some((j) => j !== i && checkedSet.has(j))) successA++;
      if (rel.matchedB.some((j) => j !== i && checkedSet.has(j))) successB++;
      if (rel.matchedC.some((j) => j !== i && checkedSet.has(j))) successC++;
    }

    const pct = (num) => ((num / total) * 100).toFixed(1);

    analysisSummary.innerHTML =
      `チェックされた線分数 / Checked segments: ${total}<br>` +
      `条件A / Condition A: ${successA}/${total} (${pct(successA)}%)<br>` +
      `条件B / Condition B: ${successB}/${total} (${pct(successB)}%)<br>` +
      `条件C / Condition C: ${successC}/${total} (${pct(successC)}%)`;
  });

  // 延長のサムネイル
  lineSegments.forEach((seg, idx) => {
    drawObjectLineThumbnail(`obj_${idx}`, seg, 0.8);
    drawAllPatternsThumbnail(`all_${idx}`, idx, lineSegments, extRelations, 0.8);
  });

  // 延長 Excel ダウンロード
  const extDownloadBtn = document.getElementById("downloadExtExcel");
  if (extDownloadBtn) {
    extDownloadBtn.addEventListener("click", () => {
      exportExtensionExcel(lineSegments, extRelations);
    });
  }

  // ===== 平行テーブル関連 =====
  const pTable = document.getElementById("parallelTable");
  const pTbody = pTable.querySelector("tbody");
  const parallelSummaryDiv = document.getElementById("parallelAnalysisSummary");
  const parallelAnalyzeBtn = document.getElementById("parallelAnalyzeBtn");

  lineSegments.forEach((seg, idx) => {
    drawObjectLineThumbnail(`pobj_${idx}`, seg, 0.8);
    drawParallelPatternsThumbnail(`pall_${idx}`, idx, lineSegments, parallelRelations, 0.8);
  });

  pTbody.querySelectorAll("input.prow-toggle").forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const tr = e.target.closest("tr");
      if (!tr) return;
      if (e.target.checked) tr.classList.remove("hidden-row");
      else tr.classList.add("hidden-row");
    });
  });

  parallelAnalyzeBtn.addEventListener("click", () => {
    const checkedIdx = [];
    pTbody.querySelectorAll("input.prow-toggle").forEach((cb) => {
      if (cb.checked) {
        const idx = parseInt(cb.dataset.index, 10);
        if (!Number.isNaN(idx)) checkedIdx.push(idx);
      }
    });

    const total = checkedIdx.length;
    if (total === 0) {
      parallelSummaryDiv.textContent = "チェックされた線分がありません。 / No line segments are checked.";
      return;
    }

    const checkedSet = new Set(checkedIdx);
    let successParallel = 0;

    for (const i of checkedIdx) {
      const pr = parallelRelations[i];
      if (!pr || !pr.matchedParallel) continue;
      if (pr.matchedParallel.some((j) => j !== i && checkedSet.has(j))) {
        successParallel++;
      }
    }

    const pct = (num) => ((num / total) * 100).toFixed(1);
    parallelSummaryDiv.innerHTML =
      `チェックされた線分数 / Checked segments: ${total}<br>` +
      `平行相手を持つ線分 / Segments with parallel match: ${successParallel}/${total} (${pct(successParallel)}%)`;
  });

  const parallelDownloadBtn = document.getElementById("downloadParallelExcel");
  if (parallelDownloadBtn) {
    parallelDownloadBtn.addEventListener("click", () => {
      exportParallelExcel(lineSegments, parallelRelations);
    });
  }

  // ===== 同一間隔テーブル関連 =====
  const iTable = document.getElementById("intervalTable");
  const iTbody = iTable.querySelector("tbody");
  const intervalSummaryDiv = document.getElementById("intervalAnalysisSummary");
  const intervalAnalyzeBtn = document.getElementById("intervalAnalyzeBtn");

  lineSegments.forEach((seg, idx) => {
    drawObjectLineThumbnail(`iobj_${idx}`, seg, 0.8);
    drawIntervalPatternsThumbnail(`iall_${idx}`, idx, lineSegments, equalIntervalRelations, 0.8);
  });

  iTbody.querySelectorAll("input.irow-toggle").forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const tr = e.target.closest("tr");
      if (!tr) return;
      if (e.target.checked) tr.classList.remove("hidden-row");
      else tr.classList.add("hidden-row");
    });
  });

  intervalAnalyzeBtn.addEventListener("click", () => {
    const checkedIdx = [];
    iTbody.querySelectorAll("input.irow-toggle").forEach((cb) => {
      if (cb.checked) {
        const idx = parseInt(cb.dataset.index, 10);
        if (!Number.isNaN(idx)) checkedIdx.push(idx);
      }
    });

    const total = checkedIdx.length;
    if (total === 0) {
      intervalSummaryDiv.textContent = "チェックされた線分がありません。 / No line segments are checked.";
      return;
    }

    const checkedSet = new Set(checkedIdx);
    let successInterval = 0;

    for (const i of checkedIdx) {
      const er = equalIntervalRelations[i];
      if (!er || !er.matchedEqual) continue;
      if (er.matchedEqual.some((j) => j !== i && checkedSet.has(j))) {
        successInterval++;
      }
    }

    const pct = (num) => ((num / total) * 100).toFixed(1);
    intervalSummaryDiv.innerHTML =
      `チェックされた線分数 / Checked segments: ${total}<br>` +
      `同一間隔パターンの相手を持つ線分 / Segments with equal-interval match: ${successInterval}/${total} (${pct(successInterval)}%)`;
  });

  const intervalDownloadBtn = document.getElementById("downloadIntervalExcel");
  if (intervalDownloadBtn) {
    intervalDownloadBtn.addEventListener("click", () => {
      exportIntervalExcel(lineSegments, equalIntervalRelations);
    });
  }

  // ===== 平行線比率テーブル関連 =====
  const rTable = document.getElementById("ratioTable");
  const rTbody = rTable.querySelector("tbody");
  const ratioSummaryDiv = document.getElementById("ratioAnalysisSummary");
  const ratioAnalyzeBtn = document.getElementById("ratioAnalyzeBtn");

  lineSegments.forEach((seg, idx) => {
    drawObjectLineThumbnail(`robj_${idx}`, seg, 0.8);
    drawRatioPatternsThumbnail(`rall_${idx}`, idx, lineSegments, ratioRelations, 0.8);
  });

  rTbody.querySelectorAll("input.rrow-toggle").forEach((cb) => {
    cb.addEventListener("change", (e) => {
      const tr = e.target.closest("tr");
      if (!tr) return;
      if (e.target.checked) tr.classList.remove("hidden-row");
      else tr.classList.add("hidden-row");
    });
  });

  ratioAnalyzeBtn.addEventListener("click", () => {
    const checkedIdx = [];
    rTbody.querySelectorAll("input.rrow-toggle").forEach((cb) => {
      if (cb.checked) {
        const idx = parseInt(cb.dataset.index, 10);
        if (!Number.isNaN(idx)) checkedIdx.push(idx);
      }
    });

    const total = checkedIdx.length;
    if (total === 0) {
      ratioSummaryDiv.textContent = "チェックされた線分がありません。 / No line segments are checked.";
      return;
    }

    const checkedSet = new Set(checkedIdx);
    let successRatio = 0;

    for (const i of checkedIdx) {
      const rr = ratioRelations[i];
      if (!rr || !rr.matchedRatio) continue;
      if (rr.matchedRatio.some((j) => j !== i && checkedSet.has(j))) {
        successRatio++;
      }
    }

    const pct = (num) => ((num / total) * 100).toFixed(1);
    ratioSummaryDiv.innerHTML =
      `チェックされた線分数 / Checked segments: ${total}<br>` +
      `比率パターンの相手を持つ線分 / Segments with ratio match: ${successRatio}/${total} (${pct(successRatio)}%)`;
  });

  const ratioDownloadBtn = document.getElementById("downloadRatioExcel");
  if (ratioDownloadBtn) {
    ratioDownloadBtn.addEventListener("click", () => {
      exportRatioExcel(lineSegments, ratioRelations);
    });
  }
}

/* ======================= ここから Excel 出力関連 ======================= */

// ピクセル → 行の高さ（pt）
// 適宜調整
function pxToRowHeight(px) {
  return px * 1;
}

// ピクセル → 列幅（"文字数"ベースの Excel 単位）
// 厳密ではないが経験的に px/7〜px/8 くらいがちょうど良いのでそこで近似
function pxToColWidth(px) {
  return Math.max(px / 7, 10);  // 最低幅10くらいは確保
}

function ensureExcelJS() {
  if (typeof ExcelJS === "undefined") {
    alert("ExcelJS が読み込まれていません。\nindex.html に ExcelJS の <script> を追加してください。");
    return false;
  }
  return true;
}

function downloadWorkbook(workbook, filename) {
  workbook.xlsx.writeBuffer()
    .then((buffer) => {
      const blob = new Blob(
        [buffer],
        {
          type:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
      );
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    })
    .catch((err) => {
      console.error("Excel 書き出しでエラー:", err);
      alert("Excel ファイルの生成中にエラーが発生しました。コンソールを確認してください。");
    });
}

// 延長一致 Excel 出力
function exportExtensionExcel(lineSegments, extRelations) {
  if (!ensureExcelJS()) return;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Extension");

  sheet.columns = [
    { header: "#", key: "idx", width: 5 },
    { header: "x1", key: "x1", width: 8 },
    { header: "y1", key: "y1", width: 8 },
    { header: "x2", key: "x2", width: 8 },
    { header: "y2", key: "y2", width: 8 },
    { header: "length", key: "length", width: 10 },
    { header: "angle(deg)", key: "angle", width: 10 },
    { header: "side1", key: "side1", width: 8 },
    { header: "side2", key: "side2", width: 8 },
    { header: "A", key: "A", width: 15 },
    { header: "B", key: "B", width: 15 },
    { header: "C", key: "C", width: 15 },
    { header: "summary", key: "summary", width: 40 },
    { header: "Object", key: "obj", width: 18 },        // col 14
    { header: "AllPatterns", key: "all", width: 18 },   // col 15
  ];

  // 各列で必要な「最低列幅」を覚えておく（Object=14, All=15）
  let maxObjWidthPx = 0;
  let maxAllWidthPx = 0;

  lineSegments.forEach((seg, idx) => {
    const rel = extRelations[idx] || {
      side1Match: false,
      side2Match: false,
      matchedA: [],
      matchedB: [],
      matchedC: [],
      summary: "",
    };

    const row = sheet.addRow({
      idx,
      x1: seg.x1,
      y1: seg.y1,
      x2: seg.x2,
      y2: seg.y2,
      length: seg.length,
      angle: seg.angle,
      side1: rel.side1Match ? "○" : "",
      side2: rel.side2Match ? "○" : "",
      A: rel.matchedA.join(","),
      B: rel.matchedB.join(","),
      C: rel.matchedC.join(","),
      summary: rel.summary,
    });

    const rowNumber = row.number; // 1-origin
    let rowHeightPx = 0;          // この行で必要な px 高さ

    // Object サムネイル
    const objCanvas = document.getElementById(`obj_${idx}`);
    if (objCanvas) {
      const dataURL = objCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({
        base64,
        extension: "png",
      });

      sheet.addImage(imgId, {
        tl: { col: 13, row: rowNumber - 1 }, // 14列目 => index 13
        ext: { width: objCanvas.width, height: objCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, objCanvas.height);
      maxObjWidthPx = Math.max(maxObjWidthPx, objCanvas.width);
    }

    // AllPatterns サムネイル
    const allCanvas = document.getElementById(`all_${idx}`);
    if (allCanvas) {
      const dataURL = allCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({
        base64,
        extension: "png",
      });

      sheet.addImage(imgId, {
        tl: { col: 14, row: rowNumber - 1 }, // 15列目 => index 14
        ext: { width: allCanvas.width, height: allCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, allCanvas.height);
      maxAllWidthPx = Math.max(maxAllWidthPx, allCanvas.width);
    }

    // 画像がある行だけ、行高を画像に合わせて拡張
    if (rowHeightPx > 0) {
      const r = sheet.getRow(rowNumber);
      const current = r.height || 0;
      const needed = pxToRowHeight(rowHeightPx);
      if (needed > current) {
        r.height = needed;
      }
    }
  });

  // 画像列の幅も、画像サイズに応じて調整
  if (maxObjWidthPx > 0) {
    sheet.getColumn(14).width = pxToColWidth(maxObjWidthPx);
  }
  if (maxAllWidthPx > 0) {
    sheet.getColumn(15).width = pxToColWidth(maxAllWidthPx);
  }

  downloadWorkbook(workbook, "extension_result.xlsx");
}

// 平行一致 Excel 出力
function exportParallelExcel(lineSegments, parallelRelations) {
  if (!ensureExcelJS()) return;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Parallel");

  sheet.columns = [
    { header: "#", key: "idx", width: 5 },
    { header: "x1", key: "x1", width: 8 },
    { header: "y1", key: "y1", width: 8 },
    { header: "x2", key: "x2", width: 8 },
    { header: "y2", key: "y2", width: 8 },
    { header: "angle(deg)", key: "angle", width: 10 },
    { header: "ParallelIdx", key: "pidx", width: 20 },
    { header: "summary", key: "summary", width: 40 },
    { header: "Object", key: "obj", width: 18 },        // col 9
    { header: "ParallelAll", key: "all", width: 18 },   // col 10
  ];

  let maxObjWidthPx = 0;
  let maxAllWidthPx = 0;

  lineSegments.forEach((seg, idx) => {
    const pr = parallelRelations[idx] || {
      matchedParallel: [],
      summary: "",
    };

    const row = sheet.addRow({
      idx,
      x1: seg.x1,
      y1: seg.y1,
      x2: seg.x2,
      y2: seg.y2,
      angle: seg.angle,
      pidx: (pr.matchedParallel || []).join(","),
      summary: pr.summary,
    });

    const rowNumber = row.number;
    let rowHeightPx = 0;

    const objCanvas = document.getElementById(`pobj_${idx}`);
    if (objCanvas) {
      const dataURL = objCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({ base64, extension: "png" });
      sheet.addImage(imgId, {
        tl: { col: 8, row: rowNumber - 1 }, // 9列目
        ext: { width: objCanvas.width, height: objCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, objCanvas.height);
      maxObjWidthPx = Math.max(maxObjWidthPx, objCanvas.width);
    }

    const allCanvas = document.getElementById(`pall_${idx}`);
    if (allCanvas) {
      const dataURL = allCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({ base64, extension: "png" });
      sheet.addImage(imgId, {
        tl: { col: 9, row: rowNumber - 1 }, // 10列目
        ext: { width: allCanvas.width, height: allCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, allCanvas.height);
      maxAllWidthPx = Math.max(maxAllWidthPx, allCanvas.width);
    }

    if (rowHeightPx > 0) {
      const r = sheet.getRow(rowNumber);
      const current = r.height || 0;
      const needed = pxToRowHeight(rowHeightPx);
      if (needed > current) r.height = needed;
    }
  });

  if (maxObjWidthPx > 0) {
    sheet.getColumn(9).width = pxToColWidth(maxObjWidthPx);
  }
  if (maxAllWidthPx > 0) {
    sheet.getColumn(10).width = pxToColWidth(maxAllWidthPx);
  }

  downloadWorkbook(workbook, "parallel_result.xlsx");
}

// 同一間隔 Excel 出力
function exportIntervalExcel(lineSegments, equalIntervalRelations) {
  if (!ensureExcelJS()) return;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Interval");

  sheet.columns = [
    { header: "#", key: "idx", width: 5 },
    { header: "x1", key: "x1", width: 8 },
    { header: "y1", key: "y1", width: 8 },
    { header: "x2", key: "x2", width: 8 },
    { header: "y2", key: "y2", width: 8 },
    { header: "angle(deg)", key: "angle", width: 10 },
    { header: "IntervalIdx", key: "iidx", width: 20 },
    { header: "summary", key: "summary", width: 40 },
    { header: "Object", key: "obj", width: 18 },          // col 9
    { header: "IntervalAll", key: "all", width: 18 },     // col 10
  ];

  let maxObjWidthPx = 0;
  let maxAllWidthPx = 0;

  lineSegments.forEach((seg, idx) => {
    const er = equalIntervalRelations[idx] || {
      matchedEqual: [],
      summary: "",
    };

    const row = sheet.addRow({
      idx,
      x1: seg.x1,
      y1: seg.y1,
      x2: seg.x2,
      y2: seg.y2,
      angle: seg.angle,
      iidx: (er.matchedEqual || []).join(","),
      summary: er.summary,
    });

    const rowNumber = row.number;
    let rowHeightPx = 0;

    const objCanvas = document.getElementById(`iobj_${idx}`);
    if (objCanvas) {
      const dataURL = objCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({ base64, extension: "png" });
      sheet.addImage(imgId, {
        tl: { col: 8, row: rowNumber - 1 },
        ext: { width: objCanvas.width, height: objCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, objCanvas.height);
      maxObjWidthPx = Math.max(maxObjWidthPx, objCanvas.width);
    }

    const allCanvas = document.getElementById(`iall_${idx}`);
    if (allCanvas) {
      const dataURL = allCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({ base64, extension: "png" });
      sheet.addImage(imgId, {
        tl: { col: 9, row: rowNumber - 1 },
        ext: { width: allCanvas.width, height: allCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, allCanvas.height);
      maxAllWidthPx = Math.max(maxAllWidthPx, allCanvas.width);
    }

    if (rowHeightPx > 0) {
      const r = sheet.getRow(rowNumber);
      const current = r.height || 0;
      const needed = pxToRowHeight(rowHeightPx);
      if (needed > current) r.height = needed;
    }
  });

  if (maxObjWidthPx > 0) {
    sheet.getColumn(9).width = pxToColWidth(maxObjWidthPx);
  }
  if (maxAllWidthPx > 0) {
    sheet.getColumn(10).width = pxToColWidth(maxAllWidthPx);
  }

  downloadWorkbook(workbook, "interval_result.xlsx");
}

// 平行線比率 Excel 出力
function exportRatioExcel(lineSegments, ratioRelations) {
  if (!ensureExcelJS()) return;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Ratio");

  sheet.columns = [
    { header: "#", key: "idx", width: 5 },
    { header: "x1", key: "x1", width: 8 },
    { header: "y1", key: "y1", width: 8 },
    { header: "x2", key: "x2", width: 8 },
    { header: "y2", key: "y2", width: 8 },
    { header: "angle(deg)", key: "angle", width: 10 },
    { header: "RatioIdx", key: "ridx", width: 20 },
    { header: "summary", key: "summary", width: 40 },
    { header: "Object", key: "obj", width: 18 },       // col 9
    { header: "RatioAll", key: "all", width: 18 },     // col 10
  ];

  let maxObjWidthPx = 0;
  let maxAllWidthPx = 0;

  lineSegments.forEach((seg, idx) => {
    const rr = ratioRelations[idx] || {
      matchedRatio: [],
      summary: "",
    };

    const row = sheet.addRow({
      idx,
      x1: seg.x1,
      y1: seg.y1,
      x2: seg.x2,
      y2: seg.y2,
      angle: seg.angle,
      ridx: (rr.matchedRatio || []).join(","),
      summary: rr.summary,
    });

    const rowNumber = row.number;
    let rowHeightPx = 0;

    const objCanvas = document.getElementById(`robj_${idx}`);
    if (objCanvas) {
      const dataURL = objCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({ base64, extension: "png" });
      sheet.addImage(imgId, {
        tl: { col: 8, row: rowNumber - 1 },
        ext: { width: objCanvas.width, height: objCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, objCanvas.height);
      maxObjWidthPx = Math.max(maxObjWidthPx, objCanvas.width);
    }

    const allCanvas = document.getElementById(`rall_${idx}`);
    if (allCanvas) {
      const dataURL = allCanvas.toDataURL("image/png");
      const base64 = dataURL.split(",")[1];
      const imgId = workbook.addImage({ base64, extension: "png" });
      sheet.addImage(imgId, {
        tl: { col: 9, row: rowNumber - 1 },
        ext: { width: allCanvas.width, height: allCanvas.height },
      });

      rowHeightPx = Math.max(rowHeightPx, allCanvas.height);
      maxAllWidthPx = Math.max(maxAllWidthPx, allCanvas.width);
    }

    if (rowHeightPx > 0) {
      const r = sheet.getRow(rowNumber);
      const current = r.height || 0;
      const needed = pxToRowHeight(rowHeightPx);
      if (needed > current) r.height = needed;
    }
  });

  if (maxObjWidthPx > 0) {
    sheet.getColumn(9).width = pxToColWidth(maxObjWidthPx);
  }
  if (maxAllWidthPx > 0) {
    sheet.getColumn(10).width = pxToColWidth(maxAllWidthPx);
  }

  downloadWorkbook(workbook, "ratio_result.xlsx");
}

/* ======================= ここまで Excel 出力関連 ======================= */

// ===== サムネイル用ヘルパー =====

function drawBaseThumbnailToCanvas(canvas, scale) {
  if (!originalImage) return { ctx: null, scale };

  const w = Math.round(originalImage.cols * scale);
  const h = Math.round(originalImage.rows * scale);
  canvas.width = w;
  canvas.height = h;

  const ctx = canvas.getContext("2d");
  if (!ctx) return { ctx: null, scale };

  const tmp = document.createElement("canvas");
  tmp.width = originalImage.cols;
  tmp.height = originalImage.rows;
  const tmpCtx = tmp.getContext("2d");

  const imgData = new ImageData(
    new Uint8ClampedArray(originalImage.data),
    originalImage.cols,
    originalImage.rows
  );
  tmpCtx.putImageData(imgData, 0, 0);

  ctx.drawImage(tmp, 0, 0, w, h);

  return { ctx, scale };
}

// Object: 自分の線だけ
function drawObjectLineThumbnail(canvasId, segment, scale = 0.3) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const { ctx, scale: s } = drawBaseThumbnailToCanvas(canvas, scale);
  if (!ctx) return;

  ctx.strokeStyle = "rgb(0,255,0)";
  ctx.lineWidth = 2;

  ctx.beginPath();
  ctx.moveTo(segment.x1 * s, segment.y1 * s);
  ctx.lineTo(segment.x2 * s, segment.y2 * s);
  ctx.stroke();
}

// AllPatterns: A/B/C すべて（延長一致用）
function drawAllPatternsThumbnail(
  canvasId,
  index,
  lineSegments,
  extRelations,
  scale = 0.3
) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const rel = extRelations[index];
  if (!rel) return;

  const { ctx, scale: s } = drawBaseThumbnailToCanvas(canvas, scale);
  if (!ctx) return;

  const seg = lineSegments[index];
  const w = canvas.width;
  const h = canvas.height;

  const cross2D = (u, v) => u[0] * v[1] - u[1] * v[0];
  const dot2D = (u, v) => u[0] * v[0] + u[1] * v[1];

  const SHIFT = 3;
  const ORTH_THR = EXT_A_MAX_ORTH;

  function drawRayToEdge(x0, y0, dir, color) {
    let [dx, dy] = dir;
    const eps = 1e-6;
    if (Math.abs(dx) < eps && Math.abs(dy) < eps) return;

    let tMax = Infinity;
    if (Math.abs(dx) > eps) {
      if (dx > 0) tMax = Math.min(tMax, (w - x0) / dx);
      else        tMax = Math.min(tMax, (0 - x0) / dx);
    }
    if (Math.abs(dy) > eps) {
      if (dy > 0) tMax = Math.min(tMax, (h - y0) / dy);
      else        tMax = Math.min(tMax, (0 - y0) / dy);
    }
    if (!isFinite(tMax) || tMax <= 0) return;

    const x1 = x0 + dx * tMax;
    const y1 = y0 + dy * tMax;

    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = 0.7;
    ctx.setLineDash([2, 2]);
    ctx.beginPath();
    ctx.moveTo(x0, y0);
    ctx.lineTo(x1, y1);
    ctx.stroke();
    ctx.restore();
  }

  const base_p1 = [seg.x1, seg.y1];
  const base_p2 = [seg.x2, seg.y2];
  const base_vec = [base_p2[0] - base_p1[0], base_p2[1] - base_p1[1]];
  const base_len = Math.hypot(base_vec[0], base_vec[1]);
  if (base_len < 1e-6) return;
  const base_u = [base_vec[0] / base_len, base_vec[1] / base_len];

  const matchedIdxSet = new Set([
    ...rel.matchedA,
    ...rel.matchedB,
    ...rel.matchedC,
  ]);

  // 評価線分
  ctx.strokeStyle = "rgb(0,255,0)";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(seg.x1 * s, seg.y1 * s);
  ctx.lineTo(seg.x2 * s, seg.y2 * s);
  ctx.stroke();

  // マッチした他線分（オレンジ）
  ctx.strokeStyle = "rgb(255,130,0)";
  ctx.lineWidth = 2;
  for (const j of matchedIdxSet) {
    const sj = lineSegments[j];
    ctx.beginPath();
    ctx.moveTo(sj.x1 * s, sj.y1 * s);
    ctx.lineTo(sj.x2 * s, sj.y2 * s);
    ctx.stroke();
  }

  const blockedFromOther = new Set();

  // ① 評価線分側からの延長
  let sidePosHit = false;
  let sideNegHit = false;

  for (const j of matchedIdxSet) {
    const sj = lineSegments[j];
    const q1 = [sj.x1, sj.y1];
    const q2 = [sj.x2, sj.y2];
    const endpoints = [q1, q2];

    for (const p of endpoints) {
      const v = [p[0] - base_p1[0], p[1] - base_p1[1]];
      const t = dot2D(v, base_u);
      const orth = Math.abs(cross2D(base_u, v));

      if (orth <= ORTH_THR) {
        blockedFromOther.add(j);

        if (t > base_len) {
          sidePosHit = true;
        } else if (t < 0) {
          sideNegHit = true;
        }
      }
    }
  }

  if (sidePosHit) {
    const startX = base_p2[0] * s;
    const startY = base_p2[1] * s;
    drawRayToEdge(startX, startY, [base_u[0], base_u[1]], "rgb(255,255,0)");
  }
  if (sideNegHit) {
    const startX = base_p1[0] * s;
    const startY = base_p1[1] * s;
    drawRayToEdge(startX, startY, [-base_u[0], -base_u[1]], "rgb(255,255,0)");
  }

  // ② 他線分側からの延長
  for (const j of matchedIdxSet) {
    if (blockedFromOther.has(j)) continue;

    const sj = lineSegments[j];
    const p1 = [sj.x1, sj.y1];
    const p2 = [sj.x2, sj.y2];
    const v  = [p2[0] - p1[0], p2[1] - p1[1]];
    const len = Math.hypot(v[0], v[1]);
    if (len < 1e-6) continue;

    const u = [v[0] / len, v[1] / len];

    const endpoints = [p1, p2];
    const targets   = [base_p1, base_p2];

    let bestEndpoint = null;
    let bestDist = Infinity;

    for (const E of endpoints) {
      for (const T of targets) {
        const dx = T[0] - E[0];
        const dy = T[1] - E[1];
        const d  = Math.hypot(dx, dy);
        if (d < bestDist) {
          bestDist = d;
          bestEndpoint = { pt: E, toEval: [dx, dy] };
        }
      }
    }

    if (!bestEndpoint) continue;

    const proj = dot2D(bestEndpoint.toEval, u);
    const dir  = (proj >= 0) ? u : [-u[0], -u[1]];

    const start = [
      (bestEndpoint.pt[0] + dir[0] * SHIFT) * s,
      (bestEndpoint.pt[1] + dir[1] * SHIFT) * s,
    ];

    drawRayToEdge(start[0], start[1], dir, "rgb(255,255,0)");
  }
}

// ===== Parallel All 用サムネイル描画 =====
function drawParallelPatternsThumbnail(
  canvasId,
  index,
  lineSegments,
  parallelRelations,
  scale = 0.3
) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const rel = parallelRelations[index];
  if (!rel) return;

  const { ctx, scale: s } = drawBaseThumbnailToCanvas(canvas, scale);
  if (!ctx) return;

  const seg = lineSegments[index];
  const w = canvas.width;
  const h = canvas.height;

  const matchedIdxSet = new Set(rel.matchedParallel || []);

  const cross2D = (u, v) => u[0] * v[1] - u[1] * v[0];

  // 線分方向・法線
  const base_p1 = [seg.x1, seg.y1];
  const base_p2 = [seg.x2, seg.y2];
  const base_vec = [base_p2[0] - base_p1[0], base_p2[1] - base_p1[1]];
  const base_len = Math.hypot(base_vec[0], base_vec[1]);
  if (base_len < 1e-6) return;

  const base_u = [base_vec[0] / base_len, base_vec[1] / base_len];  // 長手方向
  const base_n = [-base_u[1], base_u[0]];                           // 法線方向

  function drawRayToEdge(x0, y0, dir, color, lineWidth = 0.7, dashed = true) {
    let [dx, dy] = dir;
    const eps = 1e-6;
    if (Math.abs(dx) < eps && Math.abs(dy) < eps) return;

    let tMax = Infinity;
    if (Math.abs(dx) > eps) {
      if (dx > 0) tMax = Math.min(tMax, (w - x0) / dx);
      else        tMax = Math.min(tMax, (0 - x0) / dx);
    }
    if (Math.abs(dy) > eps) {
      if (dy > 0) tMax = Math.min(tMax, (h - y0) / dy);
      else        tMax = Math.min(tMax, (0 - y0) / dy);
    }
    if (!isFinite(tMax) || tMax <= 0) return;

    const x1 = x0 + dx * tMax;
    const y1 = y0 + dy * tMax;

    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(dashed ? [3, 3] : []);
    ctx.beginPath();
    ctx.moveTo(x0, y0);
    ctx.lineTo(x1, y1);
    ctx.stroke();
    ctx.restore();
  }

  // 評価線分（太い緑）
  ctx.strokeStyle = "rgb(0,255,0)";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(seg.x1 * s, seg.y1 * s);
  ctx.lineTo(seg.x2 * s, seg.y2 * s);
  ctx.stroke();

  // 平行抽出された他線分（オレンジ） + その両端から青点線
  for (const j of matchedIdxSet) {
    const sj = lineSegments[j];
    const p1 = [sj.x1, sj.y1];
    const p2 = [sj.x2, sj.y2];

    // オレンジ実線
    ctx.strokeStyle = "rgb(255,130,0)";
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(p1[0] * s, p1[1] * s);
    ctx.lineTo(p2[0] * s, p2[1] * s);
    ctx.stroke();

    // その線の方向に沿って、両端から青い点線
    const v  = [p2[0] - p1[0], p2[1] - p1[1]];
    const len = Math.hypot(v[0], v[1]);
    if (len > 1e-6) {
      const u = [v[0] / len, v[1] / len];

      drawRayToEdge(p1[0] * s, p1[1] * s, [ u[0],  u[1]], "rgb(0,0,255)", 0.7, true);
      drawRayToEdge(p1[0] * s, p1[1] * s, [-u[0], -u[1]], "rgb(0,0,255)", 0.7, true);

      drawRayToEdge(p2[0] * s, p2[1] * s, [ u[0],  u[1]], "rgb(0,0,255)", 0.7, true);
      drawRayToEdge(p2[0] * s, p2[1] * s, [-u[0], -u[1]], "rgb(0,0,255)", 0.7, true);
    }
  }

  // 抽出基準となる「2本の垂線」（青い細い実線）
  const p1s = [base_p1[0] * s, base_p1[1] * s];
  const p2s = [base_p2[0] * s, base_p2[1] * s];

  drawRayToEdge(p1s[0], p1s[1], [ base_n[0],  base_n[1]], "rgb(0,0,255)", 1.0, false);
  drawRayToEdge(p1s[0], p1s[1], [-base_n[0], -base_n[1]], "rgb(0,0,255)", 1.0, false);
  drawRayToEdge(p2s[0], p2s[1], [ base_n[0],  base_n[1]], "rgb(0,0,255)", 1.0, false);
  drawRayToEdge(p2s[0], p2s[1], [-base_n[0], -base_n[1]], "rgb(0,0,255)", 1.0, false);
}

// ===== Interval All 用サムネイル描画 =====
function drawIntervalPatternsThumbnail(
  canvasId,
  index,
  lineSegments,
  equalIntervalRelations,
  scale = 0.3
) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const rel = equalIntervalRelations[index];
  if (!rel) return;

  const { ctx, scale: s } = drawBaseThumbnailToCanvas(canvas, scale);
  if (!ctx) return;

  const seg = lineSegments[index];

  // 評価線分（緑）
  ctx.strokeStyle = "rgb(0,255,0)";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(seg.x1 * s, seg.y1 * s);
  ctx.lineTo(seg.x2 * s, seg.y2 * s);
  ctx.stroke();

  const baseMid = [
    (seg.x1 + seg.x2) / 2.0,
    (seg.y1 + seg.y2) / 2.0,
  ];

  // 他線分（オレンジ）＋中間点に青丸
  ctx.lineWidth = 2;
  for (const j of rel.matchedEqual || []) {
    const sj = lineSegments[j];

    // オレンジ線
    ctx.strokeStyle = "rgb(255,130,0)";
    ctx.beginPath();
    ctx.moveTo(sj.x1 * s, sj.y1 * s);
    ctx.lineTo(sj.x2 * s, sj.y2 * s);
    ctx.stroke();

    // 中間点（between midpoints）に青い小さな丸
    const midJ = [(sj.x1 + sj.x2) / 2.0, (sj.y1 + sj.y2) / 2.0];
    const between = [
      (baseMid[0] + midJ[0]) / 2.0,
      (baseMid[1] + midJ[1]) / 2.0,
    ];

    const bx = between[0] * s;
    const by = between[1] * s;

    ctx.fillStyle = "rgb(0,0,255)";
    ctx.beginPath();
    ctx.arc(bx, by, 2.0, 0, Math.PI * 2);
    ctx.fill();
  }
}

// ===== Ratio All 用サムネイル描画 =====
function drawRatioPatternsThumbnail(
  canvasId,
  index,
  lineSegments,
  ratioRelations,
  scale = 0.3
) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const rel = ratioRelations[index];
  if (!rel) return;

  const { ctx, scale: s } = drawBaseThumbnailToCanvas(canvas, scale);
  if (!ctx) return;

  const seg = lineSegments[index];

  // 評価線分（緑）
  ctx.strokeStyle = "rgb(0,255,0)";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(seg.x1 * s, seg.y1 * s);
  ctx.lineTo(seg.x2 * s, seg.y2 * s);
  ctx.stroke();

  // 比率パターンの相手（オレンジ）
  ctx.strokeStyle = "rgb(255,130,0)";
  ctx.lineWidth = 2;
  for (const j of rel.matchedRatio || []) {
    const sj = lineSegments[j];
    ctx.beginPath();
    ctx.moveTo(sj.x1 * s, sj.y1 * s);
    ctx.lineTo(sj.x2 * s, sj.y2 * s);
    ctx.stroke();
  }
}

// ===== 閾値パネルと再分析ボタンの処理 =====
function initThresholdInputs() {
  const minLen = document.getElementById("minLineLength");
  const aMin   = document.getElementById("extAmin");
  const aOrth  = document.getElementById("extAorth");
  const bOrth  = document.getElementById("extBorth");
  const cMin   = document.getElementById("extCmin");
  const cOrth  = document.getElementById("extCorth");

  const pAngle     = document.getElementById("parallelAngle");
  const pInstrip   = document.getElementById("parallelInstrip");
  const intervalEp = document.getElementById("intervalEps");
  const ratioEp    = document.getElementById("ratioPointEps");

  // 線分長
  if (minLen) minLen.value = MIN_LINE_LENGTH;

  // 延長 A/B/C
  if (aMin)  aMin.value  = EXT_A_MIN_PROJ;
  if (aOrth) aOrth.value = EXT_A_MAX_ORTH;
  if (bOrth) bOrth.value = EXT_B_MAX_ORTH;
  if (cMin)  cMin.value  = EXT_C_MIN_PROJ;
  if (cOrth) cOrth.value = EXT_C_MAX_ORTH;

  // 平行
  if (pAngle)   pAngle.value   = PARALLEL_ANGLE_THRESHOLD_DEG;
  if (pInstrip) pInstrip.value = PARALLEL_MIN_INSTRIP_LENGTH;

  // 同一間隔 EPS
  if (intervalEp) intervalEp.value = INTERVAL_EPS;

  // 平行線比率 EPS
  if (ratioEp) ratioEp.value = RATIO_POINT_EPS;
}

function applyThresholdFromUI() {
  const minLen = document.getElementById("minLineLength");
  const aMin   = document.getElementById("extAmin");
  const aOrth  = document.getElementById("extAorth");
  const bOrth  = document.getElementById("extBorth");
  const cMin   = document.getElementById("extCmin");
  const cOrth  = document.getElementById("extCorth");

  const pAngle     = document.getElementById("parallelAngle");
  const pInstrip   = document.getElementById("parallelInstrip");
  const intervalEp = document.getElementById("intervalEps");
  const ratioEp    = document.getElementById("ratioPointEps");

  // 線分長
  if (minLen) {
    const vMin = parseFloat(minLen.value);
    if (!isNaN(vMin) && vMin > 0) {
      MIN_LINE_LENGTH = vMin;
    }
  }

  // 延長 A/B/C
  const vAmin  = aMin  ? parseFloat(aMin.value)  : NaN;
  const vAorth = aOrth ? parseFloat(aOrth.value) : NaN;
  const vBorth = bOrth ? parseFloat(bOrth.value) : NaN;
  const vCmin  = cMin  ? parseFloat(cMin.value)  : NaN;
  const vCorth = cOrth ? parseFloat(cOrth.value) : NaN;

  if (!isNaN(vAmin))  EXT_A_MIN_PROJ = vAmin;
  if (!isNaN(vAorth)) EXT_A_MAX_ORTH = vAorth;
  if (!isNaN(vBorth)) EXT_B_MAX_ORTH = vBorth;
  if (!isNaN(vCmin))  EXT_C_MIN_PROJ = vCmin;
  if (!isNaN(vCorth)) EXT_C_MAX_ORTH = vCorth;

  // 平行
  const vPAngle   = pAngle   ? parseFloat(pAngle.value)   : NaN;
  const vPInstrip = pInstrip ? parseFloat(pInstrip.value) : NaN;

  if (!isNaN(vPAngle) && vPAngle >= 0) {
    PARALLEL_ANGLE_THRESHOLD_DEG = vPAngle;
  }
  if (!isNaN(vPInstrip) && vPInstrip >= 0) {
    PARALLEL_MIN_INSTRIP_LENGTH = vPInstrip;
  }

  // 同一間隔 EPS
  const vInterval = intervalEp ? parseFloat(intervalEp.value) : NaN;
  if (!isNaN(vInterval) && vInterval >= 0) {
    INTERVAL_EPS = vInterval;
  }

  // 平行線比率 EPS
  const vRatio = ratioEp ? parseFloat(ratioEp.value) : NaN;
  if (!isNaN(vRatio) && vRatio >= 0) {
    RATIO_POINT_EPS = vRatio;
  }
}

// 再分析ボタン
function reanalyzeWithCurrentThresholds() {
  // 画像がまだ選択されていない場合
  if (!originalImage) {
    alert("先に画像ファイルを選択してください。 / Please select an image first.");
    return;
  }

  // OpenCV 初期化チェック
  if (!cvReady) {
    alert("OpenCV.js の初期化が完了していません。 / OpenCV.js is not initialized yet.");
    return;
  }

  const statusSpan = document.getElementById("reanalyzeStatus");
  if (statusSpan) {
    statusSpan.textContent = "再分析中... / Re-analyzing...";
  }

  // ★ 再分析のたびに最新の UI 値を反映
  applyThresholdFromUI();

  // 元画像から、線分抽出＋A/B/C＋平行＋同一間隔＋平行線比率＋画像描画＋テーブルを全部やり直し
  analyzeFromOriginalImage();

  if (statusSpan) {
    statusSpan.textContent = "再分析完了 / Re-analysis completed";
    setTimeout(() => {
      statusSpan.textContent = "";
    }, 2000);
  }
}

// DOM が読み込まれたら閾値初期化とボタンイベントを設定
window.addEventListener("DOMContentLoaded", () => {
  initThresholdInputs();

  const reBtn = document.getElementById("reanalyzeBtn");
  if (reBtn) {
    reBtn.addEventListener("click", reanalyzeWithCurrentThresholds);
  }
});