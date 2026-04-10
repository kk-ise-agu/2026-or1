/**
 * Google Apps Script — 出席管理バックエンド
 *
 * 使い方:
 *   1. Google Sheets で新しいスプレッドシートを作成
 *   2. 拡張機能 → Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. デプロイ → 新しいデプロイ → ウェブアプリ
 *      - 実行ユーザー: 自分
 *      - アクセス: 全員
 *   5. デプロイURLを index.html の GAS_URL に設定
 *
 * スプレッドシート構成（自動生成）:
 *   シート「出席記録」: タイムスタンプ, 講義回, ラウンド, 学籍番号, 問題, 回答, 正誤, 経過秒, 端末ID, 重複フラグ
 *   シート「集計」: 学籍番号別・回別・ラウンド別の出席一覧
 */

// ── POST受信（出席回答を記録） ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { lecture, round, studentId, question, answer, correct, elapsedSec, deviceId } = data;
    const roundNum = round || 1;

    if (!lecture || !studentId || answer === undefined) {
      return jsonResponse({ success: false, error: "必須項目が不足しています" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("出席記録");

    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet("出席記録");
      sheet.appendRow([
        "タイムスタンプ", "講義回", "ラウンド", "学籍番号",
        "問題", "回答", "正誤", "経過秒", "端末ID", "重複疑い"
      ]);
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const timestamp = new Date();
    const isCorrect = correct ? "○" : "×";
    const devId = deviceId || "";

    // 重複チェック（同一講義回・同一ラウンド・同一学籍番号）
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (String(existingData[i][1]) === String(lecture) &&
          String(existingData[i][2]) === String(roundNum) &&
          String(existingData[i][3]) === String(studentId)) {
        // 既に回答済み → 上書きせず警告
        return jsonResponse({
          success: true,
          warning: "既に出席登録済みです",
          duplicate: true
        });
      }
    }

    // 同一端末ID（異なる学籍番号）の重複チェック
    let deviceDuplicate = "";
    if (devId) {
      for (let i = 1; i < existingData.length; i++) {
        if (String(existingData[i][1]) === String(lecture) &&
            String(existingData[i][2]) === String(roundNum) &&
            String(existingData[i][8]) === devId &&
            String(existingData[i][3]) !== String(studentId)) {
          // 同じ端末から別の学籍番号で回答 → 不正の疑い
          deviceDuplicate = "同一端末: " + String(existingData[i][3]);
          break;
        }
      }
    }

    // 記録追加
    sheet.appendRow([
      timestamp,
      lecture,
      roundNum,
      studentId,
      question,
      String(answer),
      isCorrect,
      elapsedSec,
      devId,
      deviceDuplicate
    ]);

    // 集計シートを更新
    updateSummary(ss);

    return jsonResponse({
      success: true,
      message: "出席を記録しました",
      timestamp: timestamp.toISOString()
    });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── GET受信（管理者用: 出席データ取得） ──
function doGet(e) {
  try {
    const action = e.parameter.action || "status";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "status") {
      return jsonResponse({
        success: true,
        message: "出席管理システム稼働中",
        spreadsheet: ss.getName()
      });
    }

    if (action === "data") {
      const lecture = e.parameter.lecture;
      const round = e.parameter.round;
      const sheet = ss.getSheetByName("出席記録");
      if (!sheet) {
        return jsonResponse({ success: true, data: [] });
      }

      const allData = sheet.getDataRange().getValues();
      const headers = allData[0];
      const rows = allData.slice(1);

      let filtered = rows;
      if (lecture) {
        filtered = filtered.filter(r => String(r[1]) === String(lecture));
      }
      if (round) {
        filtered = filtered.filter(r => String(r[2]) === String(round));
      }

      const result = filtered.map(row => ({
        timestamp: row[0],
        lecture: row[1],
        round: row[2],
        studentId: row[3],
        question: row[4],
        answer: row[5],
        correct: row[6],
        elapsedSec: row[7],
        deviceId: row[8] || "",
        deviceDuplicate: row[9] || ""
      }));

      return jsonResponse({ success: true, data: result });
    }

    if (action === "summary") {
      const sheet = ss.getSheetByName("出席記録");
      if (!sheet) {
        return jsonResponse({ success: true, summary: {} });
      }

      const allData = sheet.getDataRange().getValues().slice(1);
      const summary = {};

      allData.forEach(row => {
        const lec = String(row[1]);
        if (!summary[lec]) summary[lec] = { total: 0, correct: 0 };
        summary[lec].total++;
        if (row[6] === "○") summary[lec].correct++;
      });

      return jsonResponse({ success: true, summary });
    }

    return jsonResponse({ success: false, error: "不明なアクション: " + action });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── 集計シート更新 ──
function updateSummary(ss) {
  const recordSheet = ss.getSheetByName("出席記録");
  if (!recordSheet) return;

  let summarySheet = ss.getSheetByName("集計");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("集計");
  }

  const data = recordSheet.getDataRange().getValues().slice(1);

  // 講義回の一覧を取得
  const lectures = [...new Set(data.map(r => String(r[1])))].sort(
    (a, b) => parseInt(a) - parseInt(b)
  );

  // 学籍番号の一覧
  const students = [...new Set(data.map(r => String(r[3])))].sort();

  // ラウンドも考慮したヘッダー行
  const rounds = [1, 2];
  const header = ["学籍番号"];
  lectures.forEach(l => {
    rounds.forEach(r => header.push(`第${l}回R${r}`));
  });

  // データ行
  const rows = students.map(sid => {
    const row = [sid];
    lectures.forEach(lec => {
      rounds.forEach(r => {
        const record = data.find(
          d => String(d[3]) === sid && String(d[1]) === lec && String(d[2]) === String(r)
        );
        if (record) {
          row.push(record[6] === "○" ? "○" : "△"); // 正解○, 不正解△
        } else {
          row.push(""); // 未出席
        }
      });
    });
    return row;
  });

  // シートをクリアして書き込み
  summarySheet.clear();
  summarySheet.appendRow(header);
  rows.forEach(r => summarySheet.appendRow(r));

  // 書式設定
  if (header.length > 0) {
    summarySheet.getRange(1, 1, 1, header.length).setFontWeight("bold");
    summarySheet.setFrozenRows(1);
    summarySheet.setFrozenColumns(1);
  }
}

// ── JSON レスポンス ──
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
