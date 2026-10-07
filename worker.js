import * as XLSX from "xlsx";

export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return new Response("LINE Inspection Webhook OK");
    }

    try {
      // 讀取原始 Webhook 內容
      const rawBody = await request.text();

      // 取得 LINE Signature
      const signature = request.headers.get("x-line-signature");

      // 沒有 Signature，直接拒絕
      if (!signature) {
        console.log("LINE signature missing");
        return new Response("Unauthorized", { status: 401 });
      }

      // 驗證 Signature
      const valid = await verifyLineSignature(
        rawBody,
        signature,
        env.LINE_CHANNEL_SECRET
      );

      // Signature 不正確，直接拒絕
      if (!valid) {
        console.log("LINE signature invalid");
        return new Response("Unauthorized", { status: 401 });
      }

      // 驗證成功後才解析 JSON
      const body = JSON.parse(rawBody);

      for (const event of body.events || []) {
        // 群組訊息完全忽略
        if (event.source && event.source.type === "group") {
          continue;
        }

        if (event.type !== "message") {
          continue;
        }

        // 私人文字訊息
        if (event.message.type === "text") {
          await reply(
            event.replyToken,
            "Webhook 收到：" + event.message.text,
            env
          );
        }

        // 私人 Excel
        if (event.message.type === "file") {
          await excel(event, env);
        }
      }

      return new Response("OK");

    } catch (e) {
      console.error(e);
      return new Response("OK");
    }
  }
};


// ========================================
// LINE Webhook Signature 驗證
// ========================================

async function verifyLineSignature(body, signature, secret) {
  const encoder = new TextEncoder();

  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(secret),
    {
      name: "HMAC",
      hash: "SHA-256"
    },
    false,
    ["sign"]
  );

  const mac = await crypto.subtle.sign(
    "HMAC",
    key,
    encoder.encode(body)
  );

  const expectedSignature = arrayBufferToBase64(mac);

  return safeEqual(
    expectedSignature,
    signature
  );
}


function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";

  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary);
}


function safeEqual(a, b) {
  if (a.length !== b.length) {
    return false;
  }

  let result = 0;

  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }

  return result === 0;
}


// ========================================
// Excel 處理
// ========================================

async function excel(event, env) {
  const id = event.message.id;

  const r = await fetch(
    "https://api-data.line.me/v2/bot/message/" + id + "/content",
    {
      headers: {
        Authorization: "Bearer " + env.LINE_CHANNEL_ACCESS_TOKEN
      }
    }
  );

  if (!r.ok) {
    await reply(
      event.replyToken,
      "❌ Excel 下載失敗\nHTTP：" + r.status,
      env
    );
    return;
  }

  const book = XLSX.read(
    await r.arrayBuffer(),
    { type: "array" }
  );

  const sheetName = book.SheetNames.find(
    name => name.trim() === "僅異常"
  );

  if (!sheetName) {
    await reply(
      event.replyToken,
      "❌ 找不到「僅異常」工作表\n\n目前工作表：\n" +
      book.SheetNames.join("\n"),
      env
    );
    return;
  }

  const rows = XLSX.utils.sheet_to_json(
    book.Sheets[sheetName],
    {
      header: 1,
      defval: ""
    }
  );

 if (rows.length <= 1) {
  const msg =
    "✅ 化工廠巡檢完成\n\n" +
    "📅 今日巡檢無異常";

  await reply(event.replyToken, msg, env);
  await push(msg, env);
  return;
}

  const headers = rows[0];
  const col = {};

  headers.forEach((h, i) => {
    col[String(h).trim()] = i;
  });

  const data = rows
    .slice(1)
    .filter(row =>
      row.some(cell => String(cell).trim() !== "")
    );

  const get = (row, name) => {
    if (col[name] === undefined) {
      return "";
    }

    return String(
      row[col[name]] ?? ""
    ).trim();
  };

  const records = data.map(row => ({
    date: get(row, "日期"),
    inspector: get(row, "巡檢人員"),
    area: get(row, "廠區"),
    equipment: get(row, "設備"),
    item: get(row, "巡檢項目"),
    value: get(row, "數值"),
    unit: get(row, "單位"),
    lower: get(row, "下限"),
    upper: get(row, "上限"),
    status: get(row, "狀態"),
    judgment: get(row, "判定"),
    reason: get(row, "異常原因/現場狀況")
  }));


  // ========================================
  // 報告標題
  // ========================================

  let msg = "🚨 化工廠巡檢異常\n\n";

  if (records[0].date) {
    msg += "📅 " + records[0].date + "\n";
  }

  if (records[0].inspector) {
    msg += "👤 " + records[0].inspector + "\n";
  }

  msg += "⚠️ 共 " + records.length + " 項異常\n";


  // ========================================
  // 按廠區 + 設備分組
  // ========================================

  const groups = {};

  records.forEach(record => {
    const key =
      record.area +
      "|||" +
      record.equipment;

    if (!groups[key]) {
      groups[key] = {
        area: record.area,
        equipment: record.equipment,
        records: []
      };
    }

    groups[key].records.push(record);
  });


  // ========================================
  // 產生報告
  // ========================================

  let n = 1;

  Object.keys(groups).forEach(key => {
    const g = groups[key];

    msg += "\n━━━━━━━━━━━━━━\n";
    msg +=
      n +
      ". " +
      g.area +
      "｜" +
      g.equipment +
      "\n";

    g.records.forEach(record => {
      msg +=
        "\n🔸 " +
        record.item +
        "\n";

      if (record.value) {
        msg +=
          "數值：" +
          record.value;

        if (record.unit) {
          msg += " " + record.unit;
        }

        msg += "\n";
      }

      if (record.lower && record.upper) {
        msg +=
          "正常範圍：" +
          record.lower +
          "～" +
          record.upper;

        if (record.unit) {
          msg += " " + record.unit;
        }

        msg += "\n";

      } else if (record.lower) {
        msg +=
          "正常下限：" +
          record.lower;

        if (record.unit) {
          msg += " " + record.unit;
        }

        msg += "\n";

      } else if (record.upper) {
        msg +=
          "正常上限：" +
          record.upper;

        if (record.unit) {
          msg += " " + record.unit;
        }

        msg += "\n";
      }

      if (record.status) {
        msg +=
          "狀態：" +
          record.status +
          "\n";
      }

      if (record.judgment) {
        msg +=
          "判定：" +
          record.judgment +
          "\n";
      }

      if (record.reason) {
        msg +=
          "處置／現場狀況：" +
          record.reason +
          "\n";
