import * as XLSX from "xlsx";

export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return new Response("LINE Inspection Webhook OK");
    }

    try {
      const body = await request.json();

      for (const event of body.events || []) {
  if (event.source && event.source.type === "group") {
    console.log("GROUP ID:", event.source.groupId);
    continue;
  }

        if (event.type !== "message") continue;

        if (event.message.type === "text") {
          await reply(event.replyToken, "Webhook 收到：" + event.message.text, env);
        }

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
    await reply(event.replyToken, "❌ Excel 下載失敗\nHTTP：" + r.status, env);
    return;
  }

  const book = XLSX.read(await r.arrayBuffer(), { type: "array" });

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
    { header: 1, defval: "" }
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

  const data = rows.slice(1).filter(row =>
    row.some(cell => String(cell).trim() !== "")
  );

  const get = (row, name) => {
    if (col[name] === undefined) return "";
    return String(row[col[name]] ?? "").trim();
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

  let msg = "🚨 化工廠巡檢異常\n\n";

  if (records[0].date) {
    msg += "📅 " + records[0].date + "\n";
  }

  if (records[0].inspector) {
    msg += "👤 " + records[0].inspector + "\n";
  }

  msg += "⚠️ 共 " + records.length + " 項異常\n";

  const groups = {};

  records.forEach(record => {
    const key = record.area + "|||" + record.equipment;

    if (!groups[key]) {
      groups[key] = {
        area: record.area,
        equipment: record.equipment,
        records: []
      };
    }

    groups[key].records.push(record);
  });

  let n = 1;

  Object.keys(groups).forEach(key => {
    const g = groups[key];

    msg += "\n━━━━━━━━━━━━━━\n";
    msg += n + ". " + g.area + "｜" + g.equipment + "\n";

    g.records.forEach(record => {
      msg += "\n🔸 " + record.item + "\n";

      if (record.value) {
        msg += "數值：" + record.value;
        if (record.unit) msg += " " + record.unit;
        msg += "\n";
      }

      if (record.lower && record.upper) {
        msg += "正常範圍：" + record.lower + "～" + record.upper;
        if (record.unit) msg += " " + record.unit;
        msg += "\n";
      } else if (record.lower) {
        msg += "正常下限：" + record.lower;
        if (record.unit) msg += " " + record.unit;
        msg += "\n";
      } else if (record.upper) {
        msg += "正常上限：" + record.upper;
        if (record.unit) msg += " " + record.unit;
        msg += "\n";
      }

      if (record.status) {
        msg += "狀態：" + record.status + "\n";
      }

      if (record.judgment) {
        msg += "判定：" + record.judgment + "\n";
      }

      if (record.reason) {
        msg += "處置／現場狀況：" + record.reason + "\n";
      }
    });

    n++;
  });

  msg += "\n━━━━━━━━━━━━━━\n";
  msg += "📋 異常項目合計：" + records.length;

  await reply(event.replyToken, msg, env);
  await push(msg, env);
}

async function reply(token, text, env) {
  if (text.length > 4900) {
    text = text.substring(0, 4900) + "\n\n⚠️ 顯示內容過長，已截斷";
  }

  const r = await fetch(
    "https://api.line.me/v2/bot/message/reply",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + env.LINE_CHANNEL_ACCESS_TOKEN
      },
      body: JSON.stringify({
        replyToken: token,
        messages: [{ type: "text", text: text }]
      })
    }
  );

  console.log("LINE REPLY:", r.status, await r.text());
}

async function push(text, env) {
  const r = await fetch(
    "https://api.line.me/v2/bot/message/push",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + env.LINE_CHANNEL_ACCESS_TOKEN
      },
      body: JSON.stringify({
        to: env.LINE_GROUP_ID,
        messages: [{ type: "text", text: text }]
      })
    }
  );

  console.log("LINE PUSH:", r.status, await r.text());

  if (!r.ok) {
    throw new Error("LINE Push API error: " + r.status);
  }
}
