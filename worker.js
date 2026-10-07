import * as XLSX from "xlsx";

export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return new Response("LINE Inspection Webhook OK");
    }

    try {
      const body = await request.json();

      for (const event of body.events || []) {
        // 群組訊息全部忽略
        if (event.source?.type === "group") continue;

        // 私訊文字
        if (event.type === "message" && event.message?.type === "text") {
          await replyMessage(
            event.replyToken,
            `Webhook 收到：${event.message.text}`,
            env
          );
          continue;
        }

        // 私訊 Excel
        if (event.type === "message" && event.message?.type === "file") {
          await handleExcel(event, env);
        }
      }

      return new Response("OK");
    } catch (error) {
      console.error("Webhook error:", error);
      return new Response("OK");
    }
  }
};

async function handleExcel(event, env) {
  const fileName = event.message.fileName;
  const messageId = event.message.id;

  const response = await fetch(
    `https://api-data.line.me/v2/bot/message/${messageId}/content`,
    {
      headers: {
        Authorization: `Bearer ${env.LINE_CHANNEL_ACCESS_TOKEN}`
      }
    }
  );

  if (!response.ok) {
    await replyMessage(
      event.replyToken,
      `❌ Excel 下載失敗\nHTTP：${response.status}`,
      env
    );
    return;
  }

  const buffer = await response.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array"
  });

  const targetSheet = workbook.SheetNames.find(
    name => name.trim() === "僅異常"
  );

  if (!targetSheet) {
    await replyMessage(
      event.replyToken,
      `❌ 找不到「僅異常」工作表\n\n目前工作表：\n${workbook.SheetNames.join("\n")}`,
      env
    );
    return;
  }

  const rows = XLSX.utils.sheet_to_json(
    workbook.Sheets[targetSheet],
    {
      header: 1,
      defval: ""
    }
  );

  if (rows.length <= 1) {
    const message =
      `✅ ${fileName}\n\n今日沒有異常巡檢紀錄。`;

    await replyMessage(
      event.replyToken,
      message,
      env
    );

    await pushMessage(
      message,
      env
    );

    return;
  }

  const headers = rows[0];
  const col = {};

  headers.forEach((header, index) => {
    col[String(header).trim()] = index;
  });

  const records = rows
    .slice(1)
    .filter(row =>
      row.some(cell =>
        String(cell).trim() !== ""
      )
    )
    .map(row => ({
      date: getCell(row, col, "日期"),
      time: getCell(row, col, "時間"),
      inspector: getCell(row, col, "巡檢人員"),
      shift: getCell(row, col, "班別"),
      area: getCell(row, col, "廠區"),
      equipment: getCell(row, col, "設備"),
      item: getCell(row, col, "巡檢項目"),
      value: getCell(row, col, "數值"),
      status: getCell(row, col, "狀態"),
      unit: getCell(row, col, "單位"),
      lower: getCell(row, col, "下限"),
      upper: getCell(row, col, "上限"),
      judgment: getCell(row, col, "判定"),
      reason: getCell(row, col, "異常原因/現場狀況")
    }));

  const report = buildReport(records);

  // 回覆傳 Excel 的人
  await replyMessage(
    event.replyToken,
    report,
    env
  );

  // 推播到指定群組
  await pushMessage(
    report,
    env
  );
}

function buildReport(records) {
  let report =
    "🚨 化工廠巡檢異常\n\n";

  if (records[0]?.date) {
    report += `📅 ${records[0].date}\n`;
  }

  if (records[0]?.inspector) {
    report += `👤 ${records[0].inspector}\n`;
  }

  report +=
    `⚠️ 共 ${records.length} 項異常\n`;

  const groups = {};

  for (const record of records) {
    const key =
      `${record.area}|||${record.equipment}`;

    if (!groups[key]) {
      groups[key] = {
        area: record.area,
        equipment: record.equipment,
        records: []
      };
    }

    groups[key].records.push(record);
  }

  let groupNumber = 1;

  for (const key of Object.keys(groups)) {
    const group = groups[key];

    report +=
      "\n━━━━━━━━━━━━━━\n";

    report +=
      `${groupNumber}. ${group.area}｜${group.equipment}\n`;

    for (const record of group.records) {
      report +=
        `\n🔸 ${record.item}\n`;

      if (record.value !== "") {
        report +=
          `數值：${record.value}`;

        if (record.unit) {
          report +=
            ` ${record.unit}`;
        }

        report += "\n";
      }

      if (
        record.lower !== "" &&
        record.upper !== ""
      ) {
        report +=
          `正常範圍：${record.lower}～${record.upper}`;

        if (record.unit) {
          report +=
            ` ${record.unit}`;
        }

        report += "\n";

      } else if (record.lower !== "") {
        report +=
          `正常下限：${record.lower}`;

        if (record.unit) {
          report +=
            ` ${record.unit}`;
        }

        report += "\n";

      } else if (record.upper !== "") {
        report +=
          `正常上限：${record.upper}`;

        if (record.unit) {
          report +=
            ` ${record.unit}`;
        }

        report += "\n";
      }

      if (record.status) {
       
