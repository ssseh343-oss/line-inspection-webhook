import * as XLSX from "xlsx";

export default {
  async fetch(request, env) {

    if (request.method !== "POST") {
      return new Response("LINE Inspection Webhook OK", {
        status: 200
      });
    }

    try {
      const body = await request.json();

      if (!body.events || body.events.length === 0) {
        return new Response("OK", { status: 200 });
      }

      for (const event of body.events) {

        // ===== 文字訊息 =====
        if (
          event.type === "message" &&
          event.message?.type === "text"
        ) {
          await replyMessage(
  event.replyToken,
  `Webhook 收到：${event.message.text}\n\nGroup ID：${event.source?.groupId || "不是群組事件"}`,
  env
);
await pushMessage(
  "📢 化工廠巡檢\n\n這是一則群組推播測試。",
  env
);          
        }

        // ===== Excel 檔案 =====
        else if (
          event.type === "message" &&
          event.message?.type === "file"
        ) {

          const fileName = event.message.fileName;
          const messageId = event.message.id;

          // 從 LINE 下載 Excel
          const response = await fetch(
            `https://api-data.line.me/v2/bot/message/${messageId}/content`,
            {
              headers: {
                "Authorization":
                  `Bearer ${env.LINE_CHANNEL_ACCESS_TOKEN}`
              }
            }
          );

          if (!response.ok) {
            await replyMessage(
              event.replyToken,
              `❌ Excel 下載失敗\nHTTP：${response.status}`,
              env
            );
            continue;
          }

          const buffer = await response.arrayBuffer();

          // ===== 解析 Excel =====
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
            continue;
          }

          const worksheet = workbook.Sheets[targetSheet];

          const rows = XLSX.utils.sheet_to_json(
            worksheet,
            {
              header: 1,
              defval: ""
            }
          );

          // 沒有異常資料
          if (rows.length <= 1) {
            await replyMessage(
              event.replyToken,
              `✅ ${fileName}\n\n今日沒有異常巡檢紀錄。`,
              env
            );
            continue;
          }

          // ===== 建立欄位索引 =====
          const headers = rows[0];

          const col = {};

          headers.forEach((header, index) => {
            col[String(header).trim()] = index;
          });

          // ===== 讀取資料 =====
          const records = rows
            .slice(1)
            .filter(row =>
              row.some(cell => String(cell).trim() !== "")
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

          // ===== 組成漂亮報告 =====
          let report = "🚨 化工廠巡檢異常\n\n";

          if (records[0].date) {
            report += `📅 ${records[0].date}\n`;
          }

          if (records[0].inspector) {
            report += `👤 ${records[0].inspector}\n`;
          }

          report += `⚠️ 共 ${records.length} 項異常\n`;

          // ===== 依廠區 + 設備分組 =====
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

            report += "\n";
            report += `━━━━━━━━━━━━━━\n`;
            report += `${groupNumber}. ${group.area}｜${group.equipment}\n`;

            for (const record of group.records) {

              report += `\n🔸 ${record.item}\n`;

              if (record.value !== "") {
                report += `數值：${record.value}`;

                if (record.unit) {
                  report += ` ${record.unit}`;
                }

                report += "\n";
              }

              // 正常範圍
              if (
                record.lower !== "" &&
                record.upper !== ""
              ) {
                report +=
                  `正常範圍：${record.lower}～${record.upper}`;

                if (record.unit) {
                  report += ` ${record.unit}`;
                }

                report += "\n";
              }
              else if (record.lower !== "") {

                report +=
                  `正常下限：${record.lower}`;

                if (record.unit) {
                  report += ` ${record.unit}`;
                }

                report += "\n";
              }
              else if (record.upper !== "") {

                report +=
                  `正常上限：${record.upper}`;

                if (record.unit) {
                  report += ` ${record.unit}`;
                }

                report += "\n";
              }

              if (record.status) {
                report += `狀態：${record.status}\n`;
              }

              if (record.judgment) {
                report += `判定：${record.judgment}\n`;
              }

              if (record.reason) {
                report += `處置／現場狀況：${record.reason}\n`;
              }
            }

            groupNumber++;
          }

          report += "\n━━━━━━━━━━━━━━\n";
          report += `📋 異常項目合計：${records.length}`;

          await replyMessage(
            event.replyToken,
            report,
            env
          );
        }
      }

      return new Response("OK", { status: 200 });

    } catch (error) {

      console.error(error);

      return new Response(
        "❌ Worker 執行錯誤：" + error.message,
        {
          status: 200
        }
      );
    }
  }
};


// ===== 取得 Excel 儲存格 =====
function getCell(row, col, name) {

  const index = col[name];

  if (index === undefined) {
    return "";
  }

  return String(row[index] ?? "").trim();
}


// ===== 回覆 LINE =====
async function replyMessage(replyToken, text, env) {
async function pushMessage(text, env) {

  await fetch(
    "https://api.line.me/v2/bot/message/push",
    {
      method: "POST",

      headers: {
        "Content-Type": "application/json",
        "Authorization":
          `Bearer ${env.LINE_CHANNEL_ACCESS_TOKEN}`
      },

      body: JSON.stringify({
        to: env.LINE_GROUP_ID,

        messages: [
          {
            type: "text",
            text: text
          }
        ]
      })
    }
  );
}
  if (text.length > 4900) {
    text =
      text.substring(0, 4900) +
      "\n\n⚠️ 顯示內容過長，已截斷";
  }

  await fetch(
    "https://api.line.me/v2/bot/message/reply",
    {
      method: "POST",

      headers: {
        "Content-Type": "application/json",
        "Authorization":
          `Bearer ${env.LINE_CHANNEL_ACCESS_TOKEN}`
      },

      body: JSON.stringify({
        replyToken: replyToken,

        messages: [
          {
            type: "text",
            text: text
          }
        ]
      })
    }
  );
}
