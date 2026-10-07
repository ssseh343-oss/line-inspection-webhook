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
            `Webhook 收到：${event.message.text}`,
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

          const sheetNames = workbook.SheetNames;

          // 找「僅異常」
          const targetSheet = sheetNames.find(
            name => name.trim() === "僅異常"
          );

          if (!targetSheet) {

            await replyMessage(
              event.replyToken,
              `❌ 找不到「僅異常」工作表

目前 Excel 工作表：
${sheetNames.join("\n")}`,
              env
            );

            continue;
          }

          // 轉成二維陣列
          const worksheet = workbook.Sheets[targetSheet];

          const rows = XLSX.utils.sheet_to_json(
            worksheet,
            {
              header: 1,
              defval: ""
            }
          );

          // ===== 整理成文字 =====
          let output = `📊 已成功讀取「僅異常」

檔案：${fileName}

共 ${rows.length} 列

`;

          // 最多顯示前 15 列
          const displayRows = rows.slice(0, 15);

          for (const row of displayRows) {
            output += row.join(" ｜ ") + "\n";
          }

          await replyMessage(
            event.replyToken,
            output,
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


async function replyMessage(replyToken, text, env) {

  // LINE 單則文字訊息最多 5000 字元
  if (text.length > 4900) {
    text = text.substring(0, 4900) +
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
