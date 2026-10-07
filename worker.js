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

        // 文字訊息
        if (
          event.type === "message" &&
          event.message?.type === "text"
        ) {
          const userMessage = event.message.text;

          await replyMessage(
            event.replyToken,
            `Webhook 收到：${userMessage}`,
            env
          );
        }

        // 檔案訊息
        else if (
          event.type === "message" &&
          event.message?.type === "file"
        ) {

          const fileName = event.message.fileName;
          const messageId = event.message.id;

          // 從 LINE 下載檔案
          const response = await fetch(
            `https://api-data.line.me/v2/bot/message/${messageId}/content`,
            {
              method: "GET",
              headers: {
                "Authorization":
                  `Bearer ${env.LINE_CHANNEL_ACCESS_TOKEN}`
              }
            }
          );

          if (!response.ok) {

            await replyMessage(
              event.replyToken,
              `❌ Excel 下載失敗\nHTTP 狀態碼：${response.status}`,
              env
            );

            continue;
          }

          const buffer = await response.arrayBuffer();

          const size = buffer.byteLength;
          const contentType =
            response.headers.get("content-type") || "未知";

          const bytes = new Uint8Array(buffer);

          const isZip =
            bytes.length >= 2 &&
            bytes[0] === 0x50 &&
            bytes[1] === 0x4B;

          await replyMessage(
            event.replyToken,
            `📥 Excel 已成功下載

檔案：${fileName}
大小：${size.toLocaleString()} bytes
Content-Type：${contentType}

${isZip
  ? "✅ Excel ZIP 結構確認正常"
  : "⚠️ 檔案不是標準 ZIP 結構"}

下一步：讀取「僅異常」工作表`,
            env
          );
        }
      }

      return new Response("OK", { status: 200 });

    } catch (error) {

      console.error(error);

      return new Response("OK", {
        status: 200
      });
    }
  }
};


async function replyMessage(replyToken, text, env) {

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
