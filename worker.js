import * as XLSX from "xlsx";

export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return new Response("LINE Inspection Webhook OK");
    }

    try {
      // 先讀取原始內容，供 LINE Signature 驗證
      const rawBody = await request.text();

      // 驗證 LINE Webhook
      const signature = request.headers.get("x-line-signature");

      if (!signature) {
        console.log("LINE signature missing");
        return new Response("Unauthorized", { status: 401 });
      }

      const valid = await verifySignature(
        rawBody,
        signature,
        env.LINE_CHANNEL_SECRET
      );

      if (!valid) {
        console.log("LINE signature invalid");
        return new Response("Unauthorized", { status: 401 });
      }

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


// ==============================
// LINE Signature 驗證
// ==============================

async function verifySignature(body, signature, secret) {
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

  const expected = arrayBufferToBase64(mac);

  return timingSafeEqual(expected, signature);
}


function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";

  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary);
}


function timingSafeEqual(a, b) {
  if (a.length !== b.length) {
    return false;
  }

  let result = 0;

  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }

  return result === 0;
}


// ==============================
// Excel 處理
// ==============================

async function excel
