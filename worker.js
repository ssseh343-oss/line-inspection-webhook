import * as XLSX from "xlsx";

export default {
  async fetch(request, env) {

    if (request.method !== "POST") {
      return new Response("LINE Inspection Webhook OK");
    }

    try {
      const body = await request.json();

      for (const event of body.events || []) {

        // ========================================
        // 群組訊息全部忽略
        // ========================================

        if (event.source && event.source.type === "group") {
          continue;
        }


        // ========================================
        // 私訊文字
        // ========================================

        if (
          event.type === "message" &&
          event.message &&
          event.message.type === "text"
        ) {

          await replyMessage(
            event.replyToken,
            "Webhook 收到：" + event.message.text,
            env
          );

          continue;
        }


        // ========================================
        // 私訊 Excel
        // ========================================

        if (
          event.type === "message" &&
          event.message &&
          event.message.type === "file"
        ) {

          await handleExcel(event, env);
        }
      }

      return new Response("OK");

    } catch (error) {

      console.error(
        "Webhook error:",
        error
      );

      return new Response("OK");
    }
  }
};


// ================================================
// Excel 處理
// ================================================

async function handleExcel(event, env) {

  const fileName = event.message.fileName;
  const messageId = event.message.id;


  // ==============================================
  // 從 LINE 下載 Excel
  // ==============================================

  const response = await fetch(
    "https://api-data.line.me/v2/bot/message/" +
    messageId +
    "/content",
    {
      headers: {
        Authorization:
          "Bearer " +
          env.LINE_CHANNEL_ACCESS_TOKEN
      }
    }
  );


  if (!response.ok) {

    await replyMessage(
      event.replyToken,
      "❌ Excel 下載失敗\nHTTP：" +
      response.status,
      env
    );

    return;
  }


  const buffer =
    await response.arrayBuffer();


  // ==============================================
  // 讀取 Excel
  // ==============================================

  const workbook = XLSX.read(
    buffer,
    {
      type: "array"
    }
  );


  // ==============================================
  // 找「僅異常」工作表
  // ==============================================

  const targetSheet =
    workbook.SheetNames.find(
      function (name) {
        return name.trim() === "僅異常";
      }
    );


  if (!targetSheet) {

    await replyMessage(
      event.replyToken,
      "❌ 找不到「僅異常」工作表\n\n" +
      "目前工作表：\n" +
      workbook.SheetNames.join("\n"),
      env
    );

    return;
  }


  const worksheet =
    workbook.Sheets[targetSheet];


  // ==============================================
  // Excel → 陣列
  // ==============================================

  const rows =
    XLSX.utils.sheet_to_json(
      worksheet,
      {
        header: 1,
        defval: ""
      }
    );


  // ==============================================
  // 沒有異常
  // ==============================================

  if (rows.length <= 1) {

    const message =
      "✅ " +
      fileName +
      "\n\n今日沒有異常巡檢紀錄。";


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


  // ==============================================
  // 建立欄位索引
  // ==============================================

  const headers = rows[0];
  const col = {};


  headers.forEach(
    function (header, index) {

      col[
        String(header).trim()
      ] = index;

    }
  );


  // ==============================================
  // 轉換資料
  // ==============================================

  const records =
    rows
      .slice(1)

      .filter(
        function (row) {

          return row.some(
            function (cell) {

              return String(cell).trim() !== "";

            }
          );

        }
      )

      .map(
        function (row) {

          return {

            date:
              getCell(
                row,
                col,
                "日期"
              ),

            time:
              getCell(
                row,
                col,
                "時間"
              ),

            inspector:
              getCell(
                row,
                col,
                "巡檢人員"
              ),

            shift:
              getCell(
                row,
                col,
                "班別"
              ),

            area:
              getCell(
                row,
                col,
                "廠區"
              ),

            equipment:
              getCell(
                row,
                col,
                "設備"
              ),

            item:
              getCell(
                row,
                col,
                "巡檢項目"
              ),

            value:
              getCell(
                row,
                col,
                "數值"
              ),

            status:
              getCell(
                row,
                col,
                "狀態"
              ),

            unit:
              getCell(
                row,
                col,
                "單位"
              ),

            lower:
              getCell(
                row,
                col,
                "下限"
              ),

            upper:
              getCell(
                row,
                col,
                "上限"
              ),

            judgment:
              getCell(
                row,
                col,
                "判定"
              ),

            reason:
              getCell(
                row,
                col,
                "異常原因/現場狀況"
              )

          };

        }
      );


  // ==============================================
  // 建立報告
  // ==============================================

  const report =
    buildReport(records);


  // ==============================================
  // 回覆傳 Excel 的人
  // ==============================================

  await replyMessage(
    event.replyToken,
    report,
    env
  );


  // ==============================================
  //
