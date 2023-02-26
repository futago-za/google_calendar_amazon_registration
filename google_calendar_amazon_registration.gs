const gmailBaseUrl = "https://mail.google.com/mail/u/0/#inbox";
const webhook_url = "***";

// entry point
function main() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const messages = getMail(yesterday);
  deleteAllEvents()
  for (let i in messages) {
    for (let j in messages[i]) {
      const message = messages[i][j];
      parse(message);
    }
  }

  // 受け取り済みの予定は削除
  removeExpiredEvent();
}

function deleteAllEvents() {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(new Date("2019/1/1"), new Date("2024/1/1"));
  if (events.length !== 0) {
    for (const event of events) {
      if (event.getTitle().match(RegExp('Amazonから荷物受け取り'))) {
       event.deleteEvent();
      }
    }
  }
  return;
}

function getMail(date) {
  const threads = GmailApp.search(`subject:Amazon.co.jpでのご注文 AND before:${date.toLocaleDateString()}`);
  return GmailApp.getMessagesForThreads(threads).reverse();
}

function parse(message) {
  const receivedMailDate = new Date(message.getDate())
  const strMessage = message.getPlainBody();
  const mailId = message.getId();

  // Kindleは対象外
  let regexp = RegExp("コンテンツと端末の管理", 'gi');
  const isKindle = strMessage.match(regexp);
  if (isKindle) return;

  regexp = RegExp("ご注文の確認", 'gi');
  const isConfirmMail = strMessage.match(regexp);
  if (isConfirmMail) {
    parseCofirmMail(receivedMailDate, strMessage, mailId);
  } else {
    parseUpdateMail(strMessage, mailId);
  }
}

function parseCofirmMail(receivedMailDate, message, mailId) {
  const rows = message.split(/\r\n|\n/);
  let dateIndex = 0;

  // 注文番号の抽出
  const orderId = getOrderId("注文番号：", message);
  if (!orderId) return;

  for (const [index, value] of rows.entries()) {
    if (value.match(/お届け予定日：/)) {
      dateIndex = index + 1;
      if (dateIndex >= rows.length) {
        console.log("out of range");
        break;
      }
      const deliveryDate = rows[dateIndex];
      if (!hasDate(deliveryDate)) continue;

      let year = receivedMailDate.getFullYear();
      const month = getMonth(deliveryDate);
      const dayOfMonth = getDayOfMonth(deliveryDate);
      if (month < receivedMailDate.getMonth() + 1 || (month === receivedMailDate.getDate() && dayOfMonth < receivedMailDate.getDate())) {
        year += 1;
      } 

      if (isRange(deliveryDate)) {
        let fromYear = year;
        const fromMonth = month;
        const fromDayOfMonth = dayOfMonth;

        let toDateIndex = dateIndex + 1;
        if (toDateIndex > rows.length) {
          console.log("out of range");
          break;
        }
        const toDeliveryDate = rows[toDateIndex];
        let toYear = fromYear;
        const toMonth = getMonth(toDeliveryDate);
        const toDayOfMonth = getDayOfMonth(toDeliveryDate);
        // 年末年始を考慮
        if (toMonth < fromMonth) {
          toYear += 1;
        }
        // イベント作成
        const start = new Date(fromYear, fromMonth - 1, fromDayOfMonth, 0, 0, 0);
        const end = new Date(toYear, toMonth - 1, toDayOfMonth, 23, 59, 59);
        registCalenderRangeEvent(start, end, orderId, mailId);
      } else {
        // イベント作成
        const registeredDate = new Date(year, month - 1, dayOfMonth);
        registCalenderAllDayEvent(registeredDate, orderId, mailId)
      }
    }
  }
}

function parseUpdateMail(message, mailId) {
  let prevDeliveryDate = [];

  // 注文番号の抽出
  const orderId = getOrderId("ご注文番号#", message);
  if (!orderId) return;

  let prefix = "前回のお届け予定日:";
  let regexp = RegExp(prefix + '.*', 'gi');
  let result = message.match(regexp);
  if (result) {
    const temp = result[0].replace(prefix, '').trim();
    prevDeliveryDate = temp.split('-');
  }

  prefix = "新しいお届け予定日:";
  regexp = RegExp(prefix + '.*', 'gi');
  result = message.match(regexp);
  if (!result) {
    Logger.log('Not found new delivery date');
    return;
  }
  const deliveryDate = result[0].replace(prefix, '').trim();
  let newDeliveryDate = [];
  if (hasDate(deliveryDate)) {
    if (isRange(deliveryDate)) {
      newDeliveryDate.push(...deliveryDate.split('-'));
    } else {
      newDeliveryDate.push(deliveryDate);
    }
  }

  if (prevDeliveryDate.length && newDeliveryDate.length) {
    // 古い予定を削除
    if (prevDeliveryDate.length == 1) {
      removeCalendarEvent(new Date(prevDeliveryDate[0].trim()), orderId);
    } else {
      removeCalendarRanageEvent(new Date(prevDeliveryDate[0].trim()), new Date(prevDeliveryDate[1].trim()), orderId);
    }

    // 新しい予定を登録
    if (newDeliveryDate.length == 1) {
      const date = new Date(newDeliveryDate[0].trim());
      registCalenderAllDayEvent(date, orderId, mailId);
    } else {
      const start = new Date(newDeliveryDate[0].trim());
      const end = new Date(newDeliveryDate[1].trim());
      registCalenderRangeEvent(start, end, orderId, mailId);
    }
  } else {
    // Slackでの通知
    // notifySlack();
  }
}

// Amazonから"MM/dd"の形式でお届け予定日が記載されている
function hasDate(str) {
  const arr = str.match(/[0-9]+/g);
  if (arr) {
    if (arr.length >= 2) {
      return true;
    } else {
      return false;
    }
  } else {
    return false;
  }
}

function getOrderId(prefix, message) {
  const regexp = RegExp(prefix + '.*', 'gi');
  const result = message.match(regexp);
  if (!result) {
    console.log('Not found order id');
    return null;
  }
  return result[0].split(prefix)[1];
}

function isRange(str) {
  return str.match(/-/) ? true : false;
}

function getMonth(str) {
  return str.match(/[0-9]+/g)[0];
}

function getDayOfMonth(str) {
  return str.match(/[0-9]+/g)[1];
}

function registCalenderAllDayEvent(date, id, mailId) {
  const calendar = CalendarApp.getDefaultCalendar();
  const option = {
    description: `注文番号：${id}\nメール:${gmailBaseUrl}/${mailId}\n注文リンク：https://www.google.com/url?q=https://www.amazon.co.jp/gp/css/your-orders-access/ref%3DTE_tex_g&sa=D&source=calendar&usd=2&usg=AOvVaw1fp7R2H5TZATHFP83lIdJ3`,
  }

  const event = calendar.createAllDayEvent('Amazonから荷物受け取り', date, option);
  event.removeAllReminders();
}

function registCalenderRangeEvent(start, end, id, mailId) {
  const calendar = CalendarApp.getDefaultCalendar();
  const option = {
    description: `注文番号：${id}\nメール:${gmailBaseUrl}/${mailId}\n注文リンク：https://www.google.com/url?q=https://www.amazon.co.jp/gp/css/your-orders-access/ref%3DTE_tex_g&sa=D&source=calendar&usd=2&usg=AOvVaw1fp7R2H5TZATHFP83lIdJ3`,
  }

  const event = calendar.createEvent('Amazonから荷物受け取り',start, end, option);
  event.removeAllReminders();
}

function removeCalendarEvent(date, id) {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEventsForDay(date);
  if (events.length !== 0) {
    for (const event of events) {
      if (event.getDescription().match(RegExp(id))) {
       event.deleteEvent();
       return; 
      }
    }
  }
  console.log("No event found.");
  return;
}

function removeCalendarRanageEvent(start, end, id) {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(start, end);
  if (events.length !== 0) {
    for (const event of events) {
      if (event.getDescription().match(RegExp(id))) {
       event.deleteEvent();
       return; 
      }
    }
  }
  console.log("No event found.");
  return;
}

function removeExpiredEvent() {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(new Date(0), new Date());
  if (events.length !== 0) {
    for (const event of events) {
      if (event.getTitle().match(RegExp('Amazonから荷物受け取り'))) {
       event.deleteEvent();
      }
    }
  }
}

function notifySlack() {
  const userName = "GASくん"
  const icon     = ":google_apps_script:"
  let message  = `Amazonから商品のお届け日が更新されました。\nメールを確認しよう！`

  let jsonData = {
    "username": userName,
    "icon_emoji": icon,
    "text": message
  }

  let payload = JSON.stringify(jsonData)

  let options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload
  };

  UrlFetchApp.fetch(webhook_url, options);
}
