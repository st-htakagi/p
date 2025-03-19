// functions.js
Office.initialize = function () {};

// 送信イベントハンドラー
function onItemSend(event) {
  Office.context.mailbox.item.getAttachmentsAsync(function (result) {
      // 添付ファイルがあるかどうかを確認
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          var attachments = result.value;
          
          // 添付ファイルがない場合は、そのまま送信
          if (attachments.length === 0) {
              event.completed({ allowEvent: true });
              return;
          }
          
          // 添付ファイルがある場合、元のメールを送信しない
          event.completed({ allowEvent: false });
          
          // メール本文を取得
          Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
              if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                  var body = bodyResult.value;
                  var subject = Office.context.mailbox.item.subject;
                  
                  // 宛先、CC、BCCを取得
                  getRecipientsInfo(function (recipientsInfo) {
                      // 複製メールを作成して送信
                      createAndSendClonedEmail(subject, body, recipientsInfo, attachments);
                  });
              } else {
                  console.error("メール本文の取得に失敗しました。", bodyResult.error);
                  showNotification("エラー", "メール処理中にエラーが発生しました。");
              }
          });
      } else {
          console.error("添付ファイルの取得に失敗しました。", result.error);
          event.completed({ allowEvent: true }); // エラー時は元のメールを送信
      }
  });
}

// 宛先情報を取得する関数
function getRecipientsInfo(callback) {
  var recipientsInfo = { to: [], cc: [], bcc: [] };
  
  // To宛先の取得
  Office.context.mailbox.item.to.getAsync(function (toResult) {
      if (toResult.status === Office.AsyncResultStatus.Succeeded) {
          recipientsInfo.to = toResult.value.map(function (recipient) {
              return { displayName: recipient.displayName, emailAddress: recipient.emailAddress };
          });
          
          // CC宛先の取得
          Office.context.mailbox.item.cc.getAsync(function (ccResult) {
              if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
                  recipientsInfo.cc = ccResult.value.map(function (recipient) {
                      return { displayName: recipient.displayName, emailAddress: recipient.emailAddress };
                  });
                  
                  // BCC宛先の取得
                  Office.context.mailbox.item.bcc.getAsync(function (bccResult) {
                      if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
                          recipientsInfo.bcc = bccResult.value.map(function (recipient) {
                              return { displayName: recipient.displayName, emailAddress: recipient.emailAddress };
                          });
                          callback(recipientsInfo);
                      } else {
                          callback(recipientsInfo); // BCCの取得に失敗しても続行
                      }
                  });
              } else {
                  callback(recipientsInfo); // CCの取得に失敗しても続行
              }
          });
      } else {
          callback(recipientsInfo); // Toの取得に失敗しても続行
      }
  });
}

// 複製メールを作成して送信する関数
function createAndSendClonedEmail(subject, body, recipientsInfo, attachments) {
  // 添付ファイルの処理（AWS S3にアップロード）
  processAttachments(attachments, function (urls) {
      // 新しいメール本文に追記する通知
      var notification = "\n\n--------------------------------------\n";
      notification += "※ 添付ファイルは別途外部から送信します。\n";
      notification += "--------------------------------------";
      
      // 複製メールの作成と送信
      Office.context.mailbox.displayNewMessageForm({
          toRecipients: recipientsInfo.to.map(function (recipient) {
              return recipient.emailAddress;
          }),
          ccRecipients: recipientsInfo.cc.map(function (recipient) {
              return recipient.emailAddress;
          }),
          bccRecipients: recipientsInfo.bcc.map(function (recipient) {
              return recipient.emailAddress;
          }),
          subject: subject,
          htmlBody: body + notification,
          attachments: [] // 添付ファイルなし
      });
      
      // URLを含むメールを送信
      sendAttachmentUrlEmail(recipientsInfo, subject, urls);
  });
}

// 通知を表示する関数
function showNotification(title, message) {
  if (Office.context.mailbox.diagnostics.hostName === "Outlook") {
      Office.context.mailbox.item.notificationMessages.addAsync("notification", {
          type: "informationalMessage",
          message: message,
          icon: "icon-16",
          persistent: false
      });
  } else {
      // モバイルなどの場合はアラートを表示
      console.log(title + ": " + message);
  }
}

// 添付ファイルを処理する関数（ダミー実装）
function processAttachments(attachments, callback) {
  console.log("添付ファイルの処理を開始:", attachments.length + "個のファイル");
  
  // 各添付ファイルに対するURLリスト（ダミー）
  var urls = [];
  
  // 各添付ファイルを処理
  attachments.forEach(function (attachment, index) {
      // ダミーのS3 URL
      var dummyUrl = "https://your-s3-bucket.s3.amazonaws.com/files/" + attachment.id + "/" + attachment.name;
      
      // 実際にはここでAttachmentContentを取得してS3にアップロードする
      // getAttachmentContentAsync() APIを使用して添付ファイルの内容を取得
      // AWS SDK for JavaScriptを使ってS3にアップロード
      
      urls.push({
          name: attachment.name,
          url: dummyUrl
      });
      
      // 全ての添付ファイルを処理した後にコールバックを呼び出す
      if (index === attachments.length - 1) {
          // 1秒の遅延を入れて非同期処理をシミュレート
          setTimeout(function () {
              callback(urls);
          }, 1000);
      }
  });
  
  // 添付ファイルがない場合
  if (attachments.length === 0) {
      callback([]);
  }
}

// 添付ファイルのURLを含むメールを送信する関数
function sendAttachmentUrlEmail(recipientsInfo, originalSubject, urls) {
  // URLリストのHTMLを作成
  var urlListHtml = "";
  urls.forEach(function (file) {
      urlListHtml += "<li><a href='" + file.url + "'>" + file.name + "</a></li>";
  });
  
  // メール本文
  var body = "<p>以下は先ほどのメールの添付ファイルへのリンクです：</p>";
  body += "<ul>" + urlListHtml + "</ul>";
  body += "<p>リンクの有効期限は7日間です。</p>";
  
  // 新しいメールを作成して送信
  Office.context.mailbox.displayNewMessageForm({
      toRecipients: recipientsInfo.to.map(function (recipient) {
          return recipient.emailAddress;
      }),
      ccRecipients: recipientsInfo.cc.map(function (recipient) {
          return recipient.emailAddress;
      }),
      subject: "[添付ファイル] " + originalSubject,
      htmlBody: body,
      attachments: [] // 添付ファイルなし
  });
}