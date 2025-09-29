Office.onReady(() => {
  console.log("Office.js is ready");
});

async function zipAndSendEmail() {
  const item = Office.context.mailbox.item;

  const subject = item.subject;
  const bodyResult = await new Promise((resolve, reject) => {
    item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });

  const zip = new JSZip();
  zip.file("subject.txt", subject);
  zip.file("body.txt", bodyResult);

  const blob = await zip.generateAsync({ type: "blob" });

  // Send to backend
  await fetch("https://yourdomain.com/send-email", {
    method: "POST",
    headers: {
      "Content-Type": "application/zip"
    },
    body: blob
  });

  alert("Email zipped and sent to test@test.com");
}
