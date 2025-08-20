/* global Office */

type AttachmentInput = { type: "url" | "base64", name: string, value: string, isInline?: boolean };

function q<T extends HTMLElement>(id: string) {
  return document.getElementById(id) as T | null;
}

function parseRecipients(raw: string): string[] {
  return (raw || "")
    .split(/[,\n;]/)
    .map(s => s.trim())
    .filter(Boolean);
}

function guessFileNameFromUrl(url: string): string {
  try {
    const u = new URL(url);
    const last = u.pathname.split("/").filter(Boolean).pop() || "attachment";
    return decodeURIComponent(last);
  } catch { return "attachment"; }
}

function isDataUriBase64(s: string) {
  return /^data:([-\w.]+\/[-\w.+]+)?;base64,/i.test((s || "").trim());
}
function isHttpUrl(s: string) {
  return /^https?:\/\//i.test((s || "").trim());
}
function splitDataUri(dataUri: string) {
  const i = dataUri.indexOf(";base64,");
  const mime = dataUri.slice(5, i);
  const base64 = dataUri.slice(i + 8);
  return { mime, base64 };
}

async function buildAttachments(textareaValue: string): Promise<AttachmentInput[]> {
  const lines = (textareaValue || "")
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(Boolean);

  const out: AttachmentInput[] = [];
  for (const line of lines) {
    if (isDataUriBase64(line)) {
      const { mime } = splitDataUri(line);
      const name = `attachment.${(mime.split("/")[1] || "bin").replace(/[^\w.+-]/g, "")}`;
      out.push({ type: "base64", name, value: line, isInline: false });
    } else if (isHttpUrl(line)) {
      out.push({ type: "url", name: guessFileNameFromUrl(line), value: line, isInline: false });
    }
  }
  return out;
}

function setStatus(msg: string, cls: "ok" | "err" | "muted" = "muted") {
  const el = q<HTMLDivElement>("status");
  if (el) { el.className = cls; el.textContent = msg; }
}

async function openComposeWithData() {
  const to = parseRecipients(q<HTMLInputElement>("to")?.value ?? "");
  const cc = parseRecipients(q<HTMLInputElement>("cc")?.value ?? "");
  const bcc = parseRecipients(q<HTMLInputElement>("bcc")?.value ?? "");
  const subject = q<HTMLInputElement>("subject")?.value ?? "";
  const htmlBody = q<HTMLTextAreaElement>("body")?.value ?? "";
  const attachmentsText = q<HTMLTextAreaElement>("attachments")?.value ?? "";

  const attachList = await buildAttachments(attachmentsText);
  const urlAttachments = attachList.filter(a => a.type === "url").map(a => ({
    type: "file" as const,
    name: a.name,
    url: a.value,
    isInline: !!a.isInline
  }));
  const base64Attachments = attachList.filter(a => a.type === "base64");

  const mbox = (Office as any)?.context?.mailbox;

  // 1) displayNewMessageForm가 있으면 "새 메일 창"으로
  if (mbox && typeof mbox.displayNewMessageForm === "function") {
    mbox.displayNewMessageForm({
      toRecipients: to,
      ccRecipients: cc,
      bccRecipients: bcc,
      subject,
      htmlBody,
      attachments: urlAttachments
    });

    if (base64Attachments.length === 0) {
      setStatus("완료: 새 메일 창이 열렸습니다.", "ok");
      return;
    }
    // 새 창이 떠야 base64 첨부 가능하므로 잠시 대기
    await new Promise(r => setTimeout(r, 700));
    try {
      for (const a of base64Attachments) {
        const { base64 } = splitDataUri(a.value);
        await new Promise<void>((resolve, reject) => {
          Office.context.mailbox.item.addFileAttachmentFromBase64Async(
            base64,
            a.name,
            { isInline: !!a.isInline },
            res => res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error)
          );
        });
      }
      setStatus("완료: 새 메일 창 + 첨부 추가 완료.", "ok");
    } catch (e: any) {
      setStatus(`Base64 첨부 실패: ${e?.message || e}`, "err");
    }
    return;
  }

  // 2) 새 창 API가 없으면: "이미 열려있는 작성창"에 값 세팅 (Compose 전용)
  const item = (Office as any)?.context?.mailbox?.item;
  const isCompose = !!(item && item.itemType && item.saveAsync); // compose에서만 saveAsync 있음
  if (isCompose) {
    // 받는사람/제목/본문 채우기
    const setAsync = (field: any, value: any) => new Promise<void>((resolve, reject) => {
      field.setAsync(value, (res: any) => res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error));
    });

    try {
      if (to.length && item.to?.setAsync) await setAsync(item.to, to.map(a => ({ emailAddress: a })));
      if (cc.length && item.cc?.setAsync) await setAsync(item.cc, cc.map(a => ({ emailAddress: a })));
      if (bcc.length && item.bcc?.setAsync) await setAsync(item.bcc, bcc.map(a => ({ emailAddress: a })));
      if (subject && item.subject?.setAsync) await setAsync(item.subject, subject);
      if (htmlBody && item.body?.setAsync) {
        await new Promise<void>((resolve, reject) => {
          item.body.setAsync(htmlBody, { coercionType: Office.CoercionType.Html }, (res: any) =>
            res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error));
        });
      }

      // URL 첨부는 addFileAttachmentAsync
      for (const a of urlAttachments) {
        await new Promise<void>((resolve, reject) => {
          item.addFileAttachmentAsync(a.url, a.name, { isInline: a.isInline }, (res: any) =>
            res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error));
        });
      }

      // Base64 첨부
      for (const a of base64Attachments) {
        const { base64 } = splitDataUri(a.value);
        await new Promise<void>((resolve, reject) => {
          item.addFileAttachmentFromBase64Async(base64, a.name, { isInline: !!a.isInline }, (res: any) =>
            res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error));
        });
      }

      setStatus("완료: 현재 작성창에 값/첨부 적용.", "ok");
      return;
    } catch (e: any) {
      setStatus(`작성창 세팅 실패: ${e?.message || e}`, "err");
      return;
    }
  }

  // 3) 둘 다 아니면 사용 방법 안내
  setStatus("이 버튼은 Outlook 작업창에서 동작합니다. 새 메일 창을 먼저 열고 작업창을 띄운 뒤 다시 눌러주세요.", "err");
}


// ✅ Office가 준비되면 그때 버튼에 이벤트 바인딩
Office.onReady(() => {
  // 일부 구형 호스트에서도 불리는 초기화 훅
  (Office as any).initialize = (reason?: any) => {};

  // 버튼 바인딩
  const btn = q<HTMLButtonElement>("openComposeBtn");
  if (btn) {
    btn.addEventListener("click", async () => {
      try {
        await openComposeWithData();
      } catch (e: any) {
        setStatus(`오류: ${e?.message || e}`, "err");
      }
    });
  }

  // 디버깅용 상태 메시지
  setStatus("준비 완료: Outlook 작업창 연결됨.", "ok");
});

// ✅ Outlook 밖(브라우저에서 직접 URL 열었을 때) 가드
window.addEventListener("load", () => {
  if (!(window as any).Office || !(window as any).Office.context) {
    const btn = q<HTMLButtonElement>("openComposeBtn");
    if (btn) btn.disabled = true;
    setStatus("이 페이지는 Outlook Add-in 작업창 안에서만 동작합니다.", "err");
  }
});


// taskpane.ts
let dlg: Office.Dialog;

function openCrmDialog() {
  Office.context.ui.displayDialogAsync(
    "https://your-crm.example.com/compose-helper", // CRM 버튼이 있는 페이지
    { height: 60, width: 40, displayInIframe: true },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        dlg = res.value;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          if ('message' in arg) {
            const payload = JSON.parse(arg.message);
            // payload: { to:string[], cc:string[], bcc:string[], subject:string, htmlBody:string, attachments:string[] }
            // → 여기서 기존 openComposeWithData(payload) 호출해서 작성창 채우기
            openComposeWithData();
          }
          dlg.close();
        });
      }
    }
  );
}
