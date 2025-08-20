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
  console.log("=== openComposeWithData 시작 ===");
  
  const to = parseRecipients(q<HTMLInputElement>("to")?.value ?? "");
  const cc = parseRecipients(q<HTMLInputElement>("cc")?.value ?? "");
  const bcc = parseRecipients(q<HTMLInputElement>("bcc")?.value ?? "");
  const subject = q<HTMLInputElement>("subject")?.value ?? "";
  const htmlBody = q<HTMLTextAreaElement>("body")?.value ?? "";
  const attachmentsText = q<HTMLTextAreaElement>("attachments")?.value ?? "";

  console.log("입력 데이터:", { to, cc, bcc, subject, htmlBody, attachmentsText });

  const attachList = await buildAttachments(attachmentsText);
  const urlAttachments = attachList.filter(a => a.type === "url").map(a => ({
    type: "file" as const,
    name: a.name,
    url: a.value,
    isInline: !!a.isInline
  }));
  const base64Attachments = attachList.filter(a => a.type === "base64");

  console.log("첨부파일:", { urlAttachments, base64Attachments });

  const mbox = (Office as any)?.context?.mailbox;
  console.log("mailbox 객체:", mbox);

  // 1) displayNewMessageForm가 있으면 "새 메일 창"으로
  if (mbox && typeof mbox.displayNewMessageForm === "function") {
    console.log("displayNewMessageForm 실행 중...");
    try {
      mbox.displayNewMessageForm({
        toRecipients: to,
        ccRecipients: cc,
        bccRecipients: bcc,
        subject,
        htmlBody,
        attachments: urlAttachments
      });
      console.log("displayNewMessageForm 성공");

      if (base64Attachments.length === 0) {
        setStatus("완료: 새 메일 창이 열렸습니다.", "ok");
        return;
      }
      // 새 창이 떠야 base64 첨부 가능하므로 잠시 대기
      console.log("Base64 첨부 처리 시작...");
      await new Promise(r => setTimeout(r, 700));
      try {
        for (const a of base64Attachments) {
          console.log(`Base64 첨부 중: ${a.name}`);
          const { base64 } = splitDataUri(a.value);
          await new Promise<void>((resolve, reject) => {
            Office.context.mailbox.item.addFileAttachmentFromBase64Async(
              base64,
              a.name,
              { isInline: !!a.isInline },
              res => {
                console.log(`첨부 결과 (${a.name}):`, res);
                res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
              }
            );
          });
        }
        setStatus("완료: 새 메일 창 + 첨부 추가 완료.", "ok");
      } catch (e: any) {
        console.error("Base64 첨부 실패:", e);
        setStatus(`Base64 첨부 실패: ${e?.message || e}`, "err");
      }
      return;
    } catch (e: any) {
      console.error("displayNewMessageForm 실패:", e);
      setStatus(`새 메일 창 열기 실패: ${e?.message || e}`, "err");
      return;
    }
  }

  // 2) 새 창 API가 없으면: "이미 열려있는 작성창"에 값 세팅 (Compose 전용)
  const item = (Office as any)?.context?.mailbox?.item;
  console.log("item 객체:", item);
  console.log("item.itemType:", item?.itemType);
  console.log("item.saveAsync:", typeof item?.saveAsync);
  
  const isCompose = !!(item && item.itemType && item.saveAsync); // compose에서만 saveAsync 있음
  console.log("isCompose:", isCompose);
  
  if (isCompose) {
    console.log("현재 작성창에 값 세팅 시작...");
    // 받는사람/제목/본문 채우기
    const setAsync = (field: any, value: any) => new Promise<void>((resolve, reject) => {
      console.log("setAsync 호출:", { field, value });
      field.setAsync(value, (res: any) => {
        console.log("setAsync 결과:", res);
        res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
      });
    });

    try {
      if (to.length && item.to?.setAsync) {
        console.log("받는사람 설정 중...");
        const toRecipients = to.map(email => ({
          displayName: email,
          emailAddress: email
        }));
        console.log("받는사람 형식:", toRecipients);
        await setAsync(item.to, toRecipients);
      }
      if (cc.length && item.cc?.setAsync) {
        console.log("참조 설정 중...");
        const ccRecipients = cc.map(email => ({
          displayName: email,
          emailAddress: email
        }));
        console.log("참조 형식:", ccRecipients);
        await setAsync(item.cc, ccRecipients);
      }
      if (bcc.length && item.bcc?.setAsync) {
        console.log("숨은참조 설정 중...");
        const bccRecipients = bcc.map(email => ({
          displayName: email,
          emailAddress: email
        }));
        console.log("숨은참조 형식:", bccRecipients);
        await setAsync(item.bcc, bccRecipients);
      }
      if (subject && item.subject?.setAsync) {
        console.log("제목 설정 중...");
        await setAsync(item.subject, subject);
      }
      if (htmlBody && item.body?.setAsync) {
        console.log("본문 설정 중...");
        await new Promise<void>((resolve, reject) => {
          item.body.setAsync(htmlBody, { coercionType: Office.CoercionType.Html }, (res: any) => {
            console.log("본문 설정 결과:", res);
            res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
          });
        });
      }

      // URL 첨부는 addFileAttachmentAsync
      for (const a of urlAttachments) {
        console.log(`URL 첨부 중: ${a.name}`);
        await new Promise<void>((resolve, reject) => {
          item.addFileAttachmentAsync(a.url, a.name, { isInline: a.isInline }, (res: any) => {
            console.log(`URL 첨부 결과 (${a.name}):`, res);
            res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
          });
        });
      }

      // Base64 첨부
      for (const a of base64Attachments) {
        console.log(`Base64 첨부 중: ${a.name}`);
        const { base64 } = splitDataUri(a.value);
        await new Promise<void>((resolve, reject) => {
          item.addFileAttachmentFromBase64Async(base64, a.name, { isInline: !!a.isInline }, (res: any) => {
            console.log(`Base64 첨부 결과 (${a.name}):`, res);
            res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
          });
        });
      }

      setStatus("완료: 현재 작성창에 값/첨부 적용.", "ok");
      return;
    } catch (e: any) {
      console.error("작성창 세팅 중 오류:", e);
      setStatus(`작성창 세팅 실패: ${e?.message || e}`, "err");
      return;
    }
  }

  // 3) 둘 다 아니면 사용 방법 안내
  setStatus("이 버튼은 Outlook 작업창에서 동작합니다. 새 메일 창을 먼저 열고 작업창을 띄운 뒤 다시 눌러주세요.", "err");
}


// CRM 데이터 불러오기 함수
function loadCrmData() {
  try {
    setStatus("CRM 데이터 확인 중...", "muted");
    
    // URL 파라미터 확인
    const urlParams = new URLSearchParams(window.location.search);
    const source = urlParams.get('source');
    
    console.log("URL source 파라미터:", source);
    
    // localStorage에서 CRM 데이터 확인
    const crmDataStr = localStorage.getItem('outlookAddinData');
    console.log("localStorage CRM 데이터:", crmDataStr);
    
    if (crmDataStr) {
      try {
        const crmData = JSON.parse(crmDataStr);
        console.log("CRM 데이터 파싱 성공:", crmData);
        
        // 폼에 데이터 자동 입력
        if (crmData.to) q<HTMLInputElement>("to")!.value = crmData.to;
        if (crmData.cc) q<HTMLInputElement>("cc")!.value = crmData.cc;
        if (crmData.bcc) q<HTMLInputElement>("bcc")!.value = crmData.bcc;
        if (crmData.subject) q<HTMLInputElement>("subject")!.value = crmData.subject;
        if (crmData.body) q<HTMLTextAreaElement>("body")!.value = crmData.body;
        if (crmData.attachments && crmData.attachments.length > 0) {
          q<HTMLTextAreaElement>("attachments")!.value = crmData.attachments.join('\n');
        }
        
        setStatus("CRM 데이터 로드 완료!", "ok");
        
        // 데이터 사용 후 삭제 (1회용)
        localStorage.removeItem('outlookAddinData');
        
        return true;
      } catch (parseError) {
        console.error("CRM 데이터 파싱 실패:", parseError);
        setStatus("CRM 데이터 형식 오류", "err");
        return false;
      }
    } else if (source === 'crm') {
      setStatus("CRM에서 호출되었지만 데이터가 없습니다.", "err");
      return false;
    } else {
      setStatus("CRM 데이터가 없습니다.", "muted");
      return false;
    }
  } catch (error: any) {
    console.error("CRM 데이터 로드 오류:", error);
    setStatus(`CRM 데이터 로드 실패: ${error?.message || error}`, "err");
    return false;
  }
}

// 폼 초기화 함수
function clearForm() {
  q<HTMLInputElement>("to")!.value = "";
  q<HTMLInputElement>("cc")!.value = "";
  q<HTMLInputElement>("bcc")!.value = "";
  q<HTMLInputElement>("subject")!.value = "";
  q<HTMLTextAreaElement>("body")!.value = "";
  q<HTMLTextAreaElement>("attachments")!.value = "";
  setStatus("폼이 초기화되었습니다.", "muted");
}

// 자동 테스트 실행 함수
async function runAutoTest() {
  try {
    setStatus("자동 테스트 실행 중...", "muted");
    console.log("=== 자동 테스트 시작 ===");
    
    // 1단계: CRM 데이터 로드 시도
    console.log("1단계: CRM 데이터 로드 시도");
    const dataLoaded = loadCrmData();
    
    if (!dataLoaded) {
      // CRM 데이터가 없으면 테스트 데이터 직접 입력
      console.log("CRM 데이터 없음 - 테스트 데이터 직접 입력");
      q<HTMLInputElement>("to")!.value = "test1@aaa.kr";
      q<HTMLInputElement>("cc")!.value = "test2@aaa.kr";
      q<HTMLInputElement>("bcc")!.value = "test3@aaa.kr";
      q<HTMLInputElement>("subject")!.value = "test4";
      q<HTMLTextAreaElement>("body")!.value = "test5";
      q<HTMLTextAreaElement>("attachments")!.value = "http://10.1.223.25/download/quotes/pdf/quotation-20250814.pdf";
      setStatus("테스트 데이터 입력 완료", "ok");
    }
    
    // 2초 대기 후 이메일 작성 실행
    setTimeout(async () => {
      console.log("2단계: 이메일 작성 실행");
      try {
        await openComposeWithData();
        setStatus("자동 테스트 완료! 이메일이 작성되었습니다.", "ok");
      } catch (error: any) {
        console.error("이메일 작성 실패:", error);
        setStatus(`자동 테스트 실패: ${error?.message || error}`, "err");
      }
    }, 2000);
    
  } catch (error: any) {
    console.error("자동 테스트 오류:", error);
    setStatus(`자동 테스트 오류: ${error?.message || error}`, "err");
  }
}

// 자동으로 테스트 데이터 채우기 함수
async function autoFillTestData() {
  try {
    console.log("=== 자동 테스트 데이터 채우기 시작 ===");
    
    // 테스트 데이터 입력
    q<HTMLInputElement>("to")!.value = "test1@aaa.kr";
    q<HTMLInputElement>("cc")!.value = "test2@aaa.kr";
    q<HTMLInputElement>("bcc")!.value = "test3@aaa.kr";
    q<HTMLInputElement>("subject")!.value = "test4";
    q<HTMLTextAreaElement>("body")!.value = "test5";
    q<HTMLTextAreaElement>("attachments")!.value = "http://10.1.223.25/download/quotes/pdf/quotation-20250814.pdf";
    
    setStatus("테스트 데이터 자동 입력 완료", "ok");
    
    // 2초 후 자동으로 이메일 작성 실행
    setTimeout(async () => {
      console.log("=== 자동 이메일 작성 실행 ===");
      try {
        await openComposeWithData();
        setStatus("자동 이메일 작성 완료! 핀이 고정되어 다음부터는 자동으로 실행됩니다.", "ok");
      } catch (error: any) {
        console.error("자동 이메일 작성 실패:", error);
        setStatus(`자동 이메일 작성 실패: ${error?.message || error}`, "err");
      }
    }, 2000);
    
  } catch (error: any) {
    console.error("자동 테스트 데이터 채우기 실패:", error);
    setStatus(`자동 테스트 데이터 실패: ${error?.message || error}`, "err");
  }
}

// ✅ Office가 준비되면 그때 버튼에 이벤트 바인딩
Office.onReady(() => {
  // 디버깅 정보 출력
  console.log("=== Office.js 디버깅 정보 ===");
  console.log("Office.context:", Office.context);
  console.log("Office.context.mailbox:", Office.context?.mailbox);
  console.log("Office.context.mailbox.item:", Office.context?.mailbox?.item);
  console.log("item.itemType:", Office.context?.mailbox?.item?.itemType);
  console.log("item.itemClass:", Office.context?.mailbox?.item?.itemClass);
  console.log("displayNewMessageForm 지원:", typeof Office.context?.mailbox?.displayNewMessageForm);
  
  // 메인 버튼: 이메일 작성
  const btn = q<HTMLButtonElement>("openComposeBtn");
  if (btn) {
    btn.addEventListener("click", async () => {
      try {
        setStatus("처리 중...", "muted");
        console.log("=== 버튼 클릭 시작 ===");
        await openComposeWithData();
      } catch (e: any) {
        console.error("버튼 클릭 오류:", e);
        setStatus(`오류: ${e?.message || e}`, "err");
      }
    });
  }

  // CRM 데이터 불러오기 버튼
  const loadBtn = q<HTMLButtonElement>("loadCrmDataBtn");
  if (loadBtn) {
    loadBtn.addEventListener("click", () => {
      loadCrmData();
    });
  }

  // 폼 초기화 버튼
  const clearBtn = q<HTMLButtonElement>("clearFormBtn");
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      clearForm();
    });
  }

  // 자동 테스트 실행 버튼
  const autoTestBtn = q<HTMLButtonElement>("autoTestBtn");
  if (autoTestBtn) {
    autoTestBtn.addEventListener("click", async () => {
      await runAutoTest();
    });
  }

  // 페이지 로드 시 자동으로 CRM 데이터 확인
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('source') === 'crm') {
    console.log("CRM에서 호출됨 - 자동으로 데이터 로드 시도");
    setTimeout(() => {
      const loaded = loadCrmData();
      if (loaded) {
        setStatus("CRM 데이터가 자동으로 로드되었습니다. '아웃룩 새 메일 띄우기' 버튼을 클릭하세요.", "ok");
      }
    }, 500);
  }

  // 디버깅용 상태 메시지
  const itemType = Office.context?.mailbox?.item?.itemType || "unknown";
  setStatus(`준비 완료: Outlook 작업창 연결됨 (${itemType})`, "ok");
  
  // 자동으로 테스트 데이터 채우기 실행 (핀 고정 시 자동 실행)
  setTimeout(() => {
    autoFillTestData();
  }, 1000);
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
