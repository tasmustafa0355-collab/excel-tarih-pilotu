"use strict";

const liveChannel = "bestseller-live-workbook-v1";
let liveConnection;

function openLiveConnection() {
  if (liveConnection) return liveConnection.ready;
  const state = {dialog:null,pending:new Map(),ready:null};
  liveConnection = state;
  state.ready = new Promise((resolve,reject) => {
    state.rejectReady = reject;
    state.wait = window.setTimeout(() => {
      failLiveConnection(state,"Bestseller bağlantısı zaman aşımına uğradı.");
    },180000);
    const url = new URL("bestseller-connection.html?v=2.3.1",location.href);
    Office.context.ui.displayDialogAsync(url.href,{height:35,width:32,displayInIframe:false},result => {
      if (liveConnection !== state) {
        if (result.status === Office.AsyncResultStatus.Succeeded) result.value.close();
        return;
      }
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        failLiveConnection(state,"Excel bağlantı penceresini açamadı (" + (result.error?.code ?? "Office") + ").");
        return;
      }
      state.dialog = result.value;
      if (typeof state.dialog.messageChild !== "function") {
        failLiveConnection(state,"Bu Excel ortamı bağlantı penceresine veri göndermeyi desteklemiyor.");
        return;
      }
      state.dialog.addEventHandler(Office.EventType.DialogMessageReceived,event => {
        if (liveConnection !== state || (event.origin && event.origin !== location.origin)) return;
        let data;
        try { data = JSON.parse(event.message); } catch { return; }
        if (data?.channel !== liveChannel) return;
        if (data.type === "ready") {
          window.clearTimeout(state.wait);
          resolve();
        } else if (data.type === "waiting" && typeof data.message === "string" && !automaticScope) {
          autoStatus.textContent = data.message;
        } else if (data.type === "disconnected") {
          failLiveConnection(state,data.message || "Bestseller bağlantısı kapandı.");
        } else if (data.type === "response") {
          const request = state.pending.get(data.id);
          if (request) {
            window.clearTimeout(request.timer);
            state.pending.delete(data.id);
            request.resolve(data.response);
          }
        }
      });
      state.dialog.addEventHandler(Office.EventType.DialogEventReceived,event => {
        const message = event.error === 12006
          ? "Excel bağlantı penceresi kapandı veya tarayıcı bağlantıyı ayırdı."
          : "Excel bağlantı penceresi kapandı (" + event.error + ").";
        failLiveConnection(state,message);
      });
    });
  });
  return state.ready;
}

function failLiveConnection(state,message) {
  if (liveConnection !== state) return;
  liveConnection = null;
  window.clearTimeout(state.wait);
  for (const request of state.pending.values()) {
    window.clearTimeout(request.timer);
    request.resolve({error:message});
  }
  state.pending.clear();
  state.rejectReady(new Error(message));
  try { state.dialog?.close(); } catch {}
  if (automaticScope) stopAutomaticTransfer(message);
}

function closeLiveConnection() {
  if (liveConnection) failLiveConnection(liveConnection,"Otomatik aktarım kapalı.");
}

function askLiveBestseller(batch,action) {
  const state = liveConnection;
  if (!state?.dialog) return Promise.resolve({error:"Bestseller bağlantısını aç."});
  const id = crypto.randomUUID();
  return new Promise(resolve => {
    const complete = response => {
      const request = state.pending.get(id);
      if (request) window.clearTimeout(request.timer);
      state.pending.delete(id);
      resolve(response);
    };
    const timer = window.setTimeout(() => {
      complete({error:"Bestseller yanıtı gecikti. Sonucu tarihçede kontrol et."});
    },action === "sync" ? 240000 : 10000);
    state.pending.set(id,{timer,resolve});
    try {
      state.dialog.messageChild(JSON.stringify({channel:liveChannel,type:"request",id,batch,action}),
        {targetOrigin:location.origin});
    } catch {
      complete({error:"Excel bağlantı penceresine veri gönderilemedi."});
    }
  });
}
