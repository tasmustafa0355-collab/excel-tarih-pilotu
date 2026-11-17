"use strict";

const origin = "https://bestseller.ist";
const shipmentList = origin + "/#/logistic/shipment-list";
const channel = "bestseller-live-workbook-v1";
const statusMessage = document.querySelector("#connection-status");
const openButton = document.querySelector("#open-bestseller");
const token = crypto.randomUUID();
let site;
let connected = false;
let lastPong = 0;
let waitingSince = 0;

function toExcel(message) {
  Office.context.ui.messageParent(JSON.stringify({channel,...message}),{targetOrigin:location.origin});
}

function loseConnection(message) {
  toExcel({type:connected ? "disconnected" : "waiting",message});
  connected = false;
  statusMessage.textContent = message;
}

Office.onReady(() => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,event => {
    if (event.origin && event.origin !== location.origin) return;
    let data;
    try { data = JSON.parse(event.message); } catch { return; }
    if (data?.channel !== channel || data.type !== "request" || typeof data.id !== "string") return;
    if (!connected || !site || site.closed) {
      toExcel({type:"response",id:data.id,response:{error:"Bestseller bağlantısı açık değil."}});
      return;
    }
    site.postMessage({...data,token},origin);
  },result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      toExcel({type:"disconnected",message:"Excel iletişimi başlatılamadı."});
      statusMessage.textContent = "Excel iletişimi başlatılamadı.";
      return;
    }
    openButton.disabled = false;
    statusMessage.textContent = "Bestseller sekmesini açıp yeni yer işaretini çalıştır.";
  });
});

openButton.addEventListener("click",event => {
  if (!event.isTrusted) return;
  if (site && !site.closed) {
    site.focus();
    return;
  }
  site = window.open(shipmentList,"_blank");
  if (!site) {
    loseConnection("Açılır pencere engellendi. Tarayıcı Bestseller sekmesini açamadı.");
    return;
  }
  waitingSince = Date.now();
  lastPong = 0;
  statusMessage.textContent = "Bestseller’da yer işaretine tıkla ve Oturumu algıla düğmesine bas.";
});

window.addEventListener("message",event => {
  const data = event.data;
  if (event.origin !== origin || event.source !== site || data?.channel !== channel || data.token !== token) return;
  if (data.type === "pong") {
    lastPong = Date.now();
    if (data.ready === true && !connected) {
      connected = true;
      statusMessage.textContent = "Bağlantı hazır. Excel’e dön. Bu pencereyi ve Bestseller sekmesini açık bırak.";
      toExcel({type:"ready"});
    } else if (data.ready !== true) {
      loseConnection("Bestseller’da Oturumu algıla düğmesine bas. Gerekirse bir sevkiyat aç.");
    }
  } else if (data.type === "disconnected") {
    loseConnection("Bestseller sayfası kapandı veya yenilendi. Excel’den yeniden bağlan.");
  } else if (data.type === "response" && typeof data.id === "string") {
    toExcel({type:"response",id:data.id,response:data.response});
  }
});

window.setInterval(() => {
  if (!site) return;
  const elapsed = Date.now() - (lastPong || waitingSince);
  if (site.closed) {
    loseConnection("Bestseller sekmesi kapalı ya da tarayıcı pencereler arası bağlantıyı ayırdı.");
    site = null;
    return;
  }
  if (!connected) {
    if (!lastPong && elapsed > 30000) statusMessage.textContent = "Bestseller’dan bağlantı alınamadı. Yer işaretinin açık sekmede çalıştığını kontrol et.";
    site.postMessage({channel,type:"ping",token},origin);
  }
},800);
