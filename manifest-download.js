"use strict";

document.querySelector("#download").addEventListener("click", () => {
  const panel = new URL("./index.html", location.href);
  if (panel.protocol !== "https:") {
    document.querySelector("#status").textContent = "Dosyaları HTTPS sunucusunda aç.";
    return;
  }
  const url = panel.href.replaceAll("&", "&amp;").replaceAll('"', "&quot;");
  const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="TaskPaneApp">
  <Id>9d8db4b6-fd9c-4aad-b33a-c29873647647</Id>
  <Version>1.0.1.0</Version>
  <ProviderName>Bestseller Pilot</ProviderName>
  <DefaultLocale>tr-TR</DefaultLocale>
  <DisplayName DefaultValue="Bestseller Tarih Pilotu"/>
  <Description DefaultValue="Sheet2 ve Sheet3 tarih girislerini Excel panelinde gosterir."/>
  <SupportUrl DefaultValue="${url}"/>
  <Hosts><Host Name="Workbook"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.9"><Set Name="ExcelApi"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="${url}"/></DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>`;
  const download = document.createElement("a");
  const file = new Blob([manifest], { type: "application/xml" });
  const fileUrl = URL.createObjectURL(file);
  download.href = fileUrl;
  download.download = "manifest-bestseller-tarih.xml";
  download.click();
  window.setTimeout(() => URL.revokeObjectURL(fileUrl), 1000);
});
