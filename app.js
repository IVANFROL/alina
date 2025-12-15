let tnvedMap = new Map();
let tnvedLoaded = false;
let items = [];
let logoImageBuffer = null; // Кэш для логотипа

const tnvedStatusEl = document.getElementById("tnved-status");
const tnvedInfoBody = document.getElementById("tnvedInfoBody");
const tnvedInfoCard = document.getElementById("tnvedInfoCard");
const tnvedFileInput = document.getElementById("tnvedFileInput");
const tnvedFileName = document.getElementById("tnved-file-name");
const validationWarning = document.getElementById("validationWarning");
const rateHintEl = document.getElementById("rateHint");
const invoiceTotalHintEl = document.getElementById("invoiceTotalHint");

const itemsTable = document.getElementById("itemsTable");
const itemsTableBody = document.getElementById("itemsTableBody");
const itemsTotalInvoiceRub = document.getElementById("itemsTotalInvoiceRub");
const itemsTotalDomesticDeliveryRub = document.getElementById("itemsTotalDomesticDeliveryRub");
const itemsTotalIntlDeliveryRub = document.getElementById("itemsTotalIntlDeliveryRub");
const itemsTotalDutyRub = document.getElementById("itemsTotalDutyRub");
const itemsTotalVatRub = document.getElementById("itemsTotalVatRub");
const itemsTotalExciseRub = document.getElementById("itemsTotalExciseRub");
const itemsTotalCustomsRub = document.getElementById("itemsTotalCustomsRub");
const itemsTotalCertRub = document.getElementById("itemsTotalCertRub");
const itemsTotalOtherRub = document.getElementById("itemsTotalOtherRub");
const itemsTotalCostsRub = document.getElementById("itemsTotalCostsRub");

// Добавляем колонку "Действия" в заголовок, если её ещё нет
const itemsHeaderRow = itemsTable.querySelector("thead tr");
if (itemsHeaderRow && !itemsHeaderRow.querySelector(".actions-col")) {
  const th = document.createElement("th");
  th.textContent = "Действия";
  th.className = "actions-col";
  itemsHeaderRow.appendChild(th);
}

// Делегирование события на удаление товаров
itemsTableBody.addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-index]");
  if (!btn) return;
  const idx = parseInt(btn.getAttribute("data-index"), 10);
  if (!Number.isNaN(idx) && idx >= 0 && idx < items.length) {
    items.splice(idx, 1);
    renderItemsTable();
  }
});

function setTnvedStatus(ok, msg) {
  tnvedStatusEl.classList.remove("ok", "error");
  if (ok === true) tnvedStatusEl.classList.add("ok");
  if (ok === false) tnvedStatusEl.classList.add("error");
  const spanMsg = tnvedStatusEl.querySelector("span:nth-child(2)");
  if (spanMsg) spanMsg.textContent = msg;
}

function normalizeCode(raw) {
  if (!raw) return null;
  let code = String(raw).trim();
  code = code.replace(/\D/g, "");
  if (!code) return null;
  if (code.length > 10) code = code.slice(0, 10);
  if (code.length < 10) code = code.padStart(10, "0");
  return code;
}
function parseTariff(raw) {
  if (raw === undefined || raw === null || raw === "") return 0;
  let s = String(raw).replace(",", ".").replace("%", "").trim();
  let n = Number(s);
  return Number.isNaN(n) ? 0 : n;
}
function getFirst(row, keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== "") return row[k];
  }
  return null;
}

tnvedFileInput.addEventListener("change", (event) => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  tnvedFileName.textContent = file.name;
  tnvedMap.clear();
  tnvedLoaded = false;
  setTnvedStatus(null, "Чтение файла ТН ВЭД…");
  const name = file.name.toLowerCase();
  const isJson = name.endsWith(".json");

  if (isJson) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        let count = 0;
        for (const item of data) {
          const rawCode = item.code || item["Код"] || item["Код ТН ВЭД"];
          if (!rawCode) continue;
          const code = normalizeCode(rawCode);
          if (!code) continue;
          const description = item.description || item["Описание"] || item["Наименование"] || "";
          const tariff = parseTariff(item.tariff ?? item["Пошлина"] ?? item["Ставка пошлины (%)"]);
          const vat = parseTariff(item.vat ?? item["НДС"] ?? item["Ставка НДС (%)"]);
          tnvedMap.set(code, { code, description, tariff, vat });
          count++;
        }
        if (!count) {
          tnvedLoaded = false;
          setTnvedStatus(false, "JSON-файл прочитан, но кодов не найдено.");
          return;
        }
        tnvedLoaded = true;
        setTnvedStatus(true, "Справочник ТН ВЭД (JSON) загружен (" + tnvedMap.size + " кодов).");
      } catch (err) {
        console.error(err);
        tnvedLoaded = false;
        tnvedMap.clear();
        setTnvedStatus(false, "Ошибка при чтении JSON: " + err.message);
        tnvedFileName.textContent = "";
      }
    };
    reader.onerror = () => {
      tnvedLoaded = false;
      tnvedMap.clear();
      setTnvedStatus(false, "Не удалось прочитать JSON-файл.");
      tnvedFileName.textContent = "";
    };
    reader.readAsText(file, "utf-8");
  } else {
    if (typeof XLSX === "undefined") {
      setTnvedStatus(false, "Библиотека XLSX не загружена.");
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        if (!sheetName) throw new Error("В Excel нет листов.");
        const sheet = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!rows || !rows.length) throw new Error("Лист Excel пустой.");
        let count = 0;
        for (const row of rows) {
          const rawCode = getFirst(row, [
            "Код",
            "код",
            "Code",
            "code",
            "ТНВЭД",
            "ТН ВЭД",
            "Код ТН ВЭД",
            "Код ТНВЭД",
          ]);
          const code = normalizeCode(rawCode);
          if (!code) continue;
          const description =
            getFirst(row, [
              "Наименование",
              "Описание",
              "Описание кода ТН ВЭД",
              "Description",
              "description",
            ]) || "";
          const tariff = parseTariff(
            getFirst(row, ["Тариф", "Пошлина", "tariff", "Ставка пошлины (%)"])
          );
          const vat = parseTariff(getFirst(row, ["НДС", "Ставка НДС (%)"]));
          tnvedMap.set(code, { code, description, tariff, vat });
          count++;
        }
        if (!count) throw new Error("Не найдено ни одного кода. Проверьте заголовки столбцов.");
        tnvedLoaded = true;
        setTnvedStatus(true, "Справочник ТН ВЭД (Excel) загружен (" + tnvedMap.size + " кодов).");
      } catch (err) {
        console.error(err);
        tnvedLoaded = false;
        tnvedMap.clear();
        setTnvedStatus(false, "Ошибка при чтении Excel: " + err.message);
        tnvedFileName.textContent = "";
      }
    };
    reader.onerror = () => {
      tnvedLoaded = false;
      tnvedMap.clear();
      setTnvedStatus(false, "Не удалось прочитать Excel-файл.");
      tnvedFileName.textContent = "";
    };
    reader.readAsArrayBuffer(file);
  }
});

function findTnvedRecord(rawCode) {
  if (!tnvedLoaded || !rawCode) return null;
  const code = normalizeCode(rawCode);
  if (!code) return null;
  return tnvedMap.get(code) || null;
}
function updateTnvedInfo(record) {
  if (!tnvedLoaded) {
    tnvedInfoBody.innerHTML =
      "<span style='color:#b91c1c;'>Справочник ТН ВЭД не загружен.</span> Сначала выберите JSON или Excel-файл.";
    tnvedInfoCard.style.borderColor = "#fecaca";
    tnvedInfoCard.style.background = "#fef2f2";
    return;
  }
  if (!record) {
    tnvedInfoBody.innerHTML =
      "<span style='color:#b45309;'>Код не найден в вашем справочнике.</span> " +
      "Код может быть валидным в официальном классификаторе, но отсутствовать в загруженном файле. " +
      "В этом случае введите ставки пошлины и НДС вручную или добавьте код в файл.";
    tnvedInfoCard.style.borderColor = "#fbbf24";
    tnvedInfoCard.style.background = "#fffbeb";
    return;
  }
  tnvedInfoCard.style.borderColor = "#bbf7d0";
  tnvedInfoCard.style.background = "#f0fdf4";
  tnvedInfoBody.innerHTML =
    "<div><strong>Код:</strong> " +
    record.code +
    "</div>" +
    "<div><strong>Описание:</strong> " +
    (record.description || "—") +
    "</div>" +
    "<div style='margin-top:6px;'><strong>Пошлина:</strong> " +
    record.tariff +
    " % " +
    (record.vat ? "· <strong>НДС:</strong> " + record.vat + " %" : "(ставку НДС укажите вручную)") +
    "</div><div class='hint' style='margin-top:4px;'>Ставка пошлины подставлена в поле «Пошлина, %».</div>";
}

function readNumber(id) {
  const el = document.getElementById(id);
  if (!el) return 0;
  const v = String(el.value || "").replace(",", ".");
  const n = parseFloat(v);
  return Number.isNaN(n) ? 0 : n;
}
function formatMoney(val) {
  return Number(val || 0).toLocaleString("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function updateInvoiceTotalInfo() {
  const unitPrice = readNumber("invoiceValue");
  const qty = readNumber("quantity");
  const currency = document.getElementById("currency").value;
  const rate = readNumber("exchangeRate");
  if (!unitPrice || !qty) {
    invoiceTotalHintEl.textContent = "Общая сумма сделки: —";
    return;
  }
  const totalCur = unitPrice * qty;
  let text =
    "Общая сумма сделки: " +
    totalCur.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) +
    " " +
    currency;
  if (currency === "RUB") {
    invoiceTotalHintEl.textContent = text;
  } else if (rate > 0) {
    const totalRub = totalCur * rate;
    text +=
      " (" +
      formatMoney(totalRub) +
      " ₽ по курсу " +
      rate.toFixed(4).replace(".", ",") +
      ")";
    invoiceTotalHintEl.textContent = text;
  } else {
    text += " (для пересчёта в ₽ укажите курс)";
    invoiceTotalHintEl.textContent = text;
  }
}

// курсы ЦБ РФ (+3%)
async function fetchCbrRateForCurrency(currency) {
  try {
    if (!["USD", "EUR", "CNY"].includes(currency)) {
      rateHintEl.textContent = "Для валюты " + currency + " курс задаётся вручную.";
      return;
    }
    rateHintEl.textContent = "Получаем курс " + currency + "→RUB по данным ЦБ РФ…";
    const resp = await fetch("https://www.cbr-xml-daily.ru/daily_json.js");
    if (!resp.ok) throw new Error("HTTP " + resp.status);
    const data = await resp.json();
    if (!data.Valute || !data.Valute[currency]) throw new Error("Нет курса " + currency);
    const rate = Number(data.Valute[currency].Value);
    if (!rate || !isFinite(rate)) throw new Error("Курс некорректен");
    const adjusted = rate * 1.03;
    const rateInput = document.getElementById("exchangeRate");
    rateInput.value = adjusted.toFixed(4);
    rateHintEl.textContent =
      "Курс " +
      currency +
      "→RUB по ЦБ РФ: " +
      rate.toFixed(4).replace(".", ",") +
      ". В расчётах используется с надбавкой 3%: " +
      adjusted.toFixed(4).replace(".", ",") +
      ".";
    updateInvoiceTotalInfo();
  } catch (e) {
    console.error(e);
    rateHintEl.textContent =
      "Не удалось получить курс по ЦБ РФ. Введите курс " + currency + "→RUB вручную.";
  }
}

function calculateIntlDeliveryRub() {
  const deliveryType = document.getElementById("deliveryType").value;
  const currency = document.getElementById("currency").value;
  const rate = readNumber("exchangeRate");
  const volume = readNumber("volume");
  const weight = readNumber("weight");
  let intlRub = 0;
  if (deliveryType === "truck") {
    const basePrice = readNumber("tariffTruckPrice");
    const baseVolume = readNumber("tariffTruckVolume") || 1;
    if (volume > 0 && basePrice > 0) intlRub = (basePrice * volume) / baseVolume;
  } else if (deliveryType === "rail") {
    const basePrice = readNumber("tariffRailPrice");
    const baseVolume = readNumber("tariffRailVolume") || 1;
    if (volume > 0 && basePrice > 0) intlRub = (basePrice * volume) / baseVolume;
  } else if (deliveryType === "sea") {
    const basePrice = readNumber("tariffSeaPrice");
    const baseVolume = readNumber("tariffSeaVolume") || 1;
    if (volume > 0 && basePrice > 0) intlRub = (basePrice * volume) / baseVolume;
  } else if (deliveryType === "air") {
    const tPerKg = readNumber("tariffAirPerKg");
    if (weight > 0 && tPerKg > 0) {
      const costCur = weight * tPerKg;
      if (currency === "RUB") intlRub = costCur;
      else if (rate > 0) intlRub = costCur * rate;
    }
  }
  const intlInput = document.getElementById("intlDeliveryRub");
  intlInput.value = intlRub > 0 ? intlRub.toFixed(2) : "";
  return intlRub;
}

function renderItemsTable() {
  if (!items || !items.length) {
    itemsTable.style.display = "none";
    itemsTableBody.innerHTML = "";
    [
      itemsTotalInvoiceRub,
      itemsTotalDomesticDeliveryRub,
      itemsTotalIntlDeliveryRub,
      itemsTotalDutyRub,
      itemsTotalVatRub,
      itemsTotalExciseRub,
      itemsTotalCustomsRub,
      itemsTotalCertRub,
      itemsTotalOtherRub,
      itemsTotalCostsRub,
    ].forEach((el) => {
      if (el) el.textContent = "";
    });
    return;
  }
  itemsTable.style.display = "table";
  itemsTableBody.innerHTML = "";
  let totalInvoiceRub = 0,
    totalDomestic = 0,
    totalIntl = 0,
    totalDuty = 0,
    totalVat = 0,
    totalExcise = 0,
    totalCustoms = 0,
    totalCert = 0,
    totalOther = 0,
    totalCosts = 0;

  items.forEach((item, i) => {
    totalInvoiceRub += item.invoiceRub;
    totalDomestic += item.domesticDeliveryRub;
    totalIntl += item.intlDeliveryRub;
    totalDuty += item.dutyRub;
    totalVat += item.vatRub;
    totalExcise += item.exciseRub;
    totalCustoms += item.customsFeesRub;
    totalCert += item.certCostsRub;
    totalOther += item.otherCostsRub;
    totalCosts += item.totalCostsRub;

    const tr = document.createElement("tr");
    const cols = [
      i + 1,
      item.productName || "—",
      item.tnvedCode || "—",
      item.description || "—",
      item.quantity,
      item.unitPrice.toLocaleString("ru-RU", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }),
      item.totalInvoiceCurrency.toLocaleString("ru-RU", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }),
      item.currency,
      item.exchangeRate ? item.exchangeRate.toFixed(4).replace(".", ",") : "",
      formatMoney(item.invoiceRub),
      formatMoney(item.domesticDeliveryRub),
      formatMoney(item.intlDeliveryRub),
      item.dutyPercent != null ? String(item.dutyPercent).replace(".", ",") : "",
      formatMoney(item.dutyRub),
      item.vatPercent != null ? String(item.vatPercent).replace(".", ",") : "",
      formatMoney(item.vatRub),
      item.excisePercent != null ? String(item.excisePercent).replace(".", ",") : "",
      formatMoney(item.exciseRub),
      formatMoney(item.customsFeesRub),
      formatMoney(item.certCostsRub),
      formatMoney(item.otherCostsRub),
      formatMoney(item.totalCostsRub),
      formatMoney(item.unitCostRub),
    ];
    cols.forEach((v) => {
      const td = document.createElement("td");
      td.textContent = v;
      tr.appendChild(td);
    });

    // ячейка "Действия" с кнопкой удаления
    const actionTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.textContent = "✕";
    delBtn.className = "btn-delete-item";
    delBtn.setAttribute("data-index", i);
    actionTd.appendChild(delBtn);
    tr.appendChild(actionTd);

    itemsTableBody.appendChild(tr);
  });
  itemsTotalInvoiceRub.textContent = formatMoney(totalInvoiceRub);
  itemsTotalDomesticDeliveryRub.textContent = formatMoney(totalDomestic);
  itemsTotalIntlDeliveryRub.textContent = formatMoney(totalIntl);
  itemsTotalDutyRub.textContent = formatMoney(totalDuty);
  itemsTotalVatRub.textContent = formatMoney(totalVat);
  itemsTotalExciseRub.textContent = formatMoney(totalExcise);
  itemsTotalCustomsRub.textContent = formatMoney(totalCustoms);
  itemsTotalCertRub.textContent = formatMoney(totalCert);
  itemsTotalOtherRub.textContent = formatMoney(totalOther);
  itemsTotalCostsRub.textContent = formatMoney(totalCosts);
}

function exportItemsToExcel() {
  if (!items || !items.length) {
    alert("Нет данных для выгрузки.");
    return;
  }
  if (typeof XLSX === "undefined") {
    alert("Библиотека XLSX не загружена.");
    return;
  }

  const header = [
    "№",
    "Название товара",
    "Код ТН ВЭД",
    "Описание",
    "Кол-во",
    "Цена за шт (валюта)",
    "Сумма сделки (валюта)",
    "Валюта",
    "Курс (с надбавкой)",
    "Сумма сделки, RUB",
    "Дост. внутри Китая, RUB",
    "Междунар. дост. до РФ, RUB",
    "Пошлина, %",
    "Пошлина, RUB",
    "НДС, %",
    "НДС, RUB",
    "Акциз, %",
    "Акциз, RUB",
    "Тамож. сборы, RUB",
    "Сертификация, RUB",
    "Прочие издержки, RUB",
    "Итого расходов, RUB",
    "Себестоимость за шт, RUB",
  ];
  const aoa = [header];
  items.forEach((item, idx) => {
    const totalInvoiceCurrency = item.totalInvoiceCurrency;
    const row = [
      idx + 1,
      item.productName || "",
      item.tnvedCode || "",
      item.description || "",
      item.quantity,
      item.unitPrice,
      totalInvoiceCurrency,
      item.currency,
      item.exchangeRate,
      item.invoiceRub,
      item.domesticDeliveryRub,
      item.intlDeliveryRub,
      item.dutyPercent,
      item.dutyRub,
      item.vatPercent,
      item.vatRub,
      item.excisePercent,
      item.exciseRub,
      item.customsFeesRub,
      item.certCostsRub,
      item.otherCostsRub,
      item.totalCostsRub,
      item.unitCostRub,
    ];
    aoa.push(row);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Товары");
  XLSX.writeFile(wb, "sirius_items.xlsx");
}

function calculateInternal() {
  validationWarning.style.display = "none";
  validationWarning.textContent = "";
  const productName = document.getElementById("productName").value.trim();
  const unitPrice = readNumber("invoiceValue");
  const quantity = readNumber("quantity");
  const weight = readNumber("weight");
  const volume = readNumber("volume");
  const currency = document.getElementById("currency").value;
  let exchangeRate = readNumber("exchangeRate");
  const domesticDelivery = readNumber("domesticDelivery");
  const dutyPercent = readNumber("dutyPercent");
  const vatPercent = readNumber("vatPercent");
  const excisePercent = readNumber("excisePercent");
  let customsFeesRub = readNumber("customsFeesRub");
  const certCostsRub = readNumber("certCostsRub");
  const otherCostsRub = readNumber("otherCostsRub");
  const deliveryType = document.getElementById("deliveryType").value;
  const tnvedCodeVal = document.getElementById("tnvedCode").value.trim();

  const errors = [];
  const warnings = [];

  if (!productName) errors.push("Укажите название товара.");
  if (!tnvedCodeVal) errors.push("Укажите код ТН ВЭД.");
  if (!unitPrice || unitPrice <= 0) errors.push("Укажите стоимость товара за 1 шт (> 0).");
  if (!quantity || quantity <= 0) errors.push("Укажите количество (шт) (> 0).");

  if (currency !== "RUB" && (!exchangeRate || exchangeRate <= 0)) {
    errors.push("Укажите курс валюты к RUB (> 0) или дождитесь автозагрузки.");
  }

  if (deliveryType === "air") {
    if (!weight || weight <= 0) errors.push("Для авиа-доставки укажите вес (кг).");
  } else {
    if (!volume || volume <= 0) errors.push("Для авто/ЖД/морской доставки укажите объём (м³).");
  }

  if (volume > 0 && volume < 15) {
    warnings.push("Объём меньше 15 м³ — по вашим правилам такие заказы обычно не берёте.");
  }

  if (!(dutyPercent === 0 || dutyPercent > 0)) errors.push("Ставка пошлины не указана.");
  if (!(vatPercent === 0 || vatPercent > 0)) errors.push("Ставка НДС не указана (по умолчанию 20%).");

  if (!customsFeesRub || customsFeesRub <= 0) {
    customsFeesRub = 4900;
    const cf = document.getElementById("customsFeesRub");
    if (cf) cf.value = "4900";
  }

  const messages = [...errors, ...warnings];
  if (messages.length) {
    validationWarning.style.display = "block";
    validationWarning.innerHTML = messages.join("<br>");
  } else {
    validationWarning.style.display = "none";
    validationWarning.textContent = "";
  }
  if (errors.length) return { errors, warnings };

  if (currency === "RUB" && (!exchangeRate || exchangeRate <= 0)) exchangeRate = 1;

  const totalInvoiceCurrency = unitPrice * quantity;
  const invoiceRub = currency === "RUB" ? totalInvoiceCurrency : totalInvoiceCurrency * exchangeRate;
  const domesticDeliveryRub =
    currency === "RUB" ? domesticDelivery : domesticDelivery * exchangeRate;
  const intlDeliveryRub = calculateIntlDeliveryRub();

  const dutyRub = (invoiceRub * dutyPercent) / 100;
  const exciseRub = (invoiceRub * excisePercent) / 100;

  // новая база НДС: invoiceRub + dutyRub + exciseRub + domesticDeliveryRub + 20% международной логистики
  const intlVatPart = intlDeliveryRub * 0.2;
  const vatBaseRub = invoiceRub + dutyRub + exciseRub + domesticDeliveryRub + intlVatPart;
  const vatRub = (vatBaseRub * vatPercent) / 100;

  const totalDeliveryRub = domesticDeliveryRub + intlDeliveryRub;

  const totalCostsRub =
    invoiceRub +
    totalDeliveryRub +
    dutyRub +
    exciseRub +
    vatRub +
    customsFeesRub +
    certCostsRub +
    otherCostsRub;
  const unitCostRub = quantity > 0 ? totalCostsRub / quantity : 0;

  return {
    productName,
    unitPrice,
    quantity,
    weight,
    volume,
    currency,
    exchangeRate,
    domesticDelivery,
    domesticDeliveryRub,
    totalInvoiceCurrency,
    invoiceRub,
    intlDeliveryRub,
    dutyPercent,
    dutyRub,
    vatPercent,
    vatRub,
    excisePercent,
    exciseRub,
    customsFeesRub,
    certCostsRub,
    otherCostsRub,
    vatBaseRub,
    totalCostsRub,
    unitCostRub,
    deliveryType,
    errors,
    warnings,
  };
}

function showResult(calc) {
  const resultSummary = document.getElementById("resultSummary");
  const resultTable = document.getElementById("resultTable");
  const resultTableBody = document.getElementById("resultTableBody");
  const totalCostsCell = document.getElementById("totalCostsCell");
  const resultNote = document.getElementById("resultNote");
  resultSummary.innerHTML = "";
  resultTableBody.innerHTML = "";
  resultTable.style.display = "table";

  const pillAmount = document.createElement("div");
  pillAmount.className = "result-pill";
  pillAmount.innerHTML =
    "Итого расходов: <strong>" + formatMoney(calc.totalCostsRub) + " ₽</strong>";
  const pillUnit = document.createElement("div");
  pillUnit.className = "result-pill";
  pillUnit.innerHTML =
    "Себестоимость за 1 шт: <strong>" + formatMoney(calc.unitCostRub) + " ₽</strong>";
  const pillVatBase = document.createElement("div");
  pillVatBase.className = "result-pill";
  pillVatBase.innerHTML = "НДС-база: <strong>" + formatMoney(calc.vatBaseRub) + " ₽</strong>";

  if (calc.weight > 0) {
    const p = document.createElement("div");
    p.className = "result-pill";
    p.textContent = "Вес: " + calc.weight + " кг";
    resultSummary.appendChild(p);
  }
  if (calc.volume > 0) {
    const p = document.createElement("div");
    p.className = "result-pill";
    p.textContent = "Объём: " + calc.volume + " м³";
    resultSummary.appendChild(p);
  }
  resultSummary.appendChild(pillAmount);
  resultSummary.appendChild(pillUnit);
  resultSummary.appendChild(pillVatBase);

  const rows = [
    ["Стоимость товара (включая количество)", calc.invoiceRub],
    ["Доставка внутри Китая", calc.domesticDeliveryRub],
    ["Международная доставка до РФ", calc.intlDeliveryRub],
    ["Таможенная пошлина", calc.dutyRub],
    ["Акциз", calc.exciseRub],
    ["НДС", calc.vatRub],
    ["Таможенные сборы", calc.customsFeesRub],
    ["Сертификация", calc.certCostsRub],
    ["Прочие издержки", calc.otherCostsRub],
  ];
  rows.forEach(([label, value]) => {
    const tr = document.createElement("tr");
    const td1 = document.createElement("td");
    const td2 = document.createElement("td");
    td1.textContent = label;
    td2.textContent = formatMoney(value) + " ₽";
    tr.appendChild(td1);
    tr.appendChild(td2);
    resultTableBody.appendChild(tr);
  });
  totalCostsCell.textContent = formatMoney(calc.totalCostsRub) + " ₽";
  resultNote.textContent =
    "Итоги можно переносить в КП. Формулы упрощены и могут быть доработаны под вашу методику.";
}

function addItemFromCurrentInputs(clearAfter) {
  const calc = calculateInternal();
  if (calc.errors && calc.errors.length) return;
  showResult(calc);
  const tnvedCodeVal = document.getElementById("tnvedCode").value.trim();
  const tnvedRecord = findTnvedRecord(tnvedCodeVal);
  const item = {
    productName: calc.productName,
    tnvedCode: tnvedCodeVal,
    description: tnvedRecord ? tnvedRecord.description : "",
    unitPrice: calc.unitPrice,
    quantity: calc.quantity,
    currency: calc.currency,
    exchangeRate: calc.exchangeRate,
    totalInvoiceCurrency: calc.totalInvoiceCurrency,
    invoiceRub: calc.invoiceRub,
    volume: calc.volume,
    weight: calc.weight,
    intlDeliveryRub: calc.intlDeliveryRub,
    domesticDeliveryRub: calc.domesticDeliveryRub,
    dutyPercent: calc.dutyPercent,
    dutyRub: calc.dutyRub,
    vatPercent: calc.vatPercent,
    vatRub: calc.vatRub,
    excisePercent: calc.excisePercent,
    exciseRub: calc.exciseRub,
    customsFeesRub: calc.customsFeesRub,
    certCostsRub: calc.certCostsRub,
    otherCostsRub: calc.otherCostsRub,
    totalCostsRub: calc.totalCostsRub,
    unitCostRub: calc.unitCostRub,
    vatBaseRub: calc.vatBaseRub,
    deliveryType: calc.deliveryType,
  };
  if (items.length < 20) {
    items.push(item);
  } else {
    alert("Достигнуто максимальное количество товаров (20).");
  }
  renderItemsTable();
  if (clearAfter) clearMainInputs();
}

function clearMainInputs() {
  const ids = [
    "productName",
    "invoiceValue",
    "quantity",
    "tnvedCode",
    "weight",
    "volume",
    "exchangeRate",
    "domesticDelivery",
    "intlDeliveryRub",
    "dutyPercent",
    "customsFeesRub",
    "vatPercent",
    "excisePercent",
    "certCostsRub",
    "otherCostsRub",
  ];
  ids.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (id === "vatPercent") el.value = "20";
    else if (id === "certCostsRub") el.value = "30000";
    else if (id === "customsFeesRub") el.value = "4900";
    else if (id === "exchangeRate") el.value = "";
    else el.value = "";
  });
  validationWarning.style.display = "none";
  validationWarning.textContent = "";
  invoiceTotalHintEl.textContent = "Общая сумма сделки: —";
  const currencyEl = document.getElementById("currency");
  if (currencyEl && ["CNY", "USD", "EUR"].includes(currencyEl.value)) {
    fetchCbrRateForCurrency(currencyEl.value);
  }
}

function resetAll() {
  clearMainInputs();
  const deliveryTypeEl = document.getElementById("deliveryType");
  const currencyEl = document.getElementById("currency");
  if (deliveryTypeEl) deliveryTypeEl.value = "rail";
  if (currencyEl) currencyEl.value = "CNY";
  document.getElementById("tariffTruckPrice").value = "1250000";
  document.getElementById("tariffTruckVolume").value = "110";
  document.getElementById("tariffRailPrice").value = "535000"; // новый тариф ЖД
  document.getElementById("tariffRailVolume").value = "70";
  document.getElementById("tariffSeaPrice").value = "240000";
  document.getElementById("tariffSeaVolume").value = "70";
  document.getElementById("tariffAirPerKg").value = "50";
  tnvedInfoBody.innerHTML =
    "1) Загрузите JSON или Excel-файл ТН ВЭД.<br />2) Введите код ТН ВЭД и нажмите «Найти по ТН ВЭД».";
  tnvedInfoCard.style.borderColor = "#e5e7eb";
  tnvedInfoCard.style.background = "#f9fafb";
  document.getElementById("resultSummary").innerHTML = "";
  document.getElementById("resultTableBody").innerHTML = "";
  document.getElementById("resultTable").style.display = "none";
  document.getElementById("totalCostsCell").textContent = "";
  document.getElementById("resultNote").textContent = "";
  rateHintEl.textContent =
    "Для CNY / USD / EUR курс пробуем подтянуть по данным ЦБ РФ (в расчётах с надбавкой 3%).";
  items = [];
  renderItemsTable();
  fetchCbrRateForCurrency("CNY");
}

function handleFindTnved() {
  const rawCode = document.getElementById("tnvedCode").value;
  const record = findTnvedRecord(rawCode);
  if (!record) {
    updateTnvedInfo(null);
    return;
  }
  const dutyInput = document.getElementById("dutyPercent");
  const vatInput = document.getElementById("vatPercent");
  if (dutyInput) dutyInput.value = record.tariff;
  if (vatInput && record.vat) vatInput.value = record.vat;
  updateTnvedInfo(record);
}

// события
document.getElementById("tnvedCode").addEventListener("blur", () => {
  const code = document.getElementById("tnvedCode").value.trim();
  if (code.length > 0) handleFindTnved();
});
document.getElementById("btnFindTnved").addEventListener("click", handleFindTnved);
document.getElementById("btnCalculate").addEventListener("click", () =>
  addItemFromCurrentInputs(false)
);
document.getElementById("btnAddItem").addEventListener("click", () =>
  addItemFromCurrentInputs(true)
);
document.getElementById("btnReset").addEventListener("click", resetAll);
const btnClearItems = document.getElementById("btnClearItems");
if (btnClearItems)
  btnClearItems.addEventListener("click", () => {
    items = [];
    renderItemsTable();
  });
const btnExportExcel = document.getElementById("btnExportExcel");
if (btnExportExcel) btnExportExcel.addEventListener("click", exportItemsToExcel);

["invoiceValue", "quantity", "currency", "exchangeRate"].forEach((id) => {
  const el = document.getElementById(id);
  if (!el) return;
  const evt = id === "currency" ? "change" : "input";
  el.addEventListener(evt, () => {
    if (id === "currency") {
      const cur = el.value;
      if (["CNY", "USD", "EUR"].includes(cur)) fetchCbrRateForCurrency(cur);
      else rateHintEl.textContent = "Для валюты " + cur + " курс задаётся вручную.";
    }
    updateInvoiceTotalInfo();
    calculateIntlDeliveryRub();
  });
});
[
  "volume",
  "weight",
  "tariffTruckPrice",
  "tariffTruckVolume",
  "tariffRailPrice",
  "tariffRailVolume",
  "tariffSeaPrice",
  "tariffSeaVolume",
  "tariffAirPerKg",
  "deliveryType",
].forEach((id) => {
  const el = document.getElementById(id);
  if (!el) return;
  const evt = id === "deliveryType" ? "change" : "input";
  el.addEventListener(evt, () => {
    calculateIntlDeliveryRub();
  });
});

// Генерация КП
async function generateCommercialOfferExcel() {
  if (typeof ExcelJS === "undefined") {
    alert("Библиотека ExcelJS не загружена.");
    return;
  }
  if (!items || !items.length) {
    alert("Нет данных для КП. Добавьте хотя бы один товар.");
    return;
  }

  const rateInput = document.getElementById("exchangeRate");
  let cnyRate = 0;
  if (rateInput && rateInput.value) {
    cnyRate = parseFloat(String(rateInput.value).replace(",", "."));
  }
  if (!isFinite(cnyRate) || cnyRate <= 0) {
    alert("Не найден корректный курс CNY→RUB. Сначала задайте курс в калькуляторе.");
    return;
  }

  let totalVolume = 0;
  items.forEach((it) => {
    totalVolume += it.volume ? Number(it.volume) : 0;
  });

  let deliveryMarkupPercent;
  if (totalVolume < 5) deliveryMarkupPercent = 150;
  else if (totalVolume < 10) deliveryMarkupPercent = 100;
  else if (totalVolume < 15) deliveryMarkupPercent = 75;
  else deliveryMarkupPercent = 50;
  const deliveryMarkupFactor = 1 + deliveryMarkupPercent / 100;

  function readNum(id) {
    const el = document.getElementById(id);
    if (!el || el.value === undefined) return 0;
    const v = String(el.value).replace(",", ".");
    const n = parseFloat(v);
    return isFinite(n) ? n : 0;
  }
  const truckPrice = readNum("tariffTruckPrice");
  const truckVolume = readNum("tariffTruckVolume") || 1;
  const railPrice = readNum("tariffRailPrice");
  const railVolume = readNum("tariffRailVolume") || 1;

  function calcIntlDeliveryRub(mode, volume, item) {
    if (mode === "air") {
      return item.intlDeliveryRub || 0;
    }
    if (!volume || volume <= 0) return 0;
    if (mode === "truck") {
      if (truckPrice <= 0) return 0;
      return (truckPrice * volume) / truckVolume;
    }
    if (mode === "rail") {
      if (railPrice <= 0) return 0;
      return (railPrice * volume) / railVolume;
    }
    return 0;
  }

  const bankCommissionDefault = 0;
  const siriusCommissionPercent = 5;

  // Создаём рабочую книгу ExcelJS
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("КП");

  // Строка 1: Объединяем A1:O1 для логотипа и заголовка
  worksheet.mergeCells("A1:O1");
  worksheet.getRow(1).height = 103.5; // Высота строки 1 из шаблона
  worksheet.getRow(2).height = 60; // Высота строки 2 для логотипа
  
  // Загружаем логотип в A1 (точные размеры из шаблона)
  // Для онлайн-платформы: используем кэш, DOM или fetch
  const loadLogoImage = async () => {
    // 1. Используем кэшированный buffer (из input file или предыдущей загрузки)
    if (logoImageBuffer) {
      return logoImageBuffer;
    }

    // 2. Пробуем использовать изображение из DOM (уже загружено на странице)
    const logoImg = document.getElementById('logoImage') || 
                    document.querySelector('img[src*="image.png"], img[alt*="Логотип"]');
    
    if (logoImg && logoImg.complete && logoImg.naturalWidth > 0) {
      // Изображение уже загружено в DOM - конвертируем через canvas
      return new Promise((resolve, reject) => {
        try {
          const canvas = document.createElement('canvas');
          canvas.width = logoImg.naturalWidth;
          canvas.height = logoImg.naturalHeight;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(logoImg, 0, 0);
          
          canvas.toBlob((blob) => {
            if (blob) {
              const reader = new FileReader();
              reader.onload = (e) => {
                logoImageBuffer = e.target.result; // Кэшируем
                resolve(e.target.result);
              };
              reader.onerror = reject;
              reader.readAsArrayBuffer(blob);
            } else {
              reject(new Error('Не удалось создать blob из canvas'));
            }
          }, 'image/png');
        } catch (err) {
          reject(err);
        }
      });
    }

    // 3. Загружаем через fetch (основной способ для веб-сервера)
    // На онлайн-платформе это должно работать всегда
    try {
      const response = await fetch("image.png", { cache: 'no-cache' });
      if (!response.ok) {
        throw new Error('HTTP ' + response.status + ': ' + response.statusText);
      }
      const buffer = await response.arrayBuffer();
      logoImageBuffer = buffer; // Кэшируем для следующих раз
      return buffer;
    } catch (fetchErr) {
      console.warn("Ошибка загрузки логотипа через fetch:", fetchErr);
      // Если fetch не сработал, пробуем через изображение как fallback
      if (logoImg) {
        return new Promise((resolve, reject) => {
          logoImg.onload = function() {
            try {
              const canvas = document.createElement('canvas');
              canvas.width = this.naturalWidth;
              canvas.height = this.naturalHeight;
              const ctx = canvas.getContext('2d');
              ctx.drawImage(this, 0, 0);
              canvas.toBlob((blob) => {
                if (blob) {
                  const reader = new FileReader();
                  reader.onload = (e) => {
                    logoImageBuffer = e.target.result;
                    resolve(e.target.result);
                  };
                  reader.onerror = reject;
                  reader.readAsArrayBuffer(blob);
                } else {
                  reject(new Error('Не удалось создать blob из canvas'));
                }
              }, 'image/png');
            } catch (err) {
              reject(err);
            }
          };
          logoImg.onerror = () => reject(fetchErr);
          if (!logoImg.src || logoImg.src === window.location.href) {
            logoImg.src = "image.png";
          }
        });
      }
      throw fetchErr;
    }
  };

  try {
    const logoBuffer = await loadLogoImage();
    const imageId = workbook.addImage({
      buffer: logoBuffer,
      extension: "png",
    });
    // Логотип в левом верхнем углу (A1)
    // Размеры из шаблона: width: 711, height: 157 пикселей
    worksheet.addImage(imageId, {
      tl: { col: 0, row: 0 },
      ext: { width: 711, height: 157 }, // Точные размеры из шаблона
    });
  } catch (err) {
    console.warn("Не удалось загрузить логотип:", err);
    // Показываем предупреждение пользователю только если это критично
    // На онлайн-платформе обычно работает через fetch
  }

  // Логотип уже размещен в A1, текст компании не добавляем (как в шаблоне)

  // Строка 3: Дата расчёта (в колонке A, не конфликтует с адресом в D3)
  let currentRow = 3;
  const today = new Date();
  const dateCell = worksheet.getCell(`A${currentRow}`);
  dateCell.value = `Дата расчёта  ${today.toLocaleDateString("ru-RU")}`;
  dateCell.font = { bold: true, size: 11, name: "Calibri" };
  dateCell.numFmt = '#,##0.00 "₽"';
  dateCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFFFF" },
  };
  dateCell.alignment = { vertical: "bottom" };

  // Строка 5: Курс доллар (пропускаем, так как у нас только юань)
  currentRow = 5;

  // Строка 6: Курс юань +3%
  currentRow = 6;
  const cnyLabelCell = worksheet.getCell(`A${currentRow}`);
  cnyLabelCell.value = "Курс юань +3%";
  cnyLabelCell.font = { bold: true, size: 11, name: "Calibri" };
  cnyLabelCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFFFF" },
  };
  cnyLabelCell.border = {
    top: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };
  cnyLabelCell.alignment = { vertical: "bottom" };

  const cnyValueCell = worksheet.getCell(`B${currentRow}`);
  cnyValueCell.value = cnyRate;
  cnyValueCell.numFmt = "#,##0.00"; // Формат числа для курса
  cnyValueCell.font = { bold: true, size: 11, name: "Calibri" };
  cnyValueCell.border = {
    top: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };
  cnyValueCell.alignment = { horizontal: "center", vertical: "bottom" };
  
  const cnyRateRow = currentRow;
  const cnyRateCell = `$B$${cnyRateRow}`;
  // Комиссия банка и СИРИУС не используются в формулах шаблона, но оставим для совместимости
  const bankCommissionCell = "$B$7";
  const siriusCommissionCell = "$B$8";
  
  currentRow = 10; // Следующая строка для блока доставки

  const tableHeader = [
    "Модель",
    "Цена за шт (в юанях)",
    "Кол-во штук ",
    "Инвойс (юани)",
    "Инвойс (руб) с комиссией банка 2%",
    "Сбор (руб)",
    "Разрешительные документы (декларации, сертификаты), руб",
    "Пошлина (руб)",
    "НДС (руб)",
    "Таможенные платежи ИТОГО (руб)",
    "Доставка (руб)",
    "Инспекция (руб)",
    "Комиссия Сириус (руб)",
    "Итого все платежи (руб)",
    "Себестоимость единицы продукции (в руб, с НДС)",
    "", // Пустая колонка P
    "Вес",
    "Объем",
  ];

  function appendScenarioBlock(title, mode, deliveryDays) {
    // Заголовок блока (строка 10 для авто, 17 для ЖД)
    const titleCell = worksheet.getCell(`A${currentRow}`);
    titleCell.value = title;
    titleCell.font = { bold: true, size: 11, name: "Calibri" };
    titleCell.numFmt = '#,##0.00 "₽"';
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFFFF" },
    };
    titleCell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
    
    // Объединяем B и C для срока поставки
    worksheet.mergeCells(`B${currentRow}:C${currentRow}`);
    const deliveryCell = worksheet.getCell(`B${currentRow}`);
    deliveryCell.value = `Срок поставки ${deliveryDays}`;
    deliveryCell.font = { bold: true, size: 11, name: "Calibri" };
    deliveryCell.numFmt = "#,##0";
    deliveryCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFFFF" },
    };
    deliveryCell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
    
    currentRow++;

    // Заголовок таблицы (строка 11 для авто, 18 для ЖД)
    const headerRow = currentRow;
    tableHeader.forEach((header, colIdx) => {
      if (header === "") return; // Пропускаем пустые колонки
      const cell = worksheet.getCell(currentRow, colIdx + 1);
      cell.value = header;
      cell.font = { bold: true, size: 10, name: "Calibri" };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFFFF" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      // Формат для денежных колонок
      if (colIdx >= 4 && colIdx <= 14 && colIdx !== 11) {
        cell.numFmt = '#,##0.00 "₽"';
      } else if (colIdx === 11) {
        cell.numFmt = "0";
      }
    });
    currentRow++;
    const firstDataRow = currentRow;

    items.forEach((item, idx) => {
      const model = item.productName || "";
      const unitPrice = Number(item.unitPrice) || 0;
      const qty = Number(item.quantity) || 0;
      const volume = Number(item.volume) || 0;
      
      // Проверка данных для отладки
      if (!unitPrice || !qty) {
        console.warn(`Товар ${idx + 1} (${model}): unitPrice=${unitPrice}, qty=${qty}`);
      }

      const domesticDeliveryRub = item.domesticDeliveryRub || 0;
      const intlDeliveryRubBase = calcIntlDeliveryRub(mode, volume, item);
      const deliveryBaseRub = domesticDeliveryRub + intlDeliveryRubBase;
      const deliveryWithMarkup = deliveryBaseRub * deliveryMarkupFactor;

      const customsFeeBase = item.customsFeesRub || 0;
      const customsFeeKp = customsFeeBase * 2;
      const certBase = item.certCostsRub || 0;
      const certKp = certBase * 1.5;

      const dutyRub = item.dutyRub || 0;
      const vatRub = item.vatRub || 0;

      // Заполняем данные (по шаблону: формулы как D12*$B$6/0.98 для инвойса с комиссией банка 2%)
      worksheet.getCell(currentRow, 1).value = model; // A - Модель
      worksheet.getCell(currentRow, 2).value = unitPrice; // B - Цена за шт
      worksheet.getCell(currentRow, 3).value = qty; // C - Кол-во
      
      // Вычисляем значения для формул
      const invoiceCny = Number(unitPrice) * Number(qty); // D - Инвойс (юани)
      const invoiceRub = invoiceCny * Number(cnyRate) / 0.98; // E - Инвойс (руб) с комиссией банка 2%
      const dutyPercent = (item.dutyPercent != null && item.dutyPercent !== '') ? Number(item.dutyPercent) : 10;
      const duty = invoiceRub * dutyPercent / 100; // H - Пошлина
      const vat = invoiceRub * 20 / 100; // I - НДС
      const customsTotal = duty + vat + customsFeeKp + certKp; // J - Таможенные платежи
      const siriusCommission = (deliveryWithMarkup + invoiceRub + customsTotal) * 0.05; // M - Комиссия Сириус
      const totalPayments = deliveryWithMarkup + invoiceRub + customsTotal + siriusCommission + 0; // N - Итого платежи
      const unitCost = qty > 0 ? totalPayments / qty : 0; // O - Себестоимость единицы
      
      // Отладка: проверяем вычисленные значения
      if (idx === 0) {
        console.log('Первый товар - вычисленные значения:', {
          invoiceCny,
          invoiceRub,
          duty,
          vat,
          customsTotal,
          siriusCommission,
          totalPayments,
          unitCost,
          dutyPercent
        });
      }
      
      // A - Модель
      const cellA = worksheet.getCell(currentRow, 1);
      cellA.value = model;
      cellA.font = { size: 11, name: "Calibri" };
      cellA.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellA.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellA.alignment = { vertical: "middle", horizontal: "left" };
      
      // B - Цена за шт
      const cellB = worksheet.getCell(currentRow, 2);
      cellB.value = unitPrice;
      cellB.font = { size: 11, name: "Calibri" };
      cellB.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellB.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellB.alignment = { vertical: "middle", horizontal: "right" };
      cellB.numFmt = "#,##0.00";
      
      // C - Кол-во
      const cellC = worksheet.getCell(currentRow, 3);
      cellC.value = qty;
      cellC.font = { size: 11, name: "Calibri" };
      cellC.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellC.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellC.alignment = { vertical: "middle", horizontal: "right" };
      cellC.numFmt = "#,##0";
      
      // D - Инвойс (юани) = B * C
      const cellD = worksheet.getCell(currentRow, 4);
      cellD.value = { formula: `B${currentRow}*C${currentRow}`, result: Number(invoiceCny) };
      cellD.numFmt = "#,##0.00";
      cellD.font = { size: 11, name: "Calibri" };
      cellD.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellD.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellD.alignment = { vertical: "middle", horizontal: "right" };
      
      // E - Инвойс (руб) с комиссией банка 2% = D * курс / 0.98
      const cellE = worksheet.getCell(currentRow, 5);
      cellE.value = { formula: `D${currentRow}*${cnyRateCell}/0.98`, result: Number(invoiceRub) };
      cellE.numFmt = "#,##0.00";
      cellE.font = { size: 11, name: "Calibri" };
      cellE.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellE.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellE.alignment = { vertical: "middle", horizontal: "right" };
      
      // F - Сбор
      const cellF = worksheet.getCell(currentRow, 6);
      cellF.value = customsFeeKp;
      cellF.numFmt = "#,##0.00";
      cellF.font = { size: 11, name: "Calibri" };
      cellF.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellF.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellF.alignment = { vertical: "middle", horizontal: "right" };
      
      // G - Разрешительные документы
      const cellG = worksheet.getCell(currentRow, 7);
      cellG.value = certKp;
      cellG.numFmt = "#,##0.00";
      cellG.font = { size: 11, name: "Calibri" };
      cellG.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellG.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellG.alignment = { vertical: "middle", horizontal: "right" };
      
      // H - Пошлина = E * процент_пошлины / 100
      const cellH = worksheet.getCell(currentRow, 8);
      cellH.value = { formula: `E${currentRow}*${dutyPercent}/100`, result: Number(duty) };
      cellH.numFmt = "#,##0.00";
      cellH.font = { size: 11, name: "Calibri" };
      cellH.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellH.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellH.alignment = { vertical: "middle", horizontal: "right" };
      
      // I - НДС 20% = E * 20 / 100
      const cellI = worksheet.getCell(currentRow, 9);
      cellI.value = { formula: `E${currentRow}*20/100`, result: Number(vat) };
      cellI.numFmt = "#,##0.00";
      cellI.font = { size: 11, name: "Calibri" };
      cellI.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellI.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellI.alignment = { vertical: "middle", horizontal: "right" };
      
      // J - Таможенные платежи ИТОГО = H + I + F + G
      const cellJ = worksheet.getCell(currentRow, 10);
      cellJ.value = { formula: `H${currentRow}+I${currentRow}+F${currentRow}+G${currentRow}`, result: Number(customsTotal) };
      cellJ.numFmt = "#,##0.00";
      cellJ.font = { size: 11, name: "Calibri" };
      cellJ.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellJ.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellJ.alignment = { vertical: "middle", horizontal: "right" };
      
      // K - Доставка (значение, потом обновим формулой)
      const cellK = worksheet.getCell(currentRow, 11);
      cellK.value = deliveryWithMarkup;
      cellK.numFmt = "#,##0.00";
      cellK.font = { size: 11, name: "Calibri" };
      cellK.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellK.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellK.alignment = { vertical: "middle", horizontal: "right" };
      
      // L - Инспекция
      const cellL = worksheet.getCell(currentRow, 12);
      cellL.value = 0;
      cellL.numFmt = "0";
      cellL.font = { size: 11, name: "Calibri" };
      cellL.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellL.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellL.alignment = { vertical: "middle", horizontal: "right" };
      
      // M - Комиссия Сириус 5% = (K + E + J) * 0.05
      const cellM = worksheet.getCell(currentRow, 13);
      cellM.value = { formula: `(K${currentRow}+E${currentRow}+J${currentRow})*0.05`, result: Number(siriusCommission) };
      cellM.numFmt = "#,##0.00";
      cellM.font = { size: 11, name: "Calibri" };
      cellM.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellM.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellM.alignment = { vertical: "middle", horizontal: "right" };
      
      // N - Итого все платежи = K + E + J + M + L
      const cellN = worksheet.getCell(currentRow, 14);
      cellN.value = { formula: `K${currentRow}+E${currentRow}+J${currentRow}+M${currentRow}+L${currentRow}`, result: Number(totalPayments) };
      cellN.numFmt = "#,##0.00";
      cellN.font = { size: 11, name: "Calibri" };
      cellN.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellN.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellN.alignment = { vertical: "middle", horizontal: "right" };
      
      // O - Себестоимость единицы = N / C
      const cellO = worksheet.getCell(currentRow, 15);
      cellO.value = { formula: `N${currentRow}/C${currentRow}`, result: Number(unitCost) };
      cellO.numFmt = "#,##0.00";
      cellO.font = { size: 11, name: "Calibri" };
      cellO.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellO.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellO.alignment = { vertical: "middle", horizontal: "right" };
      
      // P - пустая колонка (16)
      
      // Q - Вес (17)
      const cellQ = worksheet.getCell(currentRow, 17);
      cellQ.value = item.weight || 0;
      cellQ.numFmt = "#,##0.00";
      cellQ.font = { size: 11, name: "Calibri" };
      cellQ.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellQ.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellQ.alignment = { vertical: "middle", horizontal: "left" };
      
      // R - Объем (18)
      const cellR = worksheet.getCell(currentRow, 18);
      cellR.value = item.volume || 0;
      cellR.numFmt = "#,##0.00";
      cellR.font = { size: 11, name: "Calibri" };
      cellR.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cellR.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cellR.alignment = { vertical: "middle", horizontal: "left" };

      currentRow++;
    });

    const lastDataRow = currentRow - 1;

    // Вычисляем общую доставку для всех товаров и суммы для строки ИТОГО
    let totalDeliveryValue = 0;
    let totalVolumeValue = 0;
    let totalQty = 0;
    let totalInvoiceCny = 0;
    let totalInvoiceRub = 0;
    let totalCustomsFee = 0;
    let totalCert = 0;
    let totalDuty = 0;
    let totalVat = 0;
    let totalCustoms = 0;
    
    items.forEach((item) => {
      const volume = item.volume || 0;
      const domesticDeliveryRub = item.domesticDeliveryRub || 0;
      const intlDeliveryRubBase = calcIntlDeliveryRub(mode, volume, item);
      const deliveryBaseRub = domesticDeliveryRub + intlDeliveryRubBase;
      totalDeliveryValue += deliveryBaseRub * deliveryMarkupFactor;
      totalVolumeValue += volume;
      
      // Суммируем значения для строки ИТОГО
      const unitPrice = Number(item.unitPrice) || 0;
      const qty = Number(item.quantity) || 0;
      const invoiceCny = unitPrice * qty;
      const invoiceRub = invoiceCny * cnyRate / 0.98;
      const dutyPercent = (item.dutyPercent != null && item.dutyPercent !== '') ? Number(item.dutyPercent) : 10;
      const duty = invoiceRub * dutyPercent / 100;
      const vat = invoiceRub * 20 / 100;
      const customsFeeKp = (item.customsFeesRub || 0) * 2;
      const certKp = (item.certCostsRub || 0) * 1.5;
      
      totalQty += qty;
      totalInvoiceCny += invoiceCny;
      totalInvoiceRub += invoiceRub;
      totalCustomsFee += customsFeeKp;
      totalCert += certKp;
      totalDuty += duty;
      totalVat += vat;
      totalCustoms += duty + vat + customsFeeKp + certKp;
    });
    
    // Строка ИТОГО (строка 15 для авто, 22 для ЖД)
    worksheet.getCell(currentRow, 1).value = "ИТОГО";
    worksheet.getCell(currentRow, 1).font = { bold: true, size: 11, name: "Calibri" };
    
    // Формулы для строки ИТОГО с предвычисленными значениями
    const cellC = worksheet.getCell(currentRow, 3);
    cellC.value = { formula: `SUM(C${firstDataRow}:C${lastDataRow})`, result: Number(totalQty) };
    cellC.numFmt = "#,##0";
    
    const cellD = worksheet.getCell(currentRow, 4);
    cellD.value = { formula: `SUM(D${firstDataRow}:D${lastDataRow})`, result: Number(totalInvoiceCny) };
    cellD.numFmt = "#,##0.00";
    
    const cellE = worksheet.getCell(currentRow, 5);
    cellE.value = { formula: `SUM(E${firstDataRow}:E${lastDataRow})`, result: Number(totalInvoiceRub) };
    cellE.numFmt = "#,##0.00";
    
    const cellF = worksheet.getCell(currentRow, 6);
    cellF.value = { formula: `SUM(F${firstDataRow}:F${lastDataRow})`, result: Number(totalCustomsFee) };
    cellF.numFmt = "#,##0.00";
    
    const cellG = worksheet.getCell(currentRow, 7);
    cellG.value = { formula: `SUM(G${firstDataRow}:G${lastDataRow})`, result: Number(totalCert) };
    cellG.numFmt = "#,##0.00";
    
    const cellH = worksheet.getCell(currentRow, 8);
    cellH.value = { formula: `SUM(H${firstDataRow}:H${lastDataRow})`, result: Number(totalDuty) };
    cellH.numFmt = "#,##0.00";
    
    const cellI = worksheet.getCell(currentRow, 9);
    cellI.value = { formula: `SUM(I${firstDataRow}:I${lastDataRow})`, result: Number(totalVat) };
    cellI.numFmt = "#,##0.00";
    
    const cellJ = worksheet.getCell(currentRow, 10);
    cellJ.value = { formula: `SUM(J${firstDataRow}:J${lastDataRow})`, result: Number(totalCustoms) };
    cellJ.numFmt = "#,##0.00";
    
    // Вычисляем суммы для M и N
    let totalSiriusCommission = 0;
    let totalPayments = 0;
    let totalWeight = 0;
    
    items.forEach((item, idx) => {
      const unitPrice = Number(item.unitPrice) || 0;
      const qty = Number(item.quantity) || 0;
      const invoiceCny = unitPrice * qty;
      const invoiceRub = invoiceCny * cnyRate / 0.98;
      const dutyPercent = item.dutyPercent || 10;
      const duty = invoiceRub * dutyPercent / 100;
      const vat = invoiceRub * 20 / 100;
      const customsFeeKp = (item.customsFeesRub || 0) * 2;
      const certKp = (item.certCostsRub || 0) * 1.5;
      const customsTotal = duty + vat + customsFeeKp + certKp;
      
      const volume = Number(item.volume) || 0;
      const domesticDeliveryRub = item.domesticDeliveryRub || 0;
      const intlDeliveryRubBase = calcIntlDeliveryRub(mode, volume, item);
      const deliveryBaseRub = domesticDeliveryRub + intlDeliveryRubBase;
      const deliveryWithMarkup = deliveryBaseRub * deliveryMarkupFactor;
      
      // Если есть объем, доставка будет пересчитана формулой, иначе используем вычисленное значение
      const finalDelivery = totalVolumeValue > 0 ? (totalDeliveryValue / totalVolumeValue * volume) : deliveryWithMarkup;
      
      const siriusCommission = (finalDelivery + invoiceRub + customsTotal) * 0.05;
      const payments = finalDelivery + invoiceRub + customsTotal + siriusCommission + 0;
      
      totalSiriusCommission += siriusCommission;
      totalPayments += payments;
      totalWeight += Number(item.weight) || 0;
    });
    
    // K - Доставка ИТОГО (значение или формула)
    if (totalVolumeValue > 0) {
      worksheet.getCell(currentRow, 11).value = totalDeliveryValue;
      worksheet.getCell(currentRow, 11).numFmt = "#,##0.00";
      // Обновляем формулы доставки для всех строк данных (пропорционально объёму)
      // Формула как в шаблоне: $K$15/$R$15*R12
      const totalDeliveryRow = currentRow;
      const totalVolumeRow = currentRow;
      for (let dataRow = firstDataRow; dataRow <= lastDataRow; dataRow++) {
        const item = items[dataRow - firstDataRow];
        const volume = Number(item.volume) || 0;
        const deliveryValue = totalVolumeValue > 0 ? (totalDeliveryValue / totalVolumeValue * volume) : 0;
        const cellK = worksheet.getCell(dataRow, 11);
        cellK.value = { formula: `$K$${totalDeliveryRow}/$R$${totalVolumeRow}*R${dataRow}`, result: Number(deliveryValue) };
        cellK.numFmt = "#,##0.00";
        
        // Пересчитываем M и N с новой доставкой
        const unitPrice = Number(item.unitPrice) || 0;
        const qty = Number(item.quantity) || 0;
        const invoiceCny = unitPrice * qty;
        const invoiceRub = invoiceCny * cnyRate / 0.98;
        const dutyPercent = (item.dutyPercent != null && item.dutyPercent !== '') ? Number(item.dutyPercent) : 10;
        const duty = invoiceRub * dutyPercent / 100;
        const vat = invoiceRub * 20 / 100;
        const customsFeeKp = (item.customsFeesRub || 0) * 2;
        const certKp = (item.certCostsRub || 0) * 1.5;
        const customsTotal = duty + vat + customsFeeKp + certKp;
        const siriusCommission = (deliveryValue + invoiceRub + customsTotal) * 0.05;
        const totalPayments = deliveryValue + invoiceRub + customsTotal + siriusCommission + 0;
        
        const cellM = worksheet.getCell(dataRow, 13);
        cellM.value = { formula: `(K${dataRow}+E${dataRow}+J${dataRow})*0.05`, result: Number(siriusCommission) };
        
        const cellN = worksheet.getCell(dataRow, 14);
        cellN.value = { formula: `K${dataRow}+E${dataRow}+J${dataRow}+M${dataRow}+L${dataRow}`, result: Number(totalPayments) };
      }
    } else {
      const cellK = worksheet.getCell(currentRow, 11);
      cellK.value = { formula: `SUM(K${firstDataRow}:K${lastDataRow})`, result: Number(totalDeliveryValue) };
      cellK.numFmt = "#,##0.00";
    }
    
    const cellL = worksheet.getCell(currentRow, 12);
    cellL.value = { formula: `SUM(L${firstDataRow}:L${lastDataRow})`, result: 0 };
    cellL.numFmt = "0";
    
    const cellM = worksheet.getCell(currentRow, 13);
    cellM.value = { formula: `SUM(M${firstDataRow}:M${lastDataRow})`, result: Number(totalSiriusCommission) };
    cellM.numFmt = "#,##0.00";
    
    const cellN = worksheet.getCell(currentRow, 14);
    cellN.value = { formula: `SUM(N${firstDataRow}:N${lastDataRow})`, result: Number(totalPayments) };
    cellN.numFmt = "#,##0.00";
    
    const cellQ = worksheet.getCell(currentRow, 17);
    cellQ.value = { formula: `SUM(Q${firstDataRow}:Q${lastDataRow})`, result: Number(totalWeight) }; // Вес
    cellQ.numFmt = "#,##0.00";
    
    const cellR = worksheet.getCell(currentRow, 18);
    cellR.value = { formula: `SUM(R${firstDataRow}:R${lastDataRow})`, result: Number(totalVolumeValue) }; // Объем
    cellR.numFmt = "#,##0.00";

    // Форматирование строки ИТОГО (сохраняем формулы и numFmt)
    for (let col = 1; col <= 18; col++) {
      if (col === 16) continue; // Пропускаем пустую колонку P
      const cell = worksheet.getCell(currentRow, col);
      
      // Сохраняем существующие formula и numFmt
      const existingFormula = cell.formula;
      const existingNumFmt = cell.numFmt;
      const existingResult = cell.result;
      
      cell.font = { bold: true, size: 11, name: "Calibri" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFFFF" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      
      // Восстанавливаем formula и result, если они были
      if (existingFormula) {
        cell.formula = existingFormula;
        if (existingResult !== undefined) {
          cell.result = existingResult;
        }
      }
      
      // Восстанавливаем numFmt, если он был установлен
      if (existingNumFmt) {
        cell.numFmt = existingNumFmt;
      } else if (col > 1 && col !== 2 && col !== 15) {
        // Устанавливаем numFmt только для ячеек без формул
        if (col === 12) {
          cell.numFmt = "0";
        } else {
          cell.numFmt = "#,##0";
        }
      }
      
      if (col > 1 && col !== 2 && col !== 15) {
        cell.alignment = { vertical: "middle", horizontal: "right" };
      } else if (col === 1) {
        cell.alignment = { vertical: "middle", horizontal: "left" };
      }
    }

    currentRow += 2;
  }

  // Если все товары с авиа-доставкой — делаем только авиа-блок
  const allAir = items.every((it) => it.deliveryType === "air");
  if (allAir) {
    appendScenarioBlock("Авиа доставка", "air", "7-14 дней");
  } else {
    appendScenarioBlock("Авто сборка", "truck", "28-35 дней");
    currentRow += 2; // Пропуск строки между блоками
    appendScenarioBlock("ЖД", "rail", "30-35 дней");
  }

  // Настройка ширины столбцов (по шаблону)
  worksheet.columns = [
    { width: 13.88 }, // A - Модель
    { width: 16.63 }, // B - Цена за шт
    { width: 12 }, // C - Кол-во
    { width: 15 }, // D - Инвойс (юани)
    { width: 20 }, // E - Инвойс (руб)
    { width: 12 }, // F - Сбор
    { width: 18.38 }, // G - Разрешительные документы
    { width: 15 }, // H - Пошлина
    { width: 15 }, // I - НДС
    { width: 20 }, // J - Таможенные платежи ИТОГО
    { width: 15 }, // K - Доставка
    { width: 15 }, // L - Инспекция
    { width: 18 }, // M - Комиссия Сириус
    { width: 20 }, // N - Итого все платежи
    { width: 17.75 }, // O - Себестоимость единицы
    { width: 5 }, // P - пустая
    { width: 13.88 }, // Q - Вес
    { width: 12 }, // R - Объем
  ];

  // Сохранение файла
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "KP_Sirius_Import.xlsx";
  link.click();
  window.URL.revokeObjectURL(url);
}

// Загрузка логотипа через input file (для работы при file://)
const logoFileInput = document.getElementById("logoFileInput");
const logoFileName = document.getElementById("logo-file-name");
const logoStatus = document.getElementById("logo-status");

if (logoFileInput) {
  logoFileInput.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    
    logoFileName.textContent = file.name;
    const reader = new FileReader();
    reader.onload = (e) => {
      logoImageBuffer = e.target.result;
      if (logoStatus) {
        logoStatus.style.display = "block";
        logoStatus.classList.add("ok");
        logoStatus.querySelector("span:nth-child(2)").textContent = "Логотип загружен: " + file.name;
      }
    };
    reader.onerror = () => {
      logoImageBuffer = null;
      if (logoStatus) {
        logoStatus.style.display = "block";
        logoStatus.classList.add("error");
        logoStatus.querySelector("span:nth-child(2)").textContent = "Ошибка загрузки логотипа.";
      }
      logoFileName.textContent = "";
    };
    reader.readAsArrayBuffer(file);
  });
}

// Пытаемся предзагрузить логотип при загрузке страницы
window.addEventListener("load", () => {
  const logoImg = document.getElementById("logoImage");
  if (logoImg && logoImg.complete && logoImg.naturalWidth > 0) {
    // Изображение уже загружено - конвертируем в buffer
    const canvas = document.createElement("canvas");
    canvas.width = logoImg.naturalWidth;
    canvas.height = logoImg.naturalHeight;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(logoImg, 0, 0);
    canvas.toBlob((blob) => {
      if (blob) {
        const reader = new FileReader();
        reader.onload = (e) => {
          logoImageBuffer = e.target.result;
        };
        reader.readAsArrayBuffer(blob);
      }
    }, "image/png");
  }
});

const btnGenerateKP = document.getElementById("btnGenerateKP");
if (btnGenerateKP) btnGenerateKP.addEventListener("click", generateCommercialOfferExcel);

// стартовый курс CNY
fetchCbrRateForCurrency("CNY");
