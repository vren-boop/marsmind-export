import { writeFile } from "node:fs/promises";
import { sheets } from "@googleapis/sheets";
import { GoogleAuth } from "google-auth-library";
import { chromium } from "playwright";
import * as XLSX from "xlsx";

const DATA_EXPORT_URL =
  process.env.MARSMIND_DATA_EXPORT_URL || "https://admin.marsmind.cc/data-export";
const GOOGLE_SHEET_ID =
  process.env.GOOGLE_SHEET_ID || "1nibt62SYKxDdF7QWt7tKnApUBDJ9KD14YeoCw3fJnfs";
const GOOGLE_WORKSHEET_NAME =
  process.env.GOOGLE_WORKSHEET_NAME || "Wayfair AI助理 Daily Chats";
const COMMUNICATION_WINDOW_NAME =
  process.env.COMMUNICATION_WINDOW_NAME || "wayfair AI助理";
const TIME_ZONE = process.env.TIME_ZONE || "Asia/Shanghai";

const TODAY_TEXTS = ["当天", "今天", "今日"];
const PREVIOUS_DAY_TEXTS = ["前一日", "前一天", "昨天", "昨日"];
const CUSTOM_DATE_TEXTS = ["自定义", "自定义时间", "自定义日期", "自定义范围"];
const DATE_RANGE_DROPDOWN_TEXTS = ["最近30天", "最近7天", "最近3天", "当天", "自定义"];
const DATE_PICKER_TRIGGER_SELECTORS = [
  'input[placeholder*="日期"]',
  'input[placeholder*="时间"]',
  'input[placeholder*="开始"]',
  'input[placeholder*="结束"]',
  ".ant-picker",
  ".ant-picker-range",
  ".ant-calendar-picker",
  ".el-date-editor",
  ".el-range-editor",
];
const DATE_INPUT_SELECTORS = [
  'input[placeholder*="日期"]',
  'input[placeholder*="时间"]',
  'input[placeholder*="开始"]',
  'input[placeholder*="结束"]',
  ".ant-picker input",
  ".ant-calendar-input",
  ".el-range-input",
  ".el-date-editor input",
];
const DATE_CONFIRM_TEXTS = ["确定", "确认", "应用", "Apply", "OK"];
const EXPORT_BUTTON_TEXTS = ["导出", "导 出", "导出数据", "下载", "Export"];
const COMM_WINDOW_LABEL_TEXTS = ["沟通窗口", "会话窗口", "客服窗口"];

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function getMetricDate() {
  if (process.env.METRIC_DATE) {
    return process.env.METRIC_DATE;
  }
  const today = new Date();
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(today);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function parseGoogleCredentials() {
  const rawJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  const base64Json = process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64;
  if (!rawJson && !base64Json) {
    throw new Error(
      "Missing GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_JSON_BASE64"
    );
  }
  const jsonText = rawJson
    ? rawJson
    : Buffer.from(base64Json, "base64").toString("utf8");
  const credentials = JSON.parse(jsonText);
  if (credentials.private_key) {
    credentials.private_key = credentials.private_key.replace(/\\n/g, "\n");
  }
  return credentials;
}

async function fillFirstVisible(page, selectors, value, label) {
  for (const selector of selectors) {
    const locator = page.locator(selector).first();
    if ((await locator.count()) > 0 && (await locator.isVisible().catch(() => false))) {
      await locator.fill(value);
      return;
    }
  }
  throw new Error(`Could not find ${label} input. Set ${label.toUpperCase()}_SELECTOR.`);
}

async function clickFirstVisible(page, selectors, label) {
  for (const selector of selectors) {
    const locator = page.locator(selector).first();
    if ((await locator.count()) > 0 && (await locator.isVisible().catch(() => false))) {
      await locator.click();
      return;
    }
  }
  throw new Error(`Could not find ${label}.`);
}

async function clickFirstText(page, texts, label) {
  for (const text of texts) {
    const locator = page.getByText(text, { exact: false }).first();
    if ((await locator.count()) > 0 && (await locator.isVisible().catch(() => false))) {
      await locator.click();
      return;
    }
  }
  throw new Error(`Could not find ${label}. Tried: ${texts.join(", ")}.`);
}

async function clickFirstTextIfVisible(page, texts) {
  for (const text of texts) {
    const locator = page.getByText(text, { exact: false }).first();
    if ((await locator.count()) > 0 && (await locator.isVisible().catch(() => false))) {
      await locator.click();
      return true;
    }
  }
  return false;
}

async function clickFirstVisibleOptional(page, selectors) {
  for (const selector of selectors.filter(Boolean)) {
    const locator = page.locator(selector).first();
    if ((await locator.count()) > 0 && (await locator.isVisible().catch(() => false))) {
      await locator.click();
      return true;
    }
  }
  return false;
}

async function clickNthVisibleOptional(page, selector, indexes) {
  if (!selector) {
    return false;
  }
  const locator = page.locator(selector);
  const count = await locator.count();
  for (const index of indexes) {
    if (count <= index) {
      continue;
    }
    const candidate = locator.nth(index);
    if (await candidate.isVisible().catch(() => false)) {
      await candidate.click();
      return true;
    }
  }
  return false;
}

async function fillVisibleDateInputs(page, selectors, metricDate) {
  const uniqueSelectors = [...new Set(selectors.filter(Boolean))];
  const visibleInputs = [];
  for (const selector of uniqueSelectors) {
    const locator = page.locator(selector);
    const count = await locator.count();
    for (let i = 0; i < count; i += 1) {
      const input = locator.nth(i);
      if (await input.isVisible().catch(() => false)) {
        visibleInputs.push(input);
      }
    }
  }

  if (!visibleInputs.length) {
    throw new Error("Could not find visible date input(s). Set DATE_INPUT_SELECTOR.");
  }

  const inputsToFill = visibleInputs.slice(0, Math.min(2, visibleInputs.length));
  for (const input of inputsToFill) {
    await input.click({ clickCount: 3 }).catch(() => {});
    await input.fill(metricDate);
  }
  const lastInput = inputsToFill.at(-1);
  if (lastInput) {
    await lastInput.press("Enter").catch(() => {});
  }
}

async function openDateRangeDropdown(page) {
  if (process.env.DATE_RANGE_SELECTOR) {
    await page.locator(process.env.DATE_RANGE_SELECTOR).first().click();
    return;
  }
  if (await clickNthVisibleOptional(page, ".ant-select-selector", [1, 0, 2])) {
    return;
  }
  if (await clickNthVisibleOptional(page, "[role=\"combobox\"]", [1, 0, 2])) {
    return;
  }
  if (await clickNthVisibleOptional(page, ".el-select, .el-input__wrapper", [1, 0, 2])) {
    return;
  }
  await clickFirstText(page, DATE_RANGE_DROPDOWN_TEXTS, "date-range dropdown");
}

async function selectCustomDateRange(page, metricDate) {
  const opened =
    (await clickNthVisibleOptional(page, ".ant-select-selector", [1, 0, 2])) ||
    (await clickNthVisibleOptional(page, "[role=\"combobox\"]", [1, 0, 2])) ||
    (await clickFirstVisibleOptional(page, [
      process.env.DATE_RANGE_SELECTOR,
      process.env.DATE_PICKER_SELECTOR,
      ...DATE_PICKER_TRIGGER_SELECTORS,
    ])) ||
    (await clickFirstTextIfVisible(page, DATE_RANGE_DROPDOWN_TEXTS)) ||
    (await clickFirstTextIfVisible(page, ["日期", "时间", "选择日期"]));

  if (!opened) {
    throw new Error(
      "Could not open date picker. Set DATE_PICKER_SELECTOR or DATE_INPUT_SELECTOR."
    );
  }

  await page.waitForTimeout(300);
  await clickFirstTextIfVisible(page, CUSTOM_DATE_TEXTS);
  await page.waitForTimeout(300);

  await fillVisibleDateInputs(
    page,
    [process.env.DATE_INPUT_SELECTOR, ...DATE_INPUT_SELECTORS],
    metricDate
  );
  await page.waitForTimeout(300);
  await clickFirstTextIfVisible(page, DATE_CONFIRM_TEXTS);
}

async function loginToMarsMind(page) {
  const username = requiredEnv("MARSMIND_USERNAME");
  const password = requiredEnv("MARSMIND_PASSWORD");

  await page.goto(DATA_EXPORT_URL, { waitUntil: "domcontentloaded", timeout: 60_000 });

  const usernameSelectors = [
    process.env.USERNAME_SELECTOR,
    'input[id="username"]',
    'input[type="email"]',
    'input[name="email"]',
    'input[name="username"]',
    'input[name="account"]',
    'input[placeholder*="用户名"]',
    'input[placeholder*="邮箱"]',
    'input[placeholder*="账号"]',
    'input[type="text"]',
  ].filter(Boolean);

  const passwordSelectors = [
    process.env.PASSWORD_SELECTOR,
    'input[id="password"]',
    'input[type="password"]',
    'input[name="password"]',
    'input[placeholder*="密码"]',
  ].filter(Boolean);

  const loginButtonSelectors = [
    process.env.LOGIN_BUTTON_SELECTOR,
    'button[type="submit"]',
    'button:has-text("登录")',
    'button:has-text("Login")',
    'input[type="submit"]',
  ].filter(Boolean);

  const passwordField = page.locator(passwordSelectors.join(", ")).first();
  const needsLogin =
    (await passwordField.count()) > 0 &&
    (await passwordField.isVisible().catch(() => false));

  if (!needsLogin) {
    return;
  }

  await fillFirstVisible(page, usernameSelectors, username, "username");
  await fillFirstVisible(page, passwordSelectors, password, "password");
  await clickFirstVisible(page, loginButtonSelectors, "login button");

  await page.waitForLoadState("networkidle", { timeout: 60_000 }).catch(() => {});
  await page.waitForTimeout(2_000);

  const loginPageStillVisible = await page
    .locator('input[id="password"], input[type="password"]')
    .first()
    .isVisible()
    .catch(() => false);
  if (loginPageStillVisible) {
    const pageText = await page.locator("body").innerText().catch(() => "");
    const hint = pageText
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean)
      .find((line) => /错误|失败|无效|验证码|用户名|密码/.test(line));
    throw new Error(
      `Login did not succeed${hint ? `: ${hint}` : ". Still on login page after submit."}`
    );
  }

  await page.goto(DATA_EXPORT_URL, { waitUntil: "networkidle", timeout: 60_000 }).catch(() => {});
}

async function selectPreviousDay(page, metricDate) {
  if (process.env.PREVIOUS_DAY_SELECTOR) {
    await page.locator(process.env.PREVIOUS_DAY_SELECTOR).first().click();
    return;
  }
  if (process.env.DATE_INPUT_SELECTOR) {
    await fillVisibleDateInputs(page, [process.env.DATE_INPUT_SELECTOR], metricDate);
    return;
  }
  try {
    await openDateRangeDropdown(page);
    await page.waitForTimeout(300);
    await clickFirstText(page, TODAY_TEXTS, "today range button");
    return;
  } catch (todayRangeError) {
    console.warn(
      `Today preset selection failed, trying custom date flow: ${todayRangeError.message}`
    );
  }
  try {
    await selectCustomDateRange(page, metricDate);
    return;
  } catch (customDateError) {
    console.warn(
      `Custom date selection failed, falling back to preset previous-day button: ${customDateError.message}`
    );
  }
  await clickFirstText(page, PREVIOUS_DAY_TEXTS, "previous-day range button");
}

async function selectCommunicationWindow(page) {
  const nativeSelect = page.locator("select").first();
  if (
    (await nativeSelect.count()) > 0 &&
    (await nativeSelect.isVisible().catch(() => false))
  ) {
    await nativeSelect
      .selectOption({ label: COMMUNICATION_WINDOW_NAME })
      .catch(() => {});
  }

  if (process.env.COMM_WINDOW_SELECTOR) {
    await page.locator(process.env.COMM_WINDOW_SELECTOR).first().click();
  } else {
    let opened = false;
    const candidates = [
      ...COMM_WINDOW_LABEL_TEXTS.map((t) => `text=${t}`),
      'input[placeholder*="沟通窗口"]',
      'input[placeholder*="窗口"]',
      ".ant-select-selector",
    ];
    for (const selector of candidates) {
      const locator = page.locator(selector).first();
      if ((await locator.count()) > 0 && (await locator.isVisible().catch(() => false))) {
        await locator.click();
        opened = true;
        break;
      }
    }
    if (!opened) {
      throw new Error(
        "Could not find communication-window filter. Set COMM_WINDOW_SELECTOR."
      );
    }
  }

  if (process.env.COMM_WINDOW_OPTION_SELECTOR) {
    await page.locator(process.env.COMM_WINDOW_OPTION_SELECTOR).first().click();
  } else {
    await page
      .getByText(COMMUNICATION_WINDOW_NAME, { exact: false })
      .first()
      .click({ timeout: 10_000 });
  }
}

async function applyFilters(page) {
  if (process.env.SEARCH_BUTTON_SELECTOR) {
    await page.locator(process.env.SEARCH_BUTTON_SELECTOR).first().click();
    await page.waitForLoadState("networkidle", { timeout: 30_000 }).catch(() => {});
    return;
  }
  const searchButton = page.getByRole("button", {
    name: /查询|搜索|筛选|Search|Apply/i,
  });
  if ((await searchButton.count()) > 0) {
    await searchButton.first().click();
    await page.waitForLoadState("networkidle", { timeout: 30_000 }).catch(() => {});
  }
}

async function exportAndReadRows(page, metricDate) {
  const downloadPromise = page.waitForEvent("download", { timeout: 60_000 });
  if (process.env.EXPORT_BUTTON_SELECTOR) {
    await page.locator(process.env.EXPORT_BUTTON_SELECTOR).first().click();
  } else {
    await clickFirstText(page, EXPORT_BUTTON_TEXTS, "export button");
  }
  const download = await downloadPromise;
  const suggested = download.suggestedFilename() || "export.csv";

  // Keep a copy of the raw export so it can be uploaded as a workflow artifact.
  const ext = suggested.includes(".") ? suggested.slice(suggested.lastIndexOf(".")) : ".csv";
  const savedPath = `export-${metricDate}${ext}`;
  await download.saveAs(savedPath);

  const workbook = XLSX.readFile(savedPath, { raw: false });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, {
    header: 1,
    blankrows: false,
    defval: "",
  });
  return { rows, suggested, savedPath };
}

function hasGoogleCredentials() {
  return Boolean(
    process.env.GOOGLE_SERVICE_ACCOUNT_JSON ||
      process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64
  );
}

async function getSheetsClient() {
  const auth = new GoogleAuth({
    credentials: parseGoogleCredentials(),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return sheets({ version: "v4", auth });
}

async function ensureWorksheet(sheetsClient) {
  const spreadsheet = await sheetsClient.spreadsheets.get({
    spreadsheetId: GOOGLE_SHEET_ID,
  });
  const worksheet = spreadsheet.data.sheets?.find(
    (sheet) => sheet.properties?.title === GOOGLE_WORKSHEET_NAME
  );
  if (worksheet) {
    return;
  }
  await sheetsClient.spreadsheets.batchUpdate({
    spreadsheetId: GOOGLE_SHEET_ID,
    requestBody: {
      requests: [
        { addSheet: { properties: { title: GOOGLE_WORKSHEET_NAME } } },
      ],
    },
  });
}

async function isHeaderPresent(sheetsClient) {
  const response = await sheetsClient.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `'${GOOGLE_WORKSHEET_NAME}'!A1:A1`,
  });
  return Boolean(response.data.values?.length);
}

async function appendRowsToGoogleSheet(rows) {
  if (!rows.length) {
    console.log("Export contained no rows; nothing written to Google Sheet.");
    return 0;
  }
  const sheetsClient = await getSheetsClient();
  await ensureWorksheet(sheetsClient);

  const [header, ...dataRows] = rows;
  const headerPresent = await isHeaderPresent(sheetsClient);
  const valuesToAppend = headerPresent ? dataRows : [header, ...dataRows];

  if (valuesToAppend.length) {
    await sheetsClient.spreadsheets.values.append({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `'${GOOGLE_WORKSHEET_NAME}'!A:A`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: valuesToAppend },
    });
  }
  return dataRows.length;
}

async function saveFailureArtifacts(page) {
  try {
    await page.screenshot({ path: "last-failure.png", fullPage: true });
    const html = await page.content();
    await writeFile("last-failure.html", html, "utf8");
    console.error("Saved last-failure.png and last-failure.html for debugging.");
  } catch (artifactError) {
    console.error("Could not save failure artifacts:", artifactError);
  }
}

async function main() {
  const metricDate = getMetricDate();
  const browser = await chromium.launch({ headless: true });
  let page;
  try {
    const context = await browser.newContext({ acceptDownloads: true });
    page = await context.newPage();

    await loginToMarsMind(page);
    await page.waitForTimeout(2_000);
    await selectPreviousDay(page, metricDate);
    await page.waitForTimeout(500);
    await selectCommunicationWindow(page);
    await page.waitForTimeout(500);
    await applyFilters(page);

    const { rows, suggested, savedPath } = await exportAndReadRows(page, metricDate);
    const dataRowCount = Math.max(rows.length - 1, 0);
    console.log(
      `Exported '${COMMUNICATION_WINDOW_NAME}' chats for ${metricDate} ` +
        `(file: ${suggested}, saved: ${savedPath}, ${dataRowCount} data rows).`
    );

    if (!hasGoogleCredentials()) {
      console.log(
        "No Google credentials set — skipping Google Sheet write. " +
          "The exported file is available as a workflow artifact. " +
          "Add the GOOGLE_SERVICE_ACCOUNT_JSON secret to enable writing to the sheet."
      );
      return;
    }

    const written = await appendRowsToGoogleSheet(rows);
    console.log(
      `Appended ${written} rows to Google Sheet tab '${GOOGLE_WORKSHEET_NAME}'.`
    );
  } catch (error) {
    if (page) {
      await saveFailureArtifacts(page);
    }
    throw error;
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
