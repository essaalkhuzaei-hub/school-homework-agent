
import fs from "fs";
import path from "path";
import { chromium } from "playwright";
import pdf from "pdf-parse";
import OpenAI from "openai";
import {
  Document,
  Packer,
  Paragraph,
  HeadingLevel,
  TextRun
} from "docx";

const required = [
  "SCHOOL_LOGIN_URL",
  "SCHOOL_HOMEWORK_URL",
  "SCHOOL_LOGIN_ID",
  "SCHOOL_PASSWORD",
  "OPENAI_API_KEY"
];

for (const key of required) {
  if (!process.env[key]) {
    throw new Error(`Missing required environment variable: ${key}`);
  }
}

const DOWNLOAD_DIR = path.resolve("downloads");
const OUTPUT_DIR = path.resolve("output");
fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
const model = process.env.OPENAI_MODEL || "gpt-5.6-luna";

function safeName(value) {
  return (value || "homework")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

async function fillLogin(page) {
  const password = page.locator('input[type="password"]').first();
  if (!(await password.count())) {
    throw new Error("Password field was not found. Login page selectors need adjustment.");
  }

  const loginCandidates = [
    page.locator('input[name*="login" i]').first(),
    page.locator('input[name*="user" i]').first(),
    page.locator('input[id*="login" i]').first(),
    page.locator('input[id*="user" i]').first(),
    page.locator('input[type="text"]').first()
  ];

  let login = null;
  for (const candidate of loginCandidates) {
    if (await candidate.count()) {
      login = candidate;
      break;
    }
  }
  if (!login) throw new Error("Login ID field was not found.");

  await login.fill(process.env.SCHOOL_LOGIN_ID);
  await password.fill(process.env.SCHOOL_PASSWORD);

    const signInButton = page.getByRole("button", { name: /sign in/i }).first();

  if (!(await signInButton.count())) {
    throw new Error("Sign In button was not found.");
  }

  console.log("Clicking Sign In...");
  await signInButton.click();

  // Give the school portal time to complete authentication.
  await page.waitForTimeout(3000);
  await page.waitForLoadState("domcontentloaded").catch(() => {});

  console.log("URL after login:", page.url());

  // Verify that the login form has actually disappeared.
  const passwordStillVisible = await page
    .locator('input[type="password"]')
    .first()
    .isVisible()
    .catch(() => false);

  if (passwordStillVisible) {
    throw new Error(
      "Login failed: the school login page is still visible after clicking Sign In."
    );
  }

  console.log("Login successful.");
}

async function discoverPdfLinks(page) {
  return await page.locator("a").evaluateAll((anchors) =>
    anchors
      .map((a) => ({
        href: a.href || "",
        text: (a.innerText || a.textContent || "").trim()
      }))
      .filter((x) =>
        /\.pdf(?:$|\?)/i.test(x.href) ||
        /pdf|homework|worksheet|assignment/i.test(x.text)
      )
  );
}

async function downloadPdf(context, item, index) {
  const response = await context.request.get(item.href);
  if (!response.ok()) {
    throw new Error(`Download failed (${response.status()}): ${item.href}`);
  }
  const body = await response.body();
  const contentType = response.headers()["content-type"] || "";
  if (!contentType.toLowerCase().includes("pdf") && !item.href.toLowerCase().includes(".pdf")) {
    console.warn(`Skipping non-PDF candidate: ${item.text || item.href}`);
    return null;
  }
  const filename = `${String(index + 1).padStart(2, "0")}-${safeName(item.text || "homework")}.pdf`;
  const filepath = path.join(DOWNLOAD_DIR, filename);
  fs.writeFileSync(filepath, body);
  return filepath;
}

async function extractPdfText(filepath) {
  const data = await pdf(fs.readFileSync(filepath));
  return (data.text || "").trim();
}

async function solveHomework(label, text) {
  if (!text) {
    return "تعذر استخراج نص واضح من ملف PDF. يحتاج الملف إلى مراجعة يدوية أو معالجة صور.";
  }

  const response = await openai.responses.create({
    model,
    input: [
      {
        role: "system",
        content:
          "You are a careful school homework assistant. Solve only from the supplied homework. Preserve question numbering and sections. Give clear student-ready answers. If a question is unclear or depends on an unreadable image, explicitly say it needs review rather than inventing an answer. Match the language of each question."
      },
      {
        role: "user",
        content: `Homework file: ${label}\n\n${text}`
      }
    ]
  });

  return response.output_text || "No answer was produced.";
}

async function makeDocx(results) {
  const children = [
    new Paragraph({
      text: "Weekly Homework Answers",
      heading: HeadingLevel.TITLE
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Generated: ${new Date().toLocaleString("en-GB", { timeZone: "Asia/Qatar" })}`,
          italics: true
        })
      ]
    })
  ];

  for (const result of results) {
    children.push(
      new Paragraph({ text: result.label, heading: HeadingLevel.HEADING_1 })
    );
    for (const line of result.answer.split(/\r?\n/)) {
      children.push(new Paragraph({ text: line || " " }));
    }
  }

  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  const stamp = new Date().toISOString().slice(0, 10);
  const out = path.join(OUTPUT_DIR, `homework-answers-${stamp}.docx`);
  fs.writeFileSync(out, buffer);
  return out;
}

async function main() {
  console.log("Starting School Homework Agent...");
  const browser = await chromium.launch({ headless: true });

  try {
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    console.log("Opening school login page...");
    await page.goto(process.env.SCHOOL_LOGIN_URL, {
      waitUntil: "domcontentloaded",
      timeout: 30000
    });

    console.log("Attempting login...");
    await fillLogin(page);

    console.log("Opening HomeWork Submission from portal menu...");

// Open the side menu if HomeWork Submission is not already visible.
let homeworkMenu = page.getByText("HomeWork Submission", { exact: true }).first();

if (!(await homeworkMenu.isVisible().catch(() => false))) {
  const menuButton = page.locator(
    'button:has-text("☰"), .navbar-toggler, .menu-toggle, [class*="menu"], [class*="hamburger"]'
  ).first();

  if (await menuButton.count()) {
    await menuButton.click().catch(() => {});
    await page.waitForTimeout(1000);
  }
}

homeworkMenu = page.getByText("HomeWork Submission", { exact: true }).first();

if (!(await homeworkMenu.count())) {
  throw new Error("HomeWork Submission menu item was not found.");
}

console.log("Clicking HomeWork Submission...");
await homeworkMenu.click();

await page.waitForTimeout(3000);
await page.waitForLoadState("domcontentloaded").catch(() => {});

console.log("Homework page URL:", page.url());
console.log("Homework heading visible:",
  await page.getByText("Homeworks", { exact: true }).first()
    .isVisible()
    .catch(() => false)
);
    await page.waitForTimeout(2500);

    const links = await discoverPdfLinks(page);
    console.log(`Found ${links.length} PDF/link candidate(s).`);

    if (!links.length) {
      console.log("Page title:", await page.title());
      console.log("Current URL:", page.url());
      throw new Error(
        "No homework PDF links were detected. The first Render log will help us adjust the portal-specific selectors."
      );
    }

    const results = [];
    for (let i = 0; i < links.length; i++) {
      const item = links[i];
      console.log(`Processing ${i + 1}/${links.length}: ${item.text || item.href}`);
      try {
        const filepath = await downloadPdf(context, item, i);
        if (!filepath) continue;
        const text = await extractPdfText(filepath);
        const answer = await solveHomework(path.basename(filepath), text);
        results.push({
          label: item.text || path.basename(filepath),
          answer
        });
      } catch (err) {
        console.error(`Failed item ${i + 1}:`, err.message);
      }
    }

    if (!results.length) {
      throw new Error("No homework PDFs were successfully processed.");
    }

    const output = await makeDocx(results);
    console.log(`SUCCESS: Word document created at ${output}`);
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error("AGENT FAILED:", err);
  process.exit(1);
});
