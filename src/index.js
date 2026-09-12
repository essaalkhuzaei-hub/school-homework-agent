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

/* =========================================================
   CONFIGURATION
   ========================================================= */

const required = [
  "SCHOOL_LOGIN_URL",
  "SCHOOL_LOGIN_ID",
  "SCHOOL_PASSWORD",
  "OPENAI_API_KEY"
];

for (const key of required) {
  if (!process.env[key]) {
    throw new Error(`Missing required environment variable: ${key}`);
  }
}

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

const MODEL = process.env.OPENAI_MODEL || "gpt-5.6-luna";

const DOWNLOAD_DIR = path.resolve("downloads");
const OUTPUT_DIR = path.resolve("output");

fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const STUDENTS = ["MUNEERA", "MARYAM"];

/* =========================================================
   HELPERS
   ========================================================= */

function safeName(value) {
  return String(value || "homework")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function isVisible(locator) {
  try {
    return await locator.first().isVisible();
  } catch {
    return false;
  }
}

/* =========================================================
   LOGIN
   ========================================================= */

async function fillLogin(page) {
  console.log("Attempting login...");

  const password = page.locator('input[type="password"]').first();

  if (!(await password.count())) {
    throw new Error("Password field was not found.");
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

  if (!login) {
    throw new Error("Login ID field was not found.");
  }

  await login.fill(process.env.SCHOOL_LOGIN_ID);
  await password.fill(process.env.SCHOOL_PASSWORD);

  const signInButton = page
    .getByRole("button", { name: /sign in/i })
    .first();

  if (!(await signInButton.count())) {
    throw new Error("Sign In button was not found.");
  }

  console.log("Clicking Sign In...");
  await signInButton.click();

  await page.waitForTimeout(4000);
  await page
    .waitForLoadState("domcontentloaded", { timeout: 15000 })
    .catch(() => {});

  console.log("URL after login:", page.url());

  const passwordStillVisible = await page
    .locator('input[type="password"]')
    .first()
    .isVisible()
    .catch(() => false);

  if (passwordStillVisible) {
    throw new Error(
      "Login failed: login page is still visible after clicking Sign In."
    );
  }

  console.log("Login successful.");
}

/* =========================================================
   OPEN HOMEWORK PAGE
   ========================================================= */

async function openHomeworkPage(page) {
  console.log("Opening HomeWork Submission from portal menu...");

  const menuCandidates = [
    page.getByText("HomeWork Submission", { exact: true }).first(),
    page.getByText(/HomeWork Submission/i).first(),
    page.getByText(/Homework Submission/i).first()
  ];

  let homeworkMenu = null;

  for (const candidate of menuCandidates) {
    if (await candidate.count()) {
      homeworkMenu = candidate;
      break;
    }
  }

  if (!homeworkMenu) {
    throw new Error("HomeWork Submission menu item was not found.");
  }

  await homeworkMenu.click();

  await page.waitForTimeout(3000);
  await page
    .waitForLoadState("domcontentloaded")
    .catch(() => {});

  console.log("Homework page URL:", page.url());

  const headingVisible = await page
    .getByText("Homeworks", { exact: true })
    .first()
    .isVisible()
    .catch(() => false);

  console.log("Homework heading visible:", headingVisible);

  if (!headingVisible) {
    console.log(
      "Warning: Homeworks heading was not detected, continuing anyway."
    );
  }
}

/* =========================================================
   SELECT STUDENT
   ========================================================= */

async function selectStudent(page, studentName) {
  console.log(`Selecting student: ${studentName}`);

  /*
    The portal shows the currently selected student in the
    upper-left area. Clicking the student/avatar opens My Ward(s).
  */

  const currentWardCandidates = [
    page.locator("aside").getByText(/MARYAM|MUNEERA/i).first(),
    page.getByText(/MARYAM|MUNEERA/i).first()
  ];

  let currentWard = null;

  for (const candidate of currentWardCandidates) {
    if (await candidate.count() && await isVisible(candidate)) {
      currentWard = candidate;
      break;
    }
  }

  if (!currentWard) {
    throw new Error(
      `Could not find the current student selector before selecting ${studentName}.`
    );
  }

  await currentWard.click();
  await page.waitForTimeout(1200);

  const target = page
    .getByText(new RegExp(`^${studentName}$`, "i"), { exact: true })
    .last();

  if (!(await target.count())) {
    throw new Error(
      `Student ${studentName} was not found in My Ward(s).`
    );
  }

  await target.click();

  await page.waitForTimeout(3000);
  await page
    .waitForLoadState("domcontentloaded")
    .catch(() => {});

  console.log(`${studentName} selected.`);

  /*
    Some versions of the portal return to the dashboard after
    changing student, so reopen HomeWork Submission if necessary.
  */

  const homeworkHeading = page
    .getByText("Homeworks", { exact: true })
    .first();

  if (!(await isVisible(homeworkHeading))) {
    await openHomeworkPage(page);
  }

  await page.waitForTimeout(2000);
}

/* =========================================================
   FIND HOMEWORK ROWS
   ========================================================= */

async function getHomeworkRows(page) {
  const rows = page.locator("table tbody tr");
  const count = await rows.count();

  const results = [];

  for (let i = 0; i < count; i++) {
    const row = rows.nth(i);

    const text = (await row.innerText().catch(() => "")).trim();

    if (!text) continue;

    const cells = row.locator("td");
    const cellCount = await cells.count();

    if (cellCount < 2) continue;

    let subject = "";

    /*
      In the school table the subject is normally the second
      meaningful column.
    */

    for (let c = 0; c < cellCount; c++) {
      const value = (await cells.nth(c).innerText().catch(() => "")).trim();

      if (
        value &&
        !/^\d+$/.test(value) &&
        !/^HW\d*$/i.test(value) &&
        !/submit homework/i.test(value) &&
        !/^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/.test(value)
      ) {
        subject = value;
        break;
      }
    }

    if (!subject) {
      subject = `Subject-${i + 1}`;
    }

    results.push({
      row,
      subject: safeName(subject),
      rowText: text,
      index: i
    });
  }

  console.log(`Found ${results.length} homework row(s).`);
  return results;
}

/* =========================================================
   DOWNLOAD ATTACHMENT FROM A HOMEWORK ROW
   ========================================================= */

async function downloadHomework(page, homework, studentName) {
  const { row, subject, index } = homework;

  console.log(
    `Looking for attachment: ${studentName} / ${subject}`
  );

  const studentDir = path.join(
    DOWNLOAD_DIR,
    safeName(studentName),
    safeName(subject)
  );

  fs.mkdirSync(studentDir, { recursive: true });

  /*
    The attachment in this portal is represented by the green
    download icon inside the homework row.
  */

  const clickableCandidates = [
    row.locator('a[download]').first(),
    row.locator('a[href*=".pdf" i]').first(),
    row.locator('a').filter({ has: row.locator("svg") }).first(),
    row.locator('a').filter({ has: row.locator("i") }).first(),
    row.locator('button').filter({ has: row.locator("svg") }).first(),
    row.locator('button').filter({ has: row.locator("i") }).first()
  ];

  let downloadControl = null;

  for (const candidate of clickableCandidates) {
    if (await candidate.count() && await isVisible(candidate)) {
      downloadControl = candidate;
      break;
    }
  }

  /*
    Fallback: inspect all links/buttons in the row and ignore
    "Submit Homework".
  */

  if (!downloadControl) {
    const controls = row.locator("a, button");
    const controlCount = await controls.count();

    for (let i = 0; i < controlCount; i++) {
      const control = controls.nth(i);

      const text = (
        await control.innerText().catch(() => "")
      ).trim();

      if (/submit homework/i.test(text)) continue;

      if (await isVisible(control)) {
        downloadControl = control;
        break;
      }
    }
  }

  if (!downloadControl) {
    console.log(
      `No downloadable attachment found for ${studentName} / ${subject}.`
    );
    return null;
  }

  try {
    const downloadPromise = page.waitForEvent("download", {
      timeout: 15000
    });

    await downloadControl.click();

    const download = await downloadPromise;

    let suggested = download.suggestedFilename();

    if (!suggested.toLowerCase().endsWith(".pdf")) {
      suggested = `${subject}-${index + 1}.pdf`;
    }

    const filePath = path.join(
      studentDir,
      safeName(suggested)
    );

    await download.saveAs(filePath);

    console.log(`Downloaded: ${filePath}`);

    return {
      student: studentName,
      subject,
      filePath
    };
  } catch (error) {
    console.log(
      `Direct download event not detected for ${subject}: ${error.message}`
    );

    /*
      Some school portals open the PDF in another tab instead
      of emitting a normal download event.
    */

    const href = await downloadControl
      .getAttribute("href")
      .catch(() => null);

    if (href && href !== "#" && !href.startsWith("javascript:")) {
      try {
        const absoluteUrl = new URL(href, page.url()).href;

        const response = await page.context().request.get(absoluteUrl);

        if (response.ok()) {
          const body = await response.body();

          const filePath = path.join(
            studentDir,
            `${safeName(subject)}-${index + 1}.pdf`
          );

          fs.writeFileSync(filePath, body);

          console.log(`Downloaded by authenticated request: ${filePath}`);

          return {
            student: studentName,
            subject,
            filePath
          };
        }
      } catch (fallbackError) {
        console.log(
          `Fallback download failed: ${fallbackError.message}`
        );
      }
    }

    return null;
  }
}

/* =========================================================
   READ PDF
   ========================================================= */

async function readPdf(filePath) {
  const buffer = fs.readFileSync(filePath);

  const parsed = await pdf(buffer);

  return (parsed.text || "").trim();
}

/* =========================================================
   SOLVE HOMEWORK
   ========================================================= */

async function solveHomework(student, subject, homeworkText) {
  console.log(`Solving ${student} / ${subject}...`);

  if (!homeworkText) {
    return (
      "The PDF did not contain extractable text. " +
      "The homework may be image-based and requires visual processing."
    );
  }

  const prompt = `
You are helping a school student complete homework.

Student: ${student}
Subject: ${subject}

Read the homework carefully.

Instructions:
- Answer every question.
- Keep the answers appropriate for the student's school level.
- Preserve question numbering.
- Give clear, concise answers.
- If the homework is Arabic, answer in Arabic.
- If the homework is English, answer in English.
- For mathematics, show the necessary working.
- Do not invent questions that are not in the homework.
- Return only the organized homework answers.

HOMEWORK:

${homeworkText}
`;

  const response = await openai.responses.create({
    model: MODEL,
    input: prompt
  });

  return response.output_text || "No answer generated.";
}

/* =========================================================
   CREATE WORD DOCUMENT
   ========================================================= */

async function createSubjectDocument(
  student,
  subject,
  solvedHomeworks
) {
  const children = [
    new Paragraph({
      text: `${student} - ${subject}`,
      heading: HeadingLevel.TITLE
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Student: ${student}`,
          bold: true
        })
      ]
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Subject: ${subject}`,
          bold: true
        })
      ]
    }),
    new Paragraph("")
  ];

  solvedHomeworks.forEach((item, index) => {
    children.push(
      new Paragraph({
        text: `Homework ${index + 1}`,
        heading: HeadingLevel.HEADING_1
      })
    );

    const lines = String(item.answer || "")
      .split(/\r?\n/)
      .filter(line => line.trim());

    for (const line of lines) {
      children.push(
        new Paragraph({
          text: line
        })
      );
    }

    children.push(new Paragraph(""));
  });

  const doc = new Document({
    sections: [
      {
        properties: {},
        children
      }
    ]
  });

  const studentOutputDir = path.join(
    OUTPUT_DIR,
    safeName(student)
  );

  fs.mkdirSync(studentOutputDir, { recursive: true });

  const outputPath = path.join(
    studentOutputDir,
    `${safeName(student)}-${safeName(subject)}.docx`
  );

  const buffer = await Packer.toBuffer(doc);

  fs.writeFileSync(outputPath, buffer);

  console.log(`Word document created: ${outputPath}`);

  return outputPath;
}

/* =========================================================
   PROCESS ONE STUDENT
   ========================================================= */

async function processStudent(page, studentName) {
  console.log("");
  console.log("======================================");
  console.log(`PROCESSING STUDENT: ${studentName}`);
  console.log("======================================");

  await selectStudent(page, studentName);

  await page.waitForTimeout(2500);

  const homeworkRows = await getHomeworkRows(page);

  if (!homeworkRows.length) {
    console.log(`No homework rows found for ${studentName}.`);
    return [];
  }

  const grouped = {};

  for (const homework of homeworkRows) {
    const downloaded = await downloadHomework(
      page,
      homework,
      studentName
    );

    if (!downloaded) {
      continue;
    }

    try {
      const text = await readPdf(downloaded.filePath);

      const answer = await solveHomework(
        studentName,
        downloaded.subject,
        text
      );

      if (!grouped[downloaded.subject]) {
        grouped[downloaded.subject] = [];
      }

      grouped[downloaded.subject].push({
        pdf: downloaded.filePath,
        answer
      });
    } catch (error) {
      console.log(
        `Failed processing ${downloaded.filePath}: ${error.message}`
      );
    }
  }

  const documents = [];

  for (const [subject, items] of Object.entries(grouped)) {
    const docPath = await createSubjectDocument(
      studentName,
      subject,
      items
    );

    documents.push(docPath);
  }

  console.log(
    `${studentName}: created ${documents.length} subject document(s).`
  );

  return documents;
}

/* =========================================================
   MAIN
   ========================================================= */

async function main() {
  console.log("Starting School Homework Agent...");

  const browser = await chromium.launch({
    headless: true
  });

  const context = await browser.newContext({
    acceptDownloads: true
  });

  const page = await context.newPage();

  try {
    console.log("Opening school login page...");

    await page.goto(process.env.SCHOOL_LOGIN_URL, {
      waitUntil: "domcontentloaded",
      timeout: 30000
    });

    await fillLogin(page);

    await page.waitForTimeout(2500);

    /*
      First open HomeWork Submission.
      Student processing will reopen it automatically when needed.
    */

    await openHomeworkPage(page);

    const allDocuments = [];

    /*
      Required order:
      1. Muneera
      2. Maryam
    */

    for (const student of STUDENTS) {
      try {
        const docs = await processStudent(page, student);
        allDocuments.push(...docs);
      } catch (error) {
        console.error(
          `Failed while processing ${student}:`,
          error.message
        );
      }
    }

    console.log("");
    console.log("======================================");
    console.log("HOMEWORK AGENT FINISHED");
    console.log("======================================");

    if (!allDocuments.length) {
      console.log(
        "No homework documents were created. There may be no available homework attachments."
      );
    } else {
      console.log(
        `Created ${allDocuments.length} Word document(s):`
      );

      for (const file of allDocuments) {
        console.log(` - ${file}`);
      }
    }
  } finally {
    await browser.close();
  }
}

main().catch(error => {
  console.error("AGENT FAILED:", error);
  process.exit(1);
});
