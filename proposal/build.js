const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, LevelFormat, TableLayoutType, VerticalAlign,
} = require("docx");

// ---- Theme (marine) ----
const NAVY = "0B2E4F";     // deep navy
const TEAL = "0E7C86";     // teal accent
const STEEL = "2C6E8F";
const LIGHT = "EAF2F5";    // light row
const LIGHTER = "F5F9FB";
const GOLD = "B8860B";
const GREY = "5A6B75";
const FONT = "Tahoma";

// ---- Helpers (RTL Persian) ----
function p(text, opts = {}) {
  return new Paragraph({
    bidirectional: true,
    alignment: opts.align || AlignmentType.RIGHT,
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: opts.line ?? 300 },
    indent: opts.indent,
    children: [new TextRun({
      text, font: FONT, rtl: true,
      size: opts.size ?? 22, bold: opts.bold ?? false,
      color: opts.color ?? "1A1A1A", italics: opts.italics ?? false,
    })],
    ...(opts.border ? { border: opts.border } : {}),
  });
}

function runs(children, opts = {}) {
  return new Paragraph({
    bidirectional: true,
    alignment: opts.align || AlignmentType.RIGHT,
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: opts.line ?? 300 },
    children,
  });
}

function t(text, o = {}) {
  return new TextRun({ text, font: FONT, rtl: true, size: o.size ?? 22, bold: o.bold ?? false, color: o.color ?? "1A1A1A" });
}

function h1(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 260, after: 140, line: 320 },
    border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 12, space: 6 } },
    children: [new TextRun({ text, font: FONT, rtl: true, size: 30, bold: true, color: NAVY })],
  });
}

function h2(text) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100, line: 300 },
    children: [new TextRun({ text, font: FONT, rtl: true, size: 25, bold: true, color: TEAL })],
  });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    numbering: { reference: opts.num || "bullets", level: 0 },
    spacing: { after: 70, line: 290 },
    children: [new TextRun({ text, font: FONT, rtl: true, size: 22, color: opts.color ?? "1A1A1A", bold: opts.bold ?? false })],
  });
}

function spacer(h = 120) {
  return new Paragraph({ spacing: { after: h }, children: [new TextRun("")] });
}

// ---- Table helpers ----
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const thinBorder = (c = "CFE0E7") => ({ style: BorderStyle.SINGLE, size: 4, color: c });

function cell(children, o = {}) {
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 90, right: 90 },
    children: Array.isArray(children) ? children : [children],
  });
}

function cellText(text, o = {}) {
  return cell(
    new Paragraph({
      bidirectional: true,
      alignment: o.align || AlignmentType.RIGHT,
      spacing: { after: 0, line: 280 },
      children: [new TextRun({ text, font: FONT, rtl: true, size: o.size ?? 21, bold: o.bold ?? false, color: o.color ?? "1A1A1A" })],
    }),
    { w: o.w, fill: o.fill }
  );
}

function headerRow(cells, widths) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((c, i) => cellText(c, { w: widths[i], fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.CENTER })),
  });
}

function dataTable(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [headerRow(headers, widths)];
  rows.forEach((r, ri) => {
    const fill = ri % 2 === 0 ? LIGHTER : "FFFFFF";
    trs.push(new TableRow({
      children: r.map((val, ci) => {
        const isBoldCol = o.boldFirst && ci === headers.length - 1; // last (rightmost in RTL) — but we keep simple
        return cellText(String(val), { w: widths[ci], fill, align: o.align?.[ci] || AlignmentType.CENTER, bold: o.boldCols?.includes(ci) });
      }),
    }));
  });
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    visuallyRightToLeft: true,
    borders: {
      top: thinBorder(), bottom: thinBorder(), left: thinBorder(), right: thinBorder(),
      insideHorizontal: thinBorder(), insideVertical: thinBorder(),
    },
    rows: trs,
  });
}

// callout box
function callout(title, lines, fill = LIGHT, bar = TEAL) {
  const kids = [];
  if (title) kids.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 60, line: 290 },
    children: [new TextRun({ text: title, font: FONT, rtl: true, size: 23, bold: true, color: NAVY })],
  }));
  lines.forEach(l => kids.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 40, line: 290 },
    children: [new TextRun({ text: l, font: FONT, rtl: true, size: 21, color: "1A1A1A" })],
  })));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    layout: TableLayoutType.FIXED,
    visuallyRightToLeft: true,
    borders: {
      top: noBorder, bottom: noBorder, left: noBorder, insideHorizontal: noBorder, insideVertical: noBorder,
      right: { style: BorderStyle.SINGLE, size: 24, color: bar },
    },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill },
      margins: { top: 130, bottom: 130, left: 160, right: 160 },
      children: kids,
    })] })],
  });
}

// ================= CONTENT =================
const body = [];

// ---------- COVER ----------
body.push(new Paragraph({ spacing: { before: 900 }, children: [new TextRun("")] }));
body.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 60 },
  children: [new TextRun({ text: "پروپوزال اجرای پروژه", font: FONT, rtl: true, size: 30, bold: true, color: TEAL })],
}));
body.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 40, line: 420 },
  children: [new TextRun({ text: "شفاف‌سازی و یکپارچه‌سازی سیستم اطلاعاتی", font: FONT, rtl: true, size: 46, bold: true, color: NAVY })],
}));
body.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 30, line: 380 },
  children: [new TextRun({ text: "با قدرت اکسپلور و مغایرت‌گیری بین عملیات، انبار و حسابداری", font: FONT, rtl: true, size: 26, bold: false, color: STEEL })],
}));
body.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER, spacing: { before: 40, after: 700 },
  border: { bottom: { color: GOLD, style: BorderStyle.SINGLE, size: 12, space: 10 } },
  children: [new TextRun({ text: "پلتفرم داده، مغایرت‌گیری و هوش تجاری (BI)", font: FONT, rtl: true, size: 22, color: GREY })],
}));

// cover meta table
body.push(new Table({
  width: { size: 8000, type: WidthType.DXA },
  columnWidths: [5400, 2600],
  layout: TableLayoutType.FIXED,
  alignment: AlignmentType.CENTER,
  visuallyRightToLeft: true,
  borders: {
    top: thinBorder("D8E3E8"), bottom: thinBorder("D8E3E8"), left: thinBorder("D8E3E8"), right: thinBorder("D8E3E8"),
    insideHorizontal: thinBorder("D8E3E8"), insideVertical: thinBorder("D8E3E8"),
  },
  rows: [
    ["شرکت صنعتی و خدمات مهندسی ایران", "کارفرما"],
    ["شرکت فناوران فرایند فردا", "مجری"],
    ["تیر ۱۴۰۵", "تاریخ"],
    ["نسخه ۱٫۰", "نسخه سند"],
  ].map((r, i) => new TableRow({ children: [
    cellText(r[0], { w: 5400, fill: i % 2 ? "FFFFFF" : LIGHTER, align: AlignmentType.RIGHT, bold: false }),
    cellText(r[1], { w: 2600, fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.CENTER }),
  ] })),
}));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 1. Executive summary ----------
body.push(h1("۱. خلاصهٔ مدیریتی"));
body.push(p("شرکت صنعتی و خدمات مهندسی ایران در چند شعبه، خدمات مهندسی دریایی از جمله سرویس و بازرسی «لایف‌رفت» (Life Raft / قایق نجات بادی) را ارائه می‌کند. در این فرآیند، هر عملیات در «پلتفرم عملیات» ثبت می‌شود، اقلام تعویض‌شده از «نرم‌افزار انبار» کسر می‌گردد و هزینه‌ها و درآمدهای مرتبط در «نرم‌افزار حسابداری» به ثبت می‌رسد."));
body.push(p("مشکل اصلی امروز این است که این سه سیستم داده‌های خود را به‌صورت مجزا و ناهماهنگ نگهداری می‌کنند؛ در نتیجه مدیریت هیچ تصویر واحد و قابل‌اعتمادی ندارد که نشان دهد آیا موجودی و مصرف انبار با عملیات ثبت‌شده و اسناد حسابداری هم‌خوانی دارد یا خیر، و وضعیت واقعی مالی به چه شکل است."));
body.push(spacer(60));
body.push(callout("هدف پروژه", [
  "ایجاد یک «منبع واحد حقیقت» که سه سیستم عملیات، انبار و حسابداری را یکپارچه می‌کند، مغایرت‌ها را به‌صورت خودکار آشکار می‌سازد و امکان اکسپلور بسیار بالای داده تا سطح هر تراکنش و هر شعبه را فراهم می‌آورد.",
], LIGHT, TEAL));
body.push(spacer(80));
body.push(p("این پروژه در سه فاز و مجموعاً ۴۵ روز اجرا می‌شود و پس از تحویل، تا یک ماه پشتیبانی رایگان ارائه می‌گردد.", { bold: true, color: NAVY }));
body.push(spacer(40));
body.push(dataTable(
  ["عنوان", "شرح"],
  [
    ["مدت کل اجرا", "۴۵ روز (فاز ۱: ۱۰ روز، فاز ۲: ۲۰ روز، فاز ۳: ۱۵ روز)"],
    ["مبلغ کل قرارداد اجرا", "۳۴۴٬۵۰۰٬۰۰۰ تومان"],
    ["پشتیبانی پس از تحویل", "تا یک ماه رایگان"],
    ["خروجی نهایی", "پلتفرم BI با داشبوردهای عملیاتی، انباری و مالی + گزارش مغایرت‌گیری خودکار"],
  ],
  [3400, 5960],
  { align: [AlignmentType.RIGHT, AlignmentType.RIGHT], boldCols: [] }
));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 2. Current state ----------
body.push(h1("۲. درک مسئله و وضعیت فعلی"));
body.push(p("در وضعیت کنونی، سه سامانهٔ مستقل، چرخهٔ کسب‌وکار را پوشش می‌دهند اما با یکدیگر «حرف نمی‌زنند»:"));
body.push(spacer(40));
body.push(dataTable(
  ["سامانه", "نقش در چرخهٔ کار", "دادهٔ کلیدی تولیدشده"],
  [
    ["پلتفرم عملیات", "ثبت دستور کار سرویس، تجهیزات، صورت‌حساب و هزینه/تنخواه", "دستور کار (LR/LB/LS/FF/ND/UT/CI-LT/SE/PR)، اقلام مصرفی، صورت‌حساب، هزینه، شعبه"],
    ["نرم‌افزار انبار", "کسر اقلام تعویض‌شده از موجودی", "کاردکس، ورود/خروج کالا، موجودی هر انبار"],
    ["نرم‌افزار حسابداری", "ثبت رسمی هزینه‌ها و درآمدها در دفاتر مالی", "اسناد و دفاتر مالی، درآمد، بهای تمام‌شده، سود/زیان"],
  ],
  [2200, 3260, 3900],
  { align: [AlignmentType.RIGHT, AlignmentType.RIGHT, AlignmentType.RIGHT] }
));
body.push(spacer(60));
body.push(p("خطوط سرویسِ ثبت‌شده در ماژول دستور کار: لایف‌رفت (LR)، لایف‌بوت (LB)، تجهیزات نجات جان (LS)، سیستم‌های اطفاء حریق (FF)، آزمایش‌های غیرمخرب (ND)، ضخامت‌سنجی (UT)، آزمایش‌های بارگذاری (CI/LT)، فروش قطعات (SE) و پروژه‌ها (PR).", { size: 20, color: GREY }));
body.push(spacer(120));
body.push(h2("۲٫۱. چرا این ناهماهنگی خطرناک است؟"));
body.push(p("یک عملیات واحد، هم‌زمان در هر سه سامانه ردپا دارد؛ اما چون کلید مشترک و متناظرسازی وجود ندارد، هیچ‌کس نمی‌تواند با اطمینان بگوید سه روایت با هم می‌خوانند یا نه. برای مثال:"));
body.push(bullet("عملیاتی که یک قلم را تعویض کرده، ممکن است در انبار کسر نشده باشد (نشتی موجودی)."));
body.push(bullet("قلمی که از انبار خارج شده، ممکن است سند هزینه/درآمد متناظر در حسابداری نداشته باشد (نشتی مالی)."));
body.push(bullet("درآمد ثبت‌شده در حسابداری ممکن است به عملیات مشخصی قابل‌انتساب نباشد (عدم ردیابی)."));
body.push(bullet("مقایسهٔ عملکرد و سودآوری بین شعب عملاً ناممکن است."));

body.push(spacer(120));
body.push(h2("۲٫۲. چالش‌های یکپارچه‌سازی"));
body.push(p("این‌ها همان موانعی هستند که تا امروز مانع خواندنِ خودکارِ سه سامانه شده‌اند و پروژه دقیقاً برای حل آن‌ها طراحی شده است:"));
body.push(spacer(40));
body.push(dataTable(
  ["چالش", "شرح"],
  [
    ["نبود کلید مشترک", "کد کالا در انبار، قلم در دستور کار و تفصیلی در حسابداری شناسهٔ مشترک ندارند."],
    ["داده در چند شعبه", "هر شعبه داده و گاهی نمونهٔ نرم‌افزار جداگانه دارد؛ نیازمند تجمیع و یکدست‌سازی."],
    ["اختلاف زمانی ثبت", "ثبت عملیات، خروج انبار و سند مالی هم‌زمان نیستند؛ تطبیق باید بازهٔ زمانی را لحاظ کند."],
    ["رکوردهای ناقص و تکراری", "صورت‌حساب بدون سند مالی، خروج انبار بدون دستور کار، هزینهٔ تنخواه بدون سند."],
    ["اقلام سریال‌دار و انقضادار", "پیروتکنیک، جیره و داروی لایف‌رفت با سریال و تاریخ انقضا، ردیابی ویژه می‌خواهند."],
    ["۹ خط سرویس با منطق متفاوت", "هر خط (LR/LB/FF/ND…) اقلام، چرخه و قواعد مصرف متفاوت دارد."],
    ["وضعیت و تأیید چندمرحله‌ای", "تنخواه و صورت‌حساب چند وضعیت و چند سطح تأیید دارند که باید در تطبیق لحاظ شوند."],
    ["مالیات و سامانه مؤدیان", "صورت‌حساب‌ها باید با اظهار مالیاتی و دفاتر رسمی هم بخوانند."],
  ],
  [2900, 6460],
  { align: [AlignmentType.RIGHT, AlignmentType.RIGHT], boldCols: [0] }
));

body.push(spacer(120));
body.push(callout("درد اصلی کارفرما — به بیان خودشان", [
  "«اصلاً نمی‌دانیم آیا انبار بر اساس عملیات و حسابداری درست می‌خواند یا نه، و وضعیت حسابداری به چه صورت است.»",
], "FDF3E7", GOLD));
body.push(spacer(60));
body.push(p("این پروپوزال دقیقاً همین سؤال را به یک پاسخ عددی، خودکار و لحظه‌ای تبدیل می‌کند."));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 3. Solution ----------
body.push(h1("۳. راه‌حل پیشنهادی — معماری کلی"));
body.push(p("راهکار، ساخت یک لایهٔ دادهٔ یکپارچه در بالای سه سامانهٔ موجود است، بدون آنکه لازم باشد سیستم‌های فعلی تعویض شوند. جریان کار به‌این‌صورت است:"));
body.push(spacer(40));
body.push(dataTable(
  ["لایه", "کارکرد"],
  [
    ["منابع داده", "سه سامانهٔ عملیات، انبار و حسابداری در همهٔ شعب"],
    ["انبار داده (Data Warehouse)", "گردآوری، پاک‌سازی و یکپارچه‌سازی داده‌ها در یک مخزن مرکزی"],
    ["مایگریشن و متناظرسازی", "ایجاد کلید مشترک بین عملیات ⇄ انبار ⇄ حسابداری و کشف مغایرت‌ها"],
    ["پلتفرم BI و داشبورد", "گزارش‌ها، داشبوردهای مدیریتی و قابلیت اکسپلور نامحدود"],
  ],
  [3400, 5960],
  { align: [AlignmentType.RIGHT, AlignmentType.RIGHT] }
));
body.push(spacer(120));
body.push(p("مزیت این معماری آن است که سیستم‌های عملیاتی فعلی دست‌نخورده باقی می‌مانند و لایهٔ داده به‌صورت موازی و بدون اختلال، شفافیت را فراهم می‌کند.", { bold: true, color: NAVY }));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 4. Phasing ----------
body.push(h1("۴. فازبندی اجرای پروژه"));

// Phase 1
body.push(h2("فاز اول — ساخت انبار داده (Data Warehouse) | ۱۰ روز"));
body.push(p("در این فاز، دادهٔ سه سامانه از همهٔ شعب استخراج، پاک‌سازی و در یک انبار دادهٔ مرکزی و ساخت‌یافته گردآوری می‌شود تا زیربنای یکپارچه‌سازی فراهم گردد."));
body.push(bullet("شناسایی و مستندسازی منابع داده و ساختار هر سه سامانه در تمام شعب."));
body.push(bullet("طراحی مدل دادهٔ انبار داده متناسب با چرخهٔ سرویس لایف‌رفت."));
body.push(bullet("پیاده‌سازی فرآیند استخراج و بارگذاری داده (ETL) از سه سیستم."));
body.push(bullet("پاک‌سازی، استانداردسازی و یکدست‌سازی داده‌ها (تاریخ، کد کالا، کد شعبه، مشتری)."));
body.push(bullet("راه‌اندازی انبار دادهٔ مرکزی به‌عنوان منبع واحد حقیقت."));
body.push(runs([t("خروجی فاز: ", { bold: true, color: TEAL }), t("انبار دادهٔ عملیاتیِ یکپارچه از هر سه سامانه و همهٔ شعب.")], { after: 60 }));

body.push(spacer(80));
// Phase 2
body.push(h2("فاز دوم — مایگریشن و متناظرسازی (Migration & Mapping) | ۲۰ روز"));
body.push(p("در این فاز، رکوردهای سه سامانه به یکدیگر «متناظر» می‌شوند تا برای هر عملیات، روایت انبار و روایت حسابداری کنار هم قرار گیرد و مغایرت‌ها آشکار شوند."));
body.push(bullet("تعریف کلید مشترک بین دستور کار، خروج انبار و اسناد مالی."));
body.push(bullet("پیاده‌سازی موتور مغایرت‌گیری سه‌محوره: اقلام مصرفی ⇄ انبار، صورت‌حساب/درآمد ⇄ حسابداری، هزینه و تنخواه ⇄ حسابداری."));
body.push(bullet("کشف و دسته‌بندی مغایرت‌ها: قلم کسرنشده، سند مالی مفقود، درآمد بدون دستور کار، مغایرت تنخواه، خطای ثبت."));
body.push(bullet("مایگریشن و تطبیق دادهٔ تاریخی برای تحلیل روند."));
body.push(bullet("اعتبارسنجی نتایج با نمونه‌گیری و تأیید کارفرما."));
body.push(runs([t("خروجی فاز: ", { bold: true, color: TEAL }), t("گزارش مغایرت‌گیری خودکار که دقیقاً نشان می‌دهد کجا انبار با عملیات و حسابداری نمی‌خواند.")], { after: 60 }));

body.push(spacer(80));
// Phase 3
body.push(h2("فاز سوم — پلتفرم BI و داشبوردها | ۱۵ روز"));
body.push(p("در این فاز، بر پایهٔ دادهٔ یکپارچه و متناظرشده، پلتفرم هوش تجاری با داشبوردهای مدیریتی و قابلیت اکسپلور نامحدود مستقر می‌شود."));
body.push(bullet("راه‌اندازی پلتفرم BI و اتصال آن به انبار داده."));
body.push(bullet("داشبورد عملیاتی: حجم عملیات، اقلام مصرفی، عملکرد هر شعبه."));
body.push(bullet("داشبورد انبار: موجودی، گردش کالا، هشدار نقطهٔ سفارش، مغایرت موجودی."));
body.push(bullet("داشبورد مالی: درآمد، بهای تمام‌شده، سود/زیان به تفکیک عملیات و شعبه."));
body.push(bullet("داشبورد مغایرت‌گیری: پایش لحظه‌ای هم‌خوانی سه سامانه."));
body.push(bullet("قابلیت drill-down تا سطح هر تراکنش و فیلترگذاری آزاد (اکسپلور نامحدود)."));
body.push(bullet("آموزش کاربران و تحویل مستندات."));
body.push(runs([t("خروجی فاز: ", { bold: true, color: TEAL }), t("پلتفرم BI کامل و عملیاتی با داشبوردهای مدیریتی، انباری، مالی و مغایرت‌گیری.")], { after: 60 }));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 5. Benefits ----------
body.push(h1("۵. مزایای کلیدی پروژه"));
body.push(p("مهم‌ترین دستاورد این پروژه، تبدیل ابهام به شفافیت عددی و لحظه‌ای است. مزایای اصلی عبارت‌اند از:"));
body.push(spacer(40));

const benefits = [
  ["پاسخ قطعی به پرسش اصلی مدیریت", "مشخص می‌شود آیا انبار با عملیات و حسابداری می‌خواند یا نه — به‌صورت خودکار و با عدد، نه حدس."],
  ["شفافیت کامل چرخهٔ سرویس", "ردیابی هر لایف‌رفت از ثبت عملیات تا کسر از انبار تا ثبت مالی، در یک نگاه."],
  ["کشف نشتی و مغایرت", "شناسایی اقلام کسرنشده، اسناد مالی مفقود، پرت، خطای ثبت و درآمدهای بی‌انتساب."],
  ["منبع واحد حقیقت", "یکپارچگی دادهٔ سه سامانهٔ مجزا در یک مخزن قابل‌اعتماد و بدون تناقض."],
  ["دید یکپارچه بین شعب", "مقایسهٔ عملکرد، مصرف و سودآوری همهٔ شعب روی یک معیار واحد."],
  ["سودآوری واقعی", "محاسبهٔ سود/زیان واقعی هر عملیات، هر نوع سرویس و هر شعبه."],
  ["تصمیم‌گیری داده‌محور", "داشبوردهای مدیریتی لحظه‌ای به‌جای گزارش‌های دستی و دیرهنگام."],
  ["قدرت اکسپلور بسیار بالا", "کاوش آزاد داده و drill-down تا سطح تک‌تراکنش برای پاسخ به هر پرسش تازه."],
  ["حسابرسی و ردیابی (Audit Trail)", "امکان پیگیری کامل مسیر هر رکورد برای کنترل داخلی و حسابرسی."],
  ["کاهش هزینه‌های پنهان", "جلوگیری از دوباره‌کاری، مغایرت‌های مالی و از دست رفتن درآمد."],
  ["مقیاس‌پذیری", "افزودن آسان شعب جدید یا سامانه‌های جدید به لایهٔ دادهٔ یکپارچه."],
];
body.push(dataTable(
  ["مزیت", "توضیح"],
  benefits,
  [2900, 6460],
  { align: [AlignmentType.RIGHT, AlignmentType.RIGHT], boldCols: [0] }
));

body.push(spacer(140));
body.push(callout("جمله‌ای که پس از این پروژه می‌توانید بگویید", [
  "«موجودی انبار، عملیات ثبت‌شده و اسناد حسابداری ما با هم می‌خوانند — و هرجا نخوانند، سیستم بلافاصله و با جزئیات به ما نشان می‌دهد.»",
], LIGHT, TEAL));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 6. Deliverables ----------
body.push(h1("۶. خروجی‌ها و تحویل‌شدنی‌ها"));
body.push(dataTable(
  ["فاز", "تحویل‌شدنی اصلی"],
  [
    ["فاز ۱", "انبار دادهٔ مرکزی یکپارچه + مستندات مدل داده و فرآیند ETL"],
    ["فاز ۲", "موتور و گزارش مغایرت‌گیری سه‌جانبهٔ خودکار + دادهٔ تاریخی متناظرشده"],
    ["فاز ۳", "پلتفرم BI با داشبوردهای عملیاتی، انباری، مالی و مغایرت‌گیری + آموزش کاربران"],
    ["پشتیبانی", "تا یک ماه پشتیبانی رایگان پس از تحویل فاز سوم"],
  ],
  [1800, 7560],
  { align: [AlignmentType.CENTER, AlignmentType.RIGHT], boldCols: [0] }
));

body.push(spacer(140));
// ---------- 7. Timeline ----------
body.push(h1("۷. زمان‌بندی کلی"));
body.push(dataTable(
  ["فاز", "عنوان", "مدت", "بازهٔ تجمعی"],
  [
    ["فاز ۱", "ساخت انبار داده", "۱۰ روز", "روز ۱ تا ۱۰"],
    ["فاز ۲", "مایگریشن و متناظرسازی", "۲۰ روز", "روز ۱۱ تا ۳۰"],
    ["فاز ۳", "پلتفرم BI و داشبوردها", "۱۵ روز", "روز ۳۱ تا ۴۵"],
    ["جمع", "کل اجرای پروژه", "۴۵ روز", "—"],
  ],
  [1500, 4160, 1900, 1800],
  { align: [AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.CENTER, AlignmentType.CENTER], boldCols: [0] }
));

body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 8. Support ----------
body.push(h1("۸. پشتیبانی"));
body.push(p("پس از تحویل فاز سوم، تا یک ماه پشتیبانیِ رایگان ارائه می‌شود تا بهره‌برداری اولیه به‌درستی انجام گیرد."));
body.push(spacer(40));
body.push(dataTable(
  ["عنوان", "شرح"],
  [
    ["نوع خدمت", "پشتیبانی رایگان پس از تحویل"],
    ["مدت", "تا یک ماه پس از تحویل فاز سوم — رایگان"],
    ["پوشش", "رفع اشکال، تنظیم و اصلاح داشبوردها، پاسخ‌گویی فنی و همراهی در بهره‌برداری اولیه"],
  ],
  [2900, 6460],
  { align: [AlignmentType.RIGHT, AlignmentType.RIGHT] }
));

body.push(spacer(160));
// ---------- 9. Financials ----------
body.push(h1("۹. برآورد مالی و شرایط قرارداد"));
body.push(dataTable(
  ["ردیف", "شرح", "مبلغ (تومان)"],
  [
    ["۱", "اجرای فاز اول — ساخت انبار داده (۱۰ روز)", "۸۴٬۵۰۰٬۰۰۰"],
    ["۲", "اجرای فاز دوم — مایگریشن و متناظرسازی (۲۰ روز)", "۱۴۳٬۰۰۰٬۰۰۰"],
    ["۳", "اجرای فاز سوم — پلتفرم BI و داشبوردها (۱۵ روز)", "۱۱۷٬۰۰۰٬۰۰۰"],
    ["", "جمع کل قرارداد اجرا (سه فاز)", "۳۴۴٬۵۰۰٬۰۰۰"],
    ["", "پشتیبانی تا یک ماه پس از تحویل", "رایگان"],
  ],
  [1100, 6060, 2200],
  { align: [AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.CENTER], boldCols: [] }
));
body.push(spacer(60));
body.push(p("پرداخت مبلغ هر فاز متناسب با تحویل همان فاز انجام می‌شود؛ مبلغ کل قرارداد اجرا ۳۴۴٫۵ میلیون تومان است.", { size: 20, color: GREY }));
body.push(spacer(60));
body.push(h2("۹٫۱. شرایط عمومی پیشنهادی"));
body.push(bullet("پرداخت مبلغ اجرا در سه مرحله و متناسب با تحویل هر فاز."));
body.push(bullet("همکاری کارفرما در تأمین دسترسی به سه سامانه و دادهٔ شعب."));
body.push(bullet("تا یک ماه پس از تحویل فاز سوم، پشتیبانی به‌صورت رایگان ارائه می‌شود."));
body.push(bullet("مبالغ فوق مطابق قوانین جاری مشمول مالیات بر ارزش افزوده خواهد بود (در صورت شمول)."));

body.push(spacer(160));
// ---------- 10. Why us / closing ----------
body.push(h1("۱۰. جمع‌بندی و گام بعدی"));
body.push(p("این پروپوزال یک مسیر روشن و مرحله‌ای برای عبور از وضعیت «سه سامانهٔ ناهماهنگ» به وضعیت «یک تصویر واحد، شفاف و قابل‌کاوش» ارائه می‌کند. در پایان ۴۵ روز، مدیریت شرکت صنعتی و خدمات مهندسی ایران خواهد توانست هم‌خوانی انبار، عملیات و حسابداری را به‌صورت لحظه‌ای پایش کند و برای هر پرسش تازه، خودش داده را کاوش کند."));
body.push(spacer(60));
body.push(callout("گام بعدی", [
  "۱) تأیید کلیات این پروپوزال و برنامهٔ فازبندی.",
  "۲) برگزاری جلسهٔ فنی برای دسترسی به سه سامانه و شروع فاز اول.",
  "۳) عقد قرارداد و آغاز اجرا.",
], LIGHT, TEAL));
body.push(spacer(120));
body.push(p("با احترام،", { align: AlignmentType.RIGHT }));
body.push(p("شرکت فناوران فرایند فردا", { bold: true, color: NAVY }));
body.push(p("[محل درج لوگو، مهر و امضای مجاز]", { color: GREY }));

// ================= DOCUMENT =================
const doc = new Document({
  creator: "Proposal",
  title: "پروپوزال شفاف‌سازی و یکپارچه‌سازی سیستم اطلاعاتی",
  styles: {
    default: {
      document: { run: { font: FONT, size: 22 } },
    },
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.RIGHT,
        style: { run: { color: TEAL }, paragraph: { indent: { right: 360, hanging: 200 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1100, bottom: 1100, left: 1200, right: 1200 },
      },
      bidi: true,
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/home/user/NewT/proposal/proposal.docx", buf);
  console.log("WROTE proposal.docx", buf.length, "bytes");
});
