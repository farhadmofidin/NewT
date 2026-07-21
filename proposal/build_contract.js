const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, LevelFormat, TableLayoutType, VerticalAlign,
} = require("docx");

const NAVY = "0B2E4F", TEAL = "0E7C86", STEEL = "2C6E8F";
const LIGHTER = "F5F9FB", GREY = "5A6B75", GOLD = "B8860B";
const FONT = "Tahoma";

function p(text, o = {}) {
  return new Paragraph({
    bidirectional: true, alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { after: o.after ?? 120, before: o.before ?? 0, line: o.line ?? 300 },
    indent: o.indent,
    children: [new TextRun({ text, font: FONT, rtl: true, size: o.size ?? 22, bold: o.bold ?? false, color: o.color ?? "1A1A1A" })],
  });
}
function article(num, title) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT, heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 110, line: 320 },
    border: { bottom: { color: TEAL, style: BorderStyle.SINGLE, size: 8, space: 5 } },
    children: [new TextRun({ text: `مادهٔ ${num} — ${title}`, font: FONT, rtl: true, size: 26, bold: true, color: NAVY })],
  });
}
function bullet(text, o = {}) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.JUSTIFIED,
    numbering: { reference: o.num || "b", level: 0 },
    spacing: { after: 70, line: 290 },
    children: [new TextRun({ text, font: FONT, rtl: true, size: 21, color: "1A1A1A", bold: o.bold ?? false })],
  });
}
function clause(n, text, o = {}) {
  return new Paragraph({
    bidirectional: true, alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 80, line: 300 },
    indent: { right: 360, hanging: 360 },
    children: [
      new TextRun({ text: n + "‏ ", font: FONT, rtl: true, size: 21, bold: true, color: TEAL }),
      new TextRun({ text, font: FONT, rtl: true, size: 21, color: "1A1A1A", bold: o.bold ?? false }),
    ],
  });
}
function spacer(h = 120) { return new Paragraph({ spacing: { after: h }, children: [new TextRun("")] }); }

const thin = (c = "CFE0E7") => ({ style: BorderStyle.SINGLE, size: 4, color: c });
function cellText(text, o = {}) {
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 90, right: 90 },
    children: [new Paragraph({
      bidirectional: true, alignment: o.align || AlignmentType.RIGHT, spacing: { after: 0, line: 270 },
      children: [new TextRun({ text, font: FONT, rtl: true, size: o.size ?? 20, bold: o.bold ?? false, color: o.color ?? "1A1A1A" })],
    })],
  });
}
function dataTable(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [new TableRow({ tableHeader: true, children: headers.map((c, i) => cellText(c, { w: widths[i], fill: NAVY, color: "FFFFFF", bold: true, align: AlignmentType.CENTER })) })];
  rows.forEach((r, ri) => {
    const isTotal = o.totalRow && ri === rows.length - 1;
    const fill = isTotal ? "F1E3C6" : (ri % 2 === 0 ? LIGHTER : "FFFFFF");
    trs.push(new TableRow({ children: r.map((v, ci) => cellText(String(v), { w: widths[ci], fill, align: o.align?.[ci] || AlignmentType.CENTER, bold: isTotal || o.boldCols?.includes(ci) })) }));
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    layout: TableLayoutType.FIXED, visuallyRightToLeft: true,
    borders: { top: thin(), bottom: thin(), left: thin(), right: thin(), insideHorizontal: thin(), insideVertical: thin() },
    rows: trs,
  });
}

const body = [];

// ---------- Title ----------
body.push(new Paragraph({ spacing: { before: 300, after: 40 }, alignment: AlignmentType.CENTER, bidirectional: true,
  children: [new TextRun({ text: "بسمه تعالی", font: FONT, rtl: true, size: 22, color: GREY })] }));
body.push(new Paragraph({ alignment: AlignmentType.CENTER, bidirectional: true, spacing: { before: 60, after: 30, line: 420 },
  children: [new TextRun({ text: "قرارداد اجرای پروژهٔ شفاف‌سازی و یکپارچه‌سازی سیستم اطلاعاتی", font: FONT, rtl: true, size: 34, bold: true, color: NAVY })] }));
body.push(new Paragraph({ alignment: AlignmentType.CENTER, bidirectional: true, spacing: { after: 20 },
  children: [new TextRun({ text: "(رمزگشایی دیتابیس، انبار داده، مایگریشن و متناظرسازی، و پلتفرم هوش تجاری)", font: FONT, rtl: true, size: 22, color: STEEL })] }));
body.push(new Paragraph({ alignment: AlignmentType.CENTER, bidirectional: true, spacing: { before: 30, after: 200 },
  border: { bottom: { color: GOLD, style: BorderStyle.SINGLE, size: 10, space: 8 } },
  children: [new TextRun({ text: "شمارهٔ قرارداد: ...................        تاریخ: ......../......../۱۴۰۴", font: FONT, rtl: true, size: 20, color: GREY })] }));

// ---------- Parties ----------
body.push(article("۱", "طرفین قرارداد"));
body.push(p("این قرارداد در تاریخ ......../......../۱۴۰۴ فی‌مابین طرفین زیر منعقد می‌گردد:"));
body.push(clause("۱-۱)", "کارفرما: شرکت صنعتی و خدمات مهندسی ایران، به شمارهٔ ثبت ............ و شناسهٔ ملی ............، به نشانی ............................، به نمایندگی آقای/خانم ............ (سمت: ............)، که از این پس در این قرارداد «کارفرما» نامیده می‌شود."));
body.push(clause("۱-۲)", "مجری: شرکت فناوران فرایند فردا، به شمارهٔ ثبت ............ و شناسهٔ ملی ............، به نشانی ............................، به نمایندگی آقای/خانم ............ (سمت: ............)، که از این پس در این قرارداد «مجری» نامیده می‌شود."));
body.push(p("طرفین با علم و آگاهی کامل نسبت به مفاد و آثار حقوقی آن، این قرارداد را امضا و اجرای مفاد آن را تعهد می‌نمایند."));

// ---------- Subject ----------
body.push(article("۲", "موضوع قرارداد"));
body.push(p("طراحی و اجرای پروژهٔ شفاف‌سازی و یکپارچه‌سازی سیستم اطلاعاتی کارفرما میان سه سامانهٔ عملیات، انبار و حسابداری، مشتمل بر چهار فاز به شرح مادهٔ ۴، مطابق پروپوزال مورد تأیید کارفرما که به‌عنوان «پیوست شمارهٔ ۱» جزء لاینفک این قرارداد است."));

// ---------- Duration ----------
body.push(article("۳", "مدت قرارداد"));
body.push(clause("۳-۱)", "مدت اجرای پروژه ۵۵ روز کاری از تاریخ امضای قرارداد و دریافت پیش‌پرداخت و فراهم شدن دسترسی‌های موضوع مادهٔ ۷ است."));
body.push(clause("۳-۲)", "دورهٔ پشتیبانی رایگان موضوع مادهٔ ۹ به مدت یک ماه پس از تحویل و تأیید فاز چهارم، به مدت قرارداد افزوده می‌گردد."));

// ---------- Phases ----------
body.push(article("۴", "فازها و زمان‌بندی اجرا"));
body.push(p("پروژه در چهار فاز و بر اساس جدول زیر اجرا می‌شود. زمان‌بندی هر فاز پس از تحویل و تأیید فاز پیشین آغاز می‌گردد:"));
body.push(spacer(40));
body.push(dataTable(
  ["فاز", "شرح", "مدت", "بازهٔ تجمعی"],
  [
    ["فاز ۱", "رمزگشایی دیتابیس سه نرم‌افزار", "۱۰ روز", "روز ۱ تا ۱۰"],
    ["فاز ۲", "ساخت انبار داده", "۱۰ روز", "روز ۱۱ تا ۲۰"],
    ["فاز ۳", "مایگریشن و متناظرسازی", "۲۰ روز", "روز ۲۱ تا ۴۰"],
    ["فاز ۴", "پلتفرم BI و داشبوردها", "۱۵ روز", "روز ۴۱ تا ۵۵"],
    ["جمع", "کل اجرای پروژه", "۵۵ روز", "—"],
  ],
  [1300, 4560, 1700, 1800],
  { align: [AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.CENTER, AlignmentType.CENTER], boldCols: [0], totalRow: true }
));
body.push(spacer(40));
body.push(p("شرح تفصیلی فعالیت‌ها و خروجی هر فاز مطابق پیوست شمارهٔ ۱ است.", { size: 20, color: GREY }));

// ---------- Amount & payment ----------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(article("۵", "مبلغ قرارداد و نحوهٔ پرداخت"));
body.push(clause("۵-۱)", "مبلغ کل قرارداد ۳۹۹٬۵۰۰٬۰۰۰ تومان (سیصد و نود و نه میلیون و پانصد هزار تومان) است."));
body.push(clause("۵-۲)", "پرداخت‌ها بر اساس جدول زیر و پس از تحویل و تأیید هر فاز انجام می‌شود:"));
body.push(spacer(40));
body.push(dataTable(
  ["ردیف", "مرحله", "زمان پرداخت", "مبلغ (تومان)"],
  [
    ["۱", "پیش‌پرداخت", "هنگام امضای قرارداد", "۸۰٬۰۰۰٬۰۰۰"],
    ["۲", "فاز اول — رمزگشایی دیتابیس", "پس از تحویل و تأیید", "۵۵٬۰۰۰٬۰۰۰"],
    ["۳", "فاز دوم — ساخت انبار داده", "پس از تحویل و تأیید", "۸۴٬۵۰۰٬۰۰۰"],
    ["۴", "فاز سوم — مایگریشن و متناظرسازی", "پس از تحویل و تأیید", "۱۴۳٬۰۰۰٬۰۰۰"],
    ["۵", "فاز چهارم — پلتفرم BI (به‌کسر پیش‌پرداخت)", "پس از تحویل و تأیید", "۳۷٬۰۰۰٬۰۰۰"],
    ["", "جمع کل", "", "۳۹۹٬۵۰۰٬۰۰۰"],
  ],
  [900, 4560, 2100, 1800],
  { align: [AlignmentType.CENTER, AlignmentType.RIGHT, AlignmentType.CENTER, AlignmentType.CENTER], boldCols: [0], totalRow: true }
));
body.push(spacer(60));
body.push(clause("۵-۳)", "کارفرما موظف است مبلغ هر مرحله را حداکثر ظرف ۷ روز کاری پس از تحویل و تأیید آن مرحله و در قبال ارائهٔ صورت‌حساب رسمی توسط مجری پرداخت نماید."));
body.push(clause("۵-۴)", "کلیهٔ مبالغ مشمول کسورات قانونی (از جمله مالیات تکلیفی و در صورت شمول، بیمه) مطابق قوانین جاری است. مالیات بر ارزش افزوده در صورت شمول به مبالغ فوق افزوده و در قبال صورت‌حساب رسمی توسط کارفرما پرداخت می‌شود."));
body.push(clause("۵-۵)", "کارفرما می‌تواند معادل ۱۰٪ از هر پرداخت را به‌عنوان تضمین حسن انجام کار کسر و پس از پایان دورهٔ پشتیبانی یک‌ماهه آزاد نماید. (این بند اختیاری است و در صورت توافق طرفین حذف می‌گردد.)"));

// ---------- Delivery & acceptance ----------
body.push(article("۶", "تحویل، بررسی و تأیید"));
body.push(clause("۶-۱)", "مجری هر فاز را به‌همراه مستندات مربوط، طی صورت‌جلسهٔ کتبی به کارفرما تحویل می‌دهد."));
body.push(clause("۶-۲)", "کارفرما موظف است ظرف ۵ روز کاری، فاز تحویل‌شده را بررسی و به‌صورت کتبی تأیید یا موارد اصلاحی مستند و مرتبط با موضوع قرارداد را اعلام نماید."));
body.push(clause("۶-۳)", "چنانچه کارفرما ظرف مهلت فوق نظری اعلام نکند، فاز مربوط تحویل‌شده و تأییدشده تلقی می‌گردد و مبنای پرداخت قرار می‌گیرد."));

// ---------- Contractor obligations ----------
body.push(article("۷", "تعهدات مجری"));
body.push(bullet("اجرای کامل موضوع قرارداد مطابق فازها، خروجی‌ها و کیفیت مندرج در پیوست شمارهٔ ۱."));
body.push(bullet("رعایت زمان‌بندی مادهٔ ۴ و تحویل خروجی هر فاز در موعد مقرر."));
body.push(bullet("به‌کارگیری نیروی متخصص و ابزار مناسب برای اجرای پروژه."));
body.push(bullet("حفظ محرمانگی کامل داده‌ها و اطلاعات کارفرما مطابق مادهٔ ۱۰."));
body.push(bullet("ارائهٔ مستندات فنی و آموزش کاربران در پایان فاز چهارم."));
body.push(bullet("ارائهٔ پشتیبانی رایگان یک‌ماهه پس از تحویل مطابق مادهٔ ۹."));

// ---------- Employer obligations ----------
body.push(article("۸", "تعهدات کارفرما"));
body.push(bullet("فراهم‌سازی دسترسی لازم و امن به سه سامانهٔ عملیات، انبار و حسابداری و داده‌های شعب."));
body.push(bullet("معرفی نمایندهٔ فنی و ادارای مطلع برای هماهنگی و پاسخ‌گویی به‌موقع."));
body.push(bullet("بررسی و تأیید به‌موقع خروجی هر فاز مطابق مادهٔ ۶."));
body.push(bullet("پرداخت به‌موقع مبالغ قرارداد مطابق مادهٔ ۵."));
body.push(bullet("در صورت نیاز به همکاری اشخاص ثالث (از جمله شرکت‌های ارائه‌دهندهٔ نرم‌افزارها)، هماهنگی و پیگیری لازم را انجام دهد."));

// ---------- Support ----------
body.push(article("۹", "پشتیبانی"));
body.push(clause("۹-۱)", "مجری به مدت یک ماه پس از تحویل و تأیید فاز چهارم، پشتیبانی رایگان شامل رفع اشکال، تنظیم و اصلاح داشبوردها و پاسخ‌گویی فنی را ارائه می‌دهد."));
body.push(clause("۹-۲)", "توسعه یا افزودن قابلیت‌های جدید خارج از موضوع قرارداد، پس از دورهٔ فوق، تابع توافق و قرارداد جداگانه خواهد بود."));

// ---------- Penalty ----------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(article("۱۰", "جریمهٔ تأخیر و خسارت"));
body.push(clause("۱۰-۱)", "در صورت تأخیر مجری در تحویل هر فاز نسبت به زمان‌بندی مادهٔ ۴، به‌ازای هر روز تقویمی تأخیر معادل نیم درصد (۰٫۵٪) مبلغ همان فاز به‌عنوان جریمه از مطالبات مجری کسر می‌گردد؛ سقف جریمهٔ هر فاز ۱۰٪ مبلغ آن فاز است."));
body.push(clause("۱۰-۲)", "چنانچه تأخیر مجری در تحویل یک فاز از ۱۵ روز تقویمی تجاوز کند، کارفرما حق فسخ یک‌طرفهٔ قرارداد و مطالبهٔ خسارت وارده را خواهد داشت."));
body.push(clause("۱۰-۳)", "در صورت تأخیر کارفرما در پرداخت هر مرحله بیش از مهلت بند ۵-۳، به‌ازای هر روز تأخیر معادل نیم درصد (۰٫۵٪) مبلغ آن مرحله به مطالبات مجری افزوده شده و زمان‌بندی پروژه به‌همان میزان تمدید می‌گردد."));
body.push(clause("۱۰-۴)", "تأخیرات ناشی از قصور کارفرما (از جمله عدم تأمین دسترسی، داده یا تأییدهای به‌موقع) جزو تأخیر مجری محسوب نمی‌شود و زمان‌بندی به‌همان نسبت تمدید می‌گردد."));

// ---------- Confidentiality ----------
body.push(article("۱۱", "محرمانگی"));
body.push(clause("۱۱-۱)", "مجری متعهد است کلیهٔ داده‌ها، اطلاعات مالی، عملیاتی و فنی کارفرما را که در جریان این قرارداد در اختیار می‌گیرد، کاملاً محرمانه تلقی نموده و از افشا یا استفادهٔ آن خارج از موضوع قرارداد خودداری کند."));
body.push(clause("۱۱-۲)", "تعهد محرمانگی پس از خاتمه یا فسخ قرارداد نیز به قوت خود باقی است."));

// ---------- IP & data ownership ----------
body.push(article("۱۲", "مالکیت داده و خروجی‌ها"));
body.push(clause("۱۲-۱)", "کلیهٔ داده‌های کارفرما متعلق به کارفرما است و مجری هیچ‌گونه حق مالکیتی نسبت به آن‌ها ندارد."));
body.push(clause("۱۲-۲)", "خروجی‌ها، مستندات و داشبوردهای تولیدشده در این پروژه، پس از تسویهٔ کامل مالی، به کارفرما تعلق می‌گیرد."));
body.push(clause("۱۲-۳)", "دانش فنی و ابزارهای عمومیِ مجری که مستقل از داده‌های کارفرما است، در مالکیت مجری باقی می‌ماند."));

// ---------- Termination ----------
body.push(article("۱۳", "فسخ قرارداد"));
body.push(clause("۱۳-۱)", "در صورت نقض تعهدات اساسی هر یک از طرفین و عدم رفع آن ظرف ۷ روز کاری پس از اخطار کتبی، طرف مقابل حق فسخ قرارداد را خواهد داشت."));
body.push(clause("۱۳-۲)", "در صورت فسخ، هزینهٔ فازهای تحویل‌شده و تأییدشده و نیز کار انجام‌شدهٔ فاز جاری به‌نسبت پیشرفت، پس از کسر جرایم احتمالی، تسویه می‌گردد."));

// ---------- Force majeure ----------
body.push(article("۱۴", "فورس ماژور"));
body.push(p("در صورت بروز حوادث قهری و خارج از ارادهٔ طرفین (از جمله بلایای طبیعی، قطع سراسری زیرساخت، یا تصمیمات حاکمیتی) که اجرای قرارداد را ناممکن یا متوقف سازد، تعهدات طرفین در آن بازه معلق و زمان‌بندی به‌همان میزان تمدید می‌گردد. در صورت تداوم بیش از ۳۰ روز، طرفین دربارهٔ ادامه یا خاتمهٔ قرارداد تصمیم‌گیری خواهند کرد."));

// ---------- Disputes ----------
body.push(article("۱۵", "حل اختلاف"));
body.push(p("چنانچه در تفسیر یا اجرای این قرارداد اختلافی بروز کند، طرفین ابتدا از طریق مذاکره و سازش نسبت به حل آن اقدام می‌نمایند. در صورت عدم حصول نتیجه، موضوع از طریق مراجع قضایی صالح رسیدگی خواهد شد."));

// ---------- Legal address ----------
body.push(article("۱۶", "اقامتگاه قانونی و ابلاغ"));
body.push(p("نشانی‌های مندرج در مادهٔ ۱ اقامتگاه قانونی طرفین است و کلیهٔ مکاتبات و ابلاغ‌ها به این نشانی‌ها معتبر تلقی می‌شود. هرگونه تغییر نشانی باید ظرف ۷ روز کتباً به طرف مقابل اعلام گردد."));

// ---------- Copies ----------
body.push(article("۱۷", "نسخ و پیوست‌های قرارداد"));
body.push(p("این قرارداد در ۱۷ ماده و ۱ پیوست (پروپوزال مورد تأیید) و در ۲ نسخهٔ متحدالمتن و دارای اعتبار یکسان تنظیم و پس از امضا و مبادله، برای طرفین لازم‌الاجرا است."));

// ---------- Signatures ----------
body.push(spacer(300));
body.push(new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
  layout: TableLayoutType.FIXED, visuallyRightToLeft: true,
  borders: { top: thin("FFFFFF"), bottom: thin("FFFFFF"), left: thin("FFFFFF"), right: thin("FFFFFF"), insideHorizontal: thin("FFFFFF"), insideVertical: thin("FFFFFF") },
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: 4680, type: WidthType.DXA }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [
      new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 220, line: 300 }, children: [new TextRun({ text: "مجری", font: FONT, rtl: true, size: 22, bold: true, color: NAVY })] }),
      new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "شرکت فناوران فرایند فردا", font: FONT, rtl: true, size: 20 })] }),
      new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "نام و امضا / مهر", font: FONT, rtl: true, size: 19, color: GREY })] }),
    ] }),
    new TableCell({ width: { size: 4680, type: WidthType.DXA }, margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [
      new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 220, line: 300 }, children: [new TextRun({ text: "کارفرما", font: FONT, rtl: true, size: 22, bold: true, color: NAVY })] }),
      new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "شرکت صنعتی و خدمات مهندسی ایران", font: FONT, rtl: true, size: 20 })] }),
      new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "نام و امضا / مهر", font: FONT, rtl: true, size: 19, color: GREY })] }),
    ] }),
  ] })],
}));

const doc = new Document({
  creator: "Contract",
  title: "قرارداد اجرای پروژهٔ شفاف‌سازی و یکپارچه‌سازی سیستم اطلاعاتی",
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.RIGHT, style: { run: { color: TEAL }, paragraph: { indent: { right: 500, hanging: 220 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1100, bottom: 1100, left: 1200, right: 1200 } }, bidi: true },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync("/home/user/NewT/proposal/contract.docx", buf); console.log("WROTE contract.docx", buf.length, "bytes"); });
