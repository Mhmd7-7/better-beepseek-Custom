(function(){"use strict";function X(e){const t=Array.isArray(e.skills)?e.skills.map(s=>({name:String(s&&s.name?s.name:"skill"),content:String(s&&s.content?s.content:"")})).filter(s=>s.content.trim().length>0):[],n=Array.isArray(e.memories)?e.memories.map(s=>({key:q(s&&s.key),value:String(s&&s.value?s.value:""),importance:H(s&&s.importance)})).filter(s=>s.key&&s.value.trim().length>0):[],r=M(e.activeProject),i=(Array.isArray(e.systemPromptEntries)?e.systemPromptEntries:[]).map(s=>({id:String(s&&s.id?s.id:""),content:String(s&&s.content?s.content:""),enabled:s&&typeof s.enabled=="boolean"?s.enabled:!0,schedule:F(s&&s.schedule)})).filter(s=>s.id&&s.content.trim().length>0&&s.enabled);return{systemPrompt:String(e.systemPrompt||""),systemPromptEntries:i,skills:t,memories:n,activeCharacter:e.activeCharacter||null,preferredLang:String(e.preferredLang||""),disableSystemPrompt:!!e.disableSystemPrompt,disableMemory:!!e.disableMemory,systemPromptInjectionFrequency:String(e.systemPromptInjectionFrequency||"first"),systemPromptInjectionInterval:Number(e.systemPromptInjectionInterval)||3,activeProject:r,projectRagEnabled:!!e.projectRagEnabled,projectRagLimit:Number(e.projectRagLimit)||5,injectSystemDateTime:!!e.injectSystemDateTime,deepResearch:R(e.deepResearch)}}function R(e){return!e||typeof e!="object"?{enabled:!1,runId:""}:{enabled:!!e.enabled,runId:String(e.runId||"").trim()}}function M(e){if(!e||typeof e!="object")return null;const t=String(e.name||"").trim(),n=String(e.instructions||""),r=Array.isArray(e.files)?e.files.map(o=>({name:String(o&&o.name?o.name:"file"),content:String(o&&o.content?o.content:"")})).filter(o=>o.content.length>0):[];return t?{name:t,instructions:n,files:r}:null}function F(e){if(!e||typeof e!="object")return{type:"first",everyNTurns:1};const t=String(e.type||"first");return{type:["first","always","interval"].includes(t)?t:"first",everyNTurns:Math.max(1,Math.floor(Number(e.everyNTurns)||3))}}function q(e){return String(e||"").trim().toLowerCase().replace(/[^a-z0-9_]/g,"")}function H(e){return String(e||"called").toLowerCase()==="always"?"always":"called"}const J=`
## SheetJS (XLSX) Library Reference

### GLOBAL AVAILABILITY
- XLSX is ALREADY globally available as \`window.XLSX\` in the sandbox.
- Do NOT use \`import\`, \`require\`, or \`const XLSX = ...\`.
- Just call \`XLSX.utils.book_new()\`, \`XLSX.utils.json_to_sheet()\`, etc. directly.

### CORRECT API (most common operations)

1. CREATE WORKBOOK:
   const wb = XLSX.utils.book_new();

2. CREATE SHEET FROM DATA:
   // From array of objects (column headers auto-detected):
   const ws = XLSX.utils.json_to_sheet([
     { Name: "Alice", Age: 30 },
     { Name: "Bob", Age: 25 }
   ]);
   // From array of arrays (first row = headers):
   const ws2 = XLSX.utils.aoa_to_sheet([
     ["Name", "Age"],
     ["Alice", 30],
     ["Bob", 25]
   ]);

3. APPEND SHEET TO WORKBOOK:
   XLSX.utils.book_append_sheet(wb, ws, "SheetName");

4. COLUMN WIDTHS (optional but recommended):
   ws["!cols"] = [{ wch: 20 }, { wch: 10 }];

5. SAVE \u2014 ALWAYS end with:
   XLSX.writeFile(wb, "filename.xlsx");
   // CRITICAL: This triggers the download. Without it, nothing happens.

### COMPLETE MINIMAL EXAMPLE:
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.json_to_sheet([
  { Product: "Widget", Price: 9.99, Stock: 42 },
  { Product: "Gadget", Price: 24.99, Stock: 17 }
]);
ws["!cols"] = [{ wch: 15 }, { wch: 10 }, { wch: 10 }];
XLSX.utils.book_append_sheet(wb, ws, "Products");
XLSX.writeFile(wb, "products.xlsx");

### COMMON MISTAKES TO AVOID:
- \u2717 \`const XLSX = require('xlsx')\` \u2014 NOT available, don't use require
- \u2717 \`const XLSX = ...\` \u2014 XLSX is already defined, redeclaring causes error
- \u2717 \`XLSX.write(wb, ...)\` without type \u2014 use \`XLSX.writeFile(wb, name)\` for download
- \u2717 \`for each row manually\` \u2014 use json_to_sheet or aoa_to_sheet
- \u2717 Forgetting \`XLSX.utils.book_append_sheet()\` \u2014 the sheet must be added to workbook
- \u2717 \`await XLSX.writeFile()\` \u2014 writeFile is synchronous, no await needed
- \u2717 Browser APIs like \`document.getElementById\`, \`fetch\`, \`Blob\` \u2014 NOT available in sandbox

### CELL STYLING (limited support):
// Cell object in sheet:
ws["A1"] = { t: "s", v: "Header", s: { font: { bold: true } } };
// But for simplicity, prefer json_to_sheet or aoa_to_sheet with post-processing.

### MULTIPLE SHEETS:
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data1), "Sheet1");
XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data2), "Sheet2");
XLSX.writeFile(wb, "report.xlsx");

### FORMULAS:
const ws = XLSX.utils.aoa_to_sheet([
  ["Item", "Price", "Qty", "Total"],
  ["A", 10, 2, { t: "n", f: "B2*C2" }]
]);
`.trim(),U=`
## PptxGenJS Library Reference (PowerPoint)

### GLOBAL AVAILABILITY
- PptxGenJS is ALREADY globally available as \`window.PptxGenJS\` and \`window.pptxgen\` in the sandbox.
- Do NOT use \`import\`, \`require\`, or \`const PptxGenJS = ...\`.
- Just call \`new PptxGenJS()\` directly.

### CORRECT API

1. CREATE PRESENTATION:
   const pptx = new PptxGenJS();

2. CONFIGURE (optional):
   pptx.author = "Better DeepSeek";
   pptx.title = "Presentation Title";
   pptx.layout = "LAYOUT_WIDE"; // 16:9

3. ADD A SLIDE:
   const slide = pptx.addSlide();

4. ADD CONTENT TO SLIDE:
   // Text:
   slide.addText("Hello World", { x: 1, y: 1, w: 8, h: 1, fontSize: 24 });

   // Multi-line / bullet points:
   slide.addText([
     { text: "Main Title", options: { fontSize: 28, bold: true } },
     { text: "Subtitle text", options: { fontSize: 18 } }
   ], { x: 0.5, y: 0.5, w: 9, h: 2 });

   // Table:
   slide.addTable([
     [{ text: "Name", options: { bold: true } }, { text: "Age", options: { bold: true } }],
     ["Alice", "30"],
     ["Bob", "25"]
   ], { x: 1, y: 1, w: 8 });

   // Chart (bar, line, pie, etc.):
   slide.addChart(pptx.charts.BAR, [
     { name: "Sales", labels: ["Q1","Q2","Q3","Q4"], values: [100, 150, 130, 200] }
   ], { x: 1, y: 1, w: 8, h: 4 });

   // Image from URL:
   // slide.addImage({ path: "https://example.com/image.png", x: 1, y: 1, w: 4, h: 3 });

   // Shape:
   slide.addShape(pptx.shapes.RECTANGLE, { x: 1, y: 1, w: 4, h: 3, fill: { color: "4472C4" } });

5. SAVE \u2014 ALWAYS end with:
   await pptx.writeFile({ fileName: "Presentation.pptx" });
   // CRITICAL: Without this call, no file is generated. Must be awaited.

### COMPLETE MINIMAL EXAMPLE:
const pptx = new PptxGenJS();
pptx.title = "Project Plan";
pptx.layout = "LAYOUT_WIDE";

const slide1 = pptx.addSlide();
slide1.addText("Project Plan 2026", { x: 1, y: 1.5, w: 8, h: 1.5, fontSize: 36, bold: true, color: "1e3a8a", align: "center" });
slide1.addText("Prepared by Better DeepSeek", { x: 1, y: 3.5, w: 8, h: 0.8, fontSize: 16, align: "center" });

const slide2 = pptx.addSlide();
slide2.addText("Timeline", { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 28, bold: true });
slide2.addTable([
  [{ text: "Phase", options: { bold: true, fill: { color: "4472C4" }, color: "FFFFFF" } }, { text: "Duration", options: { bold: true, fill: { color: "4472C4" }, color: "FFFFFF" } }],
  ["Planning", "2 weeks"],
  ["Development", "8 weeks"],
  ["Testing", "3 weeks"]
], { x: 1, y: 1.5, w: 8 });

const slide3 = pptx.addSlide();
slide3.addText("Budget Overview", { x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 28, bold: true });
slide3.addChart(pptx.charts.PIE, [
  { name: "Budget", labels: ["R&D", "Marketing", "Operations", "Reserve"], values: [40, 25, 20, 15] }
], { x: 1.5, y: 1.5, w: 7, h: 4 });

await pptx.writeFile({ fileName: "ProjectPlan.pptx" });

### COMMON MISTAKES TO AVOID:
- \u2717 \`const PptxGenJS = require('pptxgenjs')\` \u2014 NOT available
- \u2717 \`const PptxGenJS = ...\` \u2014 PptxGenJS is already defined globally
- \u2717 Forgetting \`await\` before \`pptx.writeFile()\` \u2014 it's async, must be awaited
- \u2717 \`pptx.save()\` \u2014 wrong method, use \`pptx.writeFile({ fileName: ... })\`
- \u2717 \`slide.addText("text", x, y, w, h)\` \u2014 wrong! Second arg is an options object
- \u2717 Using \`document.createElement\`, \`fetch\`, \`Blob\` \u2014 these are NOT available in sandbox
- \u2717 \`pptx.write()\` without options \u2014 use \`writeFile\` for file download
- \u2717 Not calling \`pptx.writeFile\` at all \u2014 the most common reason for "no output"

### POSITIONING HELP:
- Slide dimensions: LAYOUT_WIDE = 10" x 5.625", LAYOUT_STANDARD = 10" x 7.5"
- All positions in inches: { x: 0.5, y: 0.5, w: 9, h: 1 }
- (0,0) = top-left corner

### CHART TYPES:
pptx.charts.BAR, pptx.charts.COLUMN, pptx.charts.LINE, pptx.charts.PIE,
pptx.charts.DOUGHNUT, pptx.charts.SCATTER, pptx.charts.AREA, pptx.charts.RADAR

### SHAPES:
pptx.shapes.RECTANGLE, pptx.shapes.OVAL, pptx.shapes.LINE, pptx.shapes.RIGHT_TRIANGLE,
pptx.shapes.PENTAGON, pptx.shapes.HEXAGON, pptx.shapes.CHEVRON, pptx.shapes.STAR_5_POINT
`.trim(),G=`
## docx Library Reference (Word Documents)

### GLOBAL AVAILABILITY
- The \`docx\` library is ALREADY globally available as \`window.docx\`, \`window.DOCX\`, and \`window.Packer\`.
- All library exports are also available as globals: \`Document\`, \`Paragraph\`, \`TextRun\`, \`Table\`, etc.
- Do NOT use \`import\`, \`require\`, or \`const docx = ...\` / \`const DOCX = ...\`.
- Use \`DOCX.save(doc, "filename.docx")\` to trigger download.

### CORRECT API

1. DESTRUCTURE NEEDED CLASSES (optional, for cleaner code):
   const { Document, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, BorderStyle, WidthType } = DOCX;

2. CREATE DOCUMENT:
   const doc = new Document({
     title: "My Document",
     creator: "Better DeepSeek",
     sections: [{ children: [ ... ] }]
   });

3. CONTENT ELEMENTS (use inside children array):

   // Simple paragraph:
   new Paragraph({ children: [new TextRun("Hello World")] })

   // Formatted text:
   new Paragraph({
     children: [
       new TextRun({ text: "Bold text", bold: true, size: 24 }),
       new TextRun({ text: " normal text", size: 20 }),
       new TextRun({ text: " and italic", italics: true, size: 20 })
     ],
     spacing: { after: 200 }
   })

   // Heading:
   new Paragraph({
     text: "Chapter 1",
     heading: HeadingLevel.HEADING_1
   })

   // Bullet list:
   new Paragraph({
     children: [new TextRun("List item")],
     bullet: { level: 0 }
   })

   // Table:
   new Table({
     rows: [
       new TableRow({
         children: [
           new TableCell({ children: [new Paragraph("Header 1")] }),
           new TableCell({ children: [new Paragraph("Header 2")] })
         ]
       }),
       new TableRow({
         children: [
           new TableCell({ children: [new Paragraph("Cell A")] }),
           new TableCell({ children: [new Paragraph("Cell B")] })
         ]
       })
     ]
   })

   // Page break:
   new Paragraph({ pageBreakBefore: true })

4. SAVE \u2014 ALWAYS end with:
   await DOCX.save(doc, "filename.docx");
   // Alternatively: const blob = await DOCX.Packer.toBlob(doc);
   // CRITICAL: Without DOCX.save(), no file is generated.

### COMPLETE MINIMAL EXAMPLE:
const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell } = DOCX;

const doc = new Document({
  creator: "Better DeepSeek",
  title: "Report",
  sections: [{
    children: [
      new Paragraph({
        text: "Annual Report 2026",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({
        children: [
          new TextRun({ text: "This is the introduction paragraph. ", size: 22 }),
          new TextRun({ text: "Important note in bold.", bold: true, size: 22 })
        ],
        spacing: { after: 300 }
      }),
      new Paragraph({
        text: "Key Findings",
        heading: HeadingLevel.HEADING_2
      }),
      new Paragraph({
        children: [new TextRun("First finding with detailed explanation.")],
        bullet: { level: 0 }
      }),
      new Paragraph({
        children: [new TextRun("Second finding.")],
        bullet: { level: 0 }
      }),
      new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph("Metric")] }),
              new TableCell({ children: [new Paragraph("Value")] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph("Revenue")] }),
              new TableCell({ children: [new Paragraph("$1.2M")] })
            ]
          })
        ]
      })
    ]
  }]
});

await DOCX.save(doc, "AnnualReport.docx");

### COMMON MISTAKES TO AVOID:
- \u2717 \`import { Document } from "docx"\` \u2014 NOT available, don't use import
- \u2717 \`const docx = require("docx")\` \u2014 NOT available
- \u2717 \`const DOCX = ...\` or \`const docx = ...\` \u2014 DOCX/docx is already globally defined
- \u2717 \`new Docx()\` \u2014 wrong! Use \`new Document()\` from the library
- \u2717 \`doc.save("filename.docx")\` \u2014 use \`DOCX.save(doc, "filename.docx")\`
- \u2717 Forgetting \`await\` before \`DOCX.save()\` \u2014 it's async
- \u2717 \`new TextRun("text", { bold: true })\` \u2014 wrong! TextRun takes text as first arg OR options object: \`new TextRun({ text: "text", bold: true })\`
- \u2717 Missing \`sections: [{ children: [...] }]\` \u2014 Document requires at least one section
- \u2717 Using \`document.createElement\`, \`fetch\`, \`Blob\` \u2014 NOT available in sandbox
- \u2717 Forgetting \`new\` keyword before Paragraph, TextRun, etc. \u2014 these are constructors

### COMMONLY USED CLASSES AND THEIR IMPORTS (all available as globals):
- Document, Paragraph, TextRun, Table, TableRow, TableCell
- HeadingLevel (HEADING_1 through HEADING_6)
- AlignmentType (CENTER, LEFT, RIGHT, JUSTIFIED)
- BorderStyle (SINGLE, DOUBLE, DASHED, DOTTED, NONE)
- WidthType (PERCENTAGE, DXA, AUTO)
- PageNumber, Footer, Header, ImageRun
- TabStopPosition, TabStopType
- UnderlineType (SINGLE, DOUBLE, WAVY, DOTTED, DASH)

### TEXT STYLING OPTIONS (inside TextRun):
{ text: string, bold?: boolean, italics?: boolean, size?: number (half-points, e.g. 24 = 12pt),
  color?: string (hex), font?: string, underline?: { type: UnderlineType, color?: string },
  strike?: boolean, superScript?: boolean, subScript?: boolean }

### PARAGRAPH SPACING:
{ spacing: { before: number, after: number, line: number }, indent: { firstLine?: number, left?: number } }
`.trim(),A=[{name:"xlsx",keywords:["excel","spreadsheet","xlsx","xls","sheet","tabular data","workbook","cells",".xlsx"],skill:J},{name:"pptx",keywords:["powerpoint","presentation","slide","pptx",".pptx","slideshow","deck","power point"],skill:U},{name:"docx",keywords:["word","document","docx","msword","word document","doc",".docx","letter","report"],skill:G}];function z(e){if(!e||typeof e!="string")return[];const t=e.toLowerCase(),n=[];for(const r of A)for(const o of r.keywords)if(t.includes(o)){n.push(r.name);break}return n}function $(e){const t=z(e);if(!t.length)return"";const n=[];for(const r of t){const o=A.find(i=>i.name===r);o&&n.push(o.skill)}return n.length?["<BetterDeepSeek>","[OFFICE SKILL] The user wants to create an office document. Below is the API reference for the required library:","",n.join(`

`),"</BetterDeepSeek>"].join(`
`):""}const W=new Set(["the","a","an","and","or","but","if","then","else","when","at","by","for","with","about","against","is","it","was","were","are","be","been","between","into","through","during","before","after","above","below","to","from","up","down","in","out","on","off","over","under","again","further","once","here","there","all","any","both","each","few","more","most","other","some","such","no","nor","not","only","own","same","so","than","too","very","can","will","just","should","now","how","what","where","why","who","which","ve","veya","ama","fakat","lakin","ancak","ise","ki","de","da","mi","mu","m\xFC","m\u0131","bir","bu","\u015Fu","o","i\xE7in","gibi","kadar","ile","taraf\u0131ndan","hakk\u0131nda","kar\u015F\u0131","aras\u0131nda","i\xE7ine","boyunca","\xF6nce","sonra","\xFCzerinde","alt\u0131nda","yine","daha","en","t\xFCm","her","baz\u0131","hi\xE7","sadece","kendi","ayn\u0131","\xF6yle","b\xF6yle","\xE7ok","yap\u0131lan","yaparak","olan"]);function Y(e,t=800,n=5){if(!e||!e.content)return[];const r=e.content.split(/\r?\n/);if(r.length===0)return[];const o=[];let i=0;for(;i<r.length;){const s=[];let d=0;const u=i+1;for(;i<r.length&&(d<t||s.length<3);)s.push(r[i]),d+=r[i].length+1,i++;const a=i;if(o.push({fileName:e.name,content:s.join(`
`),startLine:u,endLine:a}),i>=r.length)break;i=Math.max(u,i-n)}return o}function D(e){return e?(String(e).toLowerCase().match(/[a-z0-9_\u015f\u00e7g\u00f6\u0131\u00fc]+/gi)||[]).filter(n=>n.length>=2&&!W.has(n)):[]}function V(e,t,n=5){if(!e||!t||!t.length)return[];const r=[];for(const m of t)r.push(...Y(m,800,5));if(r.length===0)return[];const o=D(e);if(o.length===0)return[];const i=r.length,s=r.map(m=>D(m.content)),d=s.map(m=>m.length),a=d.reduce((m,g)=>m+g,0)/i||1,l={};for(const m of o){l[m]=0;for(const g of s)g.includes(m)&&l[m]++}const p=1.2,x=.75,E=[];for(let m=0;m<i;m++){const g=r[m],b=s[m],k=d[m];let c=0;const f={};for(const w of b)f[w]=(f[w]||0)+1;for(const w of o){const T=f[w]||0;if(T===0)continue;const y=l[w]||0,L=Math.log(1+(i-y+.5)/(y+.5))*(T*(p+1))/(T+p*(1-x+x*(k/a)));c+=L}const h=String(g.fileName).toLowerCase();for(const w of o)h.includes(w)&&(c+=12);c>0&&E.push({...g,score:c})}return E.sort((m,g)=>g.score-m.score).slice(0,Math.max(1,n))}function K(e,t="Project"){if(!e||!e.length)return"";let n=`<BDS:PROJECT_CONTEXT>
`;n+=`You are working on the project "${t}". Based on the user's latest prompt, here are the most relevant sections of the project files:

`;for(const r of e){const o=r.fileName.split(".").pop()||"";n+=`--- [FILE: ${r.fileName} (Lines ${r.startLine}-${r.endLine})] ---
`,n+=`\`\`\`${o}
`,n+=r.content+`
`,n+="```\n\n"}return n+="</BDS:PROJECT_CONTEXT>",n}function O(e,t){t.sessionUserMsgCounts||(t.sessionUserMsgCounts={});const n=Q(e),r=Z(e);let o=1;n&&n.length>0?(o=n.filter(d=>{const u=String(d.role||d.author||"").toLowerCase();return u==="user"||u==="human"}).length,t.sessionUserMsgCounts[r]=o):typeof e.prompt=="string"&&(e.message_id===1||e.parent_message_id==null?o=1:o=(t.sessionUserMsgCounts[r]||0)+1,t.sessionUserMsgCounts[r]=o);let i=!1,s=null;if(n&&n.length>0){s=ee(n)||n[n.length-1];const d=C(s);if(d){const u=v(d),a=te(n,s);let l=!1;const p=t.config.systemPromptInjectionFrequency||"first";if(p==="always")l=!0;else if(p==="every_x"){const E=t.config.systemPromptInjectionInterval||3;(o-1)%E===0?l=!0:a||(l=!0)}else l=!a,(n.length>1||t.hasInjected&&t.hasInjected(r))&&(l=!1);const x=_(u,r,t,l,n,s);window.dispatchEvent(new CustomEvent("bds:mutation-applied",{detail:JSON.stringify({conversationId:r,injectedText:x||"",userPrompt:u})})),x?(N(s,`${x}

${u}`),i=!0):u!==d&&(N(s,u),i=!0)}}else if(typeof e.prompt=="string"){const d=v(e.prompt),u=e.message_id===1||e.parent_message_id==null,a=t.config.systemPromptInjectionFrequency||"first";let l=!1;if(a==="always")l=!0;else if(a==="every_x"){const x=t.config.systemPromptInjectionInterval||3;(u||(o-1)%x===0)&&(l=!0)}else l=u;const p=_(d,r,t,l,null,null);window.dispatchEvent(new CustomEvent("bds:mutation-applied",{detail:JSON.stringify({conversationId:r,injectedText:p||"",userPrompt:d})})),p?(e.prompt=`${p}

${d}`,i=!0):d!==e.prompt&&(e.prompt=d,i=!0)}return{changed:i,payload:e}}function Q(e){return Array.isArray(e.messages)?e.messages:e.data&&Array.isArray(e.data.messages)?e.data.messages:e.chat&&Array.isArray(e.chat.messages)?e.chat.messages:null}function Z(e){return String(e.conversation_id||e.conversationId||e.chat_session_id||e.chat_id||e.id||"default")}function ee(e){for(let t=e.length-1;t>=0;t-=1){const n=e[t];if(!n||typeof n!="object")continue;const r=String(n.role||n.author||"").toLowerCase();if(r==="user"||r==="human")return n}return null}function C(e){return e?typeof e.content=="string"?e.content:Array.isArray(e.content)?e.content.map(t=>typeof t=="string"?t:t&&typeof t.text=="string"?t.text:"").join(`
`):typeof e.prompt=="string"?e.prompt:"":""}function N(e,t){if(e){if(typeof e.content=="string"||e.content==null){e.content=t;return}if(Array.isArray(e.content)){e.content=[{type:"text",text:t}];return}if(typeof e.prompt=="string"){e.prompt=t;return}e.content=t}}function te(e,t=null){if(!Array.isArray(e))return!1;for(const n of e){if(n===t)continue;if(C(n).includes("<BetterDeepSeek>"))return!0}return!1}function _(e,t,n,r=!1,o=null,i=null){const s=[],d=ne(e,t,n);d&&s.push(d);const u=n.config.systemPromptEntries||[];if(u.length>0){const g=n.sessionUserMsgCounts[t]||1;for(const b of u)b.content.trim()&&pe(b,g,t,n)&&(s.push(`<BetterDeepSeek>
${b.content.trim()}
</BetterDeepSeek>`),n.markEntryInjected&&n.markEntryInjected(t,b.id))}else r&&n.config.systemPrompt.trim()&&!n.config.disableSystemPrompt&&(s.push(`<BetterDeepSeek>
${n.config.systemPrompt.trim()}
</BetterDeepSeek>`),n.markInjected&&n.markInjected(t));const a=I(n.config.skills);let l=null;if(!r&&o&&(l=ge(o,i)),r||a&&a!==l){const g=oe(n);g&&s.push(g)}const p=ce(e,n,o);p&&s.push(p);const x=$(e);x&&s.push(x);const E=n.config.activeCharacter;if(E){let g=o?he(o,i):null;if(!g&&n.getLastChar&&(g=n.getLastChar(t)),!g&&n.currentSessionChar&&(o==null?void 0:o.length)>1&&(g=n.currentSessionChar),r||!g||g!==E.name){const b=de(n);b&&(s.push(b),n.setLastChar&&n.setLastChar(t,E.name),n.currentSessionChar=E.name)}}n.isNextVoiceMessage&&(s.push("<BetterDeepSeek>User send this message using voice recorder tool.</BetterDeepSeek>"),n.isNextVoiceMessage=!1);const m=n.config&&n.config.activeProject;if(m){let g=null;if(!r&&o&&(g=me(o,i)),r||!g||g!==m.name){const b=ue(n);b&&s.push(b)}if(n.config.projectRagEnabled&&Array.isArray(m.files)&&m.files.length>0){const b=Number(n.config.projectRagLimit)||5,k=V(e,m.files,b);if(k&&k.length>0){const c=K(k,m.name);c&&s.push(c)}}}if(r){const g=fe(n);g&&s.push(g)}return s.join(`

`)}function ne(e,t,n){var o;const r=(o=n.config)==null?void 0:o.deepResearch;return!(r!=null&&r.enabled)||!r.runId?"":(r.enabled=!1,re(r.runId,t,e),["<BetterDeepSeek>",'[BDS:DEEP_RESEARCH] The DeepResearch toggle is enabled. Treat this exactly as the user asking: "Perform Deep Research on the following request."',`Run ID: ${r.runId}`,"","CRITICAL: In this first turn, you must ONLY produce a research plan. Do NOT browse or search. Do NOT produce an ordinary answer. Do NOT produce a direct report.",`Output ONLY a plan using: <BDS:DEEP_RESEARCH_PLAN runId="${r.runId}">JSON</BDS:DEEP_RESEARCH_PLAN>`,"After this turn, BDS will execute steps one-by-one. After each step result is provided, analyze it before continuing. Do NOT skip ahead to the final report until BDS tells you all steps are complete.","","The JSON plan must include:",'- "title": A short descriptive title for the research','- "steps": An array of research steps, each with:','  - "id": step number','  - "action": "search" or "fetch"','  - "query": a specific search query or URL to fetch','  - "purpose": why this step is needed','  - "sourceType": for search steps, one of "general", "docs", "news", "reviews", "academic", or "commerce"',"","Search steps must use narrow queries with named entities, constraints, dates or locations, product or version names, and clear source intent.","",`User research question: ${e}`,"</BetterDeepSeek>"].join(`
`))}function re(e,t,n){typeof window>"u"||!window.dispatchEvent||window.dispatchEvent(new CustomEvent("bds:deep-research-started",{detail:JSON.stringify({runId:e,conversationId:t,userPrompt:n,timestamp:Date.now()})}))}function oe(e){if(!e.config.skills.length)return"";const t=e.config.skills.map(n=>`## ${n.name}
${n.content.trim()}`).join(`

`);return`<BetterDeepSeek> <BDS:SKILLS fingerprint="${I(e.config.skills)}">
${t}
</BDS:SKILLS> </BetterDeepSeek>`}function I(e){return!Array.isArray(e)||!e.length?"":e.map(t=>`${t.name}:${(t.content||"").length}`).sort().join("|")}function se(e){if(!Array.isArray(e))return null;for(let t=e.length-1;t>=0;t--){const n=e[t];if(!n||typeof n!="object")continue;const r=String(n.role||n.author||"").toLowerCase();if(!(r==="user"||r==="human")&&(r==="assistant"||r==="ai"||r==="bot"))return n}return null}function P(e){return!e||typeof e!="string"?[]:e.split(new RegExp("[_-]|\\s+|(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])")).map(t=>t.toLowerCase().replace(/[^a-z0-9]/g,"")).filter(t=>t.length>0)}function ie(e,t){if(!e.length||!t.length)return 0;const n=new Set(t);let r=0;for(const o of e)n.has(o)&&r++;return r/e.length}function ae(e,t){return t===1?e>=1:e>=.5}function ce(e,t,n){if(t.config.disableMemory||!t.config.memories.length)return"";const r=n?se(n):null,o=r?C(r):"",i=[e,o].filter(Boolean).join(" "),s=P(i),d=[];for(const a of t.config.memories){if(a.importance==="always"){d.push(a);continue}if(!a.key)continue;const l=P(a.key);if(!l.length){i.toLowerCase().includes(a.key.toLowerCase())&&d.push(a);continue}const p=[...new Set(l)],x=ie(p,s);(ae(x,p.length)||i.toLowerCase().includes(a.key.toLowerCase()))&&d.push(a)}return d.length?`<BetterDeepSeek>
${d.map(a=>`<BDS:memory_calls importance="${a.importance}">${a.key}: ${le(a.value)}</BDS:memory_calls>`).join(`
`)}
</BetterDeepSeek>`:""}function le(e){return String(e).replace(/<\//g,"<\\/").trim()}function ue(e){const t=e.config&&e.config.activeProject;if(!t)return"";let n="";return t.instructions&&t.instructions.trim()&&(n+=t.instructions.trim()+`
`),`<BetterDeepSeek>
<BDS:PROJECT name="${t.name}">
${n}</BDS:PROJECT>
</BetterDeepSeek>`}function de(e){const t=e.config.activeCharacter;if(!t||!t.content)return"";let n=`Character Name: ${t.name}
`;return t.usage&&(n+=`Usage Domain: ${t.usage}
`),n+=`---
${t.content.trim()}`,`<BetterDeepSeek> <BDS:RP>
${n}
</BDS:RP> </BetterDeepSeek>`}function fe(e){const t=[];if(e.config.injectSystemDateTime!==!1){const r=new Date;t.push(`User's System Date & Time: ${r.toLocaleString()}`)}const n=e.config.preferredLang;return n&&n.trim()&&t.push(`Always respond in ${n.trim()}.`),t.length===0?"":`<BetterDeepSeek>
${t.join(`
`)}
</BetterDeepSeek>`}function pe(e,t,n,r){const i=(r.getInjectedEntries?r.getInjectedEntries(n):[]).includes(e.id);switch(e.schedule.type){case"first":return!i;case"always":return!0;case"interval":{const s=e.schedule.everyNTurns||3;return i?(t-1)%s===0:!0}default:return!1}}function he(e,t=null){if(!Array.isArray(e))return null;for(let n=e.length-1;n>=0;n--){const r=e[n];if(r===t)continue;const o=C(r);if(!o.includes("<BDS:RP>"))continue;const i=o.match(/Character Name:\s*(.*?)\n/);if(i&&i[1])return i[1].trim()}return null}function ge(e,t=null){if(!Array.isArray(e))return null;for(let n=e.length-1;n>=0;n--){const r=e[n];if(r===t)continue;const i=C(r).match(/<BDS:SKILLS fingerprint="(.*?)">/);if(i&&i[1])return i[1]}return null}function me(e,t=null){if(!Array.isArray(e))return null;for(let n=e.length-1;n>=0;n--){const r=e[n];if(r===t)continue;const i=C(r).match(/<BDS:PROJECT name="(.*?)">/);if(i&&i[1])return i[1]}return null}function v(e){let t=String(e||"");return t=t.replace(/<BetterDeepSeek>([\s\S]*?)<\/BetterDeepSeek>/gi,(n,r)=>r.includes("[BDS:AUTO]")||r.includes("[BDS:DEEP_RESEARCH]")||/<BDS:memory_calls[\s>]/i.test(r)?n:""),t=t.replace(/<BDS:SKILLS>[\s\S]*?<\/BDS:SKILLS>/gi,""),t=t.replace(/<BDS:memory_calls[^>]*>[\s\S]*?<\/BDS:memory_calls>/gi,""),t=t.replace(/<BDS:RP>[\s\S]*?<\/BDS:RP>/gi,""),t=t.replace(/<BDS:PROJECT[^>]*>[\s\S]*?<\/BDS:PROJECT>/gi,""),t=t.replace(/<BDS:PROJECT_CONTEXT>[\s\S]*?<\/BDS:PROJECT_CONTEXT>/gi,""),t.trim()}function we(e,t,n,r){const o=window.fetch;window.fetch=async function(s,d){try{const u=Se(s);if(!t(u))return o.apply(this,arguments);if(Ee(s,d,e),u.includes("/api/v0/chat_session/fetch_page")){const a=await o.apply(this,arguments);return a.clone().json().then(p=>{window.dispatchEvent(new CustomEvent("bds:session-data",{detail:JSON.stringify(p)}))}).catch(()=>{}),a}if(u.includes("/api/v0/chat/history_messages")){const a=await o.apply(this,arguments);return a.clone().json().then(p=>{window.dispatchEvent(new CustomEvent("bds:history-msgs",{detail:JSON.stringify(p)}))}).catch(()=>{}),a}n(u);try{const a=await ye(s,d,e);if(!a){const p=await o.apply(this,arguments);return B(p,u,a==null?void 0:a.modelName),p}const l=await o.call(this,a.input,a.init);return l&&l.status>=500&&window.dispatchEvent(new CustomEvent("bds:network-error",{detail:JSON.stringify({url:u,status:l.status,type:"fetch"})})),B(l,u,a.modelName),l}catch(a){throw window.dispatchEvent(new CustomEvent("bds:network-error",{detail:JSON.stringify({url:u,status:0,type:"fetch",error:String(a)})})),a}finally{r(u)}}catch(u){return console.warn("[BetterDeepSeek] Request patch failed:",u),o.apply(this,arguments)}}}function B(e,t,n){if(!(!e||!e.clone))try{const r=e.clone();be(r,n).catch(()=>{})}catch{}}function Se(e){return typeof e=="string"?e:e instanceof URL?e.toString():e instanceof Request?e.url:""}async function ye(e,t,n){const r=await xe(e,t);if(!r)return null;let o;try{o=JSON.parse(r)}catch{return null}const i=o.model||null,s=O(o,n);if(!s.changed)return null;const d=JSON.stringify(s.payload),u=t&&t.headers?t.headers:e instanceof Request?e.headers:void 0,a=new Headers(u||{});a.set("content-type","application/json");const l={method:t&&t.method||(e instanceof Request?e.method:"POST"),headers:a,body:d,credentials:t&&t.credentials||(e instanceof Request?e.credentials:void 0),cache:t&&t.cache||(e instanceof Request?e.cache:void 0),mode:t&&t.mode||(e instanceof Request?e.mode:void 0),redirect:t&&t.redirect||(e instanceof Request?e.redirect:void 0),referrer:t&&t.referrer||(e instanceof Request?e.referrer:void 0),referrerPolicy:t&&t.referrerPolicy||(e instanceof Request?e.referrerPolicy:void 0),keepalive:t&&t.keepalive||(e instanceof Request?e.keepalive:void 0),integrity:t&&t.integrity||(e instanceof Request?e.integrity:void 0),signal:t&&t.signal||(e instanceof Request?e.signal:void 0)};return{input:typeof e=="string"||e instanceof URL?e:e.url,init:l,modelName:i}}async function be(e,t){try{const n=e.headers.get("content-type")||"";if(n.includes("text/event-stream")||n.includes("stream"))await Te(e,t);else{const r=await e.text();try{const o=JSON.parse(r),i=(o==null?void 0:o.usage)||(o==null?void 0:o.token_usage);i&&j(i.prompt_tokens||i.input_tokens||0,i.completion_tokens||i.output_tokens||0,t)}catch{}}}catch{}}async function Te(e,t){var s;const n=(s=e.body)==null?void 0:s.getReader();if(!n)return;const r=new TextDecoder;let o="";try{for(;;){const{done:d,value:u}=await n.read();if(u&&(o+=r.decode(u,{stream:!d})),d)break}}catch{return}const i=o.split(`
`);for(let d=i.length-1;d>=0;d--){const u=i[d].trim();if(!u.startsWith("data: "))continue;const a=u.slice(6).trim();if(a!=="[DONE]")try{const l=JSON.parse(a),p=(l==null?void 0:l.usage)||(l==null?void 0:l.token_usage);if(p){j(p.prompt_tokens||p.input_tokens||0,p.completion_tokens||p.output_tokens||0,t||(l==null?void 0:l.model));break}}catch{}}}function j(e,t,n){typeof e!="number"&&typeof t!="number"||window.dispatchEvent(new CustomEvent("bds:token-usage",{detail:JSON.stringify({inputTokens:Number(e)||0,outputTokens:Number(t)||0,modelName:n||null,timestamp:Date.now()})}))}async function xe(e,t){return t&&typeof t.body=="string"?t.body:t&&t.body instanceof URLSearchParams?t.body.toString():e instanceof Request?e.clone().text():""}function Ee(e,t,n){try{let r;if(t&&t.headers){const o=t.headers;if(o instanceof Headers)r=o.get("authorization");else if(Array.isArray(o)){for(const[i,s]of o)if(i.toLowerCase()==="authorization"){r=s;break}}else typeof o=="object"&&(r=o.Authorization||o.authorization)}!r&&e instanceof Request&&(r=e.headers.get("authorization")),r&&typeof(n==null?void 0:n.setAuthToken)=="function"&&n.setAuthToken(r)}catch{}}function ke(e,t,n,r){const o=XMLHttpRequest.prototype.open,i=XMLHttpRequest.prototype.send,s=XMLHttpRequest.prototype.setRequestHeader;XMLHttpRequest.prototype.open=function(u,a){return this.__bdsRequestMeta={method:String(u||"GET").toUpperCase(),url:String(a||"")},o.apply(this,arguments)},XMLHttpRequest.prototype.setRequestHeader=function(u,a){return u&&String(u).toLowerCase()==="authorization"&&typeof(e==null?void 0:e.setAuthToken)=="function"&&e.setAuthToken(String(a||"")),s.apply(this,arguments)},XMLHttpRequest.prototype.send=function(u){try{const a=this.__bdsRequestMeta||{};if(!t(a.url))return i.call(this,u);if(a.url.includes("/api/v0/chat_session/fetch_page"))return this.addEventListener("load",()=>{try{const c=JSON.parse(this.responseText);window.dispatchEvent(new CustomEvent("bds:session-data",{detail:JSON.stringify(c)}))}catch{}}),i.call(this,u);if(a.url.includes("/api/v0/chat/history_messages"))return this.addEventListener("load",()=>{try{const c=JSON.parse(this.responseText);window.dispatchEvent(new CustomEvent("bds:history-msgs",{detail:JSON.stringify(c)}))}catch{}}),i.call(this,u);n(a.url);let l=!1;const p=()=>{l||(l=!0,(this.status>=500||this.status===0)&&window.dispatchEvent(new CustomEvent("bds:network-error",{detail:JSON.stringify({url:a.url,status:this.status,type:"xhr"})})),r(a.url))};this.addEventListener("loadend",p,{once:!0});const x=Ce(u);if(!x)return i.call(this,u);const E=JSON.parse(x),m=E.model||null,g=O(E,e);if(!g.changed)return i.call(this,u);const b=JSON.stringify(g.payload),k=this;return this.addEventListener("load",()=>{try{const c=k.responseText;c&&Le(c,k,m)}catch{}},{once:!0}),i.call(this,b)}catch(a){const l=this.__bdsRequestMeta||{};console.warn("[BetterDeepSeek] XHR patch failed:",a);try{return i.call(this,u)}catch(p){throw t(l.url)&&r(l.url),p}}}}function Ce(e){return typeof e=="string"?e:e instanceof URLSearchParams?e.toString():""}function Le(e,t,n){var r;try{if((((r=t.getResponseHeader)==null?void 0:r.call(t,"content-type"))||"").includes("text/event-stream")||e.startsWith("data: ")){const i=e.split(`
`);for(let s=i.length-1;s>=0;s--){const d=i[s].trim();if(!d.startsWith("data: "))continue;const u=d.slice(6).trim();if(u!=="[DONE]")try{const a=JSON.parse(u),l=a==null?void 0:a.usage;if(l){window.dispatchEvent(new CustomEvent("bds:token-usage",{detail:JSON.stringify({inputTokens:l.prompt_tokens||l.input_tokens||0,outputTokens:l.completion_tokens||l.output_tokens||0,modelName:n||(a==null?void 0:a.model)||null,timestamp:Date.now()})}));break}}catch{}}}}catch{}}(function(){"use strict";const e={configUpdate:"bds:config-update",deepResearchConfigUpdate:"bds:deep-research-config-update",requestConfig:"bds:request-config",markVoiceMessage:"bds:mark-voice-message",sessionData:"bds:session-data"},t="/api/v0/chat_session/fetch_page",n="/api/v0/chat/history_messages",r="/api/v0/chat/completion";function o(){try{return JSON.parse(localStorage.getItem("bds_injected_chats")||"[]")}catch{return[]}}function i(c){const f=o();f.includes(c)||(f.push(c),f.length>50&&f.shift(),localStorage.setItem("bds_injected_chats",JSON.stringify(f)))}function s(){try{return JSON.parse(localStorage.getItem("bds_injected_chars")||"{}")}catch{return{}}}function d(c,f){const h=s();h[c]=f;const w=Object.keys(h);w.length>50&&delete h[w[0]],localStorage.setItem("bds_injected_chars",JSON.stringify(h))}function u(c){try{return JSON.parse(localStorage.getItem("bds_injected_entries")||"{}")[c]||[]}catch{return[]}}function a(c,f){try{const h=JSON.parse(localStorage.getItem("bds_injected_entries")||"{}");h[c]||(h[c]=[]),h[c].includes(f)||h[c].push(f);const w=Object.keys(h);w.length>50&&delete h[w[0]],localStorage.setItem("bds_injected_entries",JSON.stringify(h))}catch{}}function l(){var c,f;try{for(let w=0;w<localStorage.length;w++){const T=localStorage.key(w);if(T&&/token|auth|session/i.test(T)){const y=localStorage.getItem(T);if(!y)continue;if(y.trim().startsWith("{"))try{const S=JSON.parse(y),L=S.token||S.accessToken||S.access_token||S.user_token||((c=S.user)==null?void 0:c.token);if(L&&typeof L=="string")return L}catch{}else if(typeof y=="string"&&y.length>20){let S=y;return S.startsWith("Bearer ")&&(S=S.substring(7)),S.startsWith('"')&&S.endsWith('"')&&(S=S.slice(1,-1)),S}}}const h=(f=document.cookie.split("; ").find(w=>w.startsWith("user_token=")||w.startsWith("token=")))==null?void 0:f.split("=")[1];if(h)return decodeURIComponent(h)}catch(h){console.warn("[BDS] Failed to search auth token in storage:",h)}return null}const p={config:{systemPrompt:"",systemPromptEntries:[],skills:[],memories:[],activeCharacter:null},hasInjected:c=>o().includes(c),markInjected:c=>i(c),getInjectedEntries:c=>u(c),markEntryInjected:(c,f)=>a(c,f),getLastChar:c=>s()[c]||null,setLastChar:(c,f)=>d(c,f),currentSessionChar:null,activeCompletionRequests:0,isNextVoiceMessage:!1,authToken:l(),setAuthToken:function(c){c&&c!==this.authToken&&(this.authToken=c)}};if(window.__bdsNetworkPatched)return;window.__bdsNetworkPatched=!0,(function(){if(window.__BDS_CONFIG__)return;let c=0;const f=new Map;window.addEventListener("bds:debug-api-response",w=>{let T=w.detail;if(typeof T=="string")try{T=JSON.parse(T)}catch{return}const y=f.get(T.id);y&&(y(T.result),f.delete(T.id))});function h(w){return function(){const T=Array.from(arguments);return new Promise(y=>{const S=++c;f.set(S,y),window.dispatchEvent(new CustomEvent("bds:debug-api-request",{detail:JSON.stringify({id:S,method:w,args:T})}))})}}window.__BDS_CONFIG__={raw:h("getRaw"),getFlag:h("getFlag"),getConfig:h("getConfig"),applyRemote:h("applyRemote"),replaceRemote:h("replaceRemote"),resetToBuiltin:h("resetToBuiltin"),detectModel:h("detectModel"),toggleDebugPanel:h("toggleDebugPanel")}})(),window.addEventListener(e.configUpdate,c=>{let f=c&&c.detail?c.detail:{};if(typeof f=="string")try{f=JSON.parse(f)}catch(h){console.error("[BDS] Failed to parse configUpdate detail:",h)}p.config=X(f||{})}),window.addEventListener(e.deepResearchConfigUpdate,c=>{let f=c&&c.detail?c.detail:{};if(typeof f=="string")try{f=JSON.parse(f)}catch(h){console.error("[BDS] Failed to parse deepResearchConfigUpdate detail:",h)}p.config.deepResearch=R(f||{})}),window.addEventListener(e.markVoiceMessage,()=>{p.isNextVoiceMessage=!0}),window.addEventListener("bds:request-history-msgs",async c=>{let f=c&&c.detail?c.detail:{};if(typeof f=="string")try{f=JSON.parse(f)}catch{return}const h=f==null?void 0:f.sessionId;if(!h)return;const w=`${n}?chat_session_id=${encodeURIComponent(h)}`,T={"Content-Type":"application/json"};p.authToken&&(T.Authorization=`Bearer ${p.authToken}`);try{const y=await x(w,{method:"GET",headers:T,credentials:"include"});if(!y.ok){console.warn("[BDS] history_mgs fetch failed:",y.status);return}const S=await y.json();S.__bdsExplicit=!0,window.dispatchEvent(new CustomEvent("bds:history-msgs",{detail:JSON.stringify(S)}))}catch(y){console.warn("[BDS] history_msgs fetch error:",y)}}),E();const x=window.fetch.bind(window);we(p,m,b,k),ke(p,m,b,k);function E(){window.dispatchEvent(new CustomEvent(e.requestConfig))}function m(c){const f=String(c||"");return f.includes("/api/v0/chat/completion")||f.includes("/api/v0/chat/edit_message")||f.includes(t)||f.includes(n)}function g(c,f){const h={status:c,url:String(f||""),activeCompletionRequests:p.activeCompletionRequests,timestamp:Date.now()};window.dispatchEvent(new CustomEvent(e.networkState,{detail:JSON.stringify(h)}))}function b(c){p.activeCompletionRequests+=1,g("start",c)}function k(c){p.activeCompletionRequests=Math.max(0,p.activeCompletionRequests-1),g("end",c)}})()})();
