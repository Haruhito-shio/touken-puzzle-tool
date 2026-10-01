#!/usr/bin/env node
/*
 * =====================================================================
 *  刀剣乱舞ぱずぎり 统计表 → 材料规划工具 数据转换器
 * =====================================================================
 *  作用：
 *   读取「刀剣乱舞ぱずぎり统计verX.X.X.xlsx」的四张工作表，
 *   转换成材料规划工具使用的数据文件，并输出：
 *     1. game-data-output.js   新生成的游戏数据（人工核对后再替换）
 *     2. convert-report.md     转换检查报告（表头/词典/缺行等警告）
 *     3. update-diff-report.md 新旧数据差异报告（对比同目录旧 game-data.js）
 *
 *  怎么用（只需要做一次的事）：
 *    1. 电脑上装 Node.js（免费，nodejs.org 下载 LTS 版，一路下一步）
 *    2. 在这个文件夹里打开终端，运行一次：  npm install xlsx
 *    3. 把新版 xlsx + 当前在用的旧 game-data.js 放到脚本同目录
 *    4. 每次表格更新后，运行：             node convert-pazugiri.js
 *    5. 打开 convert-report.md 看转换警告，再打开 update-diff-report.md 看数据差异
 *    6. 确认没问题后，把 game-data-output.js 的内容
 *       替换进项目里的 game-data.js（或直接改名覆盖，注意先备份）
 *
 *  表头变动怎么办：只改下面 CONFIG 里的文字，解析逻辑不用动。
 * =====================================================================
 */
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
/* ==================== 你可修改的配置区 ==================== */
const rootDir = path.join(__dirname, "..");
const CONFIG = {
  /* Excel 目录（和 data tool 同级） */
  excelDir: path.join(rootDir, "excel"),

  /* true = 自动选 excel 目录里日期最新的表格；
     想强制用某张表时改成 false，然后手填 excelFileName */
  autoSelectLatest: true,
  excelFileName: "刀剑乱舞ばずき统计ver0.2-20260930.xlsx",

  /* 四张工作表的名称（如果表名改了，在这里改）。
   * 注意：关卡统计表新版名字是「关卡统计（透视表用）」，脚本会做「前缀匹配」，
   *       所以即使后缀变化（如透视表用/统计用）也能自动找到。 */
  sheetNames: {
    chars: "刀男所需材料", // 表1：每个刀男的突破阶段材料需求
    recipes: "材料合成", // 表2：半成品怎么合成
    sources: "材料来源", // 表3：材料中文名↔日文名 + 掉落关卡
    stages: "关卡统计", // 表4：关卡→掉落材料清单（新版叫「关卡统计（透视表用）」）
  },

  /* 表头文字 → 列定位。
   * 如果以后 Excel 里表头改名了（比如「材料名」改成「材料名（日）」），
   * 只改下面对应的文字即可，解析逻辑完全不用动。 */
  headerTexts: {
    chars: {
      type: "类型", // 刀种（大太刀/太刀…）
      sword: "刀刀", // 刀男名字
      rarity: "稀有度", // 普 / 彩
      stage: "进阶", // 突破阶段，如 10->20
      matHdr: "材料", // 材料需求列（合并表头，从它开始向右全部算）
    },
    recipes: {
      cn: "中文名", // 成品中文名
      jp: "材料名", // 成品日文名
      ing: "合成材料", // 配料（日文名）
      qty: "数量", // 做 1 个成品需要几个
    },
    sources: {
      cn: "中文名", // 材料中文名
      jp: "材料", // 材料日文名
      stage: "关卡", // 掉落关卡（左块）
      drops: "掉落材料", // 掉落清单（右块，仅作为备用数据源）
    },
    stages: {
      cn: "中文名",
      jp: "材料",
      stage: "关卡", // 右块关卡列（取最后一个「关卡」表头）
      drops: "掉落材料", // 右块掉落清单（日文）
    },
  },

  /* 游戏货币，识别出来后单独记录成“合成花费”，不参与材料缺口计算 */
  coinName: "小判",

  /* 旧数据：根目录的 game-data.js */
  oldGameDataFile: path.join(rootDir, "game-data.js"),
  /* 新数据：直接覆盖根目录的 game-data.js */
  outputJs: path.join(rootDir, "game-data.js"),
  /* 报告文件夹：data tool/reports/ */
  reportDir: path.join(__dirname, "reports"),
  /* 报告文件名待会动态加日期前缀 */
  outputReport: "",
  diffReport: "",
};
/* ==================== 配置区结束 ==================== */


const warnings = [];          // 警告清单（写进报告）
const infos = [];             // 一般信息
const report = [];            // 报告正文行
const now = new Date();

/* 从文件名末尾提取日期，返回 YYYYMMDD 数字（用于选最新表格） */
function extractDate(filename){
  const base = filename.replace(/\.xlsx$/i, "");
  const m = base.match(/(\d{4})[-.]?(\d{2})[-.]?(\d{2})$/);
  if(!m) return 0;
  return parseInt(m[1],10)*10000 + parseInt(m[2],10)*100 + parseInt(m[3],10);
}

/* ---------- 小工具 ---------- */
function colLetter(idx){ let s=""; idx+=1; while(idx>0){ let r=(idx-1)%26; s=String.fromCharCode(65+r)+s; idx=Math.floor((idx-1)/26); } return s; }
function clean(s){ return String(s==null?"":s).replace(/\u3000/g," ").trim(); }
function cleanName(s){ return clean(s).replace(/\s+/g,"").replace(/\n/g,""); } // 去掉换行/空格（合并单元格里常有）
function isNum(s){ return /^\d+$/.test(clean(s)); }
function parseStage(text){
  const t = clean(text).replace(/→/g,"->");
  /* 阶段有两种写法：表1 是 10->20，关卡名是 3-17（单连字符），都支持 */
  const m = t.match(/^(\d+)\s*(?:->|→|—|－|-)\s*(\d+)$/);
  if(!m) return null;
  return { from: parseInt(m[1],10), to: parseInt(m[2],10) };
}
/* 解析「叶子x2 / 木材X2 / 蓝色颜料×1 / 味增汤」 → {name, qty}，没写数量默认 1 */
function parseMatCell(text){
  let t = clean(text);
  if(!t) return null;
  const m = t.match(/^(.*?)[x×X]\s*(\d+)$/);
  if(m) return { name: clean(m[1]), qty: parseInt(m[2],10) };
  return { name: t, qty: 1 };
}
/* 把带合并的二维数组展开：合并格的值复制到整片区域 */
function expandMerges(ws, rows){
  if(!ws["!merges"]) return rows;
  ws["!merges"].forEach(function(mg){
    const v = rows[mg.s.r] && rows[mg.s.r][mg.s.c] != null ? rows[mg.s.r][mg.s.c] : "";
    for(let r=mg.s.r; r<=mg.e.r; r++){
      if(!rows[r]) rows[r]=[];
      for(let c=mg.s.c; c<=mg.e.c; c++){ rows[r][c]=v; }
    }
  });
  return rows;
}
/* 读一个 sheet 成二维数组（第 0 行是表头），合并格已展开 */
function sheetRows(ws){
  if(!ws || !ws["!ref"]) return [];
  const aoa = XLSX.utils.sheet_to_json(ws, { header:1, defval:"", raw:true });
  return expandMerges(ws, aoa);
}
/* 按表头文字找列号（第 0 行匹配），返回所有匹配列；matchAll=false 只取第一个 */
function findCols(rows, headerText, matchAll){
  const out=[];
  if(!rows.length) return out;
  rows[0].forEach(function(h,c){
    if(cleanName(String(h)) === cleanName(headerText)) out.push(c);
  });
  return matchAll ? out : (out.length?[out[0]]:[]);
}
function fail(msg){ console.error("【转换失败】"+msg); process.exit(1); }

/* ---------- 自动选择最新表格 ---------- */
let excelFile;
if(CONFIG.autoSelectLatest){
  if(!fs.existsSync(CONFIG.excelDir)) fail("excel 目录不存在：" + CONFIG.excelDir);
  const files = fs.readdirSync(CONFIG.excelDir)
    .filter(function(f){ return /\.xlsx$/i.test(f) && !f.startsWith("~$"); });
  if(files.length === 0) fail("excel 目录里没有 .xlsx 文件：" + CONFIG.excelDir);
  files.sort(function(a, b){ return extractDate(a) - extractDate(b); });
  excelFile = files[files.length - 1];
  console.log("📊 自动选择最新表格：" + excelFile);
}else{
  excelFile = CONFIG.excelFileName;
  console.log("📊 使用指定表格：" + excelFile);
}

const file = path.join(CONFIG.excelDir, excelFile);
if(!fs.existsSync(file)) fail("找不到文件：" + file);

let wb;
try { wb = XLSX.readFile(file); }
catch(e){ fail("读取 Excel 失败：" + e.message + "（请确认是 .xlsx 文件、没有被占用）"); }
/* ---------- 计算报告日期：以表格日期为主，运行时分区分同一天多次 ---------- */
function extractDateFormatted(filename){
  const base = filename.replace(/\.xlsx$/i, "");
  const m = base.match(/(\d{4})[-.]?(\d{2})[-.]?(\d{2})$/);
  if(!m) return null;
  return m[1] + "-" + m[2] + "-" + m[3];
}

const excelDate = extractDateFormatted(excelFile) ||
                  (now.getFullYear() + "-" +
                   String(now.getMonth() + 1).padStart(2, "0") + "-" +
                   String(now.getDate()).padStart(2, "0"));

const timeSuffix = String(now.getHours()).padStart(2, "0") +
                   String(now.getMinutes()).padStart(2, "0");

/* 确保 reports 文件夹存在 */
if(!fs.existsSync(CONFIG.reportDir)) fs.mkdirSync(CONFIG.reportDir);

CONFIG.outputReport = path.join(CONFIG.reportDir, excelDate + "_" + timeSuffix + "_convert-report.md");
CONFIG.diffReport   = path.join(CONFIG.reportDir, excelDate + "_" + timeSuffix + "_update-diff-report.md");

console.log("📅 报告日期前缀：" + excelDate + "_" + timeSuffix);

function getSheet(name, tag){
  let ws = wb.Sheets[name];
  /* 【新增】前缀容错：表名可能带后缀（如「关卡统计（透视表用）」），
   * 先精确匹配，找不到就匹配「以配置名为开头」的工作表 */
  if(!ws){
    const allNames = Object.keys(wb.Sheets);
    const hit = allNames.find(function(n){ return n.indexOf(name)===0; });
    if(hit){ ws = wb.Sheets[hit]; }
    else{ warnings.push("【缺表】找不到工作表「"+name+"」（现有表："+(allNames.join("、")||"无")+"），本次"+(tag||"该部分")+"数据为空"); return null; }
  }
  return sheetRows(ws);
}

/* =====================================================================
 * 第 1 步：收集「日文名 ↔ 中文名」词典
 * 来源：表2（成品中文名/材料名）+ 表3左块 + 表4左块（中文名/材料）
 * ===================================================================== */
const ja2zh = {};   // 日文名 -> 中文名
const zh2ja = {};   // 中文名 -> 日文名

function addDict(zhRaw, jaRaw, from){
  const zh = cleanName(zhRaw), ja = cleanName(jaRaw);
  if(!zh || !ja) return;
  if(ja2zh[ja] && ja2zh[ja] !== zh){
    warnings.push("【词典冲突】日文「"+ja+"」同时对应中文「"+ja2zh[ja]+"」和「"+zh+"」（来自"+from+"），保留第一个，请人工核对");
    return;
  }
  ja2zh[ja]=zh; zh2ja[zh]=ja;
}

/* 表2：成品的中文名 ↔ 材料名 */
const recipesSheet = getSheet(CONFIG.sheetNames.recipes, "材料合成");
const rc = CONFIG.headerTexts.recipes;
const rcCn  = findCols(recipesSheet||[], rc.cn);
const rcJp  = findCols(recipesSheet||[], rc.jp);
const rcIng = findCols(recipesSheet||[], rc.ing);
const rcQty = findCols(recipesSheet||[], rc.qty);
if(recipesSheet && (!rcCn.length||!rcJp.length||!rcIng.length||!rcQty.length)){
  warnings.push("【表头】表「材料合成」没找到完整表头（中文名/材料名/合成材料/数量），请核对 CONFIG.headerTexts.recipes");
}

/* 表3/表4 左块：材料中文名 ↔ 材料（日文） */
function readDictBlock(sheetRowsArr, hdr, tag){
  if(!sheetRowsArr) return;
  const cn = findCols(sheetRowsArr, hdr.cn), jp = findCols(sheetRowsArr, hdr.jp);
  if(!cn.length || !jp.length){ warnings.push("【表头】表「"+tag+"」缺少「中文名/材料」表头，词典会不完整"); return; }
  for(let r=1;r<sheetRowsArr.length;r++){
    const row=sheetRowsArr[r];
    const zh=row[cn[0]], ja=row[jp[0]];
    if(clean(zh)&&clean(ja)) addDict(zh, ja, "表「"+tag+"」第"+(r+1)+"行");
  }
}
const sourcesSheet = getSheet(CONFIG.sheetNames.sources, "材料来源");
readDictBlock(sourcesSheet, CONFIG.headerTexts.sources, CONFIG.sheetNames.sources);
const stagesSheet = getSheet(CONFIG.sheetNames.stages, "关卡统计");
readDictBlock(stagesSheet, CONFIG.headerTexts.stages, CONFIG.sheetNames.stages);

/* =====================================================================
 * 第 2 步：解析表2 材料合成 → 配方 + 小判花费
 * ===================================================================== */
const products = {};   // 中文名 -> { jp, ings:[{ja,qty}], coin }
let recipeSkipRows = [];
if(recipesSheet && rcCn.length && rcJp.length && rcIng.length && rcQty.length){
  let cur=null, curCn="";
  for(let r=1;r<recipesSheet.length;r++){
    const row=recipesSheet[r];
    const cn=clean(row[rcCn[0]]), jp=clean(row[rcJp[0]]), ing=clean(row[rcIng[0]]);
    const qtyText=clean(row[rcQty[0]]);
    if(cn && !jp){
      /* 只有中文名没有材料名 = 说明文字行（如合成bug说明），跳过 */
      recipeSkipRows.push("第"+(r+1)+"行：只有中文名没有材料名（可能是说明文字，已跳过）");
      cur=null; curCn="";
      continue;
    }
    if(cn && jp && cn !== curCn){
      /* 新成品：合并单元格会把配料行也填上成品名，所以用「名字变化」判断新成品 */
      addDict(cn, jp, "表「材料合成」第"+(r+1)+"行（成品）");
      curCn = cn;
      cur = { cn:cn, jp:cleanName(jp), ings:[], coin:0 };
      products[cn]=cur;
    }
    if(ing){
      if(!cur){ recipeSkipRows.push("第"+(r+1)+"行：配料「"+ing+"」前面没有成品（已跳过）"); continue; }
      if(cleanName(ing) === cleanName(CONFIG.coinName)){
        cur.coin = parseInt(qtyText,10)||0;
      } else {
        cur.ings.push({ ja: cleanName(ing), qty: parseInt(qtyText,10)||1 });
      }
    }
    /* 全空的纯注释行：跳过不报 */
  }
}

/* =====================================================================
 * 第 3 步：解析表1 刀男所需材料 → 角色 + 突破阶段
 * ===================================================================== */
const charsSheet = getSheet(CONFIG.sheetNames.chars, "刀男所需材料");
const cc = CONFIG.headerTexts.chars;
const chars = [];                 // 去重后的刀男名
const swordTypes = {};            // 刀男名 -> 刀种
const breakthroughs = [];         // {sword,rarity,fromStage,toStage,needs}
const emptyNeeds = [];            // 需求为空的阶段
const aliasFixes = [];            // 材料别名归一化记录
const unmappedNeeds = {};         // 找不到中日映射的材料
let charSkipRows = [];
let matCellSkipRows = [];

if(charsSheet){
  const cType=findCols(charsSheet,cc.type), cSword=findCols(charsSheet,cc.sword),
        cRare=findCols(charsSheet,cc.rarity), cStage=findCols(charsSheet,cc.stage);
  const cMat=findCols(charsSheet,cc.matHdr,true);   // 材料表头可能横跨多列（合并）
  if(!cType.length||!cSword.length||!cRare.length||!cStage.length||!cMat.length){
    warnings.push("【表头】表「刀男所需材料」没找齐表头（类型/刀刀/稀有度/进阶/材料），请核对 CONFIG.headerTexts.chars");
  } else {
    /* 把材料列解析成「名字」的原始写法 → 看是否写的是日文/别名 */
    const resolveName = function(rawName, where){
      const n = cleanName(rawName);
      if(zh2ja[n]) return { key:n, ok:true };          // 中文名，直接是标准名
      if(ja2zh[n]){ aliasFixes.push("「"+rawName+"」→ 标准中文名「"+ja2zh[n]+"」（来源："+where+"）"); return { key:ja2zh[n], ok:true }; }
      if(!unmappedNeeds[n]) unmappedNeeds[n]=[];
      unmappedNeeds[n].push(where);
      return { key:n, ok:false };
    };

    for(let r=1;r<charsSheet.length;r++){
      const row=charsSheet[r];
      const st=parseStage(row[cStage[0]]);
      if(!st){
        /* 非阶段行：注释/说明行，跳过（不报错） */
        continue;
      }
      const sword=cleanName(row[cSword[0]]), rarity=clean(row[cRare[0]])||CONFIG.defaultRarity, type=cleanName(row[cType[0]]);
      if(!sword){ charSkipRows.push("第"+(r+1)+"行：有进阶但没刀男名（可能漏填，已跳过）"); continue; }
      if(chars.indexOf(sword)<0) chars.push(sword);
      if(!swordTypes[sword]) swordTypes[sword]=type||"未知";

      const needs={};
      cMat.forEach(function(c){
        const cell=clean(row[c]);
        if(!cell) return;
        const pm=parseMatCell(cell);
        if(!pm) return;
        const rn=resolveName(pm.name, "表「刀男所需材料」第"+(r+1)+"行");
        needs[rn.key]=(needs[rn.key]||0)+pm.qty;
      });
      breakthroughs.push({ sword:sword, rarity:rarity, fromStage:st.from, toStage:st.to, needs:needs });
      if(Object.keys(needs).length===0) emptyNeeds.push(sword+"("+rarity+") "+st.from+"->"+st.to);
    }
  }
}

/* =====================================================================
 * 第 4 步：解析表4 关卡统计 → 关卡掉落清单（日文 → 中文）
 * ===================================================================== */
const stages = [];
let dropUnmapped = {};   // 掉落里找不到中文名的
let stageSkipRows = [];
if(stagesSheet){
  const sd = CONFIG.headerTexts.stages;
  const allStageCols = findCols(stagesSheet, sd.stage, true);   // 左右两块都有「关卡」
  const dropCols = findCols(stagesSheet, sd.drops);
  if(!allStageCols.length || !dropCols.length){
    warnings.push("【表头】表「关卡统计」没找到「关卡/掉落材料」表头，关卡掉落数据为空，请核对 CONFIG.headerTexts.stages");
  } else {
    const rightStageCol = allStageCols[allStageCols.length-1];  // 右块的关卡列
    const dropCol = dropCols[0];
    for(let r=1;r<stagesSheet.length;r++){
      const row=stagesSheet[r];
      const st=parseStage(row[rightStageCol]);
      if(!st) continue;   // 合计行/空行
      const dropText=clean(row[dropCol]);
      if(!dropText) continue;
      const mats=[];
      String(dropText).split(/[、，,;；\/]/).forEach(function(item){
        const it=cleanName(item);
        if(!it) return;
        if(ja2zh[it]) mats.push(ja2zh[it]);
        else if(zh2ja[it]) mats.push(it);
        else { if(!dropUnmapped[it]) dropUnmapped[it]=[]; dropUnmapped[it].push(row[rightStageCol]); mats.push(it); }
      });
      const stageName = clean(row[rightStageCol]);
      if(stages.find(function(s){return s.stage===stageName;})){
        stageSkipRows.push("第"+(r+1)+"行：关卡「"+stageName+"」重复出现（已保留第一条）");
      } else {
        stages.push({ stage: stageName, materials: mats });
      }
    }
  }
} else {
  warnings.push("【缺表】没有「关卡统计」表，本工具的材料来源查询会没有数据");
}

/* =====================================================================
 * 第 5 步：组装 materials（基础材料 / 半成品）+ recipes + 小判
 * ===================================================================== */
const materials = {};
Object.keys(ja2zh).forEach(function(ja){
  const zh=ja2zh[ja];
  materials[zh]={ ja:ja, zh:zh, type:"base" };
});
Object.keys(products).forEach(function(cn){
  if(!materials[cn]) materials[cn]={ ja:products[cn].jp, zh:cn, type:"base" };
  materials[cn].type="crafted";
  if(products[cn].coin>0) materials[cn].coin=products[cn].coin;
  products[cn].ings.forEach(function(ing){
    if(ja2zh[ing.ja]) return;
    /* 配料日文名不在词典里：补一个占位材料，保证递归不中断 */
    if(!materials[ing.ja]){ materials[ing.ja]={ ja:ing.ja, zh:ing.ja, type:"base" }; warnings.push("【未映射】配料「"+ing.ja+"」没找到中文名，暂时用它自己当名字（请补词典或检查表头）"); }
  });
});
/* 表1里出现、但词典和配方都没有的材料（如颜料） */
Object.keys(unmappedNeeds).forEach(function(n){
  if(!materials[n]){
    materials[n]={ ja:n, zh:n, type:"base" };
    warnings.push("【未映射】材料「"+n+"」在合成表/来源表里都没找到（出现于："+unmappedNeeds[n].join("；")+"），暂时用它自己当名字。可能是遗漏填写，请人工核对");
  }
});

/* recipes：配料日文名 → 中文名 key */
const recipes=[];
Object.keys(products).forEach(function(cn){
  products[cn].ings.forEach(function(ing){
    const ingKey = ja2zh[ing.ja] || ing.ja;
    recipes.push({ product: cn, ingredient: ingKey, quantity: ing.qty });
  });
});

/* coinCosts：单独记录，不进缺口 */
const coinCosts={};
Object.keys(products).forEach(function(cn){
  if(products[cn].coin>0) coinCosts[cn]=products[cn].coin;
});

/* =====================================================================
 * 第 6 步：写出 game-data-output.js
 * ===================================================================== */
const dataObj = {
  version: "0.2.9.30",
  generatedAt: excelDate,
  source: excelFile,
  characters: chars,
  swordTypes: swordTypes,
  materials: materials,
  breakthroughs: breakthroughs,
  recipes: recipes,
  coinCosts: coinCosts,
  stages: stages,
};
const jsHead =
"/*\n"+
" * 刀剑乱舞材料规划工具 —— 公共游戏数据（由 convert-pazugiri.js 自动生成）\n"+
" * 来源："+excelFile+" · 生成时间："+dataObj.generatedAt+"\n"+
" * 材料显示规则：中日文一致只显示一个（水）；不一致显示 日文|中文（丸太|木材）\n"+
" * 小判是游戏货币：不进材料缺口，单独记录在 coinCosts / 材料.coin\n"+
" */\n\n"+
"window.GAME_DATA = "+JSON.stringify(dataObj,null,2)+";\n";
/* ---------- 先读旧数据（必须在覆盖前！） ---------- */
const oldRes = loadOldGameData();

/* ---------- 自动备份旧 game-data.js ---------- */
(function(){
  if(!fs.existsSync(CONFIG.oldGameDataFile)) return;
  const backupDir = path.join(__dirname, "backup");
  if(!fs.existsSync(backupDir)) fs.mkdirSync(backupDir);
  const backupName = excelDate + "_" + timeSuffix + "_game-data.js";
  fs.copyFileSync(CONFIG.oldGameDataFile, path.join(backupDir, backupName));
  console.log("📦 已备份旧数据到：backup/" + backupName);
})();

fs.writeFileSync(CONFIG.outputJs, jsHead, "utf8");

/* =====================================================================
 * 【新增】第 6.5 步：新旧数据差异报告 update-diff-report.md
 * 对比基准：同目录下的旧 game-data.js（当前工具在用的那份）
 * ===================================================================== */
function loadOldGameData(){
  const fpath = CONFIG.oldGameDataFile;
  if(!fs.existsSync(fpath)) return { ok:false, reason:"file-not-exist", data:null };
  try{
    const src = fs.readFileSync(fpath,"utf8");
    const ctx = vm.createContext({ window:{} });
    vm.runInContext(src, ctx, { filename: CONFIG.oldGameDataFile });
    const old = ctx.window.GAME_DATA;
    if(!old || typeof old !== "object") return { ok:false, reason:"empty", data:null };
    return { ok:true, data:old };
  }catch(e){
    return { ok:false, reason:"parse-error", error:e.message, data:null };
  }
}
/* 两个数组的集合差集：onlyA=只在A、onlyB=只在B、both=都有 */
function setDiff(aArr, bArr){
  const setA=new Set(aArr||[]), setB=new Set(bArr||[]);
  return { onlyA:[...setA].filter(x=>!setB.has(x)), onlyB:[...setB].filter(x=>!setA.has(x)),
           both:[...setA].filter(x=>setB.has(x)) };
}
function makeBreakthroughMap(list){
  const map=new Map();
  (list||[]).forEach(function(bt){ map.set(bt.sword+"||"+bt.rarity+"||"+bt.fromStage+"||"+bt.toStage, bt); });
  return map;
}
function makeRecipeMap(list){
  const map=new Map();
  (list||[]).forEach(function(r){ map.set(r.product+"||"+r.ingredient, r); });
  return map;
}
/* 数组按「集合」比较（不管顺序），关卡掉落用它 */
function arrEqualUnordered(a,b){
  const s1=new Set(a||[]), s2=new Set(b||[]);
  if(s1.size!==s2.size) return false;
  for(const x of s1) if(!s2.has(x)) return false;
  return true;
}
/* 生成完整差异报告 markdown */
function buildDiffReport(oldData, newData){
  const d=[];
  d.push("# 新旧数据差异报告（update-diff-report.md）");
  d.push("");
  d.push("> 基准：旧 `"+CONFIG.oldGameDataFile+"`（当前工具在用的）｜对比：本次转换的 "+CONFIG.outputJs);
  d.push("> ⚠ 请先看 convert-report.md 的转换警告，确认表格解析无误后，再审阅本差异；全部确认后再替换正式 "+CONFIG.oldGameDataFile);
  d.push("");
  /* --- 总览 --- */
  const oChar=(oldData.characters||[]).length, nChar=(newData.characters||[]).length;
  const oMat=Object.keys(oldData.materials||{}).length, nMat=Object.keys(newData.materials||{}).length;
  const oBt=(oldData.breakthroughs||[]).length, nBt=(newData.breakthroughs||[]).length;
  const oRec=(oldData.recipes||[]).length, nRec=(newData.recipes||[]).length;
  const oSt=(oldData.stages||[]).length, nSt=(newData.stages||[]).length;
  const oCoin=Object.keys(oldData.coinCosts||{}).length, nCoin=Object.keys(newData.coinCosts||{}).length;
  d.push("## 📊 总量总览");
  d.push("| 项目 | 旧数量 | 新数量 | 净变化 |");
  d.push("| --- | --- | --- | --- |");
  d.push("| 角色 | "+oChar+" | "+nChar+" | "+(nChar-oChar)+" |");
  d.push("| 材料 | "+oMat+" | "+nMat+" | "+(nMat-oMat)+" |");
  d.push("| 突破阶段记录 | "+oBt+" | "+nBt+" | "+(nBt-oBt)+" |");
  d.push("| 合成配方 | "+oRec+" | "+nRec+" | "+(nRec-oRec)+" |");
  d.push("| 关卡 | "+oSt+" | "+nSt+" | "+(nSt-oSt)+" |");
  d.push("| 小判花费配置 | "+oCoin+" | "+nCoin+" | "+(nCoin-oCoin)+" |");
  d.push("");
  /* --- 1.角色 --- */
  d.push("## 1. 角色 characters");
  const cd=setDiff(oldData.characters, newData.characters);
  d.push("- 新增角色（"+cd.onlyB.length+"）："+(cd.onlyB.length?cd.onlyB.join("、"):"无"));
  d.push("- 删除角色（"+cd.onlyA.length+"）："+(cd.onlyA.length?cd.onlyA.join("、"):"无"));
  d.push("");
  /* --- 2.材料 --- */
  d.push("## 2. 材料 materials");
  const oldMats=oldData.materials||{}, newMats=newData.materials||{};
  const mk=setDiff(Object.keys(oldMats), Object.keys(newMats));
  d.push("- 新增材料（"+mk.onlyB.length+"）："+(mk.onlyB.length?mk.onlyB.join("、"):"无"));
  d.push("- 删除材料（"+mk.onlyA.length+"）："+(mk.onlyA.length?mk.onlyA.join("、"):"无"));
  const changed=[];
  mk.both.forEach(function(k){
    const o=oldMats[k], n=newMats[k], ch=[];
    if(o.ja!==n.ja) ch.push("日文名："+o.ja+" → "+n.ja);
    if(o.zh!==n.zh) ch.push("中文名："+o.zh+" → "+n.zh);
    if(o.type!==n.type) ch.push("类型："+o.type+" → "+n.type);
    if((o.coin||0)!==(n.coin||0)) ch.push("小判："+(o.coin||0)+" → "+(n.coin||0));
    if(ch.length) changed.push(k+"："+ch.join("；"));
  });
  if(changed.length){ d.push("### 材料属性变化"); changed.forEach(function(x){ d.push("- "+x); }); }
  else d.push("- 共同存在的材料：属性无变化");
  d.push("");
  /* --- 3.突破阶段 --- */
  d.push("## 3. 突破阶段 breakthroughs");
  const oldBtMap=makeBreakthroughMap(oldData.breakthroughs), newBtMap=makeBreakthroughMap(newData.breakthroughs);
  const bk=setDiff([...oldBtMap.keys()], [...newBtMap.keys()]);
  d.push("### 3.1 整条阶段新增（"+bk.onlyB.length+"）");
  if(bk.onlyB.length) bk.onlyB.forEach(function(k){ const v=newBtMap.get(k); d.push("- "+v.sword+"【"+v.rarity+"】 "+v.fromStage+"→"+v.toStage); });
  else d.push("- 无");
  d.push("### 3.2 整条阶段删除（"+bk.onlyA.length+"）");
  if(bk.onlyA.length) bk.onlyA.forEach(function(k){ const v=oldBtMap.get(k); d.push("- "+v.sword+"【"+v.rarity+"】 "+v.fromStage+"→"+v.toStage); });
  else d.push("- 无");
  d.push("### 3.3 已有阶段内部材料需求变化");
  let hasNeed=false;
  bk.both.forEach(function(k){
    const o=oldBtMap.get(k), n=newBtMap.get(k), oN=o.needs||{}, nN=n.needs||{};
    const nk=setDiff(Object.keys(oN), Object.keys(nN));
    const lines=[];
    nk.both.forEach(function(m){ if(oN[m]!==nN[m]) lines.push(m+"："+oN[m]+" → "+nN[m]); });
    if(nk.onlyB.length) lines.push("新增材料："+nk.onlyB.join("、"));
    if(nk.onlyA.length) lines.push("移除材料："+nk.onlyA.join("、"));
    if(lines.length){ hasNeed=true; d.push("- **"+n.sword+"【"+n.rarity+"】 "+n.fromStage+"→"+n.toStage+"**"); lines.forEach(function(x){ d.push("　- "+x); }); }
  });
  if(!hasNeed) d.push("- 无");
  d.push("");
  /* --- 4.配方 --- */
  d.push("## 4. 合成配方 recipes");
  const oldRecMap=makeRecipeMap(oldData.recipes), newRecMap=makeRecipeMap(newData.recipes);
  const rk=setDiff([...oldRecMap.keys()], [...newRecMap.keys()]);
  d.push("### 4.1 新增配方（"+rk.onlyB.length+"）");
  if(rk.onlyB.length) rk.onlyB.forEach(function(k){ const r=newRecMap.get(k); d.push("- "+r.product+" ← "+r.ingredient+" ×"+r.quantity); });
  else d.push("- 无");
  d.push("### 4.2 删除配方（"+rk.onlyA.length+"）");
  if(rk.onlyA.length) rk.onlyA.forEach(function(k){ const r=oldRecMap.get(k); d.push("- "+r.product+" ← "+r.ingredient+" ×"+r.quantity); });
  else d.push("- 无");
  d.push("### 4.3 配料数量变化");
  const qch=[];
  rk.both.forEach(function(k){
    const o=oldRecMap.get(k), n=newRecMap.get(k);
    if(o.quantity!==n.quantity) qch.push(n.product+" ← "+n.ingredient+"："+o.quantity+" → "+n.quantity);
  });
  if(qch.length) qch.forEach(function(x){ d.push("- "+x); });
  else d.push("- 无");
  d.push("");
  /* --- 5.关卡 --- */
  d.push("## 5. 关卡 stages");
  const oldStMap=new Map((oldData.stages||[]).map(function(s){return [s.stage,s];}));
  const newStMap=new Map((newData.stages||[]).map(function(s){return [s.stage,s];}));
  const sk=setDiff([...oldStMap.keys()], [...newStMap.keys()]);
  d.push("- 新增关卡（"+sk.onlyB.length+"）："+(sk.onlyB.length?sk.onlyB.join("、"):"无"));
  d.push("- 删除关卡（"+sk.onlyA.length+"）："+(sk.onlyA.length?sk.onlyA.join("、"):"无"));
  d.push("### 掉落材料清单变化（集合比较，不看顺序）");
  let hasDrop=false;
  sk.both.forEach(function(k){
    const o=oldStMap.get(k), n=newStMap.get(k);
    if(!arrEqualUnordered(o.materials, n.materials)){
      hasDrop=true;
      d.push("- **"+k+"**");
      d.push("　旧掉落："+((o.materials||[]).join("、")||"(空)"));
      d.push("　新掉落："+((n.materials||[]).join("、")||"(空)"));
    }
  });
  if(!hasDrop) d.push("- 无");
  d.push("");
  /* --- 6.小判 --- */
  d.push("## 6. 小判花费 coinCosts");
  const oldCoin=oldData.coinCosts||{}, newCoin=newData.coinCosts||{};
  const ck=setDiff(Object.keys(oldCoin), Object.keys(newCoin));
  d.push("- 新增（"+ck.onlyB.length+"）："+(ck.onlyB.length?ck.onlyB.join("、"):"无"));
  d.push("- 删除（"+ck.onlyA.length+"）："+(ck.onlyA.length?ck.onlyA.join("、"):"无"));
  const cch=[];
  ck.both.forEach(function(k){ if(oldCoin[k]!==newCoin[k]) cch.push(k+"："+oldCoin[k]+" → "+newCoin[k]); });
  if(cch.length){ d.push("### 数值变化"); cch.forEach(function(x){ d.push("- "+x); }); }
  d.push("");
  d.push("---");
  d.push("*报告由转换脚本自动生成，请人工复核全部变更后再替换正式 "+CONFIG.oldGameDataFile+"。*");
  return d.join("\n");
}

let diffGenerated=false, diffWarnText="";
if(oldRes.ok){
  fs.writeFileSync(
    CONFIG.diffReport,
    buildDiffReport(oldRes.data, dataObj),
    "utf8",
  );
  diffGenerated=true;
}else if(oldRes.reason==="file-not-exist"){
  diffWarnText = "ℹ 未找到旧的 "+CONFIG.oldGameDataFile+"，跳过差异报告（首次转换属于正常现象）";
}else{
  diffWarnText = "⚠ 读取旧的 "+CONFIG.oldGameDataFile+" 失败（"+(oldRes.error||"数据为空")+"），跳过差异报告，请检查旧文件";
}

/* =====================================================================
 * 第 7 步：写转换检查报告 convert-report.md
 * ===================================================================== */
const craftedList = Object.keys(materials).filter(function(n){return materials[n].type==="crafted";});
const baseList = Object.keys(materials).filter(function(n){return materials[n].type==="base";});

report.push("# 转换检查报告（convert-pazugiri.js）");
report.push("");
report.push("**转换文件**：" + excelFile);
report.push("**转换时间**："+dataObj.generatedAt);
report.push("**输出文件**："+CONFIG.outputJs+"（人工核对后再替换 game-data.js）");
if(diffGenerated){
  report.push("**同时生成差异报告**："+CONFIG.diffReport+"，请审阅本次新旧数据变化");
}else{
  report.push(diffWarnText);
}
report.push("");
report.push("## 一、转换概览");
report.push("");
report.push("| 项目 | 数量 |");
report.push("| --- | --- |");
report.push("| 刀男角色 | "+chars.length+" 把 |");
report.push("| 突破阶段记录 | "+breakthroughs.length+" 条 |");
report.push("| 材料总数 | "+Object.keys(materials).length+" 个（基础 "+baseList.length+" / 半成品 "+craftedList.length+"） |");
report.push("| 合成配方条数 | "+recipes.length+" 条（"+Object.keys(products).length+" 个成品） |");
report.push("| 关卡（掉落清单） | "+stages.length+" 关 |");
report.push("");
report.push("### 角色与稀有度分布");
const blocks={};
breakthroughs.forEach(function(b){ const k=b.sword+"（"+b.rarity+"）"; blocks[k]=1; });
report.push("");
report.push("- 刀男总数："+chars.length+" 把");
report.push("- 稀有度组合块："+Object.keys(blocks).length+" 个（普/彩分开算；5 把刀有普、彩两套：太郎太刀、三日月宗近、一期一振、歌仙兼定、乱藤四郎）");
report.push("- 每块都是 3 个阶段（10→20 / 20→30 / 30→35），共 "+breakthroughs.length+" 条");
report.push("");
report.push("### 材料中日对照表");
report.push("");
report.push("| 中文名 | 日文名 | 显示 | 类型 | 合成花费(小判) |");
report.push("| --- | --- | --- | --- | --- |");
Object.keys(materials).sort().forEach(function(n){
  const m=materials[n];
  const show = (m.ja===m.zh) ? m.zh : m.ja+"|"+m.zh;
  report.push("| "+n+" | "+m.ja+" | "+show+" | "+m.type+" | "+(m.coin?("小判×"+m.coin):"—")+" |");
});
report.push("");
report.push("### 小判消耗清单（货币，不进材料缺口）");
report.push("");
if(Object.keys(coinCosts).length){
  Object.keys(coinCosts).forEach(function(cn){ report.push("- "+cn+"：小判×"+coinCosts[cn]); });
} else {
  report.push("- 无");
}
report.push("");
report.push("### 关卡掉落清单（"+stages.length+" 关）");
report.push("");
report.push("| 关卡 | 掉落材料 |");
report.push("| --- | --- |");
stages.forEach(function(s){ report.push("| "+s.stage+" | "+s.materials.join("、")+" |"); });
report.push("");
report.push("## 二、⚠ 警告清单（转换前请逐条核对）");
report.push("");
if(warnings.length){ warnings.forEach(function(w){ report.push("- ⚠ "+w); }); }
else { report.push("- 无警告"); }
report.push("");
if(aliasFixes.length){
  report.push("### 别名归一化（同一材料的不同写法，已统一成标准中文名）");
  report.push("");
  aliasFixes.forEach(function(a){ report.push("- "+a); });
  report.push("");
}
if(emptyNeeds.length){
  report.push("### 需求为空的阶段（数据未填写，按 0 需求处理）");
  report.push("");
  report.push("共 "+emptyNeeds.length+" 条：");
  report.push("");
  emptyNeeds.forEach(function(e){ report.push("- "+e); });
  report.push("");
}
if(dropUnmapped && Object.keys(dropUnmapped).length){
  report.push("### 关卡掉落里没找到中文名的材料");
  report.push("");
  Object.keys(dropUnmapped).forEach(function(n){ report.push("- "+n+"（出现关卡："+dropUnmapped[n].join("、")+"）"); });
  report.push("");
}
if(recipeSkipRows.length){
  report.push("### 表「材料合成」跳过的行");
  report.push("");
  recipeSkipRows.forEach(function(s){ report.push("- "+s); });
  report.push("");
}
if(charSkipRows.length){
  report.push("### 表「刀男所需材料」跳过的行");
  report.push("");
  charSkipRows.forEach(function(s){ report.push("- "+s); });
  report.push("");
}
if(stageSkipRows.length){
  report.push("### 表「关卡统计」跳过的行");
  report.push("");
  stageSkipRows.forEach(function(s){ report.push("- "+s); });
  report.push("");
}
report.push("## 三、人工检查点（建议逐条确认）");
report.push("");
report.push("- [ ] 材料数量：基础材料是否都能在关卡掉落里找到（除了颜料这类未填写来源的）");
report.push("- [ ] 半成品列表是否齐全：茶碗/围棋/双色牡丹饼/算盘/花札/团子/锄头/花蜜/锤子/纸/笔/臼与杵/书写道具/天妇罗/双六/味增汤/味噌/面包/各种心得");
report.push("- [ ] 合成递归：茶碗 → 碗(お椀) → 粘土+打火石，应能一层层拆到底");
report.push("- [ ] 突破筛选：5→35 应算 10→20、20→30、30→35 三段（起点>当前 且 起点<目标）");
report.push("- [ ] 普/彩两套的刀：太郎太刀、三日月宗近、一期一振、歌仙兼定、乱藤四郎，选稀有度后需求应不同");
report.push("- [ ] 关卡掉落是「有没有」清单，不是掉率；关卡 x/n 只表示覆盖几种材料");
report.push("");
report.push("---");
report.push("");
report.push("*报告由脚本自动生成，最终以人工核对为准。*");
fs.writeFileSync(CONFIG.outputReport, report.join("\n"), "utf8");

/* ---------- 控制台摘要 ---------- */
console.log("转换完成 ✅");
console.log("  刀男 "+chars.length+" 把 · 突破阶段 "+breakthroughs.length+" 条 · 材料 "+Object.keys(materials).length+" 个");
console.log("  配方 "+recipes.length+" 条（"+Object.keys(products).length+" 个成品）· 关卡 "+stages.length+" 关");
console.log("  警告 "+warnings.length+" 条"+(warnings.length?"（请打开 convert-report.md 查看！）":""));
if(diffGenerated) console.log("  ✅已生成新旧差异报告："+CONFIG.diffReport);
else console.log("  "+diffWarnText);
console.log("输出文件："+CONFIG.outputJs+" / "+CONFIG.outputReport+" / "+CONFIG.diffReport);
//（注：内容由AI生成）
