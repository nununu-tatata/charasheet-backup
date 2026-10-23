// ==========================================
// チャットパレット・プロフィール生成ロジック（スプレッドシート統合版）
// ==========================================

// ---- メモ欄から特徴・特記を抽出 ----
function extractTraitsFromMemo(rawMemo) {
  const memoLines = rawMemo.split(/\r\n|\r|\n/);
  const traitKeywords = [
    "君は風邪を引かない", "大きな体", "素早い", "オシャレ", "天才", "強固な意志", "勉強家", "幸運のお守り",
    "一族伝来の宝物", "予期せぬ協力者", "手先が器用", "影が薄い", "親の七光り", "愛書家", "鋭い洞察力",
    "アウトドア派", "珍しい技能", "芸術的才能", "バイリンガル", "前職", "天気予報士", "プロ・ドライバー",
    "飛ばし屋", "戦士", "銃火器の達人", "格闘センスの持ち主", "俊敏", "信頼のおける人", "スポーツ万能",
    "平凡な容姿", "目つきが悪い", "方向音痴", "異性が苦手", "動物に嫌われる", "不思議ちゃん", "寄せ餌",
    "眼鏡をかけている", "大切なもの", "暗黒の祖先", "夜に弱い", "動物に好かれる", "斜め上からの発想",
    "失敗は発明の母", "ペット", "おおらか", "異物への耐性", "潜水の名人", "大酒飲み", "ド根性", "受け身",
    "奇妙な幸運", "投擲の才能", "鋼の筋力", "実は生きていた", "急所を見抜く", "急速な回復力",
    "不屈の精神力", "マニア・コレクター", "行方不明の家族", "好意を寄せられている"
  ];

  let extractedTraits = [];
  let extractedTokki = [];
  let isTraitMode = false;
  let isTokkiMode = false;
  let consecutiveEmptyLines = 0;

  for (let i = 0; i < memoLines.length; i++) {
    const rawLine = memoLines[i];
    const tLine = rawLine.trim();

    if (!tLine) {
      consecutiveEmptyLines++;
      if (consecutiveEmptyLines >= 1) { isTraitMode = false; isTokkiMode = false; }
      continue;
    }
    consecutiveEmptyLines = 0;

    if (tLine.match(/^(.)\1{4,}$/)) { isTraitMode = false; isTokkiMode = false; continue; }

    if (tLine.match(/^[◆■▼●【\[]/) && !tLine.includes("特徴") && !tLine.includes("特記")) {
      if (!traitKeywords.some(k => tLine.includes(k))) { isTraitMode = false; isTokkiMode = false; continue; }
    }

    if (tLine.match(/^[【\[(（<《■◆▼●・\s]*特記/)) {
      isTraitMode = false; isTokkiMode = true;
      if (tLine.replace(/[【】:：\s]/g, "").length > "特記".length) extractedTokki.push(rawLine);
      continue;
    }
    if (tLine.match(/^[【\[(（<《■◆▼●・\s]*特徴/)) {
      isTokkiMode = false; isTraitMode = true;
      if (tLine.replace(/[◆■▼●【】:：\s]/g, "").length > "特徴".length) extractedTraits.push(rawLine);
      continue;
    }

    if (!isTraitMode && !isTokkiMode) {
      for (const key of traitKeywords) {
        if (tLine.includes(key)) { isTraitMode = true; if (!extractedTraits.includes(rawLine)) extractedTraits.push(rawLine); break; }
      }
      continue;
    }

    if (isTraitMode) { if (!extractedTraits.includes(rawLine)) extractedTraits.push(rawLine); }
    else if (isTokkiMode) { extractedTokki.push(rawLine); }
  }
  return { traits: extractedTraits, tokki: extractedTokki };
}

// ---- プロフィールテキスト生成 ----
function generateProfileText_sheet(data, stats, specStats, extractedData, dbValue, getStatValue, settings) {
  const { includeStatusProfile, includeSkillListProfile, includeTraits,
    includeWeaponsProfile, includeItems, includePersonalData } = settings;

  function safeArray(val) { if (!val) return []; return Array.isArray(val) ? val : [val]; }
  function chunkArray(arr, size) {
    const res = [];
    for (let i = 0; i < arr.length; i += size) res.push(arr.slice(i, i + size));
    return res;
  }
  function createWeaponLine(name, hit, dmg, rng, count, ammo, vital, notes) {
    const parts = [name];
    if (hit) parts.push(hit + "%");
    if (dmg) parts.push(dmg);
    if (rng) parts.push("射程:" + rng);
    if (count) parts.push("攻撃回数:" + count);
    if (ammo) parts.push("装弾数:" + ammo);
    if (vital) parts.push("耐久力:" + vital);
    if (notes) parts.push(notes);
    return parts.join("｜");
  }

  let profile = [];
  profile.push(data.pc_name || "名称未設定");
  profile.push(`年齢：${data.age || ""}　職業：${data.shokugyo || data.shuzoku || ""}　性別：${data.sex || ""}`);

  if (includeStatusProfile) {
    let statStrList = stats.map(s => `${s.name}:${s.val}`);
    statStrList.push(`HP:${getStatValue("NA9", "NP9")}`, `MP:${getStatValue("NA10", "NP10")}`, `SAN:${data.SAN_Left || 0}`, `DB:${dbValue || "なし"}`);
    chunkArray(statStrList, 4).forEach(chunk => profile.push(chunk.join(" ")));
  }

  if (includePersonalData) {
    profile.push(`身長:${data.pc_height || ""} 体重:${data.pc_weight || ""} 出身:${data.pc_kigen || ""}`);
    profile.push(`髪の色:${data.color_hair || ""} 瞳の色:${data.color_eye || ""} 肌の色:${data.color_skin || ""}`);
  }

  if (includeSkillListProfile) {
    profile.push("------------------------------");
    profile.push("【技能】");
    const specLine = specStats.filter(s => s.val > 0).map(s => `${s.name}(${s.val})`).join("　");
    if (specLine) profile.push(specLine);
    profile.push("{{SKILL_LIST_PLACEHOLDER}}");
  }

  profile.push("------------------------------");

  if (includeTraits) {
    profile.push("【特徴表】");
    extractedData.traits.forEach(t => profile.push(t));
    profile.push("【特記】");
    let tokkiOut = [];
    if (data.tokki) tokkiOut.push(data.tokki);
    tokkiOut = tokkiOut.concat(extractedData.tokki);
    tokkiOut.forEach(t => profile.push(t));
    profile.push("------------------------------");
  }

  if (includeWeaponsProfile) {
    profile.push("【武器】");
    const names = safeArray(data.arms_name), hits = safeArray(data.arms_hit), dmgs = safeArray(data.arms_damage);
    const rngs = safeArray(data.arms_range), cnts = safeArray(data.arms_attack_count), ammos = safeArray(data.arms_last_shot);
    const vitals = safeArray(data.arms_vitality), notes = safeArray(data.arms_sonota);
    let hasWeapon = false;
    for (let i = 0; i < names.length; i++) {
      if (names[i]) { hasWeapon = true; profile.push(createWeaponLine(names[i], hits[i], dmgs[i], rngs[i], cnts[i], ammos[i], vitals[i], notes[i])); }
    }
    if (!hasWeapon) profile.push("なし");
    profile.push("------------------------------");
  }

  if (includeItems) {
    profile.push("【持ち物】");
    const items = safeArray(data.item_name).filter(i => i);
    if (items.length > 0) chunkArray(items, 3).forEach(c => profile.push(c.join("/")));
    else profile.push("なし");
  }

  return profile.join('\n');
}

// ---- ダメージロール生成 ----
function getDamageRoll_palette(name, dbSuffix, isAddBrackets, prefix) {
  let dice = "";
  if (name.includes("こぶし") || name.includes("パンチ")) dice = "1d3";
  else if (name.includes("キック")) dice = "1d6";
  else if (name.includes("頭突き")) dice = "1d4";
  if (!dice) return null;
  const dName = isAddBrackets ? `【${name}ダメージ】` : `${name}ダメージ`;
  return `${prefix}${dice}${dbSuffix} ${dName}`;
}

// ==========================================
// メイン生成関数
// settings: {
//   diceCommand: 'CCB'|'CC'|'1D100',
//   isOutputAll: boolean,
//   isAddBrackets: boolean,
//   isUseFolding: boolean,       ← K164が'Tekey向け'の場合true
//   includeStatVariables: boolean, ← K164が'Tekey向け'の場合true（//STR=xx）
//   includeStatx5: boolean,      ← N164 ON/OFF
//   includeStatRes: boolean,     ← Q164 ON/OFF
//   isSecret: boolean,           ← T164 ON/OFF（ダイスロールの先頭にSを付けるシークレットダイス）
//   includeStatusProfile, includeSkillListProfile, includeTraits,
//   includeWeaponsProfile, includeItems, includePersonalData
// }
// returns: { palette: string, profile: string, ccfoliaJson: string }
// ==========================================
function generateChatPaletteData(data, url, settings) {
  const baseCommand = settings.diceCommand || "CCB";
  const isOutputAll = settings.isOutputAll || false;
  const isAddBrackets = settings.isAddBrackets || false;
  const isUseFolding = settings.isUseFolding || false;
  const capVal = 100;
  // シークレットダイス：ダイスロールの先頭に S を付ける（:SAN+ などのコマンドには付けない）
  const prefix = settings.isSecret ? "S" : "";
  const command = prefix + baseCommand;

  function getStatValue(naKey, npKey) {
    const cv = parseInt(data[npKey]) || 0;
    const bv = parseInt(data[naKey]) || 0;
    return cv || bv;
  }
  function safeArray(val) {
    if (!val) return [];
    if (Array.isArray(val)) return val;
    if (typeof val === 'object') return Object.values(val);
    return [val];
  }

  const dbValue = data.dmg_bonus || data.db || "";
  const isValidDB = (dbValue && dbValue !== "0" && dbValue !== "なし" && dbValue !== "");
  const dbSuffix = isValidDB ? (dbValue.trim().startsWith("-") ? "{DB}" : "+{DB}") : "";

  const extractedData = extractTraitsFromMemo(data.pc_making_memo || "");

  const stats = [
    { val: getStatValue("NA1", "NP1"), name: "STR" }, { val: getStatValue("NA2", "NP2"), name: "CON" },
    { val: getStatValue("NA3", "NP3"), name: "POW" }, { val: getStatValue("NA4", "NP4"), name: "DEX" },
    { val: getStatValue("NA5", "NP5"), name: "APP" }, { val: getStatValue("NA6", "NP6"), name: "SIZ" },
    { val: getStatValue("NA7", "NP7"), name: "INT" }, { val: getStatValue("NA8", "NP8"), name: "EDU" }
  ];
  const specStats = [
    { val: getStatValue("NA12", "NP12"), name: "アイデア" },
    { val: getStatValue("NA13", "NP13"), name: "幸運" },
    { val: getStatValue("NA14", "NP14"), name: "知識" }
  ];

  // --- チャットパレット生成 ---
  let palette = [];

  // コマンド
  if (isUseFolding) palette.push("###コマンド");
  palette.push(":SAN+", ":SAN-", ":HP+", ":HP-", ":MP+", ":MP-", ":不定=");
  if (isUseFolding) palette.push("###");

  // 空ロール・SAN
  const emptyBracket = isAddBrackets ? " 【】" : "";
  palette.push(`${command}<= ${emptyBracket}`);
  if (data.SAN_Left) {
    const sanName = isAddBrackets ? "【SAN値チェック】" : "SAN値チェック";
    palette.push(`${command}<={SAN} ${sanName}`);
  }

  // アイデア・幸運・知識
  specStats.forEach(stat => {
    if (stat.val > 0) {
      const sName = isAddBrackets ? `【${stat.name}】` : stat.name;
      palette.push(`${command}<=${stat.val} ${sName}`);
    }
  });
  palette.push('----------');

  // 技能カテゴリー
  const subSkillMap = { "製作": "seisaku_bunya", "運転": "unten_bunya", "操縦": "main_souju_norimono", "母国語": "mylang_name", "芸術": "geijutu_bunya" };
  const categories = [
    { name: "戦闘技能", key: "combat", prefix: "TB", defaults: ["回避", "キック", "組み付き", "こぶし（パンチ）", "頭突き", "投擲", "マーシャルアーツ", "拳銃", "サブマシンガン", "ショットガン", "マシンガン", "ライフル"] },
    { name: "探索技能", key: "search", prefix: "TF", defaults: ["応急手当", "鍵開け", "隠す", "隠れる", "聞き耳", "忍び歩き", "写真術", "精神分析", "追跡", "登攀", "図書館", "目星"] },
    { name: "行動技能", key: "action", prefix: "TA", defaults: ["運転", "機械修理", "重機械操作", "乗馬", "水泳", "製作", "操縦", "跳躍", "電気修理", "ナビゲート", "変装"] },
    { name: "交渉技能", key: "negotiation", prefix: "TC", defaults: ["言いくるめ", "信用", "説得", "値切り", "母国語"] },
    { name: "知識技能", key: "knowledge", prefix: "TK", defaults: ["医学", "オカルト", "化学", "クトゥルフ神話", "芸術", "経理", "考古学", "コンピューター", "心理学", "人類学", "生物学", "地質学", "電子工学", "天文学", "博物学", "物理学", "法律", "薬学", "歴史"] }
  ];

  let profileSkills = [];

  categories.forEach(cat => {
    const arrTotal = safeArray(data[cat.prefix + "AP"]);
    const arrDefault = safeArray(data[cat.prefix + "AD"]);
    const arrJob = safeArray(data[cat.prefix + "AS"]);
    const arrInterest = safeArray(data[cat.prefix + "AK"]);
    const addedNames = safeArray(data[cat.prefix + "AName"]);
    let categoryLines = [];

    for (let i = 0; i < arrTotal.length; i++) {
      let val = parseInt(arrTotal[i]);
      if (isNaN(val)) continue;
      if (val > capVal) val = capVal;

      let baseName = "", displayName = "";
      if (i < cat.defaults.length) {
        baseName = cat.defaults[i]; displayName = baseName;
        if (subSkillMap[baseName] && data[subSkillMap[baseName]]) displayName = `${baseName}（${data[subSkillMap[baseName]]}）`;
      } else {
        baseName = addedNames[i - cat.defaults.length] || "不明な技能"; displayName = baseName;
      }

      const jobPoints = parseInt(arrJob[i]) || 0;
      const interestPoints = parseInt(arrInterest[i]) || 0;
      const defaultVal = parseInt(arrDefault[i]);
      let isLearned = (jobPoints + interestPoints > 0) ||
        (!isNaN(defaultVal) && val !== defaultVal) ||
        (i >= cat.defaults.length && val > 0);

      if (isLearned) profileSkills.push(`${displayName}(${val})`);

      if (isOutputAll || isLearned) {
        const finalName = isAddBrackets ? `【${displayName}】` : displayName;
        categoryLines.push(`${command}<=${val} ${finalName}`);
        const dmgLine = getDamageRoll_palette(baseName, dbSuffix, isAddBrackets, prefix);
        if (dmgLine) categoryLines.push(dmgLine);
        if (baseName.includes("応急手当")) {
          const label = isAddBrackets ? "【応急手当回復値】" : "応急手当回復値";
          categoryLines.push(`${prefix}1d3 ${label}`);
        }
      }
    }

    // 武器パレット（戦闘技能カテゴリーに追加）
    if (cat.name === "戦闘技能" && data.arms_name) {
      const wNames = safeArray(data.arms_name), wAccs = safeArray(data.arms_hit), wDmgs = safeArray(data.arms_damage);
      for (let i = 0; i < wNames.length; i++) {
        if (!wNames[i]) continue;
        let acc = parseInt(wAccs[i]);
        if (isNaN(acc)) acc = 0;
        if (acc > capVal) acc = capVal;
        const wLabel = isAddBrackets ? `【${wNames[i]}】` : wNames[i];
        const dLabel = isAddBrackets ? `【${wNames[i]}ダメージ】` : `${wNames[i]}ダメージ`;
        if (wAccs[i]) categoryLines.push(`${command}<=${acc} ${wLabel}`);
        if (wDmgs[i]) categoryLines.push(`${prefix}${wDmgs[i].replace(/DB/gi, "{DB}")} ${dLabel}`);
      }
    }

    if (categoryLines.length > 0) {
      if (isUseFolding) {
        palette.push(`###${cat.name}`);
        palette = palette.concat(categoryLines);
        palette.push("###");
      } else {
        palette = palette.concat(categoryLines);
        palette.push('----------');
      }
    }
  });

  // --- *5 倍数ロール（N164 ON時） ---
  if (settings.includeStatx5) {
    const x5Lines = [];
    stats.forEach(stat => {
      if (!stat.val) return;
      const sName = isAddBrackets ? `【${stat.name}*5】` : `${stat.name}*5`;
      x5Lines.push(`${command}<=({${stat.name}}*5) ${sName}`);
    });
    if (x5Lines.length > 0) {
      if (isUseFolding) palette.push("###倍数ロール");
      else palette.push('----------');
      palette = palette.concat(x5Lines);
      if (isUseFolding) palette.push("###");
    }
  }

  // --- 対抗ロール RESB/RES（Q164 ON時） ---
  if (settings.includeStatRes) {
    const resCmd = baseCommand === 'CC' ? 'RES' : 'RESB';
    const resLines = [];
    stats.forEach(stat => {
      if (!stat.val) return;
      const sName = isAddBrackets ? `【${stat.name}対抗】` : `${stat.name}対抗`;
      resLines.push(`${prefix}${resCmd}({${stat.name}}-対抗値) ${sName}`);
    });
    if (resLines.length > 0) {
      if (isUseFolding) palette.push("###対抗ロール");
      else palette.push('----------');
      palette = palette.concat(resLines);
      if (isUseFolding) palette.push("###");
    }
  }

  // --- 能力値変数（K164が'Tekey向け'の場合、パレット末尾に自動追加） ---
  if (settings.includeStatVariables) {
    const varLines = [];
    stats.forEach(stat => { if (stat.val) varLines.push(`//${stat.name}=${stat.val}`); });
    if (isValidDB) varLines.push(`//DB=${dbValue}`);
    if (varLines.length > 0) {
      if (isUseFolding) palette.push("###能力値変数");
      else palette.push('----------');
      palette = palette.concat(varLines);
      if (isUseFolding) palette.push("###");
    }
  }

  // --- プロフィール生成 ---
  let skillListString = "なし";
  if (profileSkills.length > 0) {
    const chunks = [];
    for (let i = 0; i < profileSkills.length; i += 4) chunks.push(profileSkills.slice(i, i + 4).join("　"));
    skillListString = chunks.join('\n');
  }

  let profileString = generateProfileText_sheet(data, stats, specStats, extractedData, dbValue, getStatValue, settings);
  profileString = profileString.replace("{{SKILL_LIST_PLACEHOLDER}}", skillListString);

  // --- CCFOLIAデータ生成 ---
  const paramsList = stats.map(s => ({ label: s.name, value: String(s.val) }));
  if (isValidDB) paramsList.push({ label: "DB", value: dbValue });

  const ccfoliaData = {
    kind: "character",
    data: {
      name: data.pc_name || "名称未設定",
      memo: profileString,
      initiative: getStatValue("NA4", "NP4"),
      externalUrl: url,
      status: [
        { label: "HP", value: getStatValue("NA9", "NP9"), max: parseInt(data.NA9) || 0 },
        { label: "MP", value: getStatValue("NA10", "NP10"), max: parseInt(data.NA10) || 0 },
        { label: "SAN", value: parseInt(data.SAN_Left) || 0, max: parseInt(data.SAN_Left) || 0 }
      ],
      params: paramsList,
      commands: palette.join('\n')
    }
  };

  return {
    palette: palette.join('\n'),
    profile: profileString,
    ccfoliaJson: JSON.stringify(ccfoliaData)
  };
}
