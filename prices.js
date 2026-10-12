// ============================================================
// アプリ用 単価ファイル（客出し単価のみ）
// ※ このファイルは tools/build_prices.py が原価マスターから自動生成します。
//    手で直さず、master/cost-master.xlsx を直して作り直してください。
//    原価・掛率はこのファイルには含まれません。
//    生成日: 2026-10-12
// ============================================================
window.PRICE_TABLE = {
 "version": 2,
 "generatedAt": "2026-10-12",
 "settings": {
  "company": "株式会社ワイケイアート",
  "companyInfo": "",
  "overheadRate": 0.0,
  "taxRate": 0.0,
  "defaultNote": "※本書は現地調査時の概算見積です。詳細調査・商品選定後に正式見積をご提出します。\n※解体後に下地の腐食等が見つかった場合、別途費用が発生することがあります。"
 },
 "items": {
  "floor_over": {
   "name": "フローリング増し張り",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 9000
  },
  "door_replace": {
   "name": "建具交換",
   "unit": "枚",
   "price": 80000
  },
  "cloth": {
   "name": "クロス張替え",
   "unit": "㎡",
   "qtyFrom": "cloth",
   "price": 1200
  },
  "cf": {
   "name": "クッションフロア張替え",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 4000
  },
  "habaki": {
   "name": "巾木",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 1200
  },
  "mawaribuchi": {
   "name": "周り縁撤去",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 800
  },
  "window_paint": {
   "name": "窓枠塗装",
   "unit": "箇所",
   "price": 15000
  },
  "window_dinoc": {
   "name": "窓枠ダイノック",
   "unit": "箇所",
   "price": 25000
  },
  "light": {
   "name": "照明交換",
   "unit": "箇所",
   "price": 15000
  },
  "knob": {
   "name": "ドアノブ交換",
   "unit": "箇所",
   "price": 8000
  },
  "level": {
   "name": "レベル工事（床）",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 3000
  },
  "ac_outlet": {
   "name": "エアコンコンセント",
   "unit": "箇所",
   "price": 18000
  },
  "okabe": {
   "name": "大壁化（天組・壁組・床組）",
   "unit": "一式",
   "onlyIf": {
    "wall": "shinkabe"
   },
   "note": "真壁の和室のみ",
   "price": 250000
  },
  "tatami_omote": {
   "name": "畳 表替え",
   "unit": "枚",
   "qtyFrom": "tatami",
   "price": 7000
  },
  "fusuma": {
   "name": "襖 張替え",
   "unit": "枚",
   "qtyDefault": 2,
   "price": 6000
  },
  "shoji": {
   "name": "障子 張替え",
   "unit": "枚",
   "qtyDefault": 2,
   "price": 4000
  },
  "aku": {
   "name": "木部あく抜き・オイルステイン",
   "unit": "一式",
   "price": 60000
  },
  "ceil_over": {
   "name": "天井増し張り",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 6000
  },
  "wall_over": {
   "name": "壁増し張り",
   "unit": "㎡",
   "qtyFrom": "walls",
   "price": 6000
  },
  "oshiire_cf": {
   "name": "押入れCF",
   "unit": "一式",
   "price": 15000
  },
  "oshiire_over": {
   "name": "押入れ増し張り",
   "unit": "一式",
   "price": 40000
  },
  "wa_to_yo": {
   "name": "和室から洋室（一式セット）",
   "unit": "一式",
   "note": "大壁化・床・クロス等をまとめた仮単価",
   "price": 350000
  },
  "genkan_door": {
   "name": "玄関ドア交換（ワイケイ材工）",
   "unit": "一式",
   "price": 350000
  },
  "key": {
   "name": "鍵交換",
   "unit": "箇所",
   "price": 25000
  },
  "kamachi": {
   "name": "框（上がり框）",
   "unit": "一式",
   "price": 45000
  },
  "getabako": {
   "name": "下駄箱交換",
   "unit": "一式",
   "price": 80000
  },
  "kitchen": {
   "name": "システムキッチン交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 650000
    },
    {
     "name": "上位",
     "price": 900000
    }
   ]
  },
  "kitchen_faucet": {
   "name": "キッチン水栓交換",
   "unit": "箇所",
   "price": 35000
  },
  "bath_zairai_ub": {
   "name": "在来浴室 → ユニットバス化（解体含む）",
   "unit": "一式",
   "onlyIf": {
    "bath": "zairai"
   },
   "grades": [
    {
     "name": "標準",
     "price": 1000000
    },
    {
     "name": "上位",
     "price": 1300000
    }
   ]
  },
  "bath_ub": {
   "name": "ユニットバス交換",
   "unit": "一式",
   "onlyIf": {
    "bath": "ub"
   },
   "grades": [
    {
     "name": "標準",
     "price": 850000
    },
    {
     "name": "上位",
     "price": 1100000
    }
   ]
  },
  "bath_door": {
   "name": "浴室ドア交換",
   "unit": "箇所",
   "price": 70000
  },
  "bath_fan": {
   "name": "浴室換気扇交換",
   "unit": "箇所",
   "price": 35000
  },
  "toilet": {
   "name": "トイレ交換（便器＋温水洗浄便座）",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 180000
    },
    {
     "name": "上位",
     "price": 280000
    }
   ]
  },
  "toilet_interior": {
   "name": "トイレ内装（クロス・CF）",
   "unit": "一式",
   "price": 45000
  },
  "washstand": {
   "name": "洗面化粧台交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 120000
    },
    {
     "name": "上位",
     "price": 200000
    }
   ]
  },
  "wash_interior": {
   "name": "洗面所内装（クロス・CF）",
   "unit": "一式",
   "price": 50000
  },
  "ext_wall": {
   "name": "外壁のみ塗装（足場込）",
   "unit": "一式",
   "price": 500000
  },
  "ext_wall_roof": {
   "name": "外壁・屋根塗装（足場込）",
   "unit": "一式",
   "price": 562500
  },
  "ext_wash": {
   "name": "高圧洗浄のみ（足場込）",
   "unit": "一式",
   "price": 100000
  },
  "carport_rm": {
   "name": "カーポート撤去",
   "unit": "台",
   "price": 60000
  },
  "gate_rm": {
   "name": "門扉撤去",
   "unit": "一式",
   "price": 40000
  },
  "weed_sheet": {
   "name": "防草シート＋砕石",
   "unit": "㎡",
   "qtyDefault": 10,
   "price": 3500
  },
  "cut_tree": {
   "name": "伐採",
   "unit": "本",
   "price": 15000
  },
  "root_rm": {
   "name": "抜根",
   "unit": "本",
   "price": 20000
  },
  "stone_rm": {
   "name": "天然石撤去",
   "unit": "一式",
   "price": 50000
  },
  "shed_rm": {
   "name": "物置撤去",
   "unit": "台",
   "price": 40000
  },
  "nureen_rm": {
   "name": "濡れ縁撤去",
   "unit": "一式",
   "price": 30000
  },
  "od_ceiling": {
   "name": "照明器具（シーリングライト）",
   "unit": "台",
   "price": 8000
  },
  "od_downlight": {
   "name": "照明器具（ダウンライト）",
   "unit": "台",
   "price": 6000
  },
  "od_porch": {
   "name": "玄関灯・外灯",
   "unit": "台",
   "price": 12000
  },
  "od_switch": {
   "name": "スイッチ・コンセント交換",
   "unit": "箇所",
   "price": 3500
  },
  "od_panel": {
   "name": "分電盤",
   "unit": "台",
   "price": 60000
  },
  "od_fan": {
   "name": "換気扇",
   "unit": "台",
   "price": 15000
  },
  "od_intercom": {
   "name": "インターホン（モニター付）",
   "unit": "台",
   "price": 25000
  },
  "od_antenna": {
   "name": "TVアンテナ",
   "unit": "一式",
   "price": 50000
  },
  "od_aircon": {
   "name": "エアコン本体",
   "unit": "台",
   "price": 90000
  },
  "od_lumber": {
   "name": "構造材・下地材",
   "unit": "一式",
   "price": 50000
  },
  "od_flooring": {
   "name": "フローリング材",
   "unit": "㎡",
   "price": 4500
  },
  "od_habaki": {
   "name": "巾木材",
   "unit": "m",
   "price": 500
  },
  "od_mawari": {
   "name": "廻り縁材",
   "unit": "m",
   "price": 500
  },
  "od_door": {
   "name": "建具（室内ドア）",
   "unit": "枚",
   "price": 45000
  },
  "od_plywood": {
   "name": "合板・ベニヤ",
   "unit": "枚",
   "price": 2500
  },
  "od_zosaku": {
   "name": "造作材",
   "unit": "一式",
   "price": 30000
  },
  "od_cloth": {
   "name": "クロス材",
   "unit": "㎡",
   "price": 400
  },
  "od_cf": {
   "name": "クッションフロア材",
   "unit": "㎡",
   "price": 1500
  },
  "od_board": {
   "name": "石膏ボード",
   "unit": "枚",
   "price": 800
  },
  "od_insul": {
   "name": "断熱材",
   "unit": "㎡",
   "price": 1200
  },
  "od_tatami": {
   "name": "畳（新調）",
   "unit": "枚",
   "price": 12000
  },
  "od_sash": {
   "name": "サッシ・窓",
   "unit": "箇所",
   "price": 80000
  },
  "od_siding": {
   "name": "外壁材・サイディング",
   "unit": "㎡",
   "price": 5000
  },
  "od_paint": {
   "name": "塗料",
   "unit": "缶",
   "price": 25000
  },
  "od_key_cyl": {
   "name": "玄関錠交換（シリンダー）",
   "unit": "箇所",
   "price": 20000
  },
  "od_key_dim": {
   "name": "玄関錠交換（ディンプル）",
   "unit": "箇所",
   "price": 35000
  },
  "od_key_back": {
   "name": "勝手口錠交換",
   "unit": "箇所",
   "price": 15000
  },
  "od_key_room": {
   "name": "室内錠交換",
   "unit": "箇所",
   "price": 6000
  },
  "od_key_copy": {
   "name": "合鍵作成",
   "unit": "本",
   "price": 1500
  },
  "od_jk_pump": {
   "name": "浄化槽 汲み取り",
   "unit": "一式",
   "price": 30000
  },
  "od_jk_clean": {
   "name": "浄化槽 清掃・点検",
   "unit": "一式",
   "price": 15000
  },
  "od_jk_remove": {
   "name": "浄化槽 撤去・埋め戻し",
   "unit": "一式",
   "price": 150000
  },
  "od_sewer": {
   "name": "下水接続工事",
   "unit": "一式",
   "price": 300000
  },
  "od_carp": {
   "name": "大工手間",
   "unit": "人工",
   "price": 25000
  },
  "od_interior": {
   "name": "内装手間",
   "unit": "人工",
   "price": 22000
  },
  "od_misc": {
   "name": "雑工・養生",
   "unit": "人工",
   "price": 18000
  },
  "od_demo": {
   "name": "解体手間",
   "unit": "人工",
   "price": 20000
  },
  "od_haul": {
   "name": "運搬・搬出",
   "unit": "回",
   "price": 15000
  },
  "od_manage": {
   "name": "現場管理費",
   "unit": "一式",
   "price": 50000
  }
 },
 "roomTypes": [
  {
   "id": "genkan",
   "name": "玄関",
   "icon": "🚪",
   "fixed": true,
   "noSettings": true,
   "defaultTatami": 1,
   "items": [
    "genkan_door",
    "light",
    "habaki",
    "mawaribuchi",
    "key",
    "kamachi",
    "cloth",
    "getabako"
   ]
  },
  {
   "id": "hall1",
   "name": "1階廊下",
   "icon": "🚶",
   "fixed": true,
   "noSettings": true,
   "defaultTatami": 3,
   "items": [
    "floor_over",
    "cloth",
    "habaki",
    "light",
    "level"
   ]
  },
  {
   "id": "stairs",
   "name": "階段",
   "icon": "🪜",
   "fixed": true,
   "floor2": true,
   "noSettings": true,
   "defaultTatami": 3,
   "items": [
    "cloth",
    "habaki",
    "light"
   ]
  },
  {
   "id": "hall2",
   "name": "2階廊下",
   "icon": "🚶",
   "fixed": true,
   "floor2": true,
   "noSettings": true,
   "defaultTatami": 3,
   "items": [
    "floor_over",
    "cloth",
    "habaki",
    "light",
    "level"
   ]
  },
  {
   "id": "kitchen",
   "name": "キッチン",
   "icon": "🍳",
   "fixed": true,
   "size": true,
   "defaultTatami": 6,
   "items": [
    "kitchen",
    "kitchen_faucet",
    "cloth",
    "cf",
    "floor_over",
    "light"
   ]
  },
  {
   "id": "bath",
   "name": "浴室",
   "icon": "🛁",
   "fixed": true,
   "bath": true,
   "defaultTatami": 2,
   "items": [
    "bath_zairai_ub",
    "bath_ub",
    "bath_door",
    "bath_fan"
   ]
  },
  {
   "id": "toilet1",
   "name": "1階トイレ",
   "icon": "🚽",
   "fixed": true,
   "defaultTatami": 1,
   "items": [
    "toilet",
    "toilet_interior",
    "light"
   ]
  },
  {
   "id": "wash",
   "name": "洗面脱衣室",
   "icon": "🚰",
   "fixed": true,
   "defaultTatami": 2,
   "items": [
    "washstand",
    "wash_interior",
    "light"
   ]
  },
  {
   "id": "living",
   "name": "リビング",
   "icon": "🛋",
   "size": true,
   "defaultCount": 1,
   "defaultTatami": 10,
   "items": [
    "floor_over",
    "door_replace",
    "cloth",
    "habaki",
    "window_paint",
    "window_dinoc",
    "light",
    "knob",
    "level",
    "ac_outlet"
   ]
  },
  {
   "id": "yo",
   "name": "洋室",
   "icon": "🛏",
   "size": true,
   "defaultCount": 2,
   "defaultTatami": 6,
   "items": [
    "floor_over",
    "door_replace",
    "cloth",
    "habaki",
    "window_paint",
    "window_dinoc",
    "light",
    "knob",
    "level",
    "ac_outlet"
   ]
  },
  {
   "id": "wa",
   "name": "和室",
   "icon": "🎋",
   "size": true,
   "defaultCount": 2,
   "defaultTatami": 6,
   "wall": true,
   "items": [
    "okabe",
    "tatami_omote",
    "fusuma",
    "shoji",
    "aku",
    "knob",
    "light",
    "ceil_over",
    "wall_over",
    "oshiire_cf",
    "oshiire_over",
    "cloth"
   ]
  },
  {
   "id": "nando",
   "name": "納戸",
   "icon": "📦",
   "size": true,
   "defaultCount": 0,
   "defaultTatami": 3,
   "items": [
    "floor_over",
    "cf",
    "cloth",
    "habaki",
    "door_replace",
    "light",
    "knob"
   ]
  },
  {
   "id": "toilet_add",
   "name": "トイレ（追加）",
   "icon": "🚽",
   "defaultCount": 0,
   "defaultTatami": 1,
   "base": "トイレ",
   "startNo": 2,
   "secondFloorFirst": "2階トイレ",
   "items": [
    "toilet",
    "toilet_interior",
    "light"
   ]
  },
  {
   "id": "wash_add",
   "name": "洗面脱衣室（追加）",
   "icon": "🚰",
   "defaultCount": 0,
   "defaultTatami": 2,
   "base": "洗面脱衣室",
   "startNo": 2,
   "items": [
    "washstand",
    "wash_interior",
    "light"
   ]
  }
 ],
 "interiorCommon": {
  "name": "中工事 共通・その他",
  "icon": "🧰",
  "items": [
   "wa_to_yo",
   "floor_over",
   "habaki",
   "mawaribuchi"
  ]
 },
 "orders": [
  {
   "id": "elec",
   "name": "電気関係（照明など）",
   "icon": "💡",
   "items": [
    "od_ceiling",
    "od_downlight",
    "od_porch",
    "od_switch",
    "od_panel",
    "od_fan",
    "od_intercom",
    "od_antenna",
    "od_aircon"
   ]
  },
  {
   "id": "wood",
   "name": "木材関係",
   "icon": "🪵",
   "items": [
    "od_lumber",
    "od_flooring",
    "od_habaki",
    "od_mawari",
    "od_door",
    "od_plywood",
    "od_zosaku"
   ]
  },
  {
   "id": "mat",
   "name": "建材関係",
   "icon": "🧱",
   "items": [
    "od_cloth",
    "od_cf",
    "od_board",
    "od_insul",
    "od_tatami",
    "od_sash",
    "od_siding",
    "od_paint"
   ]
  },
  {
   "id": "key",
   "name": "鍵交換",
   "icon": "🔑",
   "items": [
    "od_key_cyl",
    "od_key_dim",
    "od_key_back",
    "od_key_room",
    "od_key_copy"
   ]
  },
  {
   "id": "jk",
   "name": "浄化槽・汲み取り",
   "icon": "🚰",
   "items": [
    "od_jk_pump",
    "od_jk_clean",
    "od_jk_remove",
    "od_sewer"
   ]
  },
  {
   "id": "inhouse",
   "name": "自社工事部",
   "icon": "👷",
   "items": [
    "od_carp",
    "od_interior",
    "od_misc",
    "od_demo",
    "od_haul",
    "od_manage"
   ]
  }
 ],
 "exterior": [
  {
   "id": "paint",
   "name": "外壁塗装",
   "icon": "🎨",
   "exclusive": true,
   "note": "どれか1つを選択。足場代はすべて含みます",
   "items": [
    "ext_wall",
    "ext_wall_roof",
    "ext_wash"
   ]
  },
  {
   "id": "gaikou",
   "name": "外構",
   "icon": "🌳",
   "items": [
    "carport_rm",
    "gate_rm",
    "weed_sheet",
    "cut_tree",
    "root_rm",
    "stone_rm",
    "shed_rm",
    "nureen_rm"
   ]
  }
 ]
};
