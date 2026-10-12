// ============================================================
// アプリ用 単価ファイル（客出し単価＋暗号化した原価）
// ※ tools/build_prices.py（Excelから）または アプリの原価マスター画面から生成します。
//    手で直さないでください。原価は合言葉がないと読めません。
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
 "marginTarget": 17.0,
 "items": {
  "floor_over": {
   "name": "フローリング増し張り",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 8700
  },
  "door_replace": {
   "name": "建具交換",
   "unit": "枚",
   "price": 77200
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
   "price": 3900
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
   "price": 14500
  },
  "window_dinoc": {
   "name": "窓枠ダイノック",
   "unit": "箇所",
   "price": 24100
  },
  "light": {
   "name": "照明交換",
   "unit": "箇所",
   "price": 14500
  },
  "knob": {
   "name": "ドアノブ交換",
   "unit": "箇所",
   "price": 7800
  },
  "level": {
   "name": "レベル工事（床）",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 2900
  },
  "ac_outlet": {
   "name": "エアコンコンセント",
   "unit": "箇所",
   "price": 17400
  },
  "okabe": {
   "name": "大壁化（天組・壁組・床組）",
   "unit": "一式",
   "onlyIf": {
    "wall": "shinkabe"
   },
   "note": "真壁の和室のみ",
   "price": 241000
  },
  "tatami_omote": {
   "name": "畳 表替え",
   "unit": "枚",
   "qtyFrom": "tatami",
   "price": 6800
  },
  "fusuma": {
   "name": "襖 張替え",
   "unit": "枚",
   "qtyDefault": 2,
   "price": 5800
  },
  "shoji": {
   "name": "障子 張替え",
   "unit": "枚",
   "qtyDefault": 2,
   "price": 3900
  },
  "aku": {
   "name": "木部あく抜き・オイルステイン",
   "unit": "一式",
   "price": 57900
  },
  "ceil_over": {
   "name": "天井増し張り",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 5800
  },
  "wall_over": {
   "name": "壁増し張り",
   "unit": "㎡",
   "qtyFrom": "walls",
   "price": 5800
  },
  "oshiire_cf": {
   "name": "押入れCF",
   "unit": "一式",
   "price": 14500
  },
  "oshiire_over": {
   "name": "押入れ増し張り",
   "unit": "一式",
   "price": 38600
  },
  "wa_to_yo": {
   "name": "和室から洋室（一式セット）",
   "unit": "一式",
   "note": "大壁化・床・クロス等をまとめた仮単価",
   "price": 337400
  },
  "genkan_door": {
   "name": "玄関ドア交換（ワイケイ材工）",
   "unit": "一式",
   "price": 337400
  },
  "key": {
   "name": "鍵交換",
   "unit": "箇所",
   "price": 24100
  },
  "kamachi": {
   "name": "框（上がり框）",
   "unit": "一式",
   "price": 43400
  },
  "getabako": {
   "name": "下駄箱交換",
   "unit": "一式",
   "price": 77200
  },
  "kitchen": {
   "name": "システムキッチン交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 626600
    },
    {
     "name": "上位",
     "price": 867500
    }
   ]
  },
  "kitchen_faucet": {
   "name": "キッチン水栓交換",
   "unit": "箇所",
   "price": 33800
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
     "price": 963900
    },
    {
     "name": "上位",
     "price": 1253100
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
     "price": 819300
    },
    {
     "name": "上位",
     "price": 1060300
    }
   ]
  },
  "bath_door": {
   "name": "浴室ドア交換",
   "unit": "箇所",
   "price": 67500
  },
  "bath_fan": {
   "name": "浴室換気扇交換",
   "unit": "箇所",
   "price": 33800
  },
  "toilet": {
   "name": "トイレ交換（便器＋温水洗浄便座）",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 173500
    },
    {
     "name": "上位",
     "price": 269900
    }
   ]
  },
  "toilet_interior": {
   "name": "トイレ内装（クロス・CF）",
   "unit": "一式",
   "price": 43400
  },
  "washstand": {
   "name": "洗面化粧台交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 115700
    },
    {
     "name": "上位",
     "price": 192800
    }
   ]
  },
  "wash_interior": {
   "name": "洗面所内装（クロス・CF）",
   "unit": "一式",
   "price": 48200
  },
  "ext_wall": {
   "name": "外壁のみ塗装（足場込）",
   "unit": "一式",
   "price": 482000
  },
  "ext_wall_roof": {
   "name": "外壁・屋根塗装（足場込）",
   "unit": "一式",
   "price": 542200
  },
  "ext_wash": {
   "name": "高圧洗浄のみ（足場込）",
   "unit": "一式",
   "price": 96400
  },
  "carport_rm": {
   "name": "カーポート撤去",
   "unit": "台",
   "price": 57900
  },
  "gate_rm": {
   "name": "門扉撤去",
   "unit": "一式",
   "price": 38600
  },
  "weed_sheet": {
   "name": "防草シート＋砕石",
   "unit": "㎡",
   "qtyDefault": 10,
   "price": 3400
  },
  "cut_tree": {
   "name": "伐採",
   "unit": "本",
   "price": 14500
  },
  "root_rm": {
   "name": "抜根",
   "unit": "本",
   "price": 19300
  },
  "stone_rm": {
   "name": "天然石撤去",
   "unit": "一式",
   "price": 48200
  },
  "shed_rm": {
   "name": "物置撤去",
   "unit": "台",
   "price": 38600
  },
  "nureen_rm": {
   "name": "濡れ縁撤去",
   "unit": "一式",
   "price": 29000
  },
  "od_ceiling": {
   "name": "照明器具（シーリングライト）",
   "unit": "台",
   "price": 7800
  },
  "od_downlight": {
   "name": "照明器具（ダウンライト）",
   "unit": "台",
   "price": 5800
  },
  "od_porch": {
   "name": "玄関灯・外灯",
   "unit": "台",
   "price": 11600
  },
  "od_switch": {
   "name": "スイッチ・コンセント交換",
   "unit": "箇所",
   "price": 3400
  },
  "od_panel": {
   "name": "分電盤",
   "unit": "台",
   "price": 57900
  },
  "od_fan": {
   "name": "換気扇",
   "unit": "台",
   "price": 14500
  },
  "od_intercom": {
   "name": "インターホン（モニター付）",
   "unit": "台",
   "price": 24100
  },
  "od_antenna": {
   "name": "TVアンテナ",
   "unit": "一式",
   "price": 48200
  },
  "od_aircon": {
   "name": "エアコン本体",
   "unit": "台",
   "price": 86800
  },
  "od_lumber": {
   "name": "構造材・下地材",
   "unit": "一式",
   "price": 48200
  },
  "od_flooring": {
   "name": "フローリング材",
   "unit": "㎡",
   "price": 4400
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
   "price": 43400
  },
  "od_plywood": {
   "name": "合板・ベニヤ",
   "unit": "枚",
   "price": 2500
  },
  "od_zosaku": {
   "name": "造作材",
   "unit": "一式",
   "price": 29000
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
   "price": 11600
  },
  "od_sash": {
   "name": "サッシ・窓",
   "unit": "箇所",
   "price": 77200
  },
  "od_siding": {
   "name": "外壁材・サイディング",
   "unit": "㎡",
   "price": 4900
  },
  "od_paint": {
   "name": "塗料",
   "unit": "缶",
   "price": 24100
  },
  "od_key_cyl": {
   "name": "玄関錠交換（シリンダー）",
   "unit": "箇所",
   "price": 19300
  },
  "od_key_dim": {
   "name": "玄関錠交換（ディンプル）",
   "unit": "箇所",
   "price": 33800
  },
  "od_key_back": {
   "name": "勝手口錠交換",
   "unit": "箇所",
   "price": 14500
  },
  "od_key_room": {
   "name": "室内錠交換",
   "unit": "箇所",
   "price": 5800
  },
  "od_key_copy": {
   "name": "合鍵作成",
   "unit": "本",
   "price": 1500
  },
  "od_jk_pump": {
   "name": "浄化槽 汲み取り",
   "unit": "一式",
   "price": 29000
  },
  "od_jk_clean": {
   "name": "浄化槽 清掃・点検",
   "unit": "一式",
   "price": 14500
  },
  "od_jk_remove": {
   "name": "浄化槽 撤去・埋め戻し",
   "unit": "一式",
   "price": 144600
  },
  "od_sewer": {
   "name": "下水接続工事",
   "unit": "一式",
   "price": 289200
  },
  "od_carp": {
   "name": "大工手間",
   "unit": "人工",
   "price": 24100
  },
  "od_interior": {
   "name": "内装手間",
   "unit": "人工",
   "price": 21300
  },
  "od_misc": {
   "name": "雑工・養生",
   "unit": "人工",
   "price": 17400
  },
  "od_demo": {
   "name": "解体手間",
   "unit": "人工",
   "price": 19300
  },
  "od_haul": {
   "name": "運搬・搬出",
   "unit": "回",
   "price": 14500
  },
  "od_manage": {
   "name": "現場管理費",
   "unit": "一式",
   "price": 48200
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
 ],
 "costEnc": {
  "v": 1,
  "salt": "OsMDYD/NAfxXU3ZRdQ02xA==",
  "iv": "4UABi56yOM4rvUqa",
  "data": "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"
 }
};
