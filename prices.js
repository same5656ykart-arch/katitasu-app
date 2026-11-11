// ============================================================
// アプリ用 単価ファイル（客出し単価＋暗号化した原価）
// ※ tools/build_prices.py（Excelから）または アプリの原価マスター画面から生成します。
//    手で直さないでください。原価は合言葉がないと読めません。
//    生成日: 2026-11-11
// ============================================================
window.PRICE_TABLE = {
 "version": 4,
 "generatedAt": "2026-11-11",
 "settings": {
  "company": "株式会社ワイケイアート",
  "companyInfo": "",
  "overheadRate": 0.0,
  "taxRate": 10.0,
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
   "name": "大壁化（天井4mmベニア増し張り＋壁組ボード張り）",
   "unit": "一式",
   "note": "既存が真壁の場合、洋室にするには大壁化が必要",
   "price": 138600
  },
  "veneer_cw": {
   "name": "天井＋壁 4mmベニア増し張り（砂壁の場合）",
   "unit": "一式",
   "note": "既存が砂壁なら、真壁のまま／和室のままクロスを張る場合に必要。大壁化する場合は不要",
   "price": 96400
  },
  "veneer_c": {
   "name": "天井 4mmベニア増し張り（板張り・ラミ天の場合）",
   "unit": "一式",
   "note": "天井が板張り（ラミ天）のとき。大壁化・天井＋壁増し張りをする場合は不要 ※原価は仮",
   "price": 42200
  },
  "yuka_gumi": {
   "name": "畳→床材変更の下地工事（床組）",
   "unit": "一式",
   "note": "畳を撤去してフローリング等にする場合の下地",
   "price": 68700
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
  },
  "sb_toilet": {
   "name": "トイレ交換工事",
   "unit": "台",
   "price": 21700
  },
  "sb_wash": {
   "name": "化粧台交換工事",
   "unit": "台",
   "price": 21700
  },
  "sb_laundry_drain": {
   "name": "洗濯排水（材工）",
   "unit": "個",
   "price": 6100
  },
  "sb_laundry_faucet": {
   "name": "洗濯水栓（材工）",
   "unit": "個",
   "price": 10900
  },
  "sb_kitchen_asm": {
   "name": "キッチン組み立て",
   "unit": "台",
   "price": 42200
  },
  "sb_bath_pipe": {
   "name": "浴室交換に伴う配管工事",
   "unit": "一式",
   "price": 36200
  },
  "sb_kitchen_pipe": {
   "name": "キッチン交換に伴う配管工事",
   "unit": "一式",
   "price": 30200
  },
  "sb_heater": {
   "name": "給湯器交換工事",
   "unit": "台",
   "price": 27800
  },
  "sb_bath_doma": {
   "name": "浴室土間打ち",
   "unit": "一式",
   "price": 48200
  },
  "sb_opt_kitchen_move": {
   "name": "【オプション】キッチン位置変えの配管工事（追加）",
   "unit": "一式",
   "price": 30200
  },
  "sb_opt_bath_hatsuri": {
   "name": "【オプション】浴室拡張に伴う基礎ハツリなど（追加）",
   "unit": "一式",
   "price": 18100
  },
  "dk_outlet_full": {
   "name": "コンセント・スイッチ交換（フル）",
   "unit": "一式",
   "price": 96400
  },
  "dk_sleeve": {
   "name": "スリーブキャップ交換",
   "unit": "一式",
   "price": 12100
  },
  "dk_circuit_bathfan": {
   "name": "専用回路（浴室換気扇用）※分電盤まで近い場合",
   "unit": "一式",
   "price": 18100
  },
  "dk_circuit_ac": {
   "name": "専用回路（エアコンコンセントが無い部屋）",
   "unit": "箇所",
   "price": 21700
  },
  "dk_tv": {
   "name": "TVコンセント新設・位置替え",
   "unit": "箇所",
   "price": 12100
  },
  "dk_kitchen_duct": {
   "name": "キッチン交換に伴うダクト接続・延長",
   "unit": "一式",
   "price": 7300
  },
  "dk_kitchen_wiring": {
   "name": "キッチン交換に伴う手元灯・レンジフード・コンセントの切り回し",
   "unit": "一式",
   "price": 18100
  },
  "dk_bath_duct": {
   "name": "浴室交換に伴うダクト接続・延長",
   "unit": "一式",
   "price": 7300
  },
  "dk_opt_bath_hole": {
   "name": "【オプション】浴室ダクトの穴あけ（追加）",
   "unit": "一式",
   "price": 6100
  },
  "dk_bath_wiring": {
   "name": "浴室交換に伴う照明・リモコン・電源の配線工事",
   "unit": "一式",
   "price": 30200
  },
  "dk_intercom": {
   "name": "インターホン交換（移設費含む）",
   "unit": "一式",
   "price": 9700
  },
  "dk_heater_remote": {
   "name": "給湯器リモコン取り付け（キッチン・浴室。移設費・浴室コーキング含む）",
   "unit": "一式",
   "price": 8500
  },
  "dk_panel": {
   "name": "分電盤交換（材工・移設費含む）",
   "unit": "一式",
   "price": 42200
  },
  "dk_light": {
   "name": "照明取り付け及び配線",
   "unit": "一式",
   "price": 24100
  },
  "sb_full": {
   "name": "【フル】設備工事 一式",
   "unit": "一式",
   "note": "トイレ交換・化粧台交換・洗濯排水・洗濯水栓・キッチン組み立て・浴室/キッチン交換の配管・給湯器交換・浴室土間打ち。オプション（キッチン位置変え・浴室拡張のハツリ）は別",
   "price": 244600
  },
  "dk_full": {
   "name": "【フル】電気工事 一式",
   "unit": "一式",
   "note": "コンセント・スイッチ交換、分電盤交換、照明交換、水回りの配線・ダクト工事、スリーブ交換、エアコン・TV各部屋配線",
   "price": 277200
  }
 },
 "whole": [
  {
   "id": "setsubi",
   "name": "設備工事",
   "icon": "🔧",
   "note": "【フル】を押すと、個別の項目は自動で外れます（オプションは一緒に選べます）。工事費だけで、商品代は含みません",
   "full": "sb_full",
   "fullKeeps": [
    "sb_opt_kitchen_move",
    "sb_opt_bath_hatsuri"
   ],
   "items": [
    "sb_full",
    "sb_toilet",
    "sb_wash",
    "sb_laundry_drain",
    "sb_laundry_faucet",
    "sb_kitchen_asm",
    "sb_bath_pipe",
    "sb_kitchen_pipe",
    "sb_heater",
    "sb_bath_doma",
    "sb_opt_kitchen_move",
    "sb_opt_bath_hatsuri"
   ]
  },
  {
   "id": "denki",
   "name": "電気工事",
   "icon": "⚡",
   "note": "【フル】を押すと、個別の項目は自動で外れます（オプションは一緒に選べます）",
   "full": "dk_full",
   "fullKeeps": [
    "dk_opt_bath_hole"
   ],
   "items": [
    "dk_full",
    "dk_outlet_full",
    "dk_sleeve",
    "dk_circuit_bathfan",
    "dk_circuit_ac",
    "dk_tv",
    "dk_kitchen_duct",
    "dk_kitchen_wiring",
    "dk_bath_duct",
    "dk_opt_bath_hole",
    "dk_bath_wiring",
    "dk_intercom",
    "dk_heater_remote",
    "dk_panel",
    "dk_light"
   ]
  }
 ],
 "rooms": [
  {
   "id": "genkan",
   "name": "玄関",
   "icon": "🚪",
   "defaultTatami": 1,
   "parts": [
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "玄関ドア・鍵",
     "items": [
      "genkan_door",
      "key"
     ]
    },
    {
     "name": "下駄箱・框",
     "items": [
      "getabako",
      "kamachi"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    },
    {
     "name": "巾木・廻り縁",
     "items": [
      "habaki",
      "mawaribuchi"
     ]
    }
   ]
  },
  {
   "id": "kitchen",
   "name": "キッチン",
   "icon": "🍳",
   "size": true,
   "defaultTatami": 6,
   "parts": [
    {
     "name": "キッチン本体（商品込み）",
     "items": [
      "kitchen",
      "kitchen_faucet"
     ]
    },
    {
     "name": "床",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf"
     ]
    },
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    },
    {
     "name": "巾木・廻り縁",
     "items": [
      "habaki",
      "mawaribuchi"
     ]
    }
   ],
   "note": "商品込みの金額です。カチタス支給品のときは選ばず、上の「家全体 → 設備工事」で工事費だけを選びます"
  },
  {
   "id": "living",
   "name": "リビング",
   "icon": "🛋",
   "size": true,
   "defaultTatami": 10,
   "parts": [
    {
     "name": "床",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf"
     ]
    },
    {
     "name": "床の下地",
     "items": [
      "level"
     ]
    },
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "建具",
     "items": [
      "door_replace",
      "knob"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    },
    {
     "name": "巾木・廻り縁",
     "items": [
      "habaki",
      "mawaribuchi"
     ]
    },
    {
     "name": "窓枠",
     "items": [
      "window_paint",
      "window_dinoc"
     ],
     "exclusive": true
    }
   ]
  },
  {
   "id": "toilet",
   "name": "トイレ",
   "icon": "🚽",
   "defaultTatami": 1,
   "parts": [
    {
     "name": "便器（商品込み）",
     "items": [
      "toilet"
     ]
    },
    {
     "name": "内装",
     "items": [
      "toilet_interior"
     ]
    },
    {
     "name": "建具",
     "items": [
      "door_replace"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    }
   ],
   "note": "商品込みの金額です。カチタス支給品のときは選ばず、上の「家全体 → 設備工事」で工事費だけを選びます"
  },
  {
   "id": "bath",
   "name": "浴室",
   "icon": "🛁",
   "defaultTatami": 2,
   "parts": [
    {
     "name": "浴室本体（商品込み）",
     "items": [
      "bath_zairai_ub",
      "bath_ub"
     ],
     "exclusive": true
    },
    {
     "name": "ドア・換気扇",
     "items": [
      "bath_door",
      "bath_fan"
     ]
    }
   ],
   "note": "商品込みの金額です。カチタス支給品のときは選ばず、上の「家全体 → 設備工事」で工事費だけを選びます"
  },
  {
   "id": "wash",
   "name": "洗面所",
   "icon": "🚰",
   "defaultTatami": 2,
   "parts": [
    {
     "name": "洗面台（商品込み）",
     "items": [
      "washstand"
     ]
    },
    {
     "name": "内装",
     "items": [
      "wash_interior"
     ]
    },
    {
     "name": "建具",
     "items": [
      "door_replace"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    }
   ],
   "note": "商品込みの金額です。カチタス支給品のときは選ばず、上の「家全体 → 設備工事」で工事費だけを選びます"
  },
  {
   "id": "yo",
   "name": "洋室",
   "icon": "🛏",
   "size": true,
   "defaultTatami": 6,
   "parts": [
    {
     "name": "床",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf"
     ]
    },
    {
     "name": "床の下地",
     "items": [
      "level"
     ]
    },
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "建具",
     "items": [
      "door_replace",
      "knob"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    },
    {
     "name": "巾木・廻り縁",
     "items": [
      "habaki",
      "mawaribuchi"
     ]
    },
    {
     "name": "窓枠",
     "items": [
      "window_paint",
      "window_dinoc"
     ],
     "exclusive": true
    }
   ]
  },
  {
   "id": "wa",
   "name": "和室",
   "icon": "🎋",
   "size": true,
   "defaultTatami": 6,
   "note": "洋室にするとき：真壁（柱が見える壁）なら大壁化、砂壁・板張り天井ならベニア増し張りが必要です",
   "parts": [
    {
     "name": "洋室にする下地",
     "items": [
      "okabe",
      "veneer_cw",
      "veneer_c"
     ],
     "exclusive": true
    },
    {
     "name": "床",
     "items": [
      "yuka_gumi",
      "floor_over",
      "tatami_omote"
     ]
    },
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "襖・障子",
     "items": [
      "fusuma",
      "shoji"
     ]
    },
    {
     "name": "木部",
     "items": [
      "aku"
     ]
    },
    {
     "name": "押入れ",
     "items": [
      "oshiire_cf",
      "oshiire_over"
     ],
     "exclusive": true
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    }
   ]
  },
  {
   "id": "hall",
   "name": "廊下",
   "icon": "🚶",
   "size": true,
   "defaultTatami": 3,
   "parts": [
    {
     "name": "床",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf"
     ]
    },
    {
     "name": "床の下地",
     "items": [
      "level"
     ]
    },
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    },
    {
     "name": "巾木・廻り縁",
     "items": [
      "habaki",
      "mawaribuchi"
     ]
    }
   ]
  },
  {
   "id": "stairs",
   "name": "階段",
   "icon": "🪜",
   "size": true,
   "defaultTatami": 3,
   "parts": [
    {
     "name": "壁・天井",
     "items": [
      "cloth"
     ]
    },
    {
     "name": "照明",
     "items": [
      "light"
     ]
    },
    {
     "name": "巾木",
     "items": [
      "habaki"
     ]
    }
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
  "salt": "dbn/ZNMjThR5/pw9B2MvHw==",
  "iv": "Zm6sHFEBoBlFyk7e",
  "data": "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"
 }
};
