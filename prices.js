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
   "note": "まとめて頼むときは一番上の【フル】をタップ。【フル】を選ぶと個別の項目は自動で外れます（オプションは一緒に選べます）。取付・配管の工事費で、商品代は含みません",
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
   "note": "まとめて頼むときは一番上の【フル】をタップ。【フル】を選ぶと個別の項目は自動で外れます（オプションは一緒に選べます）",
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
  "salt": "7tVNxxzAb2dUByQ03ldfLw==",
  "iv": "y94DQvbfg+Q7EGB/",
  "data": "R1j7Ff33WKYGSsHcm2VANYuRtQHYP4f+R7qGxx+Zvn1bHgx5T74x4VrQcUmSrIKXrDJW7EOhmQM9RBoekYMtYLsY1nKVWBR6aqTqZXhBzfflEiYb3DI0sIxZ/KYBtucKjyPbQrPfM0KjRWrMKy/AY0DGztl0bQOgg/G60fpe+yXHd6pp5Ktt+OcBAXZQQy3d3+HrbN9rQQ+BF7AipeK6VPFgRmsvL5PaDR5TEG0EjnZjWCYV/Si3Gz0chxAX7kQuGH5oRdDX5Q/4F9DCZikGhIZtQqgv8x7z4dXdYf0/GBl1q4NcfFNmQYYLTSEFOuXiggIYCvLG0iHep6lb1LGFevGAUtAwANJpJ5Hwsy3tHnYwcbDJDyRbiDm2YGp3gU/bUaQ1sOlpXbfk7ETCfoABhW4KKh1KdoarZzbhQtmGd4rYQxdmBzQAv+nzJ83w+wgm+xYoKLwj309beCpP/4BFHFLZ/79vcS26qVJ0HgoUuM97XNT2NOgbhOhrj77m4ql18U1ivEw8PMX93BN3Ry98qNb07t+4kv0Wvaftyw1sVNGxyiMkrCRhSL69UV9CZm3ZpGED+FZoKxSmtyXRg9RceaD18Yr/t7qxPYbe8Wegj4YHxOX7f7RN7le8s521nnO6FZ7gGDVT+GMmP0fECueAZlAOgGAco876UCdlBlSg9WfkhN8fOcr3v+Rin56bdpfa3q3nt4Mnk3ITEBEWB8iGS2GiQ/4MQea1NJIyCjUaxtRVXYswYrLPph6WIfNBfx/XdJ6CY0aC/fkZiZ59jMfmuK3jN0dVPzzoE/cvjxGFghNY9ZTcbIouYVZX4rY78FsvF4BS6oXoq+k7NsGYvvmHD/FmP3mGKbuiy6CpVHMgGrSW076CCsJqa5ora6N5i/S9NjVCvUB4cCajBHtnM0gsIEn/SWXl+3hNkaQ1TegIw2Qcr+Vf0S4Co2zkOpM95GAhMJrddDpsMUnjO1oh5ZhJGhQfGIdKdRjNZi1qyAPsx78EDyhGegfThspGuS+jBTbCr6pDpZObiSkQ9TfrOnIQ7boqRvCRhiuCx7UATku+feXmAC2HRuMOTVl8SqFXipmGmjCsZXC/ps1GCy+n4yOKt3ZeGoDqGYPliK2RpR2BZsbXeKtk7JLh/2bTNKqez34nLZUb2oDYGJXDag3nXNXWEHNJ5EJ/Rn6LxAJFv8eI0+gtgOAXbL8Tz/Ct0jCdgD+7ny+tUViR9GDbJjgJVgVJvyRTPbjeNPZhUJvBSTfWX4VPtrO3isi9VCL/3WGN2QkLgdJa3H/6BUuC0cf3ddswuzwyH2k/G7SbrTwyGkxIWaXqpn/xnFTkojRxi+OJr04Pqct+OcFLhdhqm8Rqulje2Ccg/aPMYBzQxATtZd+OdRV6vunbGj4B6S3y/YUFYwG19P39TtYLx/Zu30qWhzTREHpmhuHDNoQZkxi3pZpsEMuQlfGcSdOa55kr/lfNfsUqMD+691wozPR2OBNOoLy3rhS2hBno5ffmqXr/euOhWI6nUZmfAzTz/JV9jMRMx4fuCbutC+BpF7T36ogUTkmeWrn916Fwcnm3L+Z9JOPEcBbmqlPsU5g+9yCBxooml4oXBnTxSFwpgSjLsBXwV2z8mXtccsMNOf9OzyuWHd6sOMtJ2nT92elcDr9Z7u/j/QcGWi3FJx7bnX5JdFo+ON5l4usHuHtspawSEb+dEfrbTFGVZrsx6unKZA/dbE+/BrnEvFszoZAVCvYFUmEGJEHaXetqQQmiaDdcGuagIEFrb5s2o6AmdZlXV9NU8CFikvy6zrF1dhVSo5Zua8VUUOXfOotj+6TjQh65rrVBJQ5x0/OEeq8FA4DXT+EdZKUPOqj0N7ZWPxVGn1RdA7lpNNc1jQC4453QZRLIvJ+PsOoNJ7/Avmse9zAf4rr/8iy4I94cHMSfU0l0H5Dahw8Olck+hroa2GoXqK8gofe80Qa51P+MoguErEO90uziaq4sF2kACEBYDy0ys0tNbdedoz5mhy2goY6GkDnwhgu5oEwtYv6CRP95Qba55V7x8frckkCLXZCy5+Ehok6MlHvuzjbOI39f84G1Wa6yYtObCDUthKlAJ3H97IVWuEslfGqNKusXX+nWwPxQAAWao/ybxzWSdF4Yxd1W5RdmhXHjOBlxIcagYtF3AIYb/ct/iX1z6nHbaK50bBmLLb3KRVdpb386+n5LQ1798gkib6YEcYLEmi8jP6F+ckbxJ6DpfaDIiRc6USuM9XyvO2q0orW7HYbgT44pZgQyN3YHuSdCclJaJmfBXEi35Y7ZgcFCJ0qsGJ/7PnQNC9nWSy+TWKlcf9PAD9sBcXr7jqd241Xy4aND97iRBi+HpNNuQlkLjgKh4bfxP2stwOYEJx6veb8IYOwxHoxroSCGu6Oqq2hWFhfLrmS4Fm6+W8qGpu0bs6AslNaIKY9/PIiK5ovDhlgmU8ugAAOY3TSydrqXRa6Zh8t0ExDseNTDMXIgPEwWKdTuCWcOyvsX/pqjft4jLP9lm847o9BNbo+rxSrkr5uUMG/Mrdl9M2a36os0iqdbVl22H5qVFHWcQhpQnBtWj82mNVsjDwIuzfaERGwFDC3Cqwu0Nsoqx8nEVm3bs5dOz8k3hQRbZtRlNG8cw7FMIkF3LqLqpCPucX/nrfl9PvvvFdkJLXjYMmu8EXIXmB7c5H8Ap09AhyqlEYKJpHUp6oqwXn09TGgnaIzopzIOJgSUccuzqgT62MTVYjoweTTjKkOu8cq7RTt1OP1cmMFWOAD3R3B34IlJzMBRoHM7BR39hfS/79wOSFBXMrfKiCdKrkDN9CkAQ0Uhg+GRpiF6gwpbrbjtAlWheEGP030V8jziVFVO2yjn7JR+ImdMk8UX+g7GXbquy2oDI8S8cFGH/BCUNpFOoU7DfoqOQDVM2B8yRzHVeqXXszhLi+HfhlpBVPI5Xwh84JWJLpjPJNQdWQvDAhbqwy4jsUoBbbZtOeuFqPvR9BJt2fj0ii1of86RC1HtfBtJxIBwMEcP3morotjHQjKynWWW5/6o3aiU6JZ3YzmoEIwEnWqiZJsV/Imaah+m77rUjfxVfj4mv1Wrv6m9GoGXHVUMTNknjqBVKfukM2zGCP8K8FQXa4/1iKkYgYx9s61s8OieQX2gKekbLr/NIuqKH9OSnnlEiIPcrmlztgtukiOGtxfr8pGkP4geBJwILP/4oejwSSVOGKkTLn50hlaU2xOvHiyghcphmLy/7UngP6ableTGZTL/R/b5nyPLCRn4LQNT+3Cd8oQCUzGbkBZU+Z2HP7x/pgP+ATGR5etI0n3Hdkd3RZbYH9lrdk5SBVH0NBTVK1EbWZBQJpBl95tAD8SOs0d5gMCQi+1xLzBuvDdo35JcnqozWLKhyeuZoi+RBMPl3lB4UJdiYU0XfpjWsDDQ67Pkt4W4lflgXZUP4BlY4s5E+YBKNT17dv/miO9avCu50MAbzLCeY8D0eG7hwqYVrwvXAdltEvQOADsAcyzUUP5dqKhYS8itSC1Dlde6W4SN52VxnNdjYMwrvs3QJjWjK/gMoml3yQzjc3NhECP7Qi9V8Wi+Cequw9uYBpqPkNWC76N2gjsotWCHbAcZhgTrMbQHn/9RiTuCAPjraeBTg+0oiJ+lXu/MYyPZ2s2L+c/a8qvPQs4BxG+So6rvefJcj1ZzDrhqVaHHJ6K4Lq4k+TOmPGRSOQWkdDF24M2bXicX/AB+VhtRULlaZTwiDX6XreVCuv7OwD7bMCxZH5K57Y0C7D1HktLGg3Q8vShkbWNWEXgrZMXDGnTPkZZ9a3qIaGykr2ZHvDUGc2JqbtQ19yDmZUezKSWU6oeB7ysI54ykCMHByt8ZcJrEwpgILMwFhtXPBMXqXu0BYVvMKdkgi0GE4ac2mnjUhs8WcgD70CLq5Ob2Jm+ciROm2uY/SDRxzf2krijmBCeByGgg+x51cyeD5P6UpqUy1VnjGFvytZ5BgvatsljJSOKhmkG23U5Ysp5YBrtUR28S6LfQQ/mYyI08XODAcNxO9ghK4DY0vks30x5TYsvhxYtSD5K4LyzYHheV06scKTsLnc1P5fGVsD9jsD7YWGBxh6CO0Pw/sc3m6Tego01raJHWSiSY3MRvwtKjkMstjrvKEpqrUKbjv7Ba5ponbZJ6V6Sk8N5J+W/KOIC0intecBwa22RvdN6xOay2kS1ZyYTOPGgzV1TluNWn+piQS6NinDm1Wqm5RxTiOZBiPswmO3vuDAJG1vCpMoo+oxyZRVv21OoUop9HgCpJOaTUwN9C3SQD3C7BYyS3KxrG8RBzugvUgfqBSLieSq5q8cRFqbUx7pymZfqpjjG1OLp7JCu5lQJ2JJ80QCisrd+ZVg78iJdhQYDNYZQ0QpK+VZ9uPkjXW8I2bVDJold/8Buyj5kmxvn4u0tQHvvxHZyeWHZ7k4/0Rc9xW8CJj/DdHok6ABfW+v8kCeXNV8eN2UYLA9agKEFlvkbQe58SlgS5d/dowbBwd8ejoWbx7K4ymzTEuNO9ra8k+HxKCZU4Hk0Z68xE2sRAP8G5yrjJbvWAAu4p7nxG9DMFrVbB1WfTvVkSvt6hOpCyEgDhrGycqOZNuYbNjnUTUZGVyP1Yn3KINcMx+P5ZTJGcMrxdjgLlYkptlugKclfuIldTS9pfvGp7c9Opn+/pmx3BNgMsLlQ2pBa8L4CdwDkJxHLhke3KMh7/UA3CLeW0i8NackXEDomFSJiHF2btlaiwjNWu0hLXfVN9xACXan/i1eCtZ+OxUXAE0ErUXGYbyk+uKxTScw=="
 }
};
