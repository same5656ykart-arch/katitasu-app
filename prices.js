// ============================================================
// アプリ用 単価ファイル（客出し単価＋暗号化した原価）
// ※ tools/build_prices.py（Excelから）または アプリの原価マスター画面から生成します。
//    手で直さないでください。原価は合言葉がないと読めません。
//    生成日: 2026-11-11
// ============================================================
window.PRICE_TABLE = {
 "version": 5,
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
   "name": "巾木交換",
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
  },
  "kc_same": {
   "name": "キッチン交換（位置そのまま）",
   "unit": "一式",
   "price": 72300
  },
  "kc_move": {
   "name": "キッチン交換（位置が変わる）",
   "unit": "一式",
   "price": 102500
  },
  "wc_swap": {
   "name": "トイレ交換",
   "unit": "台",
   "price": 21700
  },
  "ub_from_zairai": {
   "name": "浴室交換（既存が在来）",
   "unit": "一式",
   "price": 84400
  },
  "ub_from_ub": {
   "name": "浴室交換（既存がユニットバス）",
   "unit": "一式",
   "price": 36200
  },
  "ub_expand": {
   "name": "浴室拡張（基礎ハツリなど）",
   "unit": "一式",
   "price": 18100
  },
  "ws_same": {
   "name": "化粧台交換（その場で交換）",
   "unit": "台",
   "price": 21700
  },
  "ws_move": {
   "name": "化粧台交換（位置替えあり）",
   "unit": "台",
   "price": 0
  },
  "laundry_new": {
   "name": "洗濯機置き場新設",
   "unit": "一式",
   "price": 16900
  },
  "floor_tile": {
   "name": "フロアタイル張り",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 0
  },
  "floor_choshaku": {
   "name": "長尺シート張り",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 0
  },
  "habaki_paint": {
   "name": "巾木塗装",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 0
  },
  "mawari_swap": {
   "name": "周り縁交換",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 0
  },
  "mawari_paint": {
   "name": "周り縁塗装",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 0
  },
  "window_clean": {
   "name": "窓枠クリーニング",
   "unit": "箇所",
   "price": 0
  },
  "fukinuke": {
   "name": "吹き抜けあり（増額）",
   "unit": "一式",
   "price": 0
  },
  "window_swap": {
   "name": "窓交換",
   "unit": "箇所",
   "price": 0
  },
  "window_close": {
   "name": "窓潰し",
   "unit": "箇所",
   "price": 0
  },
  "wa_yo_shin": {
   "name": "和室→洋室化（既存が真壁）",
   "unit": "一式",
   "price": 207300
  },
  "wa_yo_o": {
   "name": "和室→洋室化（既存が大壁）",
   "unit": "一式",
   "price": 68700
  },
  "st_tesuri": {
   "name": "階段 手すり",
   "unit": "一式",
   "price": 0
  },
  "st_paint": {
   "name": "階段 塗装",
   "unit": "一式",
   "price": 0
  },
  "st_cf": {
   "name": "階段 CF",
   "unit": "一式",
   "price": 0
  },
  "st_nonslip": {
   "name": "階段 ノンスリップ",
   "unit": "一式",
   "price": 0
  }
 },
 "common": [
  {
   "name": "クロス",
   "items": [
    "cloth"
   ],
   "labels": {
    "cloth": "クロス張替え"
   }
  },
  {
   "name": "床（全部屋まとめて）",
   "apply": [
    {
     "label": "フローリング",
     "item": "floor_over"
    },
    {
     "label": "CF",
     "item": "cf"
    },
    {
     "label": "フロアタイル",
     "item": "floor_tile"
    }
   ]
  },
  {
   "name": "建具",
   "apply": [
    {
     "label": "全部交換",
     "item": "door_replace"
    }
   ],
   "items": [
    "knob"
   ],
   "labels": {
    "knob": "ドアノブ交換"
   }
  },
  {
   "name": "巾木",
   "exclusive": true,
   "items": [
    "habaki",
    "habaki_paint"
   ],
   "labels": {
    "habaki": "交換",
    "habaki_paint": "塗装"
   }
  },
  {
   "name": "周り縁",
   "exclusive": true,
   "items": [
    "mawari_swap",
    "mawari_paint"
   ],
   "labels": {
    "mawari_swap": "交換",
    "mawari_paint": "塗装"
   }
  },
  {
   "name": "窓枠",
   "exclusive": true,
   "items": [
    "window_paint",
    "window_clean",
    "window_dinoc"
   ],
   "labels": {
    "window_paint": "塗装",
    "window_clean": "クリーニング",
    "window_dinoc": "ダイノックシート"
   }
  },
  {
   "name": "レベル工事",
   "items": [
    "level"
   ],
   "labels": {
    "level": "あり"
   }
  }
 ],
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
   "defaultTatami": 2,
   "parts": [
    {
     "items": [
      "genkan_door",
      "key",
      "getabako"
     ],
     "labels": {
      "genkan_door": "玄関ドア交換",
      "key": "鍵交換",
      "getabako": "下駄箱交換"
     }
    },
    {
     "name": "床張り",
     "exclusive": true,
     "items": [
      "floor_tile",
      "floor_choshaku"
     ],
     "labels": {
      "floor_tile": "フロアタイル",
      "floor_choshaku": "長尺"
     }
    },
    {
     "name": "吹き抜け",
     "items": [
      "fukinuke"
     ],
     "labels": {
      "fukinuke": "あり"
     }
    }
   ]
  },
  {
   "id": "kitchen",
   "name": "キッチン",
   "icon": "🍳",
   "defaultTatami": 4.5,
   "parts": [
    {
     "name": "キッチン交換",
     "exclusive": true,
     "items": [
      "kc_same",
      "kc_move"
     ],
     "labels": {
      "kc_same": "位置そのまま",
      "kc_move": "位置が変わる"
     }
    },
    {
     "items": [
      "door_replace"
     ],
     "labels": {
      "door_replace": "建具交換"
     }
    },
    {
     "name": "床張り",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf",
      "floor_tile"
     ],
     "labels": {
      "floor_over": "フローリング",
      "cf": "CF",
      "floor_tile": "フロアタイル"
     }
    }
   ]
  },
  {
   "id": "living",
   "name": "リビング",
   "icon": "🛋",
   "living": true,
   "parts": [
    {
     "items": [
      "door_replace"
     ],
     "labels": {
      "door_replace": "建具交換"
     }
    },
    {
     "name": "床張り",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf",
      "floor_tile"
     ],
     "labels": {
      "floor_over": "フローリング",
      "cf": "CF",
      "floor_tile": "フロアタイル"
     }
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
     "items": [
      "wc_swap"
     ],
     "labels": {
      "wc_swap": "トイレ交換"
     }
    },
    {
     "items": [
      "door_replace"
     ],
     "labels": {
      "door_replace": "建具交換"
     }
    },
    {
     "name": "床張り",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf",
      "floor_tile"
     ],
     "labels": {
      "floor_over": "フローリング",
      "cf": "CF",
      "floor_tile": "フロアタイル"
     }
    }
   ]
  },
  {
   "id": "bath",
   "name": "浴室",
   "icon": "🛁",
   "defaultTatami": 2,
   "parts": [
    {
     "name": "浴室交換",
     "exclusive": true,
     "items": [
      "ub_from_zairai",
      "ub_from_ub"
     ],
     "labels": {
      "ub_from_zairai": "既存が在来",
      "ub_from_ub": "既存がユニットバス"
     }
    },
    {
     "items": [
      "ub_expand",
      "window_swap",
      "window_close",
      "door_replace"
     ],
     "labels": {
      "ub_expand": "浴室拡張",
      "door_replace": "建具交換"
     }
    }
   ]
  },
  {
   "id": "wash",
   "name": "洗面室",
   "icon": "🚰",
   "defaultTatami": 2,
   "parts": [
    {
     "name": "化粧台交換",
     "exclusive": true,
     "items": [
      "ws_same",
      "ws_move"
     ],
     "labels": {
      "ws_same": "その場で交換",
      "ws_move": "位置替えあり"
     }
    },
    {
     "items": [
      "laundry_new",
      "window_swap",
      "window_close",
      "door_replace"
     ],
     "labels": {
      "door_replace": "建具交換"
     }
    },
    {
     "name": "床張り",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf",
      "floor_tile"
     ],
     "labels": {
      "floor_over": "フローリング",
      "cf": "CF",
      "floor_tile": "フロアタイル"
     }
    }
   ]
  },
  {
   "id": "yo",
   "name": "洋室",
   "icon": "🛏",
   "count": 3,
   "defaultTatami": 6,
   "parts": [
    {
     "items": [
      "door_replace"
     ],
     "labels": {
      "door_replace": "建具交換"
     }
    },
    {
     "name": "床張り",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf",
      "floor_tile"
     ],
     "labels": {
      "floor_over": "フローリング",
      "cf": "CF",
      "floor_tile": "フロアタイル"
     }
    }
   ]
  },
  {
   "id": "wa",
   "name": "和室",
   "icon": "🎋",
   "defaultTatami": 6,
   "parts": [
    {
     "name": "どうする？",
     "exclusive": true,
     "items": [
      "tatami_omote",
      "wa_yo_shin",
      "wa_yo_o"
     ],
     "labels": {
      "tatami_omote": "表替え（和室残し）",
      "wa_yo_shin": "洋室化（既存が真壁）",
      "wa_yo_o": "洋室化（既存が大壁）"
     }
    },
    {
     "name": "洋室化するときの床",
     "exclusive": true,
     "items": [
      "floor_over",
      "cf",
      "floor_tile"
     ],
     "labels": {
      "floor_over": "フローリング",
      "cf": "CF",
      "floor_tile": "フロアタイル"
     },
     "noApply": true
    }
   ],
   "note": "真壁＝柱が壁の表面に見えている。大壁＝柱が壁の中に隠れている"
  },
  {
   "id": "stairs",
   "name": "階段",
   "icon": "🪜",
   "defaultTatami": 2,
   "parts": [
    {
     "items": [
      "st_tesuri",
      "st_paint",
      "st_cf",
      "st_nonslip"
     ],
     "labels": {
      "st_tesuri": "手すり",
      "st_paint": "塗装",
      "st_cf": "CF",
      "st_nonslip": "ノンスリップ"
     }
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
 "inFull": {
  "kc_same": true,
  "kc_move": {
   "keep": "sb_opt_kitchen_move"
  },
  "wc_swap": true,
  "ub_from_zairai": true,
  "ub_from_ub": true,
  "ws_same": true,
  "laundry_new": true
 },
 "costEnc": {
  "v": 1,
  "salt": "4jv3zulD7SsFq470esmhbA==",
  "iv": "SFYoIzbR59w2YZeC",
  "data": "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"
 }
};
