// ============================================================
// アプリ用 単価ファイル（客出し単価＋暗号化した原価）
// ※ tools/build_prices.py（Excelから）または アプリの原価マスター画面から生成します。
//    手で直さないでください。原価は合言葉がないと読めません。
//    生成日: 2026-11-11
// ============================================================
window.PRICE_TABLE = {
 "version": 6,
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
   "price": 0
  },
  "door_replace": {
   "name": "建具交換",
   "unit": "枚",
   "price": 0
  },
  "cloth": {
   "name": "クロス張替え",
   "unit": "㎡",
   "qtyFrom": "cloth",
   "price": 0
  },
  "cf": {
   "name": "クッションフロア張替え",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 0
  },
  "habaki": {
   "name": "巾木交換",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 0
  },
  "mawaribuchi": {
   "name": "周り縁撤去",
   "unit": "m",
   "qtyFrom": "perim",
   "price": 0
  },
  "window_paint": {
   "name": "窓枠塗装",
   "unit": "箇所",
   "price": 0
  },
  "window_dinoc": {
   "name": "窓枠ダイノック",
   "unit": "箇所",
   "price": 0
  },
  "light": {
   "name": "照明交換",
   "unit": "箇所",
   "price": 0
  },
  "knob": {
   "name": "ドアノブ交換",
   "unit": "箇所",
   "price": 0
  },
  "level": {
   "name": "レベル工事（床）",
   "unit": "㎡",
   "qtyFrom": "floor",
   "price": 0
  },
  "ac_outlet": {
   "name": "エアコンコンセント",
   "unit": "箇所",
   "price": 0
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
   "price": 0
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
   "price": 0
  },
  "fusuma": {
   "name": "襖 張替え",
   "unit": "枚",
   "qtyDefault": 2,
   "price": 0
  },
  "shoji": {
   "name": "障子 張替え",
   "unit": "枚",
   "qtyDefault": 2,
   "price": 0
  },
  "aku": {
   "name": "木部あく抜き・オイルステイン",
   "unit": "一式",
   "price": 0
  },
  "oshiire_cf": {
   "name": "押入れCF",
   "unit": "一式",
   "price": 0
  },
  "oshiire_over": {
   "name": "押入れ増し張り",
   "unit": "一式",
   "price": 0
  },
  "genkan_door": {
   "name": "玄関ドア交換（ワイケイ材工）",
   "unit": "一式",
   "price": 0
  },
  "key": {
   "name": "鍵交換",
   "unit": "箇所",
   "price": 0
  },
  "kamachi": {
   "name": "框（上がり框）",
   "unit": "一式",
   "price": 0
  },
  "getabako": {
   "name": "下駄箱交換",
   "unit": "一式",
   "price": 0
  },
  "kitchen": {
   "name": "システムキッチン交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 0
    },
    {
     "name": "上位",
     "price": 0
    }
   ]
  },
  "kitchen_faucet": {
   "name": "キッチン水栓交換",
   "unit": "箇所",
   "price": 0
  },
  "bath_zairai_ub": {
   "name": "在来浴室 → ユニットバス化（解体含む）",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 0
    },
    {
     "name": "上位",
     "price": 0
    }
   ]
  },
  "bath_ub": {
   "name": "ユニットバス交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 0
    },
    {
     "name": "上位",
     "price": 0
    }
   ]
  },
  "bath_door": {
   "name": "浴室ドア交換",
   "unit": "箇所",
   "price": 0
  },
  "bath_fan": {
   "name": "浴室換気扇交換",
   "unit": "箇所",
   "price": 0
  },
  "toilet": {
   "name": "トイレ交換（便器＋温水洗浄便座）",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 0
    },
    {
     "name": "上位",
     "price": 0
    }
   ]
  },
  "toilet_interior": {
   "name": "トイレ内装（クロス・CF）",
   "unit": "一式",
   "price": 0
  },
  "washstand": {
   "name": "洗面化粧台交換",
   "unit": "一式",
   "grades": [
    {
     "name": "標準",
     "price": 0
    },
    {
     "name": "上位",
     "price": 0
    }
   ]
  },
  "wash_interior": {
   "name": "洗面所内装（クロス・CF）",
   "unit": "一式",
   "price": 0
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
   "price": 0
  },
  "carport_rm": {
   "name": "カーポート撤去",
   "unit": "台",
   "price": 0
  },
  "gate_rm": {
   "name": "門扉撤去",
   "unit": "一式",
   "price": 0
  },
  "weed_sheet": {
   "name": "防草シート＋砕石",
   "unit": "㎡",
   "qtyDefault": 10,
   "price": 0
  },
  "cut_tree": {
   "name": "伐採",
   "unit": "本",
   "price": 0
  },
  "root_rm": {
   "name": "抜根",
   "unit": "本",
   "price": 0
  },
  "stone_rm": {
   "name": "天然石撤去",
   "unit": "一式",
   "price": 0
  },
  "shed_rm": {
   "name": "物置撤去",
   "unit": "台",
   "price": 0
  },
  "nureen_rm": {
   "name": "濡れ縁撤去",
   "unit": "一式",
   "price": 0
  },
  "od_ceiling": {
   "name": "照明器具（シーリングライト）",
   "unit": "台",
   "price": 0
  },
  "od_downlight": {
   "name": "照明器具（ダウンライト）",
   "unit": "台",
   "price": 0
  },
  "od_porch": {
   "name": "玄関灯・外灯",
   "unit": "台",
   "price": 0
  },
  "od_switch": {
   "name": "スイッチ・コンセント交換",
   "unit": "箇所",
   "price": 0
  },
  "od_panel": {
   "name": "分電盤",
   "unit": "台",
   "price": 0
  },
  "od_fan": {
   "name": "換気扇",
   "unit": "台",
   "price": 0
  },
  "od_intercom": {
   "name": "インターホン（モニター付）",
   "unit": "台",
   "price": 0
  },
  "od_antenna": {
   "name": "TVアンテナ",
   "unit": "一式",
   "price": 0
  },
  "od_aircon": {
   "name": "エアコン本体",
   "unit": "台",
   "price": 0
  },
  "od_lumber": {
   "name": "構造材・下地材",
   "unit": "一式",
   "price": 0
  },
  "od_flooring": {
   "name": "フローリング材",
   "unit": "㎡",
   "price": 0
  },
  "od_habaki": {
   "name": "巾木材",
   "unit": "m",
   "price": 0
  },
  "od_mawari": {
   "name": "廻り縁材",
   "unit": "m",
   "price": 0
  },
  "od_door": {
   "name": "建具（室内ドア）",
   "unit": "枚",
   "price": 0
  },
  "od_plywood": {
   "name": "合板・ベニヤ",
   "unit": "枚",
   "price": 0
  },
  "od_zosaku": {
   "name": "造作材",
   "unit": "一式",
   "price": 0
  },
  "od_cloth": {
   "name": "クロス材",
   "unit": "㎡",
   "price": 0
  },
  "od_cf": {
   "name": "クッションフロア材",
   "unit": "㎡",
   "price": 0
  },
  "od_board": {
   "name": "石膏ボード",
   "unit": "枚",
   "price": 0
  },
  "od_insul": {
   "name": "断熱材",
   "unit": "㎡",
   "price": 0
  },
  "od_tatami": {
   "name": "畳（新調）",
   "unit": "枚",
   "price": 0
  },
  "od_sash": {
   "name": "サッシ・窓",
   "unit": "箇所",
   "price": 0
  },
  "od_siding": {
   "name": "外壁材・サイディング",
   "unit": "㎡",
   "price": 0
  },
  "od_paint": {
   "name": "塗料",
   "unit": "缶",
   "price": 0
  },
  "od_key_cyl": {
   "name": "玄関錠交換（シリンダー）",
   "unit": "箇所",
   "price": 0
  },
  "od_key_dim": {
   "name": "玄関錠交換（ディンプル）",
   "unit": "箇所",
   "price": 0
  },
  "od_key_back": {
   "name": "勝手口錠交換",
   "unit": "箇所",
   "price": 0
  },
  "od_key_room": {
   "name": "室内錠交換",
   "unit": "箇所",
   "price": 0
  },
  "od_key_copy": {
   "name": "合鍵作成",
   "unit": "本",
   "price": 0
  },
  "od_jk_pump": {
   "name": "浄化槽 汲み取り",
   "unit": "一式",
   "price": 0
  },
  "od_jk_clean": {
   "name": "浄化槽 清掃・点検",
   "unit": "一式",
   "price": 0
  },
  "od_jk_remove": {
   "name": "浄化槽 撤去・埋め戻し",
   "unit": "一式",
   "price": 0
  },
  "od_sewer": {
   "name": "下水接続工事",
   "unit": "一式",
   "price": 0
  },
  "od_carp": {
   "name": "大工手間",
   "unit": "人工",
   "price": 0
  },
  "od_interior": {
   "name": "内装手間",
   "unit": "人工",
   "price": 0
  },
  "od_misc": {
   "name": "雑工・養生",
   "unit": "人工",
   "price": 0
  },
  "od_demo": {
   "name": "解体手間",
   "unit": "人工",
   "price": 0
  },
  "od_haul": {
   "name": "運搬・搬出",
   "unit": "回",
   "price": 0
  },
  "od_manage": {
   "name": "現場管理費",
   "unit": "一式",
   "price": 0
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
   "price": 0
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
 "trades": [
  {
   "id": "t01",
   "name": "解体",
   "items": [
    "od_demo"
   ]
  },
  {
   "id": "t02",
   "name": "設備",
   "items": [
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
    "sb_opt_bath_hatsuri",
    "sb_full",
    "kc_same",
    "kc_move",
    "wc_swap",
    "ub_from_zairai",
    "ub_from_ub",
    "ub_expand",
    "ws_same",
    "ws_move",
    "laundry_new",
    "kitchen",
    "kitchen_faucet",
    "bath_zairai_ub",
    "bath_ub",
    "bath_fan",
    "toilet",
    "washstand",
    "od_jk_pump",
    "od_jk_clean",
    "od_jk_remove",
    "od_sewer"
   ]
  },
  {
   "id": "t03",
   "name": "電気",
   "items": [
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
    "dk_light",
    "dk_full",
    "ac_outlet",
    "od_switch",
    "od_panel",
    "od_fan",
    "od_intercom",
    "od_antenna",
    "od_aircon"
   ]
  },
  {
   "id": "t04",
   "name": "大工",
   "items": [
    "floor_over",
    "door_replace",
    "level",
    "habaki",
    "mawari_swap",
    "mawaribuchi",
    "st_tesuri",
    "window_swap",
    "window_close",
    "getabako",
    "kamachi",
    "bath_door",
    "od_carp"
   ]
  },
  {
   "id": "t05",
   "name": "内装",
   "items": [
    "cloth",
    "cf",
    "floor_tile",
    "floor_choshaku",
    "st_cf",
    "st_nonslip",
    "window_dinoc",
    "fukinuke",
    "toilet_interior",
    "wash_interior",
    "oshiire_cf",
    "oshiire_over",
    "od_interior"
   ]
  },
  {
   "id": "t06",
   "name": "クリーニング",
   "items": [
    "window_clean"
   ]
  },
  {
   "id": "t07",
   "name": "塗装",
   "items": [
    "habaki_paint",
    "mawari_paint",
    "window_paint",
    "st_paint",
    "ext_wall",
    "ext_wall_roof",
    "ext_wash",
    "aku",
    "od_paint"
   ]
  },
  {
   "id": "t08",
   "name": "足場",
   "items": []
  },
  {
   "id": "t09",
   "name": "外構",
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
  },
  {
   "id": "t10",
   "name": "鍵交換",
   "items": [
    "key",
    "od_key_cyl",
    "od_key_dim",
    "od_key_back",
    "od_key_room",
    "od_key_copy"
   ]
  },
  {
   "id": "t11",
   "name": "和室",
   "items": [
    "tatami_omote",
    "wa_yo_shin",
    "wa_yo_o",
    "okabe",
    "veneer_cw",
    "veneer_c",
    "yuka_gumi",
    "fusuma",
    "shoji"
   ]
  },
  {
   "id": "t12",
   "name": "木材",
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
   "id": "t13",
   "name": "照明",
   "items": [
    "light",
    "od_ceiling",
    "od_downlight",
    "od_porch"
   ]
  },
  {
   "id": "t14",
   "name": "建材",
   "items": [
    "genkan_door",
    "od_cloth",
    "od_cf",
    "od_board",
    "od_insul",
    "od_tatami",
    "od_sash",
    "od_siding"
   ]
  },
  {
   "id": "t15",
   "name": "雑工事",
   "items": [
    "knob",
    "od_misc",
    "od_manage"
   ]
  },
  {
   "id": "gomi_out",
   "name": "ゴミ出し",
   "manual": true,
   "items": []
  },
  {
   "id": "gomi_fee",
   "name": "ゴミ処分費",
   "manual": true,
   "items": []
  },
  {
   "id": "overhead",
   "name": "諸経費",
   "manual": true,
   "items": []
  }
 ],
 "costEnc": {
  "v": 1,
  "salt": "hL6JQ4G8aE+JPjh6bg4fNQ==",
  "iv": "cBL1D78/aaSJLOZ/",
  "data": "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"
 }
};
