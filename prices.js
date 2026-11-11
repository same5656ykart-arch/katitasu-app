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
  "salt": "XlJiYl+c30cx+9xMnsFhyA==",
  "iv": "C02xh2iza+iqJyPI",
  "data": "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"
 }
};
