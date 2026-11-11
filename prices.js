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
 "costEnc": {
  "v": 1,
  "salt": "ydS7f7klNyZL2uwb7nsAvQ==",
  "iv": "rNlH/rs3C5fHX3Df",
  "data": "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"
 }
};
