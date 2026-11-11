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
  "salt": "UNJLsHuDwBzIStTAJi0CvA==",
  "iv": "6oZXbhHpbMwHO/9+",
  "data": "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"
 }
};
