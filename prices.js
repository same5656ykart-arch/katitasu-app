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
   "unit": "一式",
   "price": 0
  },
  "mawaribuchi": {
   "name": "廻り縁撤去",
   "unit": "一式",
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
   "name": "レベル工事",
   "unit": "一式",
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
   "unit": "一式",
   "price": 0
  },
  "mawari_swap": {
   "name": "廻り縁交換",
   "unit": "一式",
   "price": 0
  },
  "mawari_paint": {
   "name": "廻り縁塗装",
   "unit": "一式",
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
   "name": "床（メイン床材）",
   "note": "リビング・洋室・キッチンの床です。量は間取りから出します。トイレ・洗面室・階段は、部屋ごとに選びます（CFなど）",
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
  },
  {
   "name": "建具",
   "items": [
    "door_replace",
    "knob"
   ],
   "labels": {
    "door_replace": "全部交換",
    "knob": "ドアノブ交換"
   },
   "note": "枚数は間取りから出します（居室＋リビング＋トイレ＋洗面室）"
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
   "name": "廻り縁",
   "exclusive": true,
   "items": [
    "mawari_swap",
    "mawari_paint",
    "mawaribuchi"
   ],
   "labels": {
    "mawari_swap": "交換",
    "mawari_paint": "塗装",
    "mawaribuchi": "撤去"
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
   ],
   "main": true
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
   ],
   "main": true
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
   ],
   "main": true
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
     }
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
  "salt": "mwZjfLSWcKwzreDBI0rZbQ==",
  "iv": "CMRYQlJSAntL40Tv",
  "data": "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"
 }
};
