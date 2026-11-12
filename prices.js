// ============================================================
// アプリ用 単価ファイル（客出し単価＋暗号化した原価）
// ※ tools/build_prices.py（Excelから）または アプリの原価マスター画面から生成します。
//    手で直さないでください。原価は合言葉がないと読めません。
//    生成日: 2026-11-12
// ============================================================
window.PRICE_TABLE = {
 "version": 6,
 "generatedAt": "2026-11-12",
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
  "salt": "tSYCHbRw5cPBLTKGpe/HMw==",
  "iv": "5IThU0j/qnJXs33Q",
  "data": "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"
 }
};
