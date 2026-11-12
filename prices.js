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
  "salt": "iDBKsrJExOTXEilMoPHAEw==",
  "iv": "+CJxXE1EATXWXCbJ",
  "data": "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"
 }
};
