import type { FamilyMember } from '../types';

/**
 * 繼承系統表 - 範本預設資料 (林吳罔家族)
 * 設計說明：
 * 本資料完全對照「完整繼承系統表.xlsx」中提供的家族繼承關係與個人資訊。
 * 結構包含四代家族關係（曾祖輩、祖父母輩、父母輩、子女輩），
 * 藉由此完整資料，可在系統啟動時即時渲染出相同的關係圖，供使用者參考與測試。
 */

// 曾孫輩 (最底層，即系統表中的「子女輩」)
// 1. 林北金與林秀英的子女
const childrenOfLinBeiJin: FamilyMember[] = [
  {
    id: 'c-lin-de-zhong',
    name: '林得鍾',
    gender: 'F',
    birthOrder: '長女',
    birthDate: '民國65年1月23日',
    deathDate: '',
    fatherName: '林北金',
    motherName: '林秀英',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-lin-xiao-ping',
    name: '林曉蘋',
    gender: 'F',
    birthOrder: '次女',
    birthDate: '民國67年1月14日',
    deathDate: '',
    fatherName: '林北金',
    motherName: '林秀英',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-lin-ya-wen',
    name: '林雅雯',
    gender: 'F',
    birthOrder: '三女',
    birthDate: '民國68年11月6日',
    deathDate: '',
    fatherName: '林北金',
    motherName: '林秀英',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-lin-ya-ling',
    name: '林雅玲',
    gender: 'F',
    birthOrder: '四女',
    birthDate: '民國71年4月5日',
    deathDate: '',
    fatherName: '林北金',
    motherName: '林秀英',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-lin-hui-yun',
    name: '林慧荺',
    gender: 'F',
    birthOrder: '五女',
    birthDate: '民國80年11月28日',
    deathDate: '',
    fatherName: '林北金',
    motherName: '林秀英',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  }
];

// 2. 莊熊豹與陳森的子女
const childrenOfZhuangXiongBao: FamilyMember[] = [
  {
    id: 'c-zhuang-xin-chang',
    name: '莊新昌',
    gender: 'M',
    birthOrder: '長男',
    birthDate: '民國48年12月17日',
    deathDate: '',
    fatherName: '莊熊豹',
    motherName: '陳森',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-chen-chang-jin',
    name: '陳昌進(早夭)',
    gender: 'M',
    birthOrder: '次男',
    birthDate: '民國50年8月15日',
    deathDate: '民國50年9月6日',
    fatherName: '莊熊豹',
    motherName: '陳森',
    successionStatus: 'none-early',
    successionStatusText: '',
    children: []
  },
  {
    id: 'c-chen-chang-ming',
    name: '陳昌銘',
    gender: 'M',
    birthOrder: '三男',
    birthDate: '民國51年6月29日',
    deathDate: '',
    fatherName: '莊熊豹',
    motherName: '陳森',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-zhuang-chang-zhen',
    name: '莊昌鎮',
    gender: 'M',
    birthOrder: '四男',
    birthDate: '民國52年8月27日',
    deathDate: '',
    fatherName: '莊熊豹',
    motherName: '陳森',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-zhuang-fang-zhu',
    name: '莊芳珠',
    gender: 'F',
    birthOrder: '長女',
    birthDate: '民國55年1月17日',
    deathDate: '',
    fatherName: '莊熊豹',
    motherName: '陳森',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'c-zhuang-shou-feng',
    name: '莊收豐',
    gender: 'M',
    birthOrder: '五男',
    birthDate: '民國63年7月28日',
    deathDate: '',
    fatherName: '莊熊豹',
    motherName: '陳森',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  }
];


// 孫輩 (即系統表中的「父母輩」)
// 1. 林文賓與林賴招月的子女
const childrenOfLinWenBin: FamilyMember[] = [
  {
    id: 'p-lin-bei-jin',
    name: '林北金',
    gender: 'M',
    birthOrder: '長男',
    birthDate: '民國40年10月22日',
    deathDate: '民國104年7月23日',
    fatherName: '林文賓',
    motherName: '林賴招月',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    // 關聯配偶
    spouseId: 'sp-lin-xiu-ying',
    spouseRelationText: '林北金配偶:林秀英',
    // 子女
    children: childrenOfLinBeiJin
  },
  {
    id: 'p-lin-xuan-hui',
    name: '林玄輝',
    gender: 'M',
    birthOrder: '次男',
    birthDate: '民國45年5月8日',
    deathDate: '',
    fatherName: '林文賓',
    motherName: '林賴招月',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'p-lin-yu-zhu',
    name: '林玉珠',
    gender: 'F',
    birthOrder: '長女',
    birthDate: '民國42年11月12日',
    deathDate: '',
    fatherName: '林文賓',
    motherName: '林賴招月',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'p-lin-yi-jun',
    name: '林意君',
    gender: 'F',
    birthOrder: '次女',
    birthDate: '民國51年10月6日',
    deathDate: '',
    fatherName: '林文賓',
    motherName: '林賴招月',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  },
  {
    id: 'p-lin-xi-qiang',
    name: '林錫強',
    gender: 'M',
    birthOrder: '三男',
    birthDate: '民國53年6月1日',
    deathDate: '',
    fatherName: '林文賓',
    motherName: '林賴招月',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    children: []
  }
];

// 2. 莊林連與莊居萬的子女
const childrenOfZhuangLinLian: FamilyMember[] = [
  {
    id: 'p-zhuang-shui-jin',
    name: '莊水衿(出養給林銀,改名林水衿)',
    gender: 'F',
    birthOrder: '長女',
    birthDate: '昭和9年5月20日',
    deathDate: '',
    fatherName: '莊居萬',
    motherName: '莊林連',
    successionStatus: 'none-adopt-end',
    successionStatusText: '無繼承權',
    children: []
  },
  {
    id: 'p-zhuang-xiong-bao',
    name: '莊熊豹',
    gender: 'M',
    birthOrder: '長男',
    birthDate: '民國20年10月10日',
    deathDate: '民國90年1月19日',
    fatherName: '莊居萬',
    motherName: '莊林連',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    spouseId: 'sp-chen-sen',
    spouseRelationText: '莊熊豹配偶:陳森',
    children: childrenOfZhuangXiongBao
  }
];


// 兒女輩 (即系統表中的「祖父母輩」)
const successorsOfLinWuWang: FamilyMember[] = [
  {
    id: 'g-lin-wen-bin',
    name: '林文賓',
    gender: 'M',
    birthOrder: '參男',
    birthDate: '民國12年1月21日',
    deathDate: '民國75年12月24日',
    fatherName: '林旺',
    motherName: '林吳罔',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    spouseId: 'sp-lin-lai-zhao-yue',
    spouseRelationText: '林文賓配偶:林賴招月',
    children: childrenOfLinWenBin
  },
  {
    id: 'g-lin-yue-e',
    name: '林月娥(昭和10年11月15日林吳罔養女)',
    gender: 'F',
    birthOrder: '次女',
    birthDate: '大正2年11月27日',
    deathDate: '民國90年3月25日',
    fatherName: '陳遠',
    motherName: '陳劉水錦',
    successionStatus: 'none-adopt-end',
    successionStatusText: '無繼承權（中止領養）',
    children: []
  },
  {
    id: 'g-lin-wan-de',
    name: '林萬得（早夭）',
    gender: 'M',
    birthOrder: '次男',
    birthDate: '大正9年1月27日',
    deathDate: '大正12年10月14日',
    fatherName: '林牛頭',
    motherName: '林吳罔',
    successionStatus: 'none-early',
    successionStatusText: '',
    children: []
  },
  {
    id: 'g-lin-ai',
    name: '林愛（夭折）',
    gender: 'F',
    birthOrder: '次女',
    birthDate: '大正6年11月6日',
    deathDate: '大正6年11月6日',
    fatherName: '林牛頭',
    motherName: '林吳罔',
    successionStatus: 'none-early',
    successionStatusText: '',
    children: []
  },
  {
    id: 'g-lin-tian-ding',
    name: '林添丁（絕嗣）',
    gender: 'M',
    birthOrder: '長男',
    birthDate: '大正3年12月30日',
    deathDate: '民國87年10月25日',
    fatherName: '林通',
    motherName: '林吳罔',
    successionStatus: 'none-extinct',
    successionStatusText: '',
    children: []
  },
  {
    id: 'g-lin-huo-you',
    name: '林火有（早夭）',
    gender: 'F',
    birthOrder: '次女',
    birthDate: '大正2年1月10日',
    deathDate: '大正8年8月28日',
    fatherName: '林通',
    motherName: '林吳罔',
    successionStatus: 'none-early',
    successionStatusText: '',
    children: []
  },
  {
    id: 'g-zhuang-lin-lian',
    name: '莊林連',
    gender: 'F',
    birthOrder: '長女',
    birthDate: '明治43年11月26日',
    deathDate: '昭和11年9月11日',
    fatherName: '林通',
    motherName: '林吳罔',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    spouseId: 'sp-zhuang-ju-wan',
    spouseRelationText: '莊林連配偶:莊居萬',
    children: childrenOfZhuangLinLian
  }
];

// 配偶池，主要為了編輯與資訊比對
export const dummySpouses: Record<string, Omit<FamilyMember, 'children'>> = {
  'sp-lin-lai-zhao-yue': {
    id: 'sp-lin-lai-zhao-yue',
    name: '林賴招月',
    gender: 'F',
    birthOrder: '五女',
    birthDate: '民國21年7月3日',
    deathDate: '民國111年3月21日',
    fatherName: '蔡分',
    motherName: '陳招治 生母:蔡陳年',
    spouseId: 'g-lin-wen-bin',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    isSpouse: true
  },
  'sp-lin-xiu-ying': {
    id: 'sp-lin-xiu-ying',
    name: '林秀英',
    gender: 'F',
    birthOrder: '長女',
    birthDate: '民國44年6月28日',
    deathDate: '民國115年2月11日',
    fatherName: '林振東',
    motherName: '林莊阿欵',
    spouseId: 'p-lin-bei-jin',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    isSpouse: true
  },
  'sp-zhuang-ju-wan': {
    id: 'sp-zhuang-ju-wan',
    name: '莊居萬',
    gender: 'M',
    birthOrder: '長男',
    birthDate: '明治32年8月12日',
    deathDate: '昭和15年12月13日',
    fatherName: '莊呆',
    motherName: '莊陳伖',
    spouseId: 'g-zhuang-lin-lian',
    successionStatus: 'sub-inherit',
    successionStatusText: '再轉繼承',
    isSpouse: true
  },
  'sp-chen-sen': {
    id: 'sp-chen-sen',
    name: '陳森',
    gender: 'F',
    birthOrder: '五女',
    birthDate: '民國23年1月26日',
    deathDate: '',
    fatherName: '楊傳永',
    motherName: '楊陳阿伴',
    spouseId: 'p-zhuang-xiong-bao',
    successionStatus: 'inherit',
    successionStatusText: '繼承',
    isSpouse: true
  }
};

// 曾祖父母輩 (最頂層，即被繼承人林吳罔及多位配偶)
// 注意：為了表示多任配偶，我們在「被繼承人」這個根節點下，除了子女，也可以把其配偶當作一種特殊節點，或者在匯出時單獨處理。
// 範本中，被繼承人有三任配偶：林通(第一任)、林牛頭(第二任)、林旺(第三任)。
export const dummyDeceasedRoot: FamilyMember = {
  id: 'root-deceased',
  name: '林吳罔',
  gender: 'F',
  birthOrder: '長女',
  birthDate: '民前21年8月28日',
  deathDate: '民國49年3月29日',
  fatherName: '吳其生',
  motherName: '吳高蕊',
  successionStatus: 'other',
  successionStatusText: '被繼承人',
  targetRow: 5,
  children: successorsOfLinWuWang
};

// 曾祖父母輩的配偶，與被繼承人「林吳罔」關聯
// 在 Excel 匯出時，第一任配偶林通排在最下面、第二任林牛頭在中間、第三任林旺在上面。
export const dummyRootSpouses: Omit<FamilyMember, 'children'>[] = [
  {
    id: 'root-sp-lin-wang',
    name: '林旺',
    gender: 'M',
    birthOrder: '長女',
    birthDate: '明治17年10月24日',
    deathDate: '大正14年9月17日',
    fatherName: '林功',
    motherName: '林辜市',
    spouseId: 'root-deceased',
    spouseRelationText: '第三任配偶(招夫):林旺',
    successionStatus: 'other',
    targetRow: 8,
    isSpouse: true
  },
  {
    id: 'root-sp-lin-niu-tou',
    name: '林牛頭',
    gender: 'M',
    birthOrder: '次男',
    birthDate: '明治10年7月15日',
    deathDate: '大正9年1月30日',
    fatherName: '林中庸',
    motherName: '劉錦',
    spouseId: 'root-deceased',
    spouseRelationText: '第二任配偶:林牛頭',
    successionStatus: 'other',
    targetRow: 12,
    isSpouse: true
  },
  {
    id: 'root-sp-lin-tong',
    name: '林通',
    gender: 'M',
    birthOrder: '五男',
    birthDate: '明治16年10月4日',
    deathDate: '大正3年10月30日',
    fatherName: '林云',
    motherName: '李愛',
    spouseId: 'root-deceased',
    spouseRelationText: '第一任配偶:林通',
    successionStatus: 'other',
    targetRow: 20,
    isSpouse: true
  }
];
