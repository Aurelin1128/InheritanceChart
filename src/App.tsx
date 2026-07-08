import React, { useState, useEffect } from 'react';
import type {
  FamilyMember,
  SuccessionStatus
} from './types';


import { exportToExcel, calculateAutoLayout, getColumnIndex, formatMemberText, resolveRowConflicts } from './utils/excelGenerator';
import {
  Download,
  Plus,
  Trash2,
  Edit3,
  User,
  Users,
  Heart,
  AlertTriangle,
  Info
} from 'lucide-react';
import './App.css';

/**
 * 繼承系統表網站 - 主應用程式 (App Component)
 * 設計說明：
 * 提供一站式互動的繼承系統表編輯與下載平台。
 * 主要功能：
 * 1. 狀態管理 (State Management)：包含被繼承人樹、多任配偶、一般成員配偶 Map，以及當前編輯中的狀態。
 * 2. 關係樹編輯 (Tree Editor)：左側的階層式清單，可新增子女、配偶，或編輯與刪除成員。
 * 3. 即時模擬預覽 (Live Preview)：右側利用 CSS Grid 依據實體 Excel 排版算法渲染出接近真實 Excel 的圖表。
 * 4. 匯出 Excel (Export)：調用 ExcelJS 生成與範本完全相同的標楷體系統表。
 * 5. 進度存檔 (JSON Backup)：支援將家族資料匯出為 JSON 或從 JSON 匯入，資料完全保留在本地。
 */

// 解析日期字串為年號、年、月、日 (適用於出生日期與死亡日期)
// 業務邏輯：解析範本或現有資料中的「民國12年1月21日」格式以載入下拉選單，並判定是否為空
const parseDateString = (dateStr: string) => {
  const defaultVal = { era: '民國', year: 1, month: 1, day: 1, isEmpty: true };
  if (!dateStr) return defaultVal;

  const regex = /^(民國|民前|明治|大正|昭和|西元)(\d+)年(\d+)月(\d+)日$/;
  const match = dateStr.match(regex);
  if (match) {
    return {
      era: match[1],
      year: parseInt(match[2]),
      month: parseInt(match[3]),
      day: parseInt(match[4]),
      isEmpty: false
    };
  }
  return { era: '民國', year: 1, month: 1, day: 1, isEmpty: false };
};

// 將年號生日自動換算為西元年月日
// 業務邏輯：依據台灣繼承實務上常見的明治、大正、昭和、民國年號，自動折算為西元年
const convertToCommonEra = (era: string, year: number, month: number, day: number): string => {
  if (isNaN(year) || isNaN(month) || isNaN(day)) return '日期無效';

  let ceYear = 0;
  if (era === '民國') ceYear = year + 1911;
  else if (era === '民前') ceYear = 1912 - year;
  else if (era === '明治') ceYear = year + 1867;
  else if (era === '大正') ceYear = year + 1911;
  else if (era === '昭和') ceYear = year + 1925;
  else if (era === '西元') ceYear = year;

  if (ceYear <= 0) {
    return `西元前 ${Math.abs(ceYear) + 1} 年 ${month} 月 ${day} 日`;
  }
  return `西元 ${ceYear} 年 ${month} 月 ${day} 日`;
};

// 將年號日期轉換為西元年整數，用於日期大小比較
// 業務邏輯：計算出生日期與死亡日期對應的西元年份整數，
// 並組合成可比較的數值 (yyyymmdd 格式整數)，供防呆邏輯判斷先後順序
const convertToCEYear = (era: string, year: number): number => {
  if (isNaN(year)) return 0;
  if (era === '民國') return year + 1911;
  if (era === '民前') return 1912 - year;
  if (era === '明治') return year + 1867;
  if (era === '大正') return year + 1911;
  if (era === '昭和') return year + 1925;
  if (era === '西元') return year;
  return 0;
};

// 將年號日期組合成可比較的整數 (yyyymmdd)
// 業務邏輯：透過 yyyymmdd 整數大小比較，可快速判斷兩個日期的先後順序
// 判斷條件：deathDateInt < birthDateInt 即為「死亡日期早於出生日期」的非法情形
const toDateInt = (era: string, year: number, month: number, day: number): number => {
  const ceYear = convertToCEYear(era, year);
  // 格式：年份 * 10000 + 月份 * 100 + 日，例如 1965/3/22 → 19650322
  return ceYear * 10000 + month * 100 + day;
};

// 取得特定年號、年分、月份之最大天數（處理大月、小月、平年、閏年）
// 業務邏輯：確保使用者在網頁上點選的日期為有效曆法日期
const getDaysInMonth = (era: string, year: number, month: number): number => {
  if (isNaN(year) || isNaN(month)) return 31;

  // 計算西元年份以判定閏年
  let ceYear = 0;
  if (era === '民國') ceYear = year + 1911;
  else if (era === '民前') ceYear = 1912 - year;
  else if (era === '明治') ceYear = year + 1867;
  else if (era === '大正') ceYear = year + 1911;
  else if (era === '昭和') ceYear = year + 1925;
  else if (era === '西元') ceYear = year;

  if ([1, 3, 5, 7, 8, 10, 12].includes(month)) return 31;
  if ([4, 6, 9, 11].includes(month)) return 30;
  if (month === 2) {
    // 閏年：四年一閏，百年不閏，四百年又閏
    const isLeap = (ceYear % 4 === 0 && ceYear % 100 !== 0) || (ceYear % 400 === 0);
    return isLeap ? 29 : 28;
  }
  return 31;
};

// 預設乾淨的被繼承人資料 (適合公開部署在網路上時做為初始狀態)
const initialRootMember: FamilyMember = {
  id: 'deceased-root',
  name: '被繼承人',
  gender: 'M',
  birthOrder: '長男',
  birthDate: '民國10年1月1日',
  deathDate: '民國90年1月1日', // 預設提供一個有效死亡日期以通過「一定要填死亡日期」防呆
  fatherName: '',
  motherName: '',
  successionStatus: 'inherit',
  children: []
};

export default function App() {
  // --- 狀態定義 ---
  const [rootMember, setRootMember] = useState<FamilyMember>(initialRootMember);
  const [rootSpouses, setRootSpouses] = useState<Omit<FamilyMember, 'children'>[]>([]);
  const [spousesMap, setSpousesMap] = useState<Record<string, Omit<FamilyMember, 'children'>>>({});

  // 編輯彈窗狀態
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [selectedMember, setSelectedMember] = useState<FamilyMember | null>(null);

  // 當前選取並高亮的成員 ID (與網頁預覽連動)
  const [highlightedId, setHighlightedId] = useState<string | null>(null);

  // 拖曳懸停的背景 Row 索引，用於即時排版放置提示
  const [activeHoverRow, setActiveHoverRow] = useState<number | null>(null);

  // 側邊欄寬度與拖曳縮放狀態
  const [sidebarWidth, setSidebarWidth] = useState(400);
  const [isResizing, setIsResizing] = useState(false);
  const [isMobile, setIsMobile] = useState(window.innerWidth <= 1024);

  // 監聽視窗尺寸，自適應調整 RWD 狀態
  useEffect(() => {
    const handleResize = () => {
      setIsMobile(window.innerWidth <= 1024);
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // 處理分隔線拖動開始
  const handleSplitterMouseDown = (e: React.MouseEvent) => {
    e.preventDefault();
    setIsResizing(true);
  };

  // 全域監聽滑鼠移動，計算側邊欄新寬度
  useEffect(() => {
    if (!isResizing) return;

    const handleMouseMove = (e: MouseEvent) => {
      // 限制寬度介於 280px 到 700px 之間
      const newWidth = Math.max(280, Math.min(700, e.clientX - 20));
      setSidebarWidth(newWidth);
    };

    const handleMouseUp = () => {
      setIsResizing(false);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isResizing]);

  // 監聽並加載/增加到訪人次
  useEffect(() => {
    // 判定是否為本地開發環境，開發環境僅讀取 (GET) 以免洗流量，生產環境才累加 (HIT)
    const isDev = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const url = isDev
      ? 'https://countapi.mileshilliard.com/api/v1/get/linda_heritage_system_table_visits_2026'
      : 'https://countapi.mileshilliard.com/api/v1/hit/linda_heritage_system_table_visits_2026';

    fetch(url)
      .then(res => res.json())
      .then(data => {
        if (data && typeof data.value === 'number') {
          setVisitorCount(data.value);
        }
      })
      .catch(err => {
        console.error('Failed to fetch visitor count:', err);
      });
  }, []);

  // 網頁預覽縮放狀態
  const [zoom, setZoom] = useState(1.0);

  // 根據螢幕尺寸動態計算最適合的縮放比例，防止手機破圖或溢出
  // 業務邏輯：Excel 表格的預設實體總寬度為 980px。
  //   在手機與平板（寬度小於 1024px）下，自動扣除兩側 padding 後計算出剛好能貼合螢幕寬度的最佳 zoom 值，
  //   使任何手機螢幕（如 iPhone 17 的 393px/430px）一進來就能完整看清圖表，不再發生右側溢出。
  useEffect(() => {
    const handleZoomInit = () => {
      const screenWidth = window.innerWidth;
      if (screenWidth < 1024) {
        // 手機與平板：動態計算
        // 扣除主版面 padding (32px) 與預覽區 padding (16px)，約 48px
        const availableWidth = screenWidth - 48;
        const autoZoom = Number((availableWidth / 980).toFixed(2));
        // 限制縮放比例介於 0.35 到 0.95 之間，保證文字不至於過小
        setZoom(Math.max(0.35, Math.min(0.95, autoZoom)));
      } else {
        // 電腦端預設 100%
        setZoom(1.0);
      }
    };
    handleZoomInit();
    
    // 監聽視窗尺寸改變，即時重算以維持最優排版
    window.addEventListener('resize', handleZoomInit);
    return () => window.removeEventListener('resize', handleZoomInit);
  }, []);

  // 出生日期拆解與換算狀態
  const [birthEra, setBirthEra] = useState('民國');
  const [birthYear, setBirthYear] = useState(1);
  const [birthMonth, setBirthMonth] = useState(1);
  const [birthDay, setBirthDay] = useState(1);

  // 監聽出生年、月、年號，自動校正天數（例如 2 月 31 日會自動下調為該月最大天數）
  useEffect(() => {
    const maxDays = getDaysInMonth(birthEra, birthYear, birthMonth);
    if (birthDay > maxDays) {
      setBirthDay(maxDays);
    }
  }, [birthEra, birthYear, birthMonth, birthDay]);

  // 死亡日期拆解、換算與是否死亡狀態
  const [deathEra, setDeathEra] = useState('民國');
  const [deathYear, setDeathYear] = useState(1);
  const [deathMonth, setDeathMonth] = useState(1);
  const [deathDay, setDeathDay] = useState(1);
  const [isDeceased, setIsDeceased] = useState(false);

  // 表單驗證錯誤狀態
  const [formErrors, setFormErrors] = useState<Record<string, string>>({});

  // 無繼承權原因狀態
  const [noInheritReason, setNoInheritReason] = useState('');

  // 到訪人次狀態
  const [visitorCount, setVisitorCount] = useState<number | null>(null);

  // 新手教學引導視窗狀態
  const [isIntroOpen, setIsIntroOpen] = useState(false);

  // 檢查是否為首次進入網站，若是則自動彈出教學
  useEffect(() => {
    const hasSeen = localStorage.getItem('has-seen-intro-guide-v1.2');
    if (!hasSeen) {
      setIsIntroOpen(true);
    }
  }, []);

  // Row 批次調整模式：固定使用等數（delta）模式
  // 業務邏輯：被繼承人或代位/再轉繼承人變更 Row 時，其影響範圍的成員以等數方式一同平移
  const rowAdjustMode: 'delta' | 'ratio' = 'delta';

  // 監聽死亡年、月、年號，自動校正天數
  useEffect(() => {
    const maxDays = getDaysInMonth(deathEra, deathYear, deathMonth);
    if (deathDay > maxDays) {
      setDeathDay(maxDays);
    }
  }, [deathEra, deathYear, deathMonth, deathDay]);

  // --- 輔助函數：深入遍歷搜尋與修改樹狀結構 ---

  // 尋找樹中的某個成員
  const findMemberInTree = (node: FamilyMember, id: string): FamilyMember | null => {
    if (node.id === id) return node;
    for (const child of node.children) {
      const found = findMemberInTree(child, id);
      if (found) return found;
    }
    return null;
  };

  // 在樹中更新某個成員的資料
  const updateMemberInTree = (node: FamilyMember, id: string, updatedData: Partial<FamilyMember>): FamilyMember => {
    if (node.id === id) {
      return { ...node, ...updatedData };
    }
    return {
      ...node,
      children: node.children.map(child => updateMemberInTree(child, id, updatedData))
    };
  };

  // 在樹中為特定父母新增子女
  const addChildToTree = (node: FamilyMember, parentId: string, newChild: FamilyMember): FamilyMember => {
    if (node.id === parentId) {
      return {
        ...node,
        children: [...node.children, newChild]
      };
    }
    return {
      ...node,
      children: node.children.map(child => addChildToTree(child, parentId, newChild))
    };
  };

  // 從樹中刪除某個成員
  const removeMemberFromTree = (node: FamilyMember, id: string): FamilyMember => {
    return {
      ...node,
      children: node.children
        .filter(child => child.id !== id)
        .map(child => removeMemberFromTree(child, id))
    };
  };

  // --- 業務邏輯操作 ---

  // 開啟編輯對話框
  const handleEditClick = (memberId: string) => {
    let memberToEdit: FamilyMember | null = null;

    // 檢查是否是根配偶
    const rootSp = rootSpouses.find(sp => sp.id === memberId);
    if (rootSp) {
      memberToEdit = { ...rootSp, children: [] } as FamilyMember;
    } else if (spousesMap[memberId]) {
      // 檢查是否是一般配偶
      memberToEdit = { ...spousesMap[memberId], children: [] } as FamilyMember;
    } else {
      // 檢查是否是樹中的一般成員或被繼承人
      memberToEdit = findMemberInTree(rootMember, memberId);
    }

    if (memberToEdit) {
      // 若為配偶且 successionStatus 為 'other' (例如舊版本或新建時的預設值)，自動遷移至 'inherit'，確保編輯選單與保存狀態一致
      const tempMember = { ...memberToEdit };
      if (tempMember.isSpouse && tempMember.successionStatus === 'other') {
        tempMember.successionStatus = 'inherit';
        tempMember.successionStatusText = '繼承';
      }
      setSelectedMember(tempMember);
      setFormErrors({});

      // 解析並載入出生日期選單狀態
      const parsedBirth = parseDateString(memberToEdit.birthDate);
      setBirthEra(parsedBirth.era);
      setBirthYear(parsedBirth.year);
      setBirthMonth(parsedBirth.month);
      setBirthDay(parsedBirth.day);

      // 解析並載入死亡日期選單狀態
      const parsedDeath = parseDateString(memberToEdit.deathDate);
      setIsDeceased(!parsedDeath.isEmpty);
      setDeathEra(parsedDeath.era);
      setDeathYear(parsedDeath.year);
      setDeathMonth(parsedDeath.month);
      setDeathDay(parsedDeath.day);

      // 解析無繼承權原因
      let reason = '';
      if (memberToEdit.successionStatus === 'no-inherit' && memberToEdit.successionStatusText) {
        const match = memberToEdit.successionStatusText.match(/無繼承權\("(.*?)"\)/);
        if (match) {
          reason = match[1];
        } else {
          const matchParen = memberToEdit.successionStatusText.match(/無繼承權\((.*?)\)/);
          if (matchParen) {
            reason = matchParen[1];
          } else {
            reason = memberToEdit.successionStatusText.replace('無繼承權', '').replace(/[()（）""]/g, '');
          }
        }
      }
      setNoInheritReason(reason);

      setIsEditModalOpen(true);
    }
  };

  // 核心儲存與級聯調整邏輯 (供表單儲存與拖曳排版共用)
  const performSaveMember = (memberId: string, updated: FamilyMember) => {
    // ===== 共用：Row 調整計算工具函數 =====
    // 業務邏輯：依照 rowAdjustMode 計算受影響成員的新 Row 值
    // - 等數 (delta)：新 Row = 舊 Row + 差值（所有人加減相同的數字）
    // - 等比例 (ratio)：新 Row = round(舊 Row * (新錨點 / 舊錨點))（整體依比率縮放）
    const computeNewRow = (memberOldRow: number, anchorOldRow: number, anchorNewRow: number): number => {
      if (anchorOldRow === anchorNewRow) return Math.max(1, memberOldRow);
      let calculated: number;
      if (rowAdjustMode === 'delta') {
        // 等數：所有受影響成員加上相同的差值
        calculated = memberOldRow + (anchorNewRow - anchorOldRow);
      } else {
        // 等比例：依新舊錨點的比率縮放所有成員的列座標
        // 判斷條件：anchorOldRow 不得為 0，防止除以零
        if (anchorOldRow === 0) return Math.max(1, memberOldRow);
        calculated = Math.round(memberOldRow * (anchorNewRow / anchorOldRow));
      }
      return Math.max(1, calculated);
    };

    // ===== 共用：遞迴套用調整到整棵子樹 =====
    // 設計說明：對子樹每個節點取其「實際目前列」(從 rowMap 查詢，含自動演算值)，
    // 計算新列後存入 targetRow（固定住新位置），以確保之後重新整理時不被演算法覆蓋
    const applyAdjustToSubtree = (
      node: FamilyMember,
      anchorOld: number,
      anchorNew: number
    ): FamilyMember => {
      const currentRow = rowMap.get(node.id);
      const newTargetRow = currentRow ? computeNewRow(currentRow, anchorOld, anchorNew) : node.targetRow;
      return {
        ...node,
        targetRow: newTargetRow,
        children: node.children.map(child => applyAdjustToSubtree(child, anchorOld, anchorNew))
      };
    };

    // ===== 共用：平移配偶 Map 中對應子樹成員的 Row =====
    const applyAdjustToSpousesInSubtree = (
      node: FamilyMember,
      anchorOld: number,
      anchorNew: number,
      nextMap: Record<string, Omit<FamilyMember, 'children'>>
    ): void => {
      // 若此節點有配偶，更新配偶的 Row
      if (node.spouseId && nextMap[node.spouseId]) {
        const spRow = rowMap.get(node.spouseId);
        if (spRow) {
          nextMap[node.spouseId] = {
            ...nextMap[node.spouseId],
            targetRow: computeNewRow(spRow, anchorOld, anchorNew)
          };
        }
      }
      node.children.forEach(child => applyAdjustToSpousesInSubtree(child, anchorOld, anchorNew, nextMap));
    };

    if (memberId === rootMember.id) {
      // ===== 修改被繼承人本人 =====
      // 判斷條件：以 rowMap 取得被繼承人真實的「當前列」作為錨點舊值
      // 即使先前沒有手動 targetRow（純自動排版），也能正確計算差值
      const anchorOld = rowMap.get(rootMember.id) ?? 0;
      const anchorNew = updated.targetRow ?? anchorOld; // 若清空（=自動排版），則無需平移
      const hasChange = anchorOld > 0 && anchorNew !== anchorOld;

      if (hasChange) {
        // 遞迴套用調整到整棵子樹（被繼承人的所有繼承人）
        const shiftedRoot = applyAdjustToSubtree(rootMember, anchorOld, anchorNew);

        // 套用個人資料更新，並以 updated.targetRow 覆蓋被繼承人自身的 Row
        setRootMember({
          ...shiftedRoot,
          name: updated.name,
          gender: updated.gender,
          birthOrder: updated.birthOrder,
          birthDate: updated.birthDate,
          deathDate: updated.deathDate,
          fatherName: updated.fatherName,
          motherName: updated.motherName,
          targetRow: updated.targetRow
        });

        // 被繼承人的多任配偶也一併調整
        setRootSpouses(prev => prev.map(sp => {
          const spRow = rowMap.get(sp.id);
          return spRow ? { ...sp, targetRow: computeNewRow(spRow, anchorOld, anchorNew) } : sp;
        }));

        // spousesMap 中的一般配偶也一併調整
        setSpousesMap(prev => {
          const next = { ...prev };
          // 遞迴遍歷整棵子樹，找出所有配偶並套用調整
          applyAdjustToSpousesInSubtree(rootMember, anchorOld, anchorNew, next);
          return next;
        });
      } else {
        // 無 Row 變動：只更新個人資料
        setRootMember({
          ...rootMember,
          name: updated.name,
          gender: updated.gender,
          birthOrder: updated.birthOrder,
          birthDate: updated.birthDate,
          deathDate: updated.deathDate,
          fatherName: updated.fatherName,
          motherName: updated.motherName,
          targetRow: updated.targetRow
        });
      }

    } else if (rootSpouses.some(sp => sp.id === memberId)) {
      // 修改被繼承人配偶（無子樹，不需要級聯調整）
      setRootSpouses(prev => prev.map(sp => sp.id === memberId ? {
        ...sp,
        name: updated.name,
        gender: updated.gender,
        birthOrder: updated.birthOrder,
        birthDate: updated.birthDate,
        deathDate: updated.deathDate,
        fatherName: updated.fatherName,
        motherName: updated.motherName,
        spouseRelationText: updated.spouseRelationText,
        successionStatus: updated.successionStatus,
        successionStatusText: updated.successionStatusText,
        targetRow: updated.targetRow
      } : sp));

    } else if (spousesMap[memberId]) {
      // 修改一般成員配偶（無子樹，不需要級聯調整）
      setSpousesMap(prev => {
        const next = { ...prev };
        next[memberId] = {
          ...next[memberId],
          name: updated.name,
          gender: updated.gender,
          birthOrder: updated.birthOrder,
          birthDate: updated.birthDate,
          deathDate: updated.deathDate,
          fatherName: updated.fatherName,
          motherName: updated.motherName,
          spouseRelationText: updated.spouseRelationText,
          successionStatus: updated.successionStatus,
          successionStatusText: updated.successionStatusText,
          targetRow: updated.targetRow
        };
        return next;
      });

    } else {
      // ===== 修改樹中一般繼承人 =====
      // 業務邏輯：若繼承情形為「代位繼承」或「再轉繼承」，其後代成員也一併調整
      // 判斷條件：只有這兩種特殊繼承情形才具有「子樹整體平移」的需求
      const isSubstitute = (
        updated.successionStatus === 'substitute-inherit' ||
        updated.successionStatus === 'sub-inherit'
      );
      const anchorOld = rowMap.get(memberId) ?? 0;
      const anchorNew = updated.targetRow ?? anchorOld;
      const hasChange = isSubstitute && anchorOld > 0 && anchorNew !== anchorOld;

      if (hasChange) {
        // 取得此成員在樹中的節點，遞迴套用調整到其所有子嗣
        const memberNode = findMemberInTree(rootMember, memberId);

        // 套用個人資料更新（含新 targetRow）
        setRootMember(prev => {
          // 先更新成員本身資料
          let newTree = updateMemberInTree(prev, memberId, {
            name: updated.name,
            gender: updated.gender,
            birthOrder: updated.birthOrder,
            birthDate: updated.birthDate,
            deathDate: updated.deathDate,
            fatherName: updated.fatherName,
            motherName: updated.motherName,
            successionStatus: updated.successionStatus,
            successionStatusText: updated.successionStatusText,
            targetRow: updated.targetRow
          });

          // 若有子嗣，也一併調整其子樹的 Row
          if (memberNode && memberNode.children.length > 0) {
            const adjustChildren = (node: FamilyMember): FamilyMember => ({
              ...node,
              children: node.children.map(child => {
                if (node.id === memberId) {
                  // 直接子嗣：遞迴套用調整
                  return applyAdjustToSubtree(child, anchorOld, anchorNew);
                }
                return adjustChildren(child);
              })
            });
            newTree = adjustChildren(newTree);
          }
          return newTree;
        });

        // 同步調整此成員子樹中的配偶 Row（從 spousesMap），以及此成員本人的配偶
        setSpousesMap(prev => {
          const next = { ...prev };

          // 業務邏輯：調整此成員「本人的配偶」Row
          // 判斷條件：成員有 spouseId，且配偶存在於 spousesMap
          const ownSpouseId = memberNode?.spouseId;
          if (ownSpouseId && next[ownSpouseId]) {
            const spRow = rowMap.get(ownSpouseId);
            if (spRow) {
              next[ownSpouseId] = {
                ...next[ownSpouseId],
                targetRow: computeNewRow(spRow, anchorOld, anchorNew)
              };
            }
          }

          // 調整子樹中每個後代成員的配偶 Row
          if (memberNode && memberNode.children.length > 0) {
            memberNode.children.forEach(child =>
              applyAdjustToSpousesInSubtree(child, anchorOld, anchorNew, next)
            );
          }

          return next;
        });
      } else {
        // 無 Row 變動或一般繼承：只更新個人資料
        setRootMember(prev => updateMemberInTree(prev, memberId, {
          name: updated.name,
          gender: updated.gender,
          birthOrder: updated.birthOrder,
          birthDate: updated.birthDate,
          deathDate: updated.deathDate,
          fatherName: updated.fatherName,
          motherName: updated.motherName,
          successionStatus: updated.successionStatus,
          successionStatusText: updated.successionStatusText,
          targetRow: updated.targetRow
        }));
      }
    }
  };

  // 儲存編輯後的成員資料（彈窗表單確認儲存時觸發）
  const handleSaveMember = (updated: FamilyMember) => {
    if (!selectedMember) return;
    performSaveMember(selectedMember.id, updated);
    setIsEditModalOpen(false);
    setSelectedMember(null);
  };

  // 輔助函式：計算單個成員的 Row 變動並將調整連鎖套用至其子樹與配偶，返回新的狀態草稿 (純 Row 與級聯調整，無修改個人資料)
  const applyRowChangeToDrafts = (
    memberId: string,
    newTargetRow: number,
    oldRow: number,
    rootNode: FamilyMember,
    rootSpList: Omit<FamilyMember, 'children'>[],
    spMap: Record<string, Omit<FamilyMember, 'children'>>,
    currentMap: Map<string, number>
  ) => {
    let nextRoot = { ...rootNode };
    let nextRootSp = [...rootSpList];
    let nextSpMap = { ...spMap };

    // 判斷是否需要連鎖調整子樹
    let isCascade = false;
    let memberObj: FamilyMember | null = null;

    if (memberId === rootNode.id) {
      isCascade = true;
      memberObj = rootNode;
    } else {
      const rootSp = rootSpList.find(sp => sp.id === memberId);
      if (rootSp) {
        memberObj = rootSp as FamilyMember;
      } else if (spMap[memberId]) {
        memberObj = spMap[memberId] as FamilyMember;
      } else {
        memberObj = findMemberInTree(rootNode, memberId);
        if (memberObj) {
          isCascade = (
            memberObj.successionStatus === 'sub-inherit' ||
            memberObj.successionStatus === 'substitute-inherit'
          );
        }
      }
    }

    if (!memberObj) return { rootNode: nextRoot, rootSpList: nextRootSp, spMap: nextSpMap };

    const delta = newTargetRow - oldRow;

    const computeNewRow = (memberOldRow: number, anchorOldRow: number, anchorNewRow: number): number => {
      if (anchorOldRow === anchorNewRow) return Math.max(1, memberOldRow);
      return Math.max(1, memberOldRow + (anchorNewRow - anchorOldRow));
    };

    const applyAdjustToSubtree = (node: FamilyMember, anchorOld: number, anchorNew: number): FamilyMember => {
      const currentRow = currentMap.get(node.id);
      const newTarget = currentRow ? computeNewRow(currentRow, anchorOld, anchorNew) : node.targetRow;
      return {
        ...node,
        targetRow: newTarget,
        children: node.children.map(child => applyAdjustToSubtree(child, anchorOld, anchorNew))
      };
    };

    const applyAdjustToSpousesInSubtree = (
      node: FamilyMember,
      anchorOld: number,
      anchorNew: number,
      nextMap: Record<string, Omit<FamilyMember, 'children'>>
    ) => {
      if (node.spouseId && nextMap[node.spouseId]) {
        const spRow = currentMap.get(node.spouseId);
        if (spRow) {
          nextMap[node.spouseId] = {
            ...nextMap[node.spouseId],
            targetRow: computeNewRow(spRow, anchorOld, anchorNew)
          };
        }
      }
      node.children.forEach(child => applyAdjustToSpousesInSubtree(child, anchorOld, anchorNew, nextMap));
    };

    if (memberId === rootNode.id) {
      if (delta !== 0 && oldRow > 0) {
        const shiftedRoot = applyAdjustToSubtree(rootNode, oldRow, newTargetRow);
        nextRoot = { ...shiftedRoot, targetRow: newTargetRow };

        nextRootSp = nextRootSp.map(sp => {
          const spRow = currentMap.get(sp.id);
          return spRow ? { ...sp, targetRow: computeNewRow(spRow, oldRow, newTargetRow) } : sp;
        });

        applyAdjustToSpousesInSubtree(rootNode, oldRow, newTargetRow, nextSpMap);
      } else {
        nextRoot.targetRow = newTargetRow;
      }
    } else if (rootSpList.some(sp => sp.id === memberId)) {
      nextRootSp = nextRootSp.map(sp => sp.id === memberId ? { ...sp, targetRow: newTargetRow } : sp);
    } else if (spMap[memberId]) {
      nextSpMap[memberId] = { ...nextSpMap[memberId], targetRow: newTargetRow };
    } else {
      if (isCascade && delta !== 0 && oldRow > 0) {
        // 更新目標節點本身及子嗣 Row
        const adjustChildren = (node: FamilyMember): FamilyMember => ({
          ...node,
          children: node.children.map(child => {
            if (node.id === memberId) {
              return applyAdjustToSubtree(child, oldRow, newTargetRow);
            }
            return adjustChildren(child);
          })
        });
        nextRoot = updateMemberInTree(nextRoot, memberId, { targetRow: newTargetRow });
        nextRoot = adjustChildren(nextRoot);

        // 更新目標節點本人配偶 Row 與子樹配偶 Row
        const ownSpouseId = memberObj.spouseId;
        if (ownSpouseId && nextSpMap[ownSpouseId]) {
          const spRow = currentMap.get(ownSpouseId);
          if (spRow) {
            nextSpMap[ownSpouseId] = {
              ...nextSpMap[ownSpouseId],
              targetRow: computeNewRow(spRow, oldRow, newTargetRow)
            };
          }
        }
        memberObj.children.forEach(child =>
          applyAdjustToSpousesInSubtree(child, oldRow, newTargetRow, nextSpMap)
        );
      } else {
        nextRoot = updateMemberInTree(nextRoot, memberId, { targetRow: newTargetRow });
      }
    }

    return { rootNode: nextRoot, rootSpList: nextRootSp, spMap: nextSpMap };
  };

  // 處理拖曳即時預覽單元格改變 Row 座標
  // 業務邏輯：讀取拖放成員的現有資料，更新其 targetRow 為拖放的新 Row，並執行級聯排版
  const handleDragDropRow = (memberId: string, newRow: number) => {
    const oldRow = rowMap.get(memberId) ?? 1;
    const { rootNode, rootSpList, spMap } = applyRowChangeToDrafts(
      memberId,
      newRow,
      oldRow,
      rootMember,
      rootSpouses,
      spousesMap,
      rowMap
    );
    setRootMember(rootNode);
    setRootSpouses(rootSpList);
    setSpousesMap(spMap);
  };

  // 處理兩個成員位置互換 (拖曳 A 置放於 B 上)
  // 業務邏輯：讀取 A 與 B 的舊列座標，以兩階段原子化級聯更新，將 A 的 Row 變為 B 的 Row，B 的 Row 變為 A 的 Row
  const handleDragDropSwap = (dragId: string, dropId: string) => {
    if (dragId === dropId) return;
    const dragOldRow = rowMap.get(dragId);
    const dropOldRow = rowMap.get(dropId);
    if (dragOldRow === undefined || dropOldRow === undefined) return;

    // 第一步：將拖動的 A 移動至 B 舊位置
    const step1 = applyRowChangeToDrafts(
      dragId,
      dropOldRow,
      dragOldRow,
      rootMember,
      rootSpouses,
      spousesMap,
      rowMap
    );
    // 第二步：將被置放的 B 移動至 A 舊位置 (以第一步產生的結果作為新基準)
    const step2 = applyRowChangeToDrafts(
      dropId,
      dragOldRow,
      dropOldRow,
      step1.rootNode,
      step1.rootSpList,
      step1.spMap,
      rowMap
    );

    setRootMember(step2.rootNode);
    setRootSpouses(step2.rootSpList);
    setSpousesMap(step2.spMap);
  };

  // 新增子女
  const handleAddChild = (parentId: string) => {
    const parent = parentId === rootMember.id ? rootMember : findMemberInTree(rootMember, parentId);
    if (!parent) return;

    // 自動估算下一個子女的姓名/性別與排序：採一男一女交替順序
    const numChildren = parent.children.length;
    const gender = numChildren % 2 === 0 ? 'M' : 'F';

    // 計算目前的男/女人數來估算正確的出生序別
    const prefixMap = ['長', '次', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二', '十三', '十四', '十五'];
    let birthOrder = '';
    if (gender === 'M') {
      const numSons = parent.children.filter(c => c.gender === 'M').length;
      const prefix = prefixMap[numSons] || `${numSons + 1}`;
      birthOrder = `${prefix}男`;
    } else {
      const numDaughters = parent.children.filter(c => c.gender === 'F').length;
      const prefix = prefixMap[numDaughters] || `${numDaughters + 1}`;
      birthOrder = `${prefix}女`;
    }

    const newChild: FamilyMember = {
      id: `child-${Date.now()}`,
      name: `新子女${numChildren + 1}`,
      gender: gender,
      birthOrder: birthOrder,
      birthDate: '民國 年 月 日',
      deathDate: '',
      fatherName: parent.gender === 'M' ? parent.name : '',
      motherName: parent.gender === 'F' ? parent.name : '',
      successionStatus: 'inherit',
      successionStatusText: '繼承',
      children: []
    };

    if (parentId === rootMember.id) {
      setRootMember(prev => ({
        ...prev,
        children: [...prev.children, newChild]
      }));
    } else {
      setRootMember(prev => addChildToTree(prev, parentId, newChild));
    }
  };

  // 新增配偶
  const handleAddSpouse = (memberId: string) => {
    const isRoot = memberId === rootMember.id;
    const newSpouseId = `spouse-${Date.now()}`;

    if (isRoot) {
      // 被繼承人可以有多位配偶
      const nextSpouseIndex = rootSpouses.length + 1;
      const newSp: Omit<FamilyMember, 'children'> = {
        id: newSpouseId,
        name: `新配偶${nextSpouseIndex}`,
        gender: rootMember.gender === 'M' ? 'F' : 'M',
        birthOrder: '長女',
        birthDate: '民國 年 月 日',
        deathDate: '',
        fatherName: '',
        motherName: '',
        spouseId: rootMember.id,
        spouseRelationText: `配偶:新配偶${nextSpouseIndex}`,
        successionStatus: 'inherit',
        successionStatusText: '繼承',
        isSpouse: true
      };
      setRootSpouses(prev => [...prev, newSp]);
    } else {
      const member = findMemberInTree(rootMember, memberId);
      if (!member) return;

      // 一般繼承人僅可有一位配偶
      const newSp: Omit<FamilyMember, 'children'> = {
        id: newSpouseId,
        name: `配偶`,
        gender: member.gender === 'M' ? 'F' : 'M',
        birthOrder: '長女',
        birthDate: '民國 年 月 日',
        deathDate: '',
        fatherName: '',
        motherName: '',
        spouseId: member.id,
        spouseRelationText: `${member.name}配偶:新配偶`,
        successionStatus: 'inherit',
        successionStatusText: '繼承',
        isSpouse: true
      };

      setSpousesMap(prev => ({ ...prev, [newSpouseId]: newSp }));
      setRootMember(prev => updateMemberInTree(prev, memberId, { spouseId: newSpouseId }));
    }
  };

  // 刪除成員
  const handleDeleteMember = (memberId: string) => {
    if (memberId === rootMember.id) {
      alert('無法刪除被繼承人！');
      return;
    }

    if (window.confirm('確定要刪除此成員及其底下的所有繼承人關係嗎？')) {
      // 檢查是否是根配偶
      if (rootSpouses.some(sp => sp.id === memberId)) {
        setRootSpouses(prev => prev.filter(sp => sp.id !== memberId));
        return;
      }

      // 檢查是否是一般配偶
      if (spousesMap[memberId]) {
        const associatedMemberId = spousesMap[memberId].spouseId;
        setSpousesMap(prev => {
          const next = { ...prev };
          delete next[memberId];
          return next;
        });
        if (associatedMemberId) {
          setRootMember(prev => updateMemberInTree(prev, associatedMemberId, { spouseId: undefined }));
        }
        return;
      }

      // 刪除一般繼承人
      setRootMember(prev => removeMemberFromTree(prev, memberId));
    }
  };



  // 清除所有資料並重新開始 (僅保留一個空的被繼承人節點)
  // 業務邏輯：提供使用者一鍵清空所有歷史範本數據、建立全新家族樹的工具，含防呆確認彈窗
  const handleClearAllData = () => {
    const confirmClear = window.confirm('⚠️ 警告：您確定要清除所有家族成員資料並重新開始嗎？此操作將會刪除當前建立的所有節點，且無法復原。');
    if (confirmClear) {
      setRootMember({
        id: 'deceased-root',
        name: '被繼承人',
        gender: 'M',
        birthOrder: '長男',
        birthDate: '民國1年1月1日',
        deathDate: '民國90年1月1日',
        fatherName: '',
        motherName: '',
        successionStatus: 'inherit',
        children: [] as FamilyMember[]
      });
      setRootSpouses([]);
      setSpousesMap({});
      setHighlightedId(null);
      setSelectedMember(null);
    }
  };

  // --- 觸發 Excel 下載 ---
  const handleDownloadExcel = () => {
    exportToExcel(rootMember, rootSpouses, spousesMap);
  };

  // --- 關係樹編輯元件 (遞迴渲染) ---
  const renderTreeEditorNode = (member: FamilyMember, depth: number) => {
    const hasSpouse = !!member.spouseId && !!spousesMap[member.spouseId];
    const spouse = hasSpouse && member.spouseId ? spousesMap[member.spouseId] : null;

    return (
      <div key={member.id} className="tree-node-container" style={{ margin: '4px 0' }}>
        <div
          className={`tree-node-item ${highlightedId === member.id ? 'selected' : ''}`}
          onClick={(e) => {
            e.stopPropagation();
            setHighlightedId(member.id);
          }}
          style={{ borderLeft: highlightedId === member.id ? '4px solid var(--primary)' : undefined }}
        >
          <div className="tree-node-header">
            <div className="tree-node-info">
              <span className={`gender-badge ${member.gender.toLowerCase()}`}>
                {member.gender === 'M' ? '男' : '女'}
              </span>
              <span className="node-name">{member.name}</span>
              <span className="node-birth-order" style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                {member.birthOrder}
              </span>
              {member.successionStatusText && member.successionStatusText !== '被繼承人' && (
                <span className={`status-badge status-${member.successionStatus.replace('none-', 'none')}`}>
                  {member.successionStatusText}
                </span>
              )}
            </div>

            <div className="node-actions" onClick={e => e.stopPropagation()}>
              <button className="node-btn" onClick={() => handleEditClick(member.id)} title="編輯詳細資料">
                <Edit3 size={12} />
              </button>
              {/* 
                業務邏輯：繼承情形為「繼承」時，隱藏「新增子女」與「新增配偶」按鈕
                設計說明：繼承情形=繼承，代表此人直接繼承，不需要再往下建立代位/再轉的子嗣鏈
                判斷條件：successionStatus !== 'inherit' 才顯示新增按鈕
              */}
              {member.successionStatus !== 'inherit' && (
                <button className="node-btn" onClick={() => handleAddChild(member.id)} title="新增子女">
                  <Plus size={12} /> 子
                </button>
              )}
              {member.successionStatus !== 'inherit' && !member.spouseId && (
                <button className="node-btn" onClick={() => handleAddSpouse(member.id)} title="新增配偶">
                  <Heart size={12} /> 配
                </button>
              )}
              {member.id !== rootMember.id && (
                <button className="node-btn node-btn-danger" onClick={() => handleDeleteMember(member.id)} title="刪除成員">
                  <Trash2 size={12} />
                </button>
              )}
            </div>
          </div>

          {/* 顯示配偶資訊 */}
          {spouse && (
            <div
              className="tree-node-spouse-badge"
              style={{
                marginTop: '6px',
                padding: '6px 10px',
                background: 'rgba(239, 68, 68, 0.05)',
                borderRadius: '6px',
                borderLeft: '2px solid rgba(239, 68, 68, 0.4)',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center'
              }}
            >
              <span style={{ fontSize: '0.85rem', fontWeight: 500 }}>
                💕 {spouse.name} ({spouse.birthOrder})
              </span>
              <div className="node-actions" onClick={e => e.stopPropagation()}>
                <button className="node-btn" onClick={() => handleEditClick(spouse.id)}>
                  <Edit3 size={10} />
                </button>
                <button className="node-btn node-btn-danger" onClick={() => handleDeleteMember(spouse.id)}>
                  <Trash2 size={10} />
                </button>
              </div>
            </div>
          )}
        </div>

        {/* 遞迴顯示子節點 */}
        {member.children && member.children.length > 0 && (
          <div className="tree-node-children">
            {member.children.map(child => renderTreeEditorNode(child, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  // --- 計算預覽所需佈局與二維網格大小 ---

  // 計算已定義的最大 Row 索引，並產生預覽網格
  // 業務邏輯修正：先執行自動排版演算法取得完整佈局，再將使用者手動指定的 targetRow 蓋回（覆蓋 auto 值）
  // 如此確保：(1) 所有成員都能有預設位置，(2) 已手動設定 Row 的成員會精確顯示在指定列
  const getPreviewLayout = () => {
    // 步驟一：先以自動演算法計算出所有成員的預設列位置
    const rowMap = calculateAutoLayout(rootMember, rootSpouses, spousesMap);

    // 步驟二：蒐集所有已手動指定 targetRow 的成員，並以手動值覆蓋自動計算的結果
    const applyManualRows = (member: FamilyMember) => {
      // 若成員有手動設定 targetRow，以手動值為準
      if (member.targetRow) rowMap.set(member.id, member.targetRow);
      // 若成員有配偶且配偶有手動設定 targetRow
      if (member.spouseId && spousesMap[member.spouseId]) {
        const sp = spousesMap[member.spouseId];
        if (sp.targetRow) rowMap.set(sp.id, sp.targetRow);
      }
      member.children.forEach(applyManualRows);
    };

    // 對被繼承人本人進行檢查（含其整棵子樹）
    applyManualRows(rootMember);
    // 對被繼承人多任配偶進行檢查
    rootSpouses.forEach(sp => {
      if (sp.targetRow) rowMap.set(sp.id, sp.targetRow);
    });

    // 步驟三：呼叫排解衝突與間距規範函式，排解重疊與非配偶間隔
    resolveRowConflicts(rowMap, rootMember, rootSpouses, spousesMap);

    return rowMap;
  };

  const rowMap = getPreviewLayout();
  const maxRow = Array.from(rowMap.values()).reduce((max, r) => Math.max(max, r), 0) + 2; // 多預留幾行

  // 繪製垂直線格子的 Set (記錄要在哪些 row/col 的左邊畫粗線)
  const verticalLines = new Set<string>();
  const horizontalLines = new Map<string, string>();

  const computeGridLines = (member: FamilyMember, depth: number) => {
    const rIdx = rowMap.get(member.id);
    const cIdx = getColumnIndex(depth);

    if (rIdx && member.children && member.children.length > 0) {
      const childRows = member.children
        .map(c => rowMap.get(c.id))
        .filter((r): r is number => r !== undefined);

      if (childRows.length > 0) {
        const minChildR = Math.min(...childRows);
        const maxChildR = Math.max(...childRows);

        // 業務邏輯：縱向連接線的範圍由「父親本人 Row」到「最後一個子女 Row」決定
        // 設計說明：配偶本身已經有水平線（'-----'）接入縱線，不需將配偶 Row 納入縱線範圍
        // 判斷條件：對一般繼承人（非被繼承人），配偶如果在父親下方，
        //   臨舉配偶 Row 將導致縱線多畫一格到配偶行，產生多餘的、沒有子女分支的縱線縮長
        // 例外：被繼承人（depth=1）的配偶（rootSpouses）仍納入範圍，因為配偶列可能在子女之間需要被縱線覆蓋
        let minR: number;
        let maxR: number;
        if (member.id === rootMember.id) {
          // 被繼承人：納入所有配偶 Row ，因為配偶列與子女列可能交錯，縱線需覆蓋整個範圍
          let minCoupleR = rIdx;
          let maxCoupleR = rIdx;
          rootSpouses.forEach(sp => {
            const spRow = rowMap.get(sp.id);
            if (spRow) {
              minCoupleR = Math.min(minCoupleR, spRow);
              maxCoupleR = Math.max(maxCoupleR, spRow);
            }
          });
          minR = Math.min(minChildR, minCoupleR);
          maxR = Math.max(maxChildR, maxCoupleR);
        } else {
          // 一般繼承人：縱線僅從「父親自身 Row」到「子女最遠 Row」，不包含配偶
          // 判斷條件：配偶透過水平線自行接入縱線，不需將配偶 Row 展長縱線，否則相命中只有一個子女時會多畫一段多餘縱線
          minR = Math.min(rIdx, minChildR);
          maxR = Math.max(rIdx, maxChildR);
        }
        const nextC = getColumnIndex(depth + 1);

        // 垂直線範圍
        for (let r = minR; r <= maxR; r++) {
          verticalLines.add(`${r}-${nextC}`);
        }

        // 水平線
        horizontalLines.set(`${rIdx}-${cIdx + 1}`, cIdx + 1 === 4 ? '-----' : cIdx + 1 === 6 ? '-----' : '----');
      }
    }
    member.children.forEach(child => computeGridLines(child, depth + 1));
  };

  computeGridLines(rootMember, 1);

  // 被繼承人多任配偶的水平線
  rootSpouses.forEach(sp => {
    const r = rowMap.get(sp.id);
    if (r) horizontalLines.set(`${r}-4`, '-----');
  });

  return (
    <div className="app-container">

      {/* 頂部導覽列 */}
      <header className="header-container">
        <div className="header-title-area">
          <div className="logo-icon" style={{ background: 'transparent' }}>
            <img src="/logo.png" alt="Logo" style={{ width: '100%', height: '100%', objectFit: 'contain' }} />
          </div>
          <div>
            <h1 className="header-title">繼承系統表生成系統</h1>
            <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', textAlign: 'left' }}>
              繼承系統視覺編輯 · 前端沙盒隱私處理· 符合民法繼承規範  · 輕鬆下載可編輯Excel！
            </p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '12px' }}>
          <button
            className="btn"
            onClick={() => setIsIntroOpen(true)}
            style={{
              background: 'rgba(99, 102, 241, 0.1)',
              color: 'var(--primary)',
              border: '1px solid rgba(99, 102, 241, 0.2)',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <Info size={18} /> 使用教學
          </button>
          <button className="btn btn-danger" onClick={handleClearAllData}>
            <Trash2 size={18} /> 清除所有資料
          </button>
          <button className="btn btn-primary" onClick={handleDownloadExcel}>
            <Download size={18} /> 下載 Excel 系統表
          </button>
        </div>
      </header>

      {/* 主體雙欄排版 */}
      <main
        className="main-layout"
        style={{
          gridTemplateColumns: isMobile ? '1fr' : `${sidebarWidth}px 6px 1fr`,
          gap: isMobile ? '20px' : '0px' // 拖動時不需要間距，分隔線即是間隔
        }}
      >

        {/* 左側：關係樹編輯面板 */}
        <section className="glass-panel sidebar-panel">
          <h2 className="panel-title">
            <Users size={20} color="var(--primary)" /> 家族關係編輯
          </h2>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>

            {/* 被繼承人及多配偶編輯區 */}
            <div style={{
              background: 'var(--primary-light)',
              padding: '12px',
              borderRadius: '12px',
              border: '1px solid var(--primary)'
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontWeight: 'bold', color: 'var(--primary)' }}>👑 被繼承人</span>
                <div className="node-actions">
                  <button className="node-btn" onClick={() => handleEditClick(rootMember.id)}>
                    <Edit3 size={12} /> 編輯
                  </button>
                  <button className="node-btn" onClick={() => handleAddSpouse(rootMember.id)}>
                    <Heart size={12} /> 加配偶
                  </button>
                  <button className="node-btn" onClick={() => handleAddChild(rootMember.id)}>
                    <Plus size={12} /> 加子女
                  </button>
                </div>
              </div>
              <div style={{ marginTop: '6px', fontSize: '0.95rem', fontWeight: 600 }}>
                {rootMember.name} (生: {rootMember.birthDate})
              </div>

              {/* 顯示被繼承人的多位配偶 */}
              {rootSpouses.map((sp) => (
                <div key={sp.id} style={{
                  marginTop: '8px',
                  padding: '6px 8px',
                  background: 'rgba(255,255,255,0.6)',
                  borderRadius: '6px',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  fontSize: '0.85rem'
                }}>
                  <span>💍 {sp.spouseRelationText || `配偶: ${sp.name}`}</span>
                  <div className="node-actions">
                    <button className="node-btn" onClick={() => handleEditClick(sp.id)}>
                      <Edit3 size={10} />
                    </button>
                    <button className="node-btn node-btn-danger" onClick={() => handleDeleteMember(sp.id)}>
                      <Trash2 size={10} />
                    </button>
                  </div>
                </div>
              ))}
            </div>

            {/* 子女樹狀編輯區 */}
            <div className="tree-view-container" style={{ marginTop: '10px' }}>
              <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '8px', fontWeight: 'bold' }}>
                繼承人分支樹：
              </div>
              {rootMember.children.map(child => renderTreeEditorNode(child, 2))}
            </div>

          </div>
        </section>

        {!isMobile && (
          <div
            className="layout-splitter"
            onMouseDown={handleSplitterMouseDown}
          />
        )}

        {/* 右側：即時網頁預覽 */}
        <section className="glass-panel preview-panel" style={{ padding: '24px' }}>
          <div className="preview-header" style={{ display: 'flex', flexDirection: 'column', gap: '8px', alignItems: 'flex-start' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', flexWrap: 'wrap', gap: '12px' }}>
              <h2 className="panel-title" style={{ margin: 0, border: 'none', padding: 0 }}>
                <Info size={20} color="var(--info)" /> Excel 排版即時預覽
              </h2>
              {/* RWD 縮放控制器 */}
              <div className="zoom-controller" style={{ display: 'flex', alignItems: 'center', gap: '8px', background: '#f1f5f9', padding: '4px 12px', borderRadius: '20px', border: '1px solid #e2e8f0' }}>
                <span style={{ fontSize: '0.75rem', fontWeight: 600, color: '#64748b' }}>預覽比例：</span>
                <button 
                  onClick={() => setZoom(prev => Math.max(0.4, Number((prev - 0.1).toFixed(1))))}
                  style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: '1rem', padding: '0 4px', fontWeight: 'bold', color: '#64748b' }}
                >
                  -
                </button>
                <input 
                  type="range" 
                  min="0.4" 
                  max="1.5" 
                  step="0.1" 
                  value={zoom} 
                  onChange={e => setZoom(parseFloat(e.target.value))}
                  style={{ width: '80px', height: '4px', cursor: 'pointer' }}
                />
                <button 
                  onClick={() => setZoom(prev => Math.min(1.5, Number((prev + 0.1).toFixed(1))))}
                  style={{ border: 'none', background: 'none', cursor: 'pointer', fontSize: '1rem', padding: '0 4px', fontWeight: 'bold', color: '#64748b' }}
                >
                  +
                </button>
                <span style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--primary)', minWidth: '40px', textAlign: 'right' }}>
                  {Math.round(zoom * 100)}%
                </span>
                <button
                  onClick={() => {
                    const width = window.innerWidth;
                    if (width < 640) setZoom(0.5);
                    else if (width < 1024) setZoom(0.75);
                    else setZoom(1.0);
                  }}
                  style={{ border: 'none', background: 'rgba(99, 102, 241, 0.1)', color: 'var(--primary)', padding: '2px 8px', borderRadius: '4px', fontSize: '0.7rem', cursor: 'pointer', fontWeight: 600 }}
                >
                  重設
                </button>
              </div>
            </div>
            <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: '4px' }}>
              <AlertTriangle size={14} color="var(--warning)" />
              預覽文字採用標楷體呈現，點擊方框可直接選中編輯；在手機平板上可向下滑動或拖拉比例查看。
            </span>
          </div>

          <div className="preview-viewport" style={{ overflow: 'auto', position: 'relative' }}>
            <div 
              style={{
                transform: `scale(${zoom})`,
                transformOrigin: 'top left',
                width: `${980 * zoom}px`,
                height: `${(maxRow * 98 + 320) * zoom}px`,
                transition: 'transform 0.15s ease, width 0.15s ease, height 0.15s ease'
              }}
            >
              <div className="excel-grid-preview" style={{ margin: 0 }}>
              <div style={{ textAlign: 'center', fontSize: '2.5rem', marginBottom: '20px', fontFamily: 'var(--font-kai)' }}>
                {rootMember.name} 繼承系統表
              </div>

              {/* 輩分標頭 (保持留白) */}
              <div className="excel-gen-headers">
                <div className="excel-gen-title"></div>
                <div></div>
                <div className="excel-gen-title"></div>
                <div></div>
                <div className="excel-gen-title"></div>
                <div></div>
                <div className="excel-gen-title"></div>
              </div>

              {/* 模擬 Excel Grid */}
              <div
                className="excel-canvas"
                style={{
                  gridTemplateRows: `repeat(${maxRow}, 98px)` // 每一行固定 Excel 高度
                }}
              >
                {/* 背景格線與 Row 號標示標頭 */}
                {Array.from({ length: maxRow }, (_, rIdx) => rIdx + 1).map(rowNum => {
                  const isHovered = activeHoverRow === rowNum;
                  return (
                    <React.Fragment key={`row-header-group-${rowNum}`}>
                      {/* Row 序號標頭 */}
                      <div
                        style={{
                          gridColumn: 1,
                          gridRow: rowNum,
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          background: isHovered ? '#cbd5e1' : '#f8fafc',
                          borderRight: '2px solid #cbd5e1',
                          borderBottom: '1px dashed #cbd5e1',
                          fontSize: '0.7rem',
                          color: '#64748b',
                          fontWeight: 'bold',
                          userSelect: 'none',
                          height: '98px',
                          boxSizing: 'border-box',
                          transition: 'background-color 0.15s ease'
                        }}
                      >
                        Row {rowNum}
                      </div>
                      {/* 背景橫向格線 (作為 Drag 置放目標) */}
                      <div
                        onDragOver={(e) => {
                          e.preventDefault(); // 必須 preventDefault 瀏覽器才會允許 drop
                          setActiveHoverRow(rowNum);
                        }}
                        onDragLeave={() => {
                          setActiveHoverRow(null);
                        }}
                        onDrop={(e) => {
                          e.preventDefault();
                          setActiveHoverRow(null);
                          const memberId = e.dataTransfer.getData('text/plain');
                          if (memberId) {
                            handleDragDropRow(memberId, rowNum);
                          }
                        }}
                        style={{
                          gridColumn: '2 / span 7',
                          gridRow: rowNum,
                          borderBottom: '1px dashed #cbd5e1',
                          height: '98px',
                          boxSizing: 'border-box',
                          background: isHovered ? 'rgba(15, 23, 42, 0.05)' : 'transparent',
                          borderTop: isHovered ? '2px solid var(--text-main)' : undefined,
                          borderBottomStyle: isHovered ? 'solid' : 'dashed',
                          borderBottomColor: isHovered ? 'var(--text-main)' : '#cbd5e1',
                          borderBottomWidth: isHovered ? '2px' : '1px',
                          transition: 'background-color 0.15s ease, border 0.15s ease'
                        }}
                      />
                    </React.Fragment>
                  );
                })}

                {/* 1. 渲染被繼承人 (平移至第 2 欄，即 C 欄) */}
                {(() => {
                  const r = rowMap.get(rootMember.id);
                  if (!r) return null;
                  return (
                    <div
                      draggable={true}
                      onDragStart={(e) => {
                        e.dataTransfer.setData('text/plain', rootMember.id);
                      }}
                      onDragOver={(e) => e.preventDefault()}
                      onDrop={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        const dragId = e.dataTransfer.getData('text/plain');
                        if (dragId && dragId !== rootMember.id) {
                          handleDragDropSwap(dragId, rootMember.id);
                        }
                      }}
                      className={`excel-cell-box border-r-none ${highlightedId === rootMember.id ? 'selected' : ''}`}
                      style={{ gridRow: r, gridColumn: 2 }}
                      onClick={() => handleEditClick(rootMember.id)}
                    >
                      {formatMemberText(rootMember)}
                    </div>
                  );
                })()}

                {/* 2. 渲染被繼承人的配偶們 (平移至第 2 欄，即 C 欄) */}
                {rootSpouses.map(sp => {
                  const r = rowMap.get(sp.id);
                  if (!r) return null;
                  return (
                    <div
                      key={sp.id}
                      draggable={true}
                      onDragStart={(e) => {
                        e.dataTransfer.setData('text/plain', sp.id);
                      }}
                      onDragOver={(e) => e.preventDefault()}
                      onDrop={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        const dragId = e.dataTransfer.getData('text/plain');
                        if (dragId && dragId !== sp.id) {
                          handleDragDropSwap(dragId, sp.id);
                        }
                      }}
                      className={`excel-cell-box border-r-none ${highlightedId === sp.id ? 'selected' : ''}`}
                      style={{ gridRow: r, gridColumn: 2 }}
                      onClick={() => handleEditClick(sp.id)}
                    >
                      {formatMemberText(sp as FamilyMember)}
                    </div>
                  );
                })}

                {/* 3. 遞迴渲染所有一般成員與其配偶 (平移 1 欄) */}
                {(() => {
                  const cells: React.ReactNode[] = [];

                  const collectCells = (member: FamilyMember, depth: number) => {
                    const r = rowMap.get(member.id);
                    const c = getColumnIndex(depth);

                    if (r) {
                      // 主要成員：原本的 Col (1, 3, 5, 7) 平移為 (2, 4, 6, 8)
                      const colIdx = c === 3 ? 2 : c === 5 ? 4 : c === 7 ? 6 : 8;
                      let borderClass = '';
                      if (colIdx > 2) borderClass += ' border-l-thick';
                      if (member.children.length > 0 && colIdx === 4) borderClass += ' border-r-thick';

                      cells.push(
                        <div
                          key={member.id}
                          draggable={true}
                          onDragStart={(e) => {
                            e.dataTransfer.setData('text/plain', member.id);
                          }}
                          onDragOver={(e) => e.preventDefault()}
                          onDrop={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            const dragId = e.dataTransfer.getData('text/plain');
                            if (dragId && dragId !== member.id) {
                              handleDragDropSwap(dragId, member.id);
                            }
                          }}
                          className={`excel-cell-box ${borderClass} ${highlightedId === member.id ? 'selected' : ''}`}
                          style={{ gridRow: r, gridColumn: colIdx }}
                          onClick={() => handleEditClick(member.id)}
                        >
                          {formatMemberText(member)}
                        </div>
                      );

                      // 配偶
                      if (member.spouseId && spousesMap[member.spouseId]) {
                        const sp = spousesMap[member.spouseId];
                        const spR = rowMap.get(sp.id);
                        if (spR) {
                          cells.push(
                            <div
                              key={sp.id}
                              draggable={true}
                              onDragStart={(e) => {
                                e.dataTransfer.setData('text/plain', sp.id);
                              }}
                              onDragOver={(e) => e.preventDefault()}
                              onDrop={(e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                const dragId = e.dataTransfer.getData('text/plain');
                                if (dragId && dragId !== sp.id) {
                                  handleDragDropSwap(dragId, sp.id);
                                }
                              }}
                              className={`excel-cell-box ${borderClass} ${highlightedId === sp.id ? 'selected' : ''}`}
                              style={{ gridRow: spR, gridColumn: colIdx }}
                              onClick={() => handleEditClick(sp.id)}
                            >
                              {formatMemberText(sp as FamilyMember)}
                            </div>
                          );
                        }
                      }
                    }
                    member.children.forEach(child => collectCells(child, depth + 1));
                  };

                  rootMember.children.forEach(child => collectCells(child, 2));
                  return cells;
                })()}

                {/* 4. 渲染連線與邊框延伸線 (平移 1 欄) */}
                {Array.from({ length: maxRow }, (_, rIdx) => rIdx + 1).map(r => {
                  // C=2, D=3, E=4, F=5, G=6, H=7, I=8
                  return [3, 5, 7].map(colIdx => {
                    const hasHLine = horizontalLines.has(`${r}-${colIdx === 3 ? 4 : colIdx === 5 ? 6 : 8}`);
                    const hasVLine = verticalLines.has(`${r}-${colIdx === 3 ? 5 : colIdx === 5 ? 7 : 9}`);

                    return (
                      <div
                        key={`${r}-${colIdx}`}
                        className="excel-empty-cell"
                        style={{ gridRow: r, gridColumn: colIdx }}
                      >
                        {/* 繪製水平連接線 (網頁預覽改用一條實體線示意，避免文字換行) */}
                        {hasHLine && (
                          <div className="excel-line-cell">
                            <span className="excel-line-graphic" />
                          </div>
                        )}
                        {/* 繪製垂直線模擬邊框 */}
                        {hasVLine && (
                          <div className="vertical-line-overlay" style={{ right: '0px' }} />
                        )}
                      </div>
                    );
                  });
                })}
              </div>

              {/* 底端宣告 */}
              <div style={{
                marginTop: '40px',
                borderTop: '1px solid #d1d5db',
                paddingTop: '20px',
                fontSize: '1.25rem',
                textAlign: 'left',
                fontFamily: 'var(--font-kai)'
              }}>
                本系統表係依民法相關規定訂立，如有錯誤或遺漏致他人受損害者，申請人願負法律上一切責任。
              </div>
            </div>
          </div>
        </div>

          <div className="preview-tip">
            💡 系統會自動為沒有手動指定 Row 的成員計算最緊湊且不重疊的擺放位置，保證下載的 Excel 線條美觀。
          </div>
        </section>

      </main>

      <footer className="footer-container">
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
          <span style={{ fontSize: '1rem' }}>🔒</span>
          <span style={{ textAlign: 'left' }}>
            <strong>個資安全與隱私聲明</strong>：本系統採用純前端沙盒技術，<strong>本站自身不會上傳或儲存</strong>您輸入的任何家族名冊或個資，數據僅存於瀏覽器暫存中。本站可能使用第三方 Cookie（如 Google 廣告服務）來優化投放，您可於瀏覽器設定中隨時停用。
          </span>
        </div>

        {/* 到訪人次與版本控制資訊 */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
          {visitorCount !== null && (
            <div style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              background: 'rgba(99, 102, 241, 0.06)',
              border: '1px solid rgba(99, 102, 241, 0.12)',
              padding: '4px 12px',
              borderRadius: '16px',
              fontSize: '0.75rem',
              color: 'var(--primary)',
              fontWeight: 600,
              userSelect: 'none'
            }}>
              <span className="visitor-pulse-dot"></span>
              瀏覽人次：{visitorCount.toLocaleString()}
            </div>
          )}

          <div style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '4px',
            background: 'rgba(74, 85, 104, 0.06)',
            border: '1px solid rgba(74, 85, 104, 0.12)',
            padding: '4px 12px',
            borderRadius: '16px',
            fontSize: '0.75rem',
            color: '#4a5568',
            fontWeight: 600,
            userSelect: 'none'
          }}>
            <span>⚙️</span> 版本：v1.3.0 (Build 20260707)
          </div>

          <div style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '4px',
            background: 'rgba(99, 102, 241, 0.06)',
            border: '1px solid rgba(99, 102, 241, 0.12)',
            padding: '4px 12px',
            borderRadius: '16px',
            fontSize: '0.75rem',
            color: 'var(--primary)',
            fontWeight: 600
          }}>
            <span>✉️</span> 意見信箱：<a href="mailto:aurelian1128@gmail.com" style={{ color: 'var(--primary)', textDecoration: 'none' }}>aurelian1128@gmail.com</a>
          </div>
        </div>
      </footer>

      {/* --- 成員編輯表單 Modal --- */}
      {isEditModalOpen && selectedMember && (
        <div className="modal-overlay" onClick={() => setIsEditModalOpen(false)}>
          <div className="glass-panel modal-content" onClick={e => e.stopPropagation()}>
            <h3 style={{ fontSize: '1.25rem', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <User size={20} color="var(--primary)" /> 編輯個人檔案
            </h3>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>

              <div className="form-row-2">
                <div className="form-group">
                  <label>姓名</label>
                  <input
                    type="text"
                    className="form-input"
                    value={selectedMember.name}
                    onChange={e => setSelectedMember({ ...selectedMember, name: e.target.value })}
                  />
                </div>
                <div className="form-group">
                  <label>性別</label>
                  {/* 性別選單：變更時同步重設出生別，避免「男性選到長女」的邏輯衝突 */}
                  <select
                    className="form-select"
                    value={selectedMember.gender}
                    onChange={e => {
                      const newGender = e.target.value as 'M' | 'F';
                      // 業務邏輯：性別切換時，自動將出生別重設為該性別的「長男/長女」預設值
                      // 避免儲存後出現「性別:男，出生別:長女」的矛盾資料
                      const defaultBirthOrder = newGender === 'M' ? '長男' : '長女';
                      setSelectedMember({ ...selectedMember, gender: newGender, birthOrder: defaultBirthOrder });
                    }}
                  >
                    <option value="M">男</option>
                    <option value="F">女</option>
                  </select>
                </div>
              </div>

              <div className="form-group">
                <label>出生別</label>
                {/* 
                  業務邏輯：出生別選單依性別動態過濾
                  - 男性(M)：顯示長男、次男、三男...十五男（共 15 個選項）+ 養子
                  - 女性(F)：顯示長女、次女、三女...十五女（共 15 個選項）+ 養女
                  數字排序採用中文數字（長、次、三～十五），符合台灣繼承實務慣例
                */}
                {(() => {
                  // 中文序數對應表（索引 0 = 長, 1 = 次, 2 以後為 三、四...十五）
                  const orderPrefixes = ['長', '次', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二', '十三', '十四', '十五'];
                  const suffix = selectedMember.gender === 'M' ? '男' : '女';
                  const adoptedOption = selectedMember.gender === 'M' ? '養子' : '養女';
                  return (
                    <select
                      className="form-select"
                      value={selectedMember.birthOrder}
                      onChange={e => setSelectedMember({ ...selectedMember, birthOrder: e.target.value })}
                    >
                      {/* 依性別動態生成長男～十五男 或 長女～十五女 */}
                      {orderPrefixes.map(prefix => (
                        <option key={prefix + suffix} value={prefix + suffix}>
                          {prefix + suffix}
                        </option>
                      ))}
                      {/* 養子 或 養女（依性別顯示對應選項） */}
                      <option value={adoptedOption}>{adoptedOption}</option>
                    </select>
                  );
                })()}
              </div>

              {selectedMember.isSpouse && (
                <div className="form-group">
                  <label>配偶關係說明 (例如: "配偶:林賴招月")</label>
                  <input
                    type="text"
                    className="form-input"
                    value={selectedMember.spouseRelationText || ''}
                    onChange={e => setSelectedMember({ ...selectedMember, spouseRelationText: e.target.value })}
                  />
                </div>
              )}

              {/* 出生日期選單與西元自動換算組件 */}
              <div style={{
                display: 'flex',
                flexDirection: 'column',
                gap: '10px',
                border: '1px solid var(--border-light)',
                padding: '16px',
                borderRadius: '12px',
                background: 'rgba(255, 255, 255, 0.4)'
              }}>
                <label style={{ fontSize: '0.9rem', fontWeight: 600, color: 'var(--text-main)' }}>出生日期</label>
                <div style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  flexWrap: 'wrap'
                }}>
                  {/* 年號 */}
                  <select
                    className="form-select"
                    style={{ flex: '1.5', minWidth: '75px', padding: '8px 4px', fontSize: '0.85rem' }}
                    value={birthEra}
                    onChange={e => setBirthEra(e.target.value)}
                  >
                    <option value="民國">民國</option>
                    <option value="民前">民前</option>
                    <option value="明治">明治</option>
                    <option value="大正">大正</option>
                    <option value="昭和">昭和</option>
                    <option value="西元">西元</option>
                  </select>

                  {/* 年份數字輸入 */}
                  <input
                    type="number"
                    min="1"
                    className="form-input"
                    style={{ flex: '1.2', minWidth: '55px', padding: '8px 6px', fontSize: '0.85rem', textAlign: 'center' }}
                    value={birthYear}
                    onChange={e => setBirthYear(parseInt(e.target.value) || 1)}
                  />
                  <span style={{ fontSize: '0.85rem', fontWeight: 500, color: 'var(--text-muted)' }}>年</span>

                  {/* 月份選單 */}
                  <select
                    className="form-select"
                    style={{ flex: '1.1', minWidth: '50px', padding: '8px 4px', fontSize: '0.85rem' }}
                    value={birthMonth}
                    onChange={e => setBirthMonth(parseInt(e.target.value) || 1)}
                  >
                    {Array.from({ length: 12 }, (_, i) => i + 1).map(m => (
                      <option key={m} value={m}>{m}</option>
                    ))}
                  </select>
                  <span style={{ fontSize: '0.85rem', fontWeight: 500, color: 'var(--text-muted)' }}>月</span>

                  {/* 日期選單 */}
                  <select
                    className="form-select"
                    style={{ flex: '1.1', minWidth: '50px', padding: '8px 4px', fontSize: '0.85rem' }}
                    value={birthDay}
                    onChange={e => setBirthDay(parseInt(e.target.value) || 1)}
                  >
                    {Array.from({ length: getDaysInMonth(birthEra, birthYear, birthMonth) }, (_, i) => i + 1).map(d => (
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                  <span style={{ fontSize: '0.85rem', fontWeight: 500, color: 'var(--text-muted)' }}>日</span>
                </div>

                {/* 自動換算西元顯示區 */}
                <div style={{
                  marginTop: '4px',
                  fontSize: '0.85rem',
                  color: 'var(--primary)',
                  fontWeight: 600,
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  background: 'var(--primary-light)',
                  padding: '6px 12px',
                  borderRadius: '8px'
                }}>
                  <Info size={14} />
                  <span>西元換算：{convertToCommonEra(birthEra, birthYear, birthMonth, birthDay)}</span>
                </div>
              </div>

              {/* 死亡日期切換與選單組件 */}
              <div style={{
                display: 'flex',
                flexDirection: 'column',
                gap: '10px',
                border: formErrors.deathDate ? '2px solid #ef4444' : '1px solid var(--border-light)',
                padding: '16px',
                borderRadius: '12px',
                background: formErrors.deathDate ? '#fdf2f2' : 'rgba(255, 255, 255, 0.4)',
                transition: 'all 0.2s ease-in-out'
              }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <label style={{ fontSize: '0.9rem', fontWeight: 600, color: formErrors.deathDate ? '#b91c1c' : 'var(--text-main)' }}>死亡日期</label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '0.85rem', color: 'var(--text-main)' }}>
                    <input
                      type="checkbox"
                      checked={isDeceased}
                      onChange={e => setIsDeceased(e.target.checked)}
                    />
                    <span>已歿 (註記死亡日期)</span>
                  </label>
                </div>

                {isDeceased ? (
                  <>
                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      flexWrap: 'wrap'
                    }}>
                      {/* 年號 */}
                      <select
                        className="form-select"
                        style={{ flex: '1.5', minWidth: '75px', padding: '8px 4px', fontSize: '0.85rem' }}
                        value={deathEra}
                        onChange={e => setDeathEra(e.target.value)}
                      >
                        <option value="民國">民國</option>
                        <option value="民前">民前</option>
                        <option value="明治">明治</option>
                        <option value="大正">大正</option>
                        <option value="昭和">昭和</option>
                        <option value="西元">西元</option>
                      </select>

                      {/* 年份數字輸入 */}
                      <input
                        type="number"
                        min="1"
                        className="form-input"
                        style={{ flex: '1.2', minWidth: '55px', padding: '8px 6px', fontSize: '0.85rem', textAlign: 'center' }}
                        value={deathYear}
                        onChange={e => setDeathYear(parseInt(e.target.value) || 1)}
                      />
                      <span style={{ fontSize: '0.85rem', fontWeight: 500, color: 'var(--text-muted)' }}>年</span>

                      {/* 月份選單 */}
                      <select
                        className="form-select"
                        style={{ flex: '1.1', minWidth: '50px', padding: '8px 4px', fontSize: '0.85rem' }}
                        value={deathMonth}
                        onChange={e => setDeathMonth(parseInt(e.target.value) || 1)}
                      >
                        {Array.from({ length: 12 }, (_, i) => i + 1).map(m => (
                          <option key={m} value={m}>{m}</option>
                        ))}
                      </select>
                      <span style={{ fontSize: '0.85rem', fontWeight: 500, color: 'var(--text-muted)' }}>月</span>

                      {/* 日期選單 */}
                      <select
                        className="form-select"
                        style={{ flex: '1.1', minWidth: '50px', padding: '8px 4px', fontSize: '0.85rem' }}
                        value={deathDay}
                        onChange={e => setDeathDay(parseInt(e.target.value) || 1)}
                      >
                        {Array.from({ length: getDaysInMonth(deathEra, deathYear, deathMonth) }, (_, i) => i + 1).map(d => (
                          <option key={d} value={d}>{d}</option>
                        ))}
                      </select>
                      <span style={{ fontSize: '0.85rem', fontWeight: 500, color: 'var(--text-muted)' }}>日</span>
                    </div>

                    {/* 自動換算西元顯示區 */}
                    <div style={{
                      marginTop: '4px',
                      fontSize: '0.85rem',
                      color: 'var(--danger)',
                      fontWeight: 600,
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      background: 'rgba(239, 68, 68, 0.08)',
                      padding: '6px 12px',
                      borderRadius: '8px'
                    }}>
                      <Info size={14} />
                      <span>西元換算：{convertToCommonEra(deathEra, deathYear, deathMonth, deathDay)}</span>
                    </div>

                    {/* 即時日期邏輯警告：死亡日期不得早於出生日期 */}
                    {/* 業務邏輯：比較出生與死亡的 yyyymmdd 整數值，若死亡日期 < 出生日期則顯示橘色即時警示 */}
                    {/* 判斷條件：toDateInt(death) < toDateInt(birth)，且兩個日期均為有效數值時才觸發 */}
                    {(() => {
                      const birthInt = toDateInt(birthEra, birthYear, birthMonth, birthDay);
                      const deathInt = toDateInt(deathEra, deathYear, deathMonth, deathDay);
                      // 只有死亡日期確實早於出生日期時才顯示（等於同一天允許，如出生當天死亡）
                      if (deathInt < birthInt) {
                        return (
                          <div style={{
                            marginTop: '6px',
                            fontSize: '0.85rem',
                            color: '#92400e',
                            fontWeight: 600,
                            display: 'flex',
                            alignItems: 'center',
                            gap: '6px',
                            background: 'rgba(251, 191, 36, 0.15)',
                            border: '1px solid rgba(251, 191, 36, 0.5)',
                            padding: '8px 12px',
                            borderRadius: '8px'
                          }}>
                            <AlertTriangle size={14} color="#d97706" />
                            <span>⚠️ 注意：死亡日期（{convertToCommonEra(deathEra, deathYear, deathMonth, deathDay)}）早於出生日期（{convertToCommonEra(birthEra, birthYear, birthMonth, birthDay)}），請確認是否填寫正確！</span>
                          </div>
                        );
                      }
                      return null;
                    })()}
                  </>
                ) : (
                  <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', fontStyle: 'italic', padding: '4px 0' }}>
                    存活中，不顯示死亡日期 (Excel 中會保留空白)。
                  </div>
                )}

                {formErrors.deathDate && (
                  <div style={{
                    color: '#b91c1c',
                    fontSize: '0.85rem',
                    fontWeight: 600,
                    marginTop: '4px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px',
                    background: 'rgba(239, 68, 68, 0.05)',
                    padding: '6px 12px',
                    borderRadius: '6px'
                  }}>
                    <AlertTriangle size={14} color="#ef4444" />
                    <span>{formErrors.deathDate}</span>
                  </div>
                )}
              </div>

              <div className="form-row-2">
                <div className="form-group">
                  <label>父親姓名</label>
                  <input
                    type="text"
                    className="form-input"
                    value={selectedMember.fatherName}
                    onChange={e => setSelectedMember({ ...selectedMember, fatherName: e.target.value })}
                  />
                </div>
                <div className="form-group">
                  <label>母親姓名</label>
                  <input
                    type="text"
                    className="form-input"
                    value={selectedMember.motherName}
                    onChange={e => setSelectedMember({ ...selectedMember, motherName: e.target.value })}
                  />
                </div>
              </div>

              {selectedMember.id !== rootMember.id && (
                <div className="form-group" style={{
                  background: formErrors.successionStatus ? '#fdf2f2' : 'transparent',
                  border: formErrors.successionStatus ? '2px solid #ef4444' : 'none',
                  padding: formErrors.successionStatus ? '16px' : '0px',
                  borderRadius: formErrors.successionStatus ? '12px' : '0px',
                  transition: 'all 0.2s ease-in-out'
                }}>
                  <label style={{ color: formErrors.successionStatus ? '#b91c1c' : 'var(--text-main)', fontWeight: formErrors.successionStatus ? 600 : undefined }}>繼承情形</label>
                  <select
                    className="form-select"
                    style={{ borderColor: formErrors.successionStatus ? '#ef4444' : undefined }}
                    value={selectedMember.successionStatus}
                    onChange={e => {
                      const status = e.target.value as SuccessionStatus;
                      let text = '繼承';
                      if (status === 'substitute-inherit') text = '代位繼承';
                      else if (status === 'sub-inherit') text = '再轉繼承';
                      else if (status === 'no-inherit') text = `無繼承權("${noInheritReason.trim()}")`;

                      setSelectedMember({
                        ...selectedMember,
                        successionStatus: status,
                        successionStatusText: text
                      });
                    }}
                  >
                    <option value="inherit">繼承</option>
                    <option value="substitute-inherit">代位繼承</option>
                    <option value="sub-inherit">再轉繼承</option>
                    <option value="no-inherit">無繼承權</option>
                  </select>

                  {/* 
                    業務邏輯：即時警告 - 繼承人死亡日期早於被繼承人死亡日期，但未選擇代位繼承
                    設計說明：當前編輯成員非被繼承人且非配偶（即為血親繼承人），且已標記為死亡時，
                    如果其死亡日期早於被繼承人之死亡日期，民法規定應由其直系卑親屬「代位繼承」。
                    如果此時繼承狀態未選定為「代位繼承」，顯示警告提示以提醒使用者修改。
                    判斷條件：
                    1. 編輯主體非被繼承人 (isSuccessor = true)
                    2. 編輯主體非配偶 (!selectedMember.isSpouse)
                    3. 編輯主體已歿 (isDeceased = true)
                    4. 繼承狀態並非代位繼承 (successionStatus !== 'substitute-inherit')
                    5. 被繼承人的死亡日期已設定 (rootDeathParsed.isEmpty = false)
                    6. 編輯主體死亡日期整數值小於被繼承人死亡日期整數值 (deathInt < rootDeathInt)
                  */}
                  {(() => {
                    const isSuccessor = selectedMember.id !== rootMember.id;
                    if (isSuccessor && !selectedMember.isSpouse && isDeceased && selectedMember.successionStatus !== 'substitute-inherit') {
                      const rootDeathParsed = parseDateString(rootMember.deathDate);
                      if (!rootDeathParsed.isEmpty) {
                        const rootDeathInt = toDateInt(
                          rootDeathParsed.era,
                          rootDeathParsed.year,
                          rootDeathParsed.month,
                          rootDeathParsed.day
                        );
                        const deathInt = toDateInt(deathEra, deathYear, deathMonth, deathDay);
                        if (deathInt < rootDeathInt) {
                          return (
                            <div style={{
                              marginTop: '6px',
                              fontSize: '0.85rem',
                              color: '#b91c1c',
                              fontWeight: 600,
                              display: 'flex',
                              alignItems: 'center',
                              gap: '6px',
                              background: 'rgba(239, 68, 68, 0.05)',
                              border: '1px solid rgba(239, 68, 68, 0.2)',
                              padding: '8px 12px',
                              borderRadius: '8px'
                            }}>
                              <AlertTriangle size={14} color="#ef4444" />
                              <span>⚠️ 繼承人死亡日期（{convertToCommonEra(deathEra, deathYear, deathMonth, deathDay)}）早於被繼承人死亡日期（{rootMember.deathDate}），繼承狀態僅能選擇「代位繼承」！</span>
                            </div>
                          );
                        }
                      }
                    }
                    return null;
                  })()}

                  {formErrors.successionStatus && (
                    <div style={{
                      color: '#b91c1c',
                      fontSize: '0.85rem',
                      fontWeight: 600,
                      marginTop: '6px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      background: 'rgba(239, 68, 68, 0.05)',
                      padding: '6px 12px',
                      borderRadius: '6px'
                    }}>
                      <AlertTriangle size={14} color="#ef4444" />
                      <span>{formErrors.successionStatus}</span>
                    </div>
                  )}
                </div>
              )}

              {/* 無繼承權原因輸入區 */}
              {selectedMember.id !== rootMember.id && selectedMember.successionStatus === 'no-inherit' && (
                <div className="form-group" style={{
                  background: formErrors.noInheritReason ? '#fdf2f2' : 'transparent',
                  border: formErrors.noInheritReason ? '2px solid #ef4444' : 'none',
                  padding: formErrors.noInheritReason ? '16px' : '0px',
                  borderRadius: formErrors.noInheritReason ? '12px' : '0px',
                  transition: 'all 0.2s ease-in-out'
                }}>
                  <label style={{ color: formErrors.noInheritReason ? '#b91c1c' : 'var(--text-main)', fontWeight: 600 }}>無繼承權原因</label>
                  <input
                    type="text"
                    className="form-input"
                    style={{ borderColor: formErrors.noInheritReason ? '#ef4444' : undefined }}
                    placeholder="請輸入原因（如：拋棄繼承、聲明喪失等）"
                    value={noInheritReason}
                    onChange={e => setNoInheritReason(e.target.value)}
                  />
                  {formErrors.noInheritReason && (
                    <div style={{
                      color: '#b91c1c',
                      fontSize: '0.85rem',
                      fontWeight: 600,
                      marginTop: '6px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      background: 'rgba(239, 68, 68, 0.05)',
                      padding: '6px 12px',
                      borderRadius: '6px'
                    }}>
                      <AlertTriangle size={14} color="#ef4444" />
                      <span>{formErrors.noInheritReason}</span>
                    </div>
                  )}
                </div>
              )}

              <div className="form-group">
                <label>Excel Row 列座標</label>
                {/* 
                  業務邏輯：
                  - value 優先使用 selectedMember.targetRow（使用者手動輸入的值）
                  - 若尚未設定，則從 rowMap 取得演算法自動計算的 Row 值作為初始顯示
                  - 如此使用者開啟視窗即可直接看到目前排在第幾列，不需再切換回預覽確認
                  - 儲存時以輸入框內的數字為準（統一使用 targetRow 欄位儲存）
                  - 若修改的是被繼承人，其他成員的 Row 將以等數方式（delta）一起調整
                  - 若修改的是代位繼承/再轉繼承成員，其後代 Row 也會以等數方式一起調整
                */}
                <input
                  type="number"
                  min={1}
                  className="form-input"
                  value={selectedMember.targetRow ?? rowMap.get(selectedMember.id) ?? ''}
                  onChange={e => {
                    // 判斷條件：有輸入就轉為整數，清空則設為 undefined（回歸演算法自動排版）
                    const v = e.target.value ? parseInt(e.target.value) : undefined;
                    setSelectedMember({ ...selectedMember, targetRow: v });
                  }}
                />
              </div>

            </div>

            <div className="modal-actions" style={{ justifyContent: 'space-between', width: '100%' }}>
              <div>
                {selectedMember.id !== rootMember.id && (
                  <button
                    className="btn btn-danger"
                    onClick={() => {
                      if (window.confirm(`確定要刪除 ${selectedMember.name} 嗎？`)) {
                        handleDeleteMember(selectedMember.id);
                        setIsEditModalOpen(false);
                        setSelectedMember(null);
                      }
                    }}
                  >
                    刪除此成員
                  </button>
                )}
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button className="btn btn-secondary" onClick={() => setIsEditModalOpen(false)}>
                  取消
                </button>
                <button className="btn btn-primary" onClick={() => {
                  const errors: Record<string, string> = {};

                  // 業務邏輯：被繼承人一定要填寫死亡日期
                  if (selectedMember.id === rootMember.id && !isDeceased) {
                    errors.deathDate = '被繼承人為繼承關係之主體，必須填寫死亡日期（請勾選「已歿」並設定日期）！';
                  }

                  // 業務邏輯：代位繼承 或 再轉繼承 必須設定死亡日期
                  const isSubstitute = (
                    selectedMember.successionStatus === 'substitute-inherit' ||
                    selectedMember.successionStatus === 'sub-inherit'
                  );
                  if (isSubstitute && !isDeceased) {
                    errors.deathDate = '繼承情形為「代位繼承」或「再轉繼承」時，必須填寫死亡日期！';
                    errors.successionStatus = '請變更繼承情形，或勾選「已歿」填寫死亡日期！';
                  }

                  // 業務邏輯：無繼承權 必須設定原因
                  if (selectedMember.successionStatus === 'no-inherit' && !noInheritReason.trim()) {
                    errors.noInheritReason = '繼承情形為「無繼承權」時，必須填寫原因！';
                    errors.successionStatus = '請填寫下方的無繼承權原因！';
                  }

                  // 業務邏輯：勾選已歿（註記死亡日期）時，繼承情形不能是普通的「繼承」
                  const isSuccessor = selectedMember.id !== rootMember.id;
                  if (isSuccessor && isDeceased && selectedMember.successionStatus === 'inherit') {
                    errors.successionStatus = '成員已歿時，繼承情形不能為「繼承」，請選擇「代位繼承」或「再轉繼承」！';
                    errors.deathDate = '成員已確認死亡，繼承情形請改為「代位繼承」或「再轉繼承」！';
                  }

                   // 業務邏輯：死亡日期不得早於出生日期
                  // 設計說明：將兩個年號日期轉換為可比較的 yyyymmdd 整數後進行大小比較，
                  // 若 deathDateInt < birthDateInt 表示死亡日期早於出生日期，此為邏輯錯誤應阻止儲存
                  // 判斷條件：僅在「已歿」勾選時執行（未勾選代表存活，無死亡日期不需比較）
                  if (isDeceased) {
                    const birthInt = toDateInt(birthEra, birthYear, birthMonth, birthDay);
                    const deathInt = toDateInt(deathEra, deathYear, deathMonth, deathDay);
                    if (deathInt < birthInt) {
                      errors.deathDate = `死亡日期（${convertToCommonEra(deathEra, deathYear, deathMonth, deathDay)}）不得早於出生日期（${convertToCommonEra(birthEra, birthYear, birthMonth, birthDay)}），請重新確認！`;
                    }
                  }

                  // 業務邏輯：繼承人死亡日期早於被繼承人死亡日期時，僅能選擇「代位繼承」
                  // 設計說明：當非配偶之繼承人先於被繼承人死亡時，依法在辦理繼承時僅能由其直系卑親屬代位繼承。
                  // 如果此時繼承狀態選為其他，則視為填寫錯誤應阻止儲存。
                  // 判斷條件：
                  // 1. 當前編輯者非被繼承人且非配偶 (isSuccessor && !selectedMember.isSpouse)
                  // 2. 當前編輯者已歿 (isDeceased = true)
                  // 3. 被繼承人有死亡日期 (!rootDeathParsed.isEmpty)
                  // 4. 繼承人死亡日期早於被繼承人死亡日期 (deathInt < rootDeathInt)
                  // 5. 繼承狀態非「代位繼承」 (selectedMember.successionStatus !== 'substitute-inherit')
                  if (isSuccessor && !selectedMember.isSpouse && isDeceased) {
                    const rootDeathParsed = parseDateString(rootMember.deathDate);
                    if (!rootDeathParsed.isEmpty) {
                      const rootDeathInt = toDateInt(
                        rootDeathParsed.era,
                        rootDeathParsed.year,
                        rootDeathParsed.month,
                        rootDeathParsed.day
                      );
                      const deathInt = toDateInt(deathEra, deathYear, deathMonth, deathDay);
                      if (deathInt < rootDeathInt && selectedMember.successionStatus !== 'substitute-inherit') {
                        errors.successionStatus = '繼承人死亡日期早於被繼承人死亡日期，繼承狀態僅能選擇「代位繼承」！';
                        errors.deathDate = `繼承人死亡日期（${convertToCommonEra(deathEra, deathYear, deathMonth, deathDay)}）早於被繼承人死亡日期（${rootMember.deathDate}），僅能選擇「代位繼承」！`;
                      }
                    }
                  }

                  if (Object.keys(errors).length > 0) {
                    setFormErrors(errors);
                    alert('儲存失敗！表單填寫有誤，請查看粉紅色警示欄位。');
                    return;
                  }

                  setFormErrors({});

                  const formattedBirthDate = `${birthEra}${birthYear}年${birthMonth}月${birthDay}日`;
                  const formattedDeathDate = isDeceased ? `${deathEra}${deathYear}年${deathMonth}月${deathDay}日` : '';

                  let finalSuccessionStatusText = selectedMember.successionStatusText;
                  if (isSuccessor) {
                    if (selectedMember.successionStatus === 'no-inherit') {
                      finalSuccessionStatusText = `無繼承權("${noInheritReason.trim()}")`;
                    } else if (selectedMember.successionStatus === 'substitute-inherit') {
                      finalSuccessionStatusText = '代位繼承';
                    } else if (selectedMember.successionStatus === 'sub-inherit') {
                      finalSuccessionStatusText = '再轉繼承';
                    } else if (selectedMember.successionStatus === 'inherit') {
                      finalSuccessionStatusText = '繼承';
                    }
                  }

                  // 業務邏輯：判斷 targetRow 是否為「真正的手動指定」
                  // 若使用者打開視窗後未改動數字（selectedMember.targetRow 等於自動演算 autoRow），
                  // 就不儲存為 targetRow（維持 undefined），讓成員繼續使用自動排版
                  // 只有使用者主動改成不同數字，才視為「手動鎖定 Row」
                  const autoRow = rowMap.get(selectedMember.id);
                  const inputRow = selectedMember.targetRow;
                  // 判斷條件：inputRow 存在，且與 autoRow 不同 → 視為手動指定；否則清空
                  const effectiveTargetRow = (inputRow !== undefined && inputRow !== autoRow) ? inputRow : undefined;

                  handleSaveMember({
                    ...selectedMember,
                    birthDate: formattedBirthDate,
                    deathDate: formattedDeathDate,
                    successionStatusText: finalSuccessionStatusText,
                    targetRow: effectiveTargetRow
                  });
                }}>
                  儲存變更
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* --- 新手使用教學引導 Modal --- */}
      {isIntroOpen && (
        <div className="modal-overlay" onClick={() => setIsIntroOpen(false)}>
          <div className="glass-panel modal-content" onClick={e => e.stopPropagation()} style={{ maxWidth: '600px', width: '90%' }}>
            <h3 style={{ fontSize: '1.4rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px', borderBottom: '2px solid rgba(99, 102, 241, 0.1)', paddingBottom: '12px', margin: '0 0 16px 0', color: 'var(--primary)' }}>
              🌟 歡迎使用「線上繼承系統表產生器」！
            </h3>

            <p style={{ fontSize: '0.9rem', color: 'var(--text-main)', lineHeight: '1.6', marginBottom: '16px' }}>
              這是一個完全免費、<strong>個資安全不落地</strong>的專業系統工具。只需簡單三個步驟，您就能輕鬆繪製出符合民法繼承規範的系統表：
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginBottom: '20px' }}>
              <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
                <span style={{ fontSize: '1.25rem', padding: '6px', background: 'rgba(99, 102, 241, 0.1)', borderRadius: '50%', display: 'flex', justifyContent: 'center', alignItems: 'center', width: '32px', height: '32px', flexShrink: 0 }}>1️⃣</span>
                <div>
                  <h4 style={{ fontSize: '0.95rem', fontWeight: 600, color: 'var(--text-main)', margin: '0 0 4px 0' }}>設定「被繼承人」基本資料</h4>
                  <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', margin: 0 }}>
                    請先點選畫面<strong>左上方的第一個節點（被繼承人）</strong>，編輯其姓名與死亡日期等基本資料並儲存。這是一切繼承順位的起點。
                  </p>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
                <span style={{ fontSize: '1.25rem', padding: '6px', background: 'rgba(99, 102, 241, 0.1)', borderRadius: '50%', display: 'flex', justifyContent: 'center', alignItems: 'center', width: '32px', height: '32px', flexShrink: 0 }}>2️⃣</span>
                <div>
                  <h4 style={{ fontSize: '0.95rem', fontWeight: 600, color: 'var(--text-main)', margin: '0 0 4px 0' }}>依序新增「配偶」與「子女」</h4>
                  <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', margin: 0 }}>
                    儲存被繼承人後，點選該卡片下方的<strong>「+配偶」</strong>或<strong>「+子女」</strong>按鈕，便可依序往下建立二代、三代之家族繼承樹，並可隨時編輯個別成員的繼承情形（如拋棄、代位、無繼承權等）。
                  </p>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
                <span style={{ fontSize: '1.25rem', padding: '6px', background: 'rgba(99, 102, 241, 0.1)', borderRadius: '50%', display: 'flex', justifyContent: 'center', alignItems: 'center', width: '32px', height: '32px', flexShrink: 0 }}>3️⃣</span>
                <div>
                  <h4 style={{ fontSize: '0.95rem', fontWeight: 600, color: 'var(--text-main)', margin: '0 0 4px 0' }}>靈活調整版面並匯出</h4>
                  <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', margin: 0 }}>
                    您可以隨時在右側的 Excel 排版預覽中，直接<strong>「拖曳卡片對調位置 (Swap)」</strong>微調直屬成員的行位順序。確認無誤後，點擊右上角的<strong>「下載 Excel 系統表」</strong>即可取得完美排版的 Excel 檔案！
                  </p>
                </div>
              </div>
            </div>

            <div style={{ borderTop: '1px solid var(--border-light)', paddingTop: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '0.8rem', color: 'var(--primary)', fontWeight: 600 }}>
                <span>🔒</span> 純前端防護：本站自身不儲存個資，第三方服務可能使用 Cookie。
              </div>
              <button
                className="btn btn-primary"
                onClick={() => {
                  setIsIntroOpen(false);
                  localStorage.setItem('has-seen-intro-guide-v1.2', 'true');
                }}
                style={{ padding: '8px 20px', borderRadius: '6px' }}
              >
                我知道了，開始使用！
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
