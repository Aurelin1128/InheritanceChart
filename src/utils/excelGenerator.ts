import ExcelJS from 'exceljs';
import type { FamilyMember } from '../types';

/**
 * 繼承系統表 - Excel 生成與匯出工具
 * 設計說明：
 * 本模組負責將樹狀的家族繼承人結構，轉換為帶有特定框線、字型、合併儲存格與連線的 Excel 報表。
 * 核心業務邏輯包含：
 * 1. 遞迴或根據 targetRow 計算每位成員在 Excel 中精確的 Row 與 Column 座標。
 * 2. 格式化個人與配偶的資訊文字。
 * 3. 使用 ExcelJS API 設定標楷體字型、欄寬、列高、背景色與框線。
 * 4. 繪製垂直與水平連線，完整還原樹狀關係。
 */

// 根據繼承人深度取得所在的 Excel 欄位索引 (1-based)
// 業務邏輯：曾祖輩在 C 欄，祖父母輩在 E 欄，父母輩在 G 欄，子女輩在 I 欄
export const getColumnIndex = (depth: number): number => {
  switch (depth) {
    case 1: return 3; // C 欄 (曾祖輩)
    case 2: return 5; // E 欄 (祖父母輩)
    case 3: return 7; // G 欄 (父母輩)
    case 4: return 9; // I 欄 (子女輩)
    default: return 3 + (depth - 1) * 2;
  }
};

// 格式化成員的 Excel 顯示文字
// 業務邏輯：每位成員的儲存格需包含：姓名、出生別、生卒年月日、父母姓名與繼承情形
export const formatMemberText = (member: FamilyMember): string => {
  const parts = [];
  
  // 第一行：姓名 (配偶會加註是誰的配偶)
  if (member.isSpouse && member.spouseRelationText) {
    parts.push(member.spouseRelationText);
  } else if (member.successionStatus === 'other' && member.successionStatusText === '被繼承人') {
    parts.push(`被繼承人:${member.name}`);
  } else {
    parts.push(member.name);
  }

  // 第二行：出生別
  parts.push(`出生別:${member.birthOrder}`);

  // 第三行：生卒年
  const deathStr = member.deathDate ? ` 亡:${member.deathDate}` : '';
  parts.push(`生:${member.birthDate}${deathStr}`);

  // 第四行：父母姓名 (配偶通常不顯示此行，但範本中配偶有顯示父母姓名)
  parts.push(`父:${member.fatherName} 母:${member.motherName}`);

  // 第五行：繼承情形 (若有設定才顯示)
  if (member.successionStatusText && member.successionStatusText !== '被繼承人') {
    parts.push(`繼承情形:${member.successionStatusText}`);
  }

  return parts.join('\n');
};

/**
 * 遍歷整棵家族樹，動態計算所有成員的 rowIndex (如果沒有預先指定 targetRow)
 * 設計說明：
 * 採用自適應緊湊排版演算法，確保在使用者新增自訂成員時，生成的 Excel 依然美觀且線條不交叉。
 */
/**
 * 解決行座標衝突與間距規範
 * 業務邏輯規範：
 * 1. 同一欄（世代）的資料不可以重疊。
 * 2. 最低 Row 必須為 1，不可低於 1。
 * 3. 夫妻（配偶關係）在同一欄中必須緊鄰（Row 相差 1）。
 * 4. 非夫妻關係的兩位成員，在同一欄中不可緊鄰，中間必須至少空一格（Row 相差 >= 2）。
 */
export const resolveRowConflicts = (
  rowMap: Map<string, number>,
  root: FamilyMember,
  rootSpouses: Omit<FamilyMember, 'children'>[],
  spousesMap: Record<string, Omit<FamilyMember, 'children'>>
) => {
  // 1. 按世代深度 (depth) 將所有成員與配偶分群
  const depthGroups = new Map<number, string[]>();

  const collectNodes = (member: FamilyMember, depth: number) => {
    if (!depthGroups.has(depth)) {
      depthGroups.set(depth, []);
    }
    if (!depthGroups.get(depth)!.includes(member.id)) {
      depthGroups.get(depth)!.push(member.id);
    }
    if (member.spouseId && spousesMap[member.spouseId]) {
      if (!depthGroups.get(depth)!.includes(member.spouseId)) {
        depthGroups.get(depth)!.push(member.spouseId);
      }
    }
    member.children.forEach(child => collectNodes(child, depth + 1));
  };

  collectNodes(root, 1);
  rootSpouses.forEach(sp => {
    if (!depthGroups.has(1)) depthGroups.set(1, []);
    if (!depthGroups.get(1)!.includes(sp.id)) {
      depthGroups.get(1)!.push(sp.id);
    }
  });

  // 2. 對每個世代深度獨立進行衝突排除與間距調整
  depthGroups.forEach((nodeIds, depth) => {
    interface PlacementUnit {
      id: string;
      isCouple: boolean;
      memberId?: string;
      spouseId?: string;
      preferredRow: number;
      topId?: string;
      bottomId?: string;
    }

    const units: PlacementUnit[] = [];
    const visited = new Set<string>();

    nodeIds.forEach(id => {
      if (visited.has(id)) return;

      if (depth === 1) {
        // 第一代（被繼承人與所有配偶）全部視為獨立單元，強制彼此間隔至少 1 列 (Row 差 >= 2)
        const rowVal = rowMap.get(id) ?? 1;
        units.push({
          id,
          isCouple: false,
          preferredRow: Math.max(1, rowVal)
        });
        visited.add(id);
      } else {
        // 其他代：判斷此節點是否與其他節點構成夫妻單元
        let spouseId: string | undefined;
        let memberId: string | undefined;

        if (spousesMap[id]) {
          spouseId = id;
          memberId = spousesMap[id].spouseId;
        } else {
          const spKey = Object.keys(spousesMap).find(k => spousesMap[k].spouseId === id);
          if (spKey) {
            spouseId = spKey;
            memberId = id;
          }
        }

        if (memberId && spouseId && nodeIds.includes(memberId) && nodeIds.includes(spouseId)) {
          // 夫妻單元 (size = 2)，在同欄中必須相差 1 緊鄰
          const rowM = rowMap.get(memberId) ?? 1;
          const rowS = rowMap.get(spouseId) ?? 1;
          const pref = Math.min(rowM, rowS);

          const topId = rowM <= rowS ? memberId : spouseId;
          const bottomId = rowM <= rowS ? spouseId : memberId;

          units.push({
            id: `${memberId}_spouse`,
            isCouple: true,
            memberId,
            spouseId,
            preferredRow: Math.max(1, pref),
            topId,
            bottomId
          });
          visited.add(memberId);
          visited.add(spouseId);
        } else {
          // 獨立單元
          const rowVal = rowMap.get(id) ?? 1;
          units.push({
            id,
            isCouple: false,
            preferredRow: Math.max(1, rowVal)
          });
          visited.add(id);
        }
      }
    });

    // 3. 排序所有單元（按 preferredRow 升冪排序，若相同則用 ID 穩定防呆排序）
    units.sort((a, b) => {
      if (a.preferredRow !== b.preferredRow) {
        return a.preferredRow - b.preferredRow;
      }
      return a.id.localeCompare(b.id);
    });

    // 4. 依序重新排定各單元 Row 座標，落實最低 Row 1 與非配偶間隔至少 1 列 (差值 >= 2) 的要求
    let nextAvailableRow = 1;
    units.forEach(unit => {
      const startRow = Math.max(unit.preferredRow, nextAvailableRow);
      if (unit.isCouple) {
        rowMap.set(unit.topId!, startRow);
        rowMap.set(unit.bottomId!, startRow + 1);
        nextAvailableRow = startRow + 3; // 夫妻單元本身佔 2 列，下一單元必須 >= startRow + 3 (即與 bottomId 相差 >= 2)
      } else {
        rowMap.set(unit.id, startRow);
        nextAvailableRow = startRow + 2; // 單人單元本身佔 1 列，下一單元必須 >= startRow + 2 (即與此單元相差 >= 2)
      }
    });
  });
};

export const calculateAutoLayout = (
  root: FamilyMember,
  rootSpouses: Omit<FamilyMember, 'children'>[],
  spousesMap: Record<string, Omit<FamilyMember, 'children'>>
): Map<string, number> => {
  const rowMap = new Map<string, number>();
  let currentRow = 2; // 起始列（從 Row 2 開始，與被繼承人同行對齊）

  // 遞迴函數，為成員及其子樹分配行座標
  const assignRows = (member: FamilyMember, isLastSibling: boolean, depth: number) => {
    const hasSpouse = !!member.spouseId && !!spousesMap[member.spouseId];
    
    // 計算該節點夫妻在當前列所佔用的保護高度 (成員 + 配偶 + 間隔)
    let protectedHeight = 1;
    if (hasSpouse) protectedHeight += 1;
    if (!isLastSibling) protectedHeight += 1; // 非最後一個兄弟，多留一空白列作間隔

    const startRow = currentRow;
    rowMap.set(member.id, startRow);

    if (hasSpouse && member.spouseId) {
      rowMap.set(member.spouseId, startRow + 1);
    }

    if (member.children && member.children.length > 0) {
      // 有子女時，子女排在右側。首個子女的起始列與父母對齊
      let childStartRow = startRow;
      
      member.children.forEach((child, index) => {
        const isLastChild = index === member.children.length - 1;
        
        if (index > 0) {
          // 後續子女的起始列，必須在「前一個子樹的結束列」之後，且不能與「父母夫妻保護高度」重疊
          const prevEndRow = Array.from(rowMap.values()).reduce((max, r) => Math.max(max, r), 0);
          childStartRow = Math.max(prevEndRow + 2, startRow + protectedHeight);
        }
        
        // 遞迴分配子節點
        currentRow = childStartRow;
        assignRows(child, isLastChild, depth + 1);
      });
    } else {
      // 葉子節點直接累加當前列指針
      currentRow = startRow + protectedHeight;
    }
  };

  // 分配第一代繼承人 (祖父母輩)
  root.children.forEach((child, index) => {
    const isLast = index === root.children.length - 1;
    assignRows(child, isLast, 2);
  });

  // 設定被繼承人本人的 Row（固定在 Row 2，版面較緊湊）
  // 業務邏輯：被繼承人為繼承表的根節點，固定在 Row 2，使標題與內容不會過於空曠
  rowMap.set(root.id, 2);

  // 設定被繼承人多任配偶的 Row (若有手動 targetRow 則直接用，否則自動分散)
  // 業務邏輯：第一任配偶預設在 Row 4，第二任在 Row 6，依此類推（間距 2，與非配偶間隔規範一致）
  rootSpouses.forEach((spouse, index) => {
    if (spouse.targetRow) {
      rowMap.set(spouse.id, spouse.targetRow);
    } else {
      // 預設將多任配偶均勻排在被繼承人下方，每任間距 2 列
      rowMap.set(spouse.id, 2 + (index + 1) * 2);
    }
  });

  return rowMap;
};

/**
 * 匯出 Excel 主函數
 */
export const exportToExcel = async (
  root: FamilyMember,
  rootSpouses: Omit<FamilyMember, 'children'>[],
  spousesMap: Record<string, Omit<FamilyMember, 'children'>>
) => {
  // 1. 初始化活頁簿與工作表
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('工作表1', {
    views: [{ showGridLines: true }]
  });

  // 2. 設定欄寬 (比照範本的精確設定)
  sheet.columns = [
    { key: 'A', width: 8.88 },
    { key: 'B', width: 13.0 },
    { key: 'C', width: 54.66 }, // 曾祖輩
    { key: 'D', width: 8.88 },  // 連線
    { key: 'E', width: 57.77 }, // 祖父母輩
    { key: 'F', width: 6.77 },  // 連線
    { key: 'G', width: 54.66 }, // 父母輩
    { key: 'H', width: 6.1 },   // 連線
    { key: 'I', width: 54.66 }  // 子女輩
  ];

  // 3. 設定預設列高
  for (let i = 1; i <= 50; i++) {
    sheet.getRow(i).height = 25.2; // 預設高度
  }
  sheet.getRow(2).height = 66.6;  // 標題高度

  // 4. 計算所有成員的 Row 座標
  // 業務邏輯修正：與即時預覽使用完全相同的策略
  //   步驟一：先以自動演算法取得所有成員的預設列位置（確保每個人都有位置）
  //   步驟二：再將手動指定的 targetRow 蓋回（覆蓋 auto 值）
  // 此順序確保：(1) 無手動設定的成員有預設位置，(2) 手動設定的成員精確呈現，(3) Excel 與預覽一致
  const rowMap = calculateAutoLayout(root, rootSpouses, spousesMap);

  // 步驟二：遞迴蒐集所有手動指定 targetRow 並覆蓋自動計算結果
  const applyManualTargetRows = (member: FamilyMember) => {
    if (member.targetRow) rowMap.set(member.id, member.targetRow);
    if (member.spouseId && spousesMap[member.spouseId]) {
      const sp = spousesMap[member.spouseId];
      if (sp.targetRow) rowMap.set(sp.id, sp.targetRow);
    }
    member.children.forEach(applyManualTargetRows);
  };
  applyManualTargetRows(root);
  rootSpouses.forEach(sp => {
    if (sp.targetRow) rowMap.set(sp.id, sp.targetRow);
  });

  // 步驟三：呼叫排解衝突與間距規範函式，保障「最低 Row 1」、「無重疊」、「非配偶間隔 >= 2」與「配偶緊鄰 1」
  resolveRowConflicts(rowMap, root, rootSpouses, spousesMap);

  // 5. 動態計算標題列位置
  // 業務邏輯：標題「繼承系統表」應自動排在所有繼承人內容的正上方，且與內容間空一列
  // 設計說明：
  //   - 找出 rowMap 中所有成員所在的最小列（minContentRow）
  //   - 標題列 = minContentRow - 2（標題 + 1 空白列 + 內容）
  //   - 若空間不足（成員起始列 < 3），則將整個 rowMap 向下平移，確保標題至少在 Row 1
  const allContentRows = Array.from(rowMap.values());
  const minContentRow = allContentRows.length > 0 ? Math.min(...allContentRows) : 5;
  
  // 標題需要 3 列空間：[標題列] [空白列] [內容起始列]
  // 判斷條件：成員起始列至少要 >= 3 才能放得下標題+空白
  const neededGap = 3;
  if (minContentRow < neededGap) {
    // 成員起始列太小，將所有成員向下平移以騰出標題空間
    const shift = neededGap - minContentRow;
    const shiftedEntries = Array.from(rowMap.entries()).map(([k, v]) => [k, v + shift] as [string, number]);
    rowMap.clear();
    shiftedEntries.forEach(([k, v]) => rowMap.set(k, v));
  }

  // 重新取得（可能更新後的）最小列，並計算標題列
  const finalMinRow = Math.min(...Array.from(rowMap.values()));
  // 標題列位於內容起始列的上方 2 列（中間空 1 列作間隔）
  const titleRow = finalMinRow - 2;

  // 設定標題列的列高
  sheet.getRow(titleRow).height = 66.6;

  // 寫入大標題「[被繼承人姓名] 繼承系統表」，水平跨越 C ~ I 欄
  // 業務邏輯：加底線（underline: true）以區隔標題與空白列；中間的空白列 (titleRow+1) 由預設列高自動填充
  sheet.mergeCells(`C${titleRow}:I${titleRow}`);
  const titleCell = sheet.getCell(`C${titleRow}`);
  titleCell.value = `${root.name} 繼承系統表`;
  titleCell.font = { name: '標楷體', size: 48, bold: false, underline: true };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // 6. 輩分列保持留白 (不填寫 曾祖輩/祖父母輩 等字樣)

  // 輔助函式：套用個人資訊儲存格樣式
  const styleMemberCell = (cell: ExcelJS.Cell, member: FamilyMember, colIdx: number) => {
    cell.value = formatMemberText(member);
    cell.font = { name: '標楷體', size: 18, bold: false };
    cell.alignment = { wrapText: true, horizontal: 'left', vertical: 'top' };
    
    // 設定高列高以容納多行文字 (預設為 98pt)
    const rowIdx = Number(cell.row);
    sheet.getRow(rowIdx).height = 98;

    // 邊框邏輯 (依據欄位套用對稱樣式)
    let borderStyle: Partial<ExcelJS.Borders>;

    if (colIdx === 3) {
      // 被繼承人與其配偶 (C 欄)：左中、上中、下中邊框 (右側留空以便水平線連接)
      borderStyle = {
        top: { style: 'medium' },
        bottom: { style: 'medium' },
        left: { style: 'medium' },
        right: undefined
      };
      // 特例配偶右側粗邊框
      if (member.id === 'root-sp-lin-tong') {
        borderStyle.right = { style: 'thick' };
      }
    } else {
      // 所有子代欄位 (E 欄、G 欄、I 欄)：左側一律為粗邊框 (縱向連接線)，其餘為中邊框
      borderStyle = {
        top: { style: 'medium' },
        bottom: { style: 'medium' },
        left: { style: 'thick' },
        right: { style: 'medium' }
      };
    }

    cell.border = borderStyle;
  };

  // 7. 遍歷並寫入所有成員
  const writtenRows = new Set<number>();
  
  const writeMemberToExcel = (member: FamilyMember, depth: number) => {
    const rIdx = rowMap.get(member.id);
    const cIdx = getColumnIndex(depth);

    if (rIdx) {
      const cell = sheet.getCell(rIdx, cIdx);
      styleMemberCell(cell, member, cIdx);
      writtenRows.add(rIdx);

      // 處理配偶
      if (member.spouseId && spousesMap[member.spouseId]) {
        const spouse = spousesMap[member.spouseId];
        const spRIdx = rowMap.get(spouse.id);
        if (spRIdx) {
          const spCell = sheet.getCell(spRIdx, cIdx);
          styleMemberCell(spCell, spouse as FamilyMember, cIdx);
          writtenRows.add(spRIdx);
        }
      }

      // 遞迴寫入子女
      member.children.forEach(child => writeMemberToExcel(child, depth + 1));
    }
  };

  // 寫入被繼承人本人
  writeMemberToExcel(root, 1);

  // 寫入被繼承人的配偶們
  rootSpouses.forEach(sp => {
    const rIdx = rowMap.get(sp.id);
    if (rIdx) {
      const cell = sheet.getCell(rIdx, 3);
      styleMemberCell(cell, sp as FamilyMember, 3);
      writtenRows.add(rIdx);
    }
  });

  // 8. 繪製連接線與邊框延續 (畫樹狀關係)
  
  // (A) 水平連接線 (填寫 '-----')
  const drawHorizontalLines = (member: FamilyMember, depth: number) => {
    const rIdx = rowMap.get(member.id);
    const cIdx = getColumnIndex(depth);

    if (rIdx && member.children && member.children.length > 0) {
      // 在主要儲存格的右邊一欄寫入連線文字 '-----'
      const connCell = sheet.getCell(rIdx, cIdx + 1);
      
      // 根據欄位微調連線長度以符合範本
      if (cIdx + 1 === 4) connCell.value = '-----'; // D 欄
      else if (cIdx + 1 === 6) connCell.value = '-----'; // F 欄
      else if (cIdx + 1 === 8) connCell.value = '----';   // H 欄
      else connCell.value = '----';

      connCell.font = { name: '標楷體', size: 18, bold: false };
      connCell.alignment = { horizontal: 'center', vertical: 'middle' };
      
      // 業務邏輯：被繼承人有子女時，其所有配偶（不論幾任）都需要在 D 欄畫水平連線
      // 設計說明：原本硬寫死特定 ID，改為動態遍歷 rootSpouses 陣列，確保新增第三、四任配偶也能正確畫線
      // 判斷條件：member.id 與 root.id 相同才觸發配偶連線邏輯
      if (member.id === root.id) {
        rootSpouses.forEach(sp => {
          const spRow = rowMap.get(sp.id);
          if (spRow) {
            const c = sheet.getCell(spRow, 4); // D 欄（曾祖輩右側的連線欄）
            c.value = '-----';
            c.font = { name: '標楷體', size: 18 };
            c.alignment = { horizontal: 'center', vertical: 'middle' };
          }
        });
      }
    }
    member.children.forEach(child => drawHorizontalLines(child, depth + 1));
  };
  drawHorizontalLines(root, 1);

  // (B) 垂直連線 (藉由將特定列區間的左邊框設為 thick 繪製)
  const drawVerticalLines = (member: FamilyMember, depth: number) => {
    const rIdx = rowMap.get(member.id);
    const nextColIdx = getColumnIndex(depth + 1);

    if (rIdx && member.children && member.children.length > 0) {
      // 找出該成員的所有子女中，Row 座標的最小值與最大值
      const childRows = member.children
        .map(c => rowMap.get(c.id))
        .filter((r): r is number => r !== undefined);
      
      if (childRows.length > 0) {
        const minChildRow = Math.min(...childRows);
        const maxChildRow = Math.max(...childRows);

        // 業務邏輯：計算父母本人的 Row 以及配偶的 Row，以便將縱向線延伸過去與水平線相交
        let minCoupleRow = rIdx;
        let maxCoupleRow = rIdx;
        if (member.id === root.id) {
          rootSpouses.forEach(sp => {
            const spRow = rowMap.get(sp.id);
            if (spRow) {
              minCoupleRow = Math.min(minCoupleRow, spRow);
              maxCoupleRow = Math.max(maxCoupleRow, spRow);
            }
          });
        } else if (member.spouseId && spousesMap[member.spouseId]) {
          const spRow = rowMap.get(member.spouseId);
          if (spRow) {
            minCoupleRow = Math.min(minCoupleRow, spRow);
            maxCoupleRow = Math.max(maxCoupleRow, spRow);
          }
        }

        const minRow = Math.min(minChildRow, minCoupleRow);
        const maxRow = Math.max(maxChildRow, maxCoupleRow);

        // 將此區間內該欄位 (子代所在欄) 的左邊框皆設為 thick，以畫出連貫的垂直線
        for (let r = minRow; r <= maxRow; r++) {
          const cell = sheet.getCell(r, nextColIdx);
          const currentBorder = cell.border || {};
          
          cell.border = {
            ...currentBorder,
            left: { style: 'thick' }
          };
        }
      }
    }
    member.children.forEach(child => drawVerticalLines(child, depth + 1));
  };
  drawVerticalLines(root, 1);

  // 9. 寫入底端宣告法律聲明
  // 業務邏輯：找尋所有已寫入資料的最大列，在其下方空一列後寫入聲明（+2 = 空 1 列間隔）
  // 判斷條件：至少在 Row 33 以下（避免內容太少時聲明貼太近標題）
  const maxUsedRow = Array.from(writtenRows).reduce((max, r) => Math.max(max, r), 0);
  const declarationRow = Math.max(maxUsedRow + 2, 33); // +2：空 1 列後寫入聲明
  
  sheet.mergeCells(`C${declarationRow}:I${declarationRow}`);
  const declCell = sheet.getCell(`C${declarationRow}`);
  declCell.value = '本系統表係依民法相關規定訂立，如有錯誤或遺漏致他人受損害者，申請人願負法律上一切責任。';
  declCell.font = { name: '標楷體', size: 26, bold: false };
  declCell.alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getRow(declarationRow).height = 103.8; // 設定聲明列的列高

  // 10. 產生二進位資料並以原生 Web 錨點觸發標準 Excel (.xlsx) 檔案下載
  // 業務邏輯：利用瀏覽器原生下載機制，確保副檔名及 MIME 類型在任何瀏覽器中皆能正確保存為 Excel 格式
  // 防禦性設計：過濾姓名中的非法檔名字元（如斜線、問號、點等），確保下載檔名及副檔名 (.xlsx) 永遠能被系統正確解析
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  
  const sanitizedName = root.name.replace(/[\\/:*?"<>|.]/g, '_');
  anchor.download = `完整繼承系統表_${sanitizedName}.xlsx`;
  
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.URL.revokeObjectURL(url);
};
