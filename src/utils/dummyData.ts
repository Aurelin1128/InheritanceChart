import type { FamilyMember } from '../types';

/**
 * 範本預設資料 (空白初始狀態)
 * 設計說明：
 * 本資料為空白初始狀態，系統啟動時不預載任何個人資料。
 * 所有資料皆由使用者自行輸入，且僅儲存於瀏覽器記憶體中，關閉後自動銷毀。
 */

// 空白配偶資料集（初始無任何配偶）
export const dummySpouses: Record<string, Omit<FamilyMember, 'children'>> = {};

// 空白被繼承人根節點
export const dummyDeceasedRoot: FamilyMember = {
  id: 'root-deceased',
  name: '被繼承人姓名',
  gender: 'M',
  birthOrder: '長男',
  birthDate: '民國10年1月1日',
  deathDate: '民國90年1月1日',
  fatherName: '',
  motherName: '',
  successionStatus: 'inherit',
  successionStatusText: '被繼承人',
  children: []
};

// 空白根節點配偶陣列（初始無任何配偶）
export const dummyRootSpouses: Omit<FamilyMember, 'children'>[] = [];
