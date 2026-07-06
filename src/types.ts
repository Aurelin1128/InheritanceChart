/**
 * 繼承系統表 - 成員資料型別定義
 * 設計說明：
 * 用於表示繼承關係中的每位家族成員（包括被繼承人、配偶及繼承人）。
 * 為了應對複雜的台灣繼承實務（例如再轉繼承、拋棄繼承、多任配偶等），
 * 設計了詳盡的屬性以描述個人資訊及繼承狀態。
 */

// 繼承情形的列舉型別
// 業務邏輯：不同的繼承情形會影響在系統表中的文字顯示，也可能影響繼承權的判定
export type SuccessionStatus =
  | 'inherit'       // 繼承：正常的繼承人，且目前存活或有繼承事實
  | 'sub-inherit'   // 再轉繼承：繼承人於被繼承人死亡後、辦理登記前死亡，由其繼承人再行繼承
  | 'substitute-inherit' // 代位繼承：繼承人於被繼承人死亡前死亡，由其直系血親卑親屬代位繼承其應得之分
  | 'no-inherit'    // 無繼承權
  | 'none-early'    // 早夭/夭折：在被繼承人死亡前即死亡，且無代位繼承人，或無繼承權
  | 'none-adopt-end'// 中止領養：曾經收養但已終止收養關係，無繼承權
  | 'none-waive'    // 拋棄繼承：法定繼承人聲明拋棄繼承權
  | 'none-extinct'  // 絕嗣：死亡且無子嗣繼承
  | 'other';        // 其他自訂情形

export interface FamilyMember {
  id: string;             // 唯一識別碼 (可以使用 UUID 或時間戳記)
  name: string;           // 姓名
  gender: 'M' | 'F';      // 性別 ('M' 為男，'F' 為女)
  birthOrder: string;     // 出生別 (例如: 長女、次男、參男、五女)
  
  // 生卒年月日 (在繼承實務上通常需要完整呈現，如「民前21年8月28日」或「明治17年10月24日」)
  birthDate: string;      // 出生年月日描述
  deathDate: string;      // 死亡年月日描述 (若仍存活則留空或填寫「存活」)
  
  fatherName: string;     // 父親姓名
  motherName: string;     // 母親姓名
  
  // 配偶資訊
  // 業務邏輯：在系統表中，配偶通常會排列在該成員的下方，並註明是誰的配偶。
  // 一個成員可能有配偶，若是被繼承人，可能有多任配偶（如林吳罔有三任配偶）。
  isSpouse?: boolean;     // 標記此節點本身是否為配偶節點
  spouseId?: string;      // 若是配偶，關聯的配偶成員 ID；若是一般成員，可關聯其配偶 ID
  spouseRelationText?: string; // 配偶關係說明 (例如: "第一任配偶:林通", "配偶:林賴招月")

  successionStatus: SuccessionStatus; // 繼承情形
  successionStatusText?: string;       // 繼承情形的自訂文字說明 (例如: "無繼承權（中止領養）")
  
  targetRow?: number;     // 為了還原 Excel 特定排版，可手動指定或演算法算出的目標 Row 索引
  notes?: string;         // 其他備註資訊
  
  // 子女列表 (樹狀結構的子節點)
  children: FamilyMember[];
}

/**
 * 輩分層級定義
 * 用於計算 Excel 中的輩分標題（如曾祖輩、祖父母輩等）
 */
export interface GenerationLevel {
  depth: number;          // 深度 (1, 2, 3, 4)
  title: string;          // 輩分標題 (如：曾祖輩、祖父母輩、父母輩、子女輩)
}
