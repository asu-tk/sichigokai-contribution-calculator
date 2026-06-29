export type ActivityKind = 'inspection' | 'training' | 'fire';

export type HalfKey = 'first' | 'second';

export type PeriodSelection = 'first-half' | 'second-half' | 'annual';

export type RoundingUnit = 1000 | 5000;

export type ParseErrorCode =
  | 'excel-read-failed'
  | 'google-sheets-url-invalid'
  | 'google-sheets-read-failed'
  | 'no-target-sheets'
  | 'no-activity-columns'
  | 'no-member-names'
  | 'no-calculation-data';

export interface CalculatorSettings {
  inspectionPoint: number;
  trainingPoint: number;
  firePoint: number;
  fireCap: number;
  cityAnnualCompensation: number;
  membershipFee: number;
  contributionBaseAmount: number;
  roundingUnit: RoundingUnit;
}

export interface ActivityCounts {
  inspection: number;
  training: number;
  fire: number;
}

export interface ActivityColumn {
  index: number;
  kind: ActivityKind;
}

export interface SheetMemberRecord {
  no: string;
  name: string;
  counts: ActivityCounts;
}

export interface MonthSheetScan {
  sheetName: string;
  month: number;
  half: HalfKey;
  activityColumns: ActivityColumn[];
  members: SheetMemberRecord[];
  eventCounts: ActivityCounts;
}

export interface WorkbookScan {
  sheets: MonthSheetScan[];
}

export interface MemberContributionRow {
  no: string;
  name: string;
  inspection: number;
  training: number;
  fire: number;
  earnedPoints: number;
  contributionRate: number;
  suggestionAmount: number;
}

export interface ReportSummary {
  periodLabel: string;
  sheetCount: number;
  inspectionEvents: number;
  trainingEvents: number;
  fireEvents: number;
  basicActivityPoints: number;
  memberCount: number;
  suggestionTotal: number;
}

export interface ContributionReport {
  summary: ReportSummary;
  members: MemberContributionRow[];
  scannedSheetNames: string[];
}

export interface ReportError {
  code: ParseErrorCode;
  message: string;
}

export type ReportResult =
  | { ok: true; report: ContributionReport }
  | { ok: false; error: ReportError };
