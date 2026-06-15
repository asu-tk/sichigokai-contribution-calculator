import { read, utils } from 'xlsx';
import type {
  ActivityCounts,
  ActivityKind,
  CalculatorSettings,
  ContributionReport,
  HalfKey,
  MonthSheetScan,
  PeriodSelection,
  ReportResult,
  SheetMemberRecord,
  WorkbookScan,
} from './types';

type CellValue = string | number | boolean | Date | null | undefined;

export const DEFAULT_SETTINGS: CalculatorSettings = {
  inspectionPoint: 1,
  trainingPoint: 2,
  firePoint: 3,
  fireCap: 10,
  maxSuggestion: 45000,
};

export const PERIOD_LABELS: Record<PeriodSelection, string> = {
  'first-half': '前期（4月〜9月）',
  'second-half': '後期（10月〜3月）',
  annual: '年間合計',
};

export const ERROR_MESSAGES = {
  'excel-read-failed':
    'Excelファイルを読み取れませんでした。xlsx/xls形式か、ファイルが破損していないか確認してください。',
  'no-target-sheets':
    '集計期間に該当する月別シートが見つかりません。シート名を「4月」「11月」などにしてください。',
  'no-activity-columns':
    '対象シートの3行目に「点検」「訓練」「火災」の列が見つかりません。C列以降の3行目を確認してください。',
  'no-member-names':
    'B列に団員名が入った行が見つかりません。6行目以降のB列を確認してください。',
  'no-calculation-data':
    '計算対象データがありません。対象期間の点検・訓練・火災の列と団員行を確認してください。',
} as const;

const emptyCounts = (): ActivityCounts => ({
  inspection: 0,
  training: 0,
  fire: 0,
});

export const normalizeDigits = (value: string): string =>
  value.replace(/[０-９]/g, (char) =>
    String.fromCharCode(char.charCodeAt(0) - 0xfee0),
  );

const normalizeCellText = (value: CellValue): string =>
  normalizeDigits(String(value ?? ''))
    .replace(/\s+/g, '')
    .trim();

const parseMonthFromSheetName = (sheetName: string): number | null => {
  const normalized = normalizeDigits(sheetName).replace(/\s+/g, '');
  const exact = normalized.match(/^(1[0-2]|[1-9])月$/);
  const loose = normalized.match(/(?:^|[^0-9])(1[0-2]|[1-9])月/);
  const match = exact ?? loose;

  if (!match) {
    return null;
  }

  const month = Number(match[1]);
  return month >= 1 && month <= 12 ? month : null;
};

const getHalf = (month: number): HalfKey =>
  month >= 4 && month <= 9 ? 'first' : 'second';

const detectActivityKind = (cell: CellValue): ActivityKind | null => {
  const text = normalizeCellText(cell);

  if (!text || text.includes('計')) {
    return null;
  }

  if (text.includes('点検')) {
    return 'inspection';
  }

  if (text.includes('訓練')) {
    return 'training';
  }

  if (text.includes('火災')) {
    return 'fire';
  }

  return null;
};

const isPresentMark = (cell: CellValue): boolean => {
  if (typeof cell === 'number') {
    return cell === 1;
  }

  if (typeof cell === 'string') {
    return cell.trim() === '1';
  }

  return false;
};

const addCounts = (target: ActivityCounts, source: ActivityCounts): void => {
  target.inspection += source.inspection;
  target.training += source.training;
  target.fire += source.fire;
};

const selectedSheetsForPeriod = (
  sheets: MonthSheetScan[],
  period: PeriodSelection,
): MonthSheetScan[] => {
  if (period === 'annual') {
    return sheets;
  }

  const half: HalfKey = period === 'first-half' ? 'first' : 'second';
  return sheets.filter((sheet) => sheet.half === half);
};

const calculateBucketPoints = (
  counts: ActivityCounts,
  settings: CalculatorSettings,
): number => {
  const basePoints =
    counts.inspection * settings.inspectionPoint +
    counts.training * settings.trainingPoint;
  const cappedFirePoints = Math.min(
    counts.fire * settings.firePoint,
    settings.fireCap,
  );

  return basePoints + cappedFirePoints;
};

const amountForRate = (
  contributionRate: number,
  settings: CalculatorSettings,
): number => {
  if (contributionRate >= 80) {
    return 0;
  }

  if (contributionRate >= 60) {
    return 10000;
  }

  if (contributionRate >= 40) {
    return 20000;
  }

  if (contributionRate >= 20) {
    return 35000;
  }

  return settings.maxSuggestion;
};

export const parseWorkbook = (buffer: ArrayBuffer): WorkbookScan => {
  const workbook = read(buffer, { type: 'array', cellDates: false });
  const sheets: MonthSheetScan[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const month = parseMonthFromSheetName(sheetName);

    if (!month) {
      return;
    }

    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet) {
      return;
    }

    const rows = utils.sheet_to_json<CellValue[]>(worksheet, {
      header: 1,
      raw: true,
      defval: '',
      blankrows: true,
    }) as CellValue[][];
    const typeRow = rows[2] ?? [];
    const activityColumns = typeRow
      .map((cell, index) => ({ index, kind: detectActivityKind(cell) }))
      .filter(
        (column): column is { index: number; kind: ActivityKind } =>
          column.index >= 2 && column.kind !== null,
      );
    const eventCounts = emptyCounts();

    activityColumns.forEach((column) => {
      eventCounts[column.kind] += 1;
    });

    const members: SheetMemberRecord[] = rows.slice(5).flatMap((row) => {
      const name = String(row[1] ?? '').trim();

      if (!name) {
        return [];
      }

      const counts = emptyCounts();

      activityColumns.forEach((column) => {
        if (isPresentMark(row[column.index])) {
          counts[column.kind] += 1;
        }
      });

      return [
        {
          no: String(row[0] ?? '').trim(),
          name,
          counts,
        },
      ];
    });

    sheets.push({
      sheetName,
      month,
      half: getHalf(month),
      activityColumns,
      members,
      eventCounts,
    });
  });

  return { sheets };
};

export const buildContributionReport = (
  scan: WorkbookScan,
  period: PeriodSelection,
  settings: CalculatorSettings,
): ReportResult => {
  const selectedSheets = selectedSheetsForPeriod(scan.sheets, period);

  if (selectedSheets.length === 0) {
    return {
      ok: false,
      error: {
        code: 'no-target-sheets',
        message: ERROR_MESSAGES['no-target-sheets'],
      },
    };
  }

  if (!selectedSheets.some((sheet) => sheet.activityColumns.length > 0)) {
    return {
      ok: false,
      error: {
        code: 'no-activity-columns',
        message: ERROR_MESSAGES['no-activity-columns'],
      },
    };
  }

  if (!selectedSheets.some((sheet) => sheet.members.length > 0)) {
    return {
      ok: false,
      error: {
        code: 'no-member-names',
        message: ERROR_MESSAGES['no-member-names'],
      },
    };
  }

  const eventCounts = emptyCounts();
  const memberMap = new Map<
    string,
    { no: string; name: string; first: ActivityCounts; second: ActivityCounts }
  >();

  selectedSheets.forEach((sheet) => {
    addCounts(eventCounts, sheet.eventCounts);

    sheet.members.forEach((member) => {
      const existing =
        memberMap.get(member.name) ??
        {
          no: member.no,
          name: member.name,
          first: emptyCounts(),
          second: emptyCounts(),
        };

      if (!existing.no && member.no) {
        existing.no = member.no;
      }

      addCounts(existing[sheet.half], member.counts);
      memberMap.set(member.name, existing);
    });
  });

  const totalEventCount =
    eventCounts.inspection + eventCounts.training + eventCounts.fire;

  if (totalEventCount === 0 || memberMap.size === 0) {
    return {
      ok: false,
      error: {
        code: 'no-calculation-data',
        message: ERROR_MESSAGES['no-calculation-data'],
      },
    };
  }

  const basicActivityPoints =
    eventCounts.inspection * settings.inspectionPoint +
    eventCounts.training * settings.trainingPoint;

  const members = Array.from(memberMap.values())
    .map((member) => {
      const combinedCounts = emptyCounts();
      let earnedPoints = 0;

      if (period === 'first-half') {
        addCounts(combinedCounts, member.first);
        earnedPoints = calculateBucketPoints(member.first, settings);
      } else if (period === 'second-half') {
        addCounts(combinedCounts, member.second);
        earnedPoints = calculateBucketPoints(member.second, settings);
      } else {
        addCounts(combinedCounts, member.first);
        addCounts(combinedCounts, member.second);
        earnedPoints =
          calculateBucketPoints(member.first, settings) +
          calculateBucketPoints(member.second, settings);
      }

      const rawRate =
        basicActivityPoints > 0
          ? (earnedPoints / basicActivityPoints) * 100
          : 0;
      const contributionRate = Math.min(rawRate, 100);

      return {
        no: member.no,
        name: member.name,
        inspection: combinedCounts.inspection,
        training: combinedCounts.training,
        fire: combinedCounts.fire,
        earnedPoints,
        contributionRate,
        suggestionAmount: amountForRate(contributionRate, settings),
      };
    })
    .sort((a, b) => {
      const noA = Number(a.no);
      const noB = Number(b.no);

      if (Number.isFinite(noA) && Number.isFinite(noB) && noA !== noB) {
        return noA - noB;
      }

      return a.name.localeCompare(b.name, 'ja');
    });

  const suggestionTotal = members.reduce(
    (total, member) => total + member.suggestionAmount,
    0,
  );

  const report: ContributionReport = {
    summary: {
      periodLabel: PERIOD_LABELS[period],
      sheetCount: selectedSheets.length,
      inspectionEvents: eventCounts.inspection,
      trainingEvents: eventCounts.training,
      fireEvents: eventCounts.fire,
      basicActivityPoints,
      memberCount: members.length,
      suggestionTotal,
    },
    members,
    scannedSheetNames: selectedSheets.map((sheet) => sheet.sheetName),
  };

  return { ok: true, report };
};
