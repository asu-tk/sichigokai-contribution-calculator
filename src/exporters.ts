import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import type { CalculatorSettings, ContributionReport } from './types';
import {
  getDeductedBaseAmount,
  getHalfContributionBaseAmount,
  NATIONAL_STANDARD_ANNUAL_COMPENSATION,
} from './domain';
import { formatMoney, formatPercent } from './formatters';

const reportTitle = '七護會 任意献金目安額 集計資料';
const reportNote =
  '本資料に記載された金額は、七護會への任意献金を検討するための目安であり、確定事項ではありません。';

const downloadBlob = (blob: Blob, fileName: string): void => {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
};

const escapeCsv = (value: string | number): string => {
  const text = String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};

const ymd = (date = new Date()): string => {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, '0');
  const day = `${date.getDate()}`.padStart(2, '0');
  return `${year}-${month}-${day}`;
};

export const exportCsv = (report: ContributionReport): void => {
  const rows: Array<Array<string | number>> = [
    [
      'No',
      '団員名',
      '点検',
      '訓練',
      '火災',
      '獲得P',
      '活動貢献率',
      '七護會 任意献金目安額',
    ],
    ...report.members.map((member) => [
      member.no,
      member.name,
      member.inspection,
      member.training,
      member.fire,
      member.earnedPoints,
      formatPercent(member.contributionRate),
      formatMoney(member.suggestionAmount),
    ]),
  ];
  const csv = `\uFEFF${rows.map((row) => row.map(escapeCsv).join(',')).join('\r\n')}`;
  downloadBlob(
    new Blob([csv], { type: 'text/csv;charset=utf-8' }),
    `nanagokai-contribution-${ymd()}.csv`,
  );
};

const paragraph = (text: string): Paragraph =>
  new Paragraph({
    children: [new TextRun(text)],
    spacing: { after: 120 },
  });

const tableCell = (text: string | number, bold = false): TableCell =>
  new TableCell({
    children: [
      new Paragraph({
        children: [new TextRun({ text: String(text), bold })],
      }),
    ],
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
  });

const buildSummaryTable = (report: ContributionReport): Table =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      ['対象期間', report.summary.periodLabel],
      ['読み込んだシート数', report.summary.sheetCount],
      ['点検回数', report.summary.inspectionEvents],
      ['訓練回数', report.summary.trainingEvents],
      ['火災出動件数', report.summary.fireEvents],
      ['基本活動ポイント', report.summary.basicActivityPoints],
      ['対象団員数', report.summary.memberCount],
      ['任意献金目安額の合計', formatMoney(report.summary.suggestionTotal)],
    ].map(
      ([label, value]) =>
        new TableRow({
          children: [tableCell(label, true), tableCell(value)],
        }),
    ),
  });

const buildAmountSettingsTable = (settings: CalculatorSettings): Table =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      [
        '国標準年額報酬',
        `${formatMoney(NATIONAL_STANDARD_ANNUAL_COMPENSATION)}（参考表示）`,
      ],
      ['調布市団員年額報酬', formatMoney(settings.cityAnnualCompensation)],
      ['七護會年会費', formatMoney(settings.membershipFee)],
      [
        '差引後基準額',
        `${formatMoney(getDeductedBaseAmount(settings))}（参考金額）`,
      ],
      ['献金算定基準額', formatMoney(settings.contributionBaseAmount)],
      ['半期算定基準額', formatMoney(getHalfContributionBaseAmount(settings))],
      [
        '端数処理',
        `${settings.roundingUnit.toLocaleString('ja-JP')}円単位で四捨五入`,
      ],
    ].map(
      ([label, value]) =>
        new TableRow({
          children: [tableCell(label, true), tableCell(value)],
        }),
    ),
  });

const buildMemberTable = (report: ContributionReport): Table =>
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          'No',
          '団員名',
          '点検',
          '訓練',
          '火災',
          '獲得P',
          '活動貢献率',
          '七護會 任意献金目安額',
        ].map((header) => tableCell(header, true)),
      }),
      ...report.members.map(
        (member) =>
          new TableRow({
            children: [
              tableCell(member.no),
              tableCell(member.name),
              tableCell(member.inspection),
              tableCell(member.training),
              tableCell(member.fire),
              tableCell(member.earnedPoints),
              tableCell(formatPercent(member.contributionRate)),
              tableCell(formatMoney(member.suggestionAmount)),
            ],
          }),
      ),
    ],
  });

export const exportDocx = async (
  report: ContributionReport,
  settings: CalculatorSettings,
): Promise<void> => {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: reportTitle,
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: { after: 240 },
          }),
          paragraph(`作成日：${new Date().toLocaleDateString('ja-JP')}`),
          paragraph(`対象期間：${report.summary.periodLabel}`),
          new Paragraph({
            text: '計算ルール',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 },
          }),
          paragraph(
            `点検：1回${settings.inspectionPoint}ポイント、訓練：1回${settings.trainingPoint}ポイント、火災出動：1回${settings.firePoint}ポイント。`,
          ),
          paragraph(
            `火災ポイントは半期ごとに上限${settings.fireCap}ポイントとし、年間合計では前期・後期それぞれに上限を適用します。`,
          ),
          paragraph(
            '七護會 任意献金目安額は、半期算定基準額に活動貢献率ごとの割合を掛け、選択した単位で四捨五入します。',
          ),
          paragraph(
            '活動貢献率80%以上：0%、60%以上80%未満：25%、40%以上60%未満：50%、20%以上40%未満：75%、20%未満：100%。',
          ),
          buildAmountSettingsTable(settings),
          new Paragraph({
            text: 'サマリー',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 },
          }),
          buildSummaryTable(report),
          new Paragraph({
            text: '団員別集計表',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 },
          }),
          buildMemberTable(report),
          new Paragraph({
            text: '注意書き',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 },
          }),
          paragraph(reportNote),
        ],
      },
    ],
  });
  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, `nanagokai-contribution-${ymd()}.docx`);
};
