import { ChangeEvent, useMemo, useState } from 'react';
import {
  Download,
  FileSpreadsheet,
  FileText,
  Link as LinkIcon,
  Loader2,
  Printer,
  Settings,
  Upload,
} from 'lucide-react';
import {
  buildContributionReport,
  DEFAULT_SETTINGS,
  ERROR_MESSAGES,
  parseWorkbook,
  PERIOD_LABELS,
} from './domain';
import { exportCsv, exportDocx } from './exporters';
import { formatMoney, formatPercent } from './formatters';
import { fetchGoogleSheetWorkbook } from './googleSheets';
import type {
  CalculatorSettings,
  ContributionReport,
  PeriodSelection,
  WorkbookScan,
} from './types';

const periodOptions: PeriodSelection[] = [
  'first-half',
  'second-half',
  'annual',
];

const numericFields: Array<{
  key: keyof CalculatorSettings;
  label: string;
  suffix: string;
  step: number;
}> = [
  { key: 'inspectionPoint', label: '点検P', suffix: 'P', step: 1 },
  { key: 'trainingPoint', label: '訓練P', suffix: 'P', step: 1 },
  { key: 'firePoint', label: '火災P', suffix: 'P', step: 1 },
  { key: 'fireCap', label: '火災上限', suffix: 'P', step: 1 },
  { key: 'maxSuggestion', label: '最大目安額', suffix: '円', step: 1000 },
];

const SummaryItem = ({
  label,
  value,
}: {
  label: string;
  value: string | number;
}) => (
  <div className="summary-item">
    <span>{label}</span>
    <strong>{value}</strong>
  </div>
);

const EmptyState = () => (
  <section className="empty-state">
    <FileSpreadsheet aria-hidden="true" />
    <h2>ExcelまたはGoogleスプレッドシートURLを読み込むと集計結果が表示されます</h2>
    <p>
      月別シートの3行目にある「点検」「訓練」「火災」の列だけを読み取り、B列の団員名ごとに計算します。
    </p>
  </section>
);

const Results = ({ report }: { report: ContributionReport }) => (
  <div className="results-stack">
    <section className="panel summary-panel">
      <div className="section-heading">
        <h2>サマリー</h2>
        <p>読み込んだシート：{report.scannedSheetNames.join('、')}</p>
      </div>
      <div className="summary-grid">
        <SummaryItem label="対象期間" value={report.summary.periodLabel} />
        <SummaryItem
          label="読み込んだシート数"
          value={report.summary.sheetCount}
        />
        <SummaryItem label="点検回数" value={report.summary.inspectionEvents} />
        <SummaryItem label="訓練回数" value={report.summary.trainingEvents} />
        <SummaryItem label="火災出動件数" value={report.summary.fireEvents} />
        <SummaryItem
          label="基本活動ポイント"
          value={report.summary.basicActivityPoints}
        />
        <SummaryItem label="対象団員数" value={report.summary.memberCount} />
        <SummaryItem
          label="任意献金目安額の合計"
          value={formatMoney(report.summary.suggestionTotal)}
        />
      </div>
    </section>

    <section className="panel table-panel">
      <div className="section-heading">
        <h2>団員別集計表</h2>
        <p>活動貢献率は小数第1位、金額はカンマ付きで表示します。</p>
      </div>
      <div className="table-wrap">
        <table className="result-table">
          <thead>
            <tr>
              <th>No</th>
              <th>団員名</th>
              <th>点検</th>
              <th>訓練</th>
              <th>火災</th>
              <th>獲得P</th>
              <th>活動貢献率</th>
              <th>七護會 任意献金目安額</th>
            </tr>
          </thead>
          <tbody>
            {report.members.map((member) => (
              <tr key={`${member.no}-${member.name}`}>
                <td>{member.no || '-'}</td>
                <td className="member-name">{member.name}</td>
                <td>{member.inspection}</td>
                <td>{member.training}</td>
                <td>{member.fire}</td>
                <td>{member.earnedPoints}</td>
                <td>{formatPercent(member.contributionRate)}</td>
                <td>{formatMoney(member.suggestionAmount)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  </div>
);

function App() {
  const [scan, setScan] = useState<WorkbookScan | null>(null);
  const [fileName, setFileName] = useState('');
  const [sheetUrl, setSheetUrl] = useState('');
  const [period, setPeriod] = useState<PeriodSelection>('first-half');
  const [settings, setSettings] =
    useState<CalculatorSettings>(DEFAULT_SETTINGS);
  const [readError, setReadError] = useState('');
  const [isExportingWord, setIsExportingWord] = useState(false);
  const [isImportingSheet, setIsImportingSheet] = useState(false);

  const reportResult = useMemo(
    () =>
      scan ? buildContributionReport(scan, period, settings) : undefined,
    [period, scan, settings],
  );
  const report = reportResult?.ok ? reportResult.report : undefined;
  const activeError =
    readError || (reportResult && !reportResult.ok ? reportResult.error.message : '');

  const updateNumber = (key: keyof CalculatorSettings, rawValue: string) => {
    const value = Number(rawValue);
    setSettings((current) => ({
      ...current,
      [key]: Number.isFinite(value) && value >= 0 ? value : 0,
    }));
  };

  const loadWorkbook = (buffer: ArrayBuffer, sourceName: string) => {
    setScan(parseWorkbook(buffer));
    setFileName(sourceName);
  };

  const handleFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    setReadError('');

    try {
      const buffer = await file.arrayBuffer();
      loadWorkbook(buffer, file.name);
    } catch (error) {
      console.error(error);
      setScan(null);
      setReadError(ERROR_MESSAGES['excel-read-failed']);
    }
  };

  const handleGoogleSheetImport = async () => {
    const value = sheetUrl.trim();

    if (!value) {
      setReadError(ERROR_MESSAGES['google-sheets-url-invalid']);
      return;
    }

    setReadError('');
    setIsImportingSheet(true);

    try {
      const workbook = await fetchGoogleSheetWorkbook(value);
      loadWorkbook(workbook.buffer, workbook.label);
    } catch (error) {
      console.error(error);
      setScan(null);
      setReadError(
        error instanceof Error && error.message === 'google-sheets-url-invalid'
          ? ERROR_MESSAGES['google-sheets-url-invalid']
          : ERROR_MESSAGES['google-sheets-read-failed'],
      );
    } finally {
      setIsImportingSheet(false);
    }
  };

  const handleWordExport = async () => {
    if (!report) {
      return;
    }

    setIsExportingWord(true);
    try {
      await exportDocx(report, settings);
    } finally {
      setIsExportingWord(false);
    }
  };

  return (
    <main className="app-shell">
      <header className="hero">
        <div>
          <h1>
            七護會 任意献金目安額 <span>計算ツール</span>
          </h1>
          <p>
            このツールは消防団出欠管理Excelを読み込み、点検・訓練・火災出動の回数から、七護會
            任意献金目安額を算出する計算ツールです。アップロードされたファイルは保存されません。
          </p>
        </div>
      </header>

      <section className="panel controls-panel no-print" aria-label="読み込みと集計期間">
        <div className="import-stack">
          <div className="upload-box">
            <label htmlFor="attendance-file" className="upload-label">
              <Upload aria-hidden="true" />
              <span>Excelアップロード</span>
            </label>
            <input
              id="attendance-file"
              type="file"
              accept=".xlsx,.xls,.xlsm"
              onChange={handleFileChange}
            />
            <p>{fileName ? `読み込み中の資料：${fileName}` : 'xlsx / xls / xlsm に対応'}</p>
          </div>

          <div className="url-import-box">
            <label htmlFor="google-sheet-url">GoogleスプレッドシートURL</label>
            <div className="url-import-row">
              <input
                id="google-sheet-url"
                type="url"
                value={sheetUrl}
                placeholder="https://docs.google.com/spreadsheets/d/..."
                onChange={(event) => setSheetUrl(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    void handleGoogleSheetImport();
                  }
                }}
              />
              <button
                type="button"
                onClick={handleGoogleSheetImport}
                disabled={isImportingSheet}
              >
                {isImportingSheet ? (
                  <Loader2 className="spin" aria-hidden="true" />
                ) : (
                  <LinkIcon aria-hidden="true" />
                )}
                {isImportingSheet ? '取り込み中' : 'URLから取り込み'}
              </button>
            </div>
            <p>
              共有済みURLまたはウェブ公開URLに対応します。非公開シートはExcel形式でダウンロードしてアップロードしてください。
            </p>
          </div>
        </div>

        <div className="period-control" role="group" aria-label="集計期間選択">
          {periodOptions.map((option) => (
            <button
              key={option}
              type="button"
              className={period === option ? 'active' : ''}
              onClick={() => setPeriod(option)}
            >
              {PERIOD_LABELS[option].replace(/（.*）/, '')}
            </button>
          ))}
        </div>
      </section>

      <section className="panel settings-panel no-print" aria-label="計算設定">
        <div className="section-heading compact">
          <div className="heading-with-icon">
            <Settings aria-hidden="true" />
            <h2>計算設定</h2>
          </div>
          <p>必要に応じてポイントや最大目安額を変更できます。</p>
        </div>
        <div className="settings-grid">
          {numericFields.map((field) => (
            <label key={field.key} className="setting-field">
              <span>{field.label}</span>
              <div>
                <input
                  type="number"
                  min="0"
                  step={field.step}
                  value={settings[field.key]}
                  onChange={(event) =>
                    updateNumber(field.key, event.target.value)
                  }
                />
                <em>{field.suffix}</em>
              </div>
            </label>
          ))}
        </div>
      </section>

      <section className="toolbar no-print" aria-label="出力">
        <button type="button" disabled={!report} onClick={() => report && exportCsv(report)}>
          <Download aria-hidden="true" />
          CSVダウンロード
        </button>
        <button type="button" disabled={!report || isExportingWord} onClick={handleWordExport}>
          <FileText aria-hidden="true" />
          {isExportingWord ? 'Word作成中' : 'Wordダウンロード'}
        </button>
        <button type="button" disabled={!report} onClick={() => window.print()}>
          <Printer aria-hidden="true" />
          印刷用表示
        </button>
      </section>

      {activeError ? (
        <section className="alert" role="alert">
          {activeError}
        </section>
      ) : null}

      <div className="print-main">{report ? <Results report={report} /> : <EmptyState />}</div>
    </main>
  );
}

export default App;
