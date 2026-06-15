export interface GoogleSheetWorkbook {
  buffer: ArrayBuffer;
  label: string;
}

interface GoogleSheetSource {
  downloadUrl: string;
  label: string;
}

const GOOGLE_SHEETS_HOST = 'docs.google.com';
const GOOGLE_SHEET_ID_PATTERN = /^[a-zA-Z0-9-_]{20,}$/;

const parseContentDispositionFilename = (header: string | null): string => {
  if (!header) {
    return '';
  }

  const encodedMatch = header.match(/filename\*=UTF-8''([^;]+)/i);

  if (encodedMatch?.[1]) {
    try {
      return decodeURIComponent(encodedMatch[1].replace(/"/g, '')).trim();
    } catch {
      return encodedMatch[1].replace(/"/g, '').trim();
    }
  }

  const plainMatch = header.match(/filename="?([^";]+)"?/i);
  return plainMatch?.[1]?.trim() ?? '';
};

export const parseGoogleSheetUrl = (rawValue: string): GoogleSheetSource | null => {
  const value = rawValue.trim();

  if (!value) {
    return null;
  }

  if (GOOGLE_SHEET_ID_PATTERN.test(value)) {
    return {
      downloadUrl: `https://${GOOGLE_SHEETS_HOST}/spreadsheets/d/${value}/export?format=xlsx`,
      label: 'Googleスプレッドシート',
    };
  }

  let url: URL;

  try {
    url = new URL(value);
  } catch {
    return null;
  }

  if (url.hostname !== GOOGLE_SHEETS_HOST) {
    return null;
  }

  const publishedMatch = url.pathname.match(/\/spreadsheets\/d\/e\/([^/]+)/);

  if (publishedMatch?.[1]) {
    return {
      downloadUrl: `https://${GOOGLE_SHEETS_HOST}/spreadsheets/d/e/${publishedMatch[1]}/pub?output=xlsx`,
      label: '公開Googleスプレッドシート',
    };
  }

  const sheetMatch = url.pathname.match(/\/spreadsheets\/d\/([^/]+)/);

  if (!sheetMatch?.[1]) {
    return null;
  }

  return {
    downloadUrl: `https://${GOOGLE_SHEETS_HOST}/spreadsheets/d/${sheetMatch[1]}/export?format=xlsx`,
    label: 'Googleスプレッドシート',
  };
};

export const fetchGoogleSheetWorkbook = async (
  sheetUrl: string,
): Promise<GoogleSheetWorkbook> => {
  const source = parseGoogleSheetUrl(sheetUrl);

  if (!source) {
    throw new Error('google-sheets-url-invalid');
  }

  const response = await fetch(source.downloadUrl, {
    method: 'GET',
    redirect: 'follow',
  });

  if (!response.ok) {
    throw new Error('google-sheets-read-failed');
  }

  const contentType = response.headers.get('content-type')?.toLowerCase() ?? '';

  if (contentType.includes('text/html')) {
    throw new Error('google-sheets-read-failed');
  }

  const buffer = await response.arrayBuffer();

  if (buffer.byteLength === 0) {
    throw new Error('google-sheets-read-failed');
  }

  const filename = parseContentDispositionFilename(
    response.headers.get('content-disposition'),
  );

  return {
    buffer,
    label: filename || source.label,
  };
};
