import * as XLSX from 'xlsx';
import type { NumericValue, PortfolioItem, StockData } from '../types';

const parseNumericCell = (value: unknown, scale = 1): NumericValue => {
  if (value === null || value === undefined) return null;

  if (typeof value === 'string') {
    const trimmedValue = value.trim();
    if (trimmedValue === '' || trimmedValue.toLowerCase() === 'n/a') return null;
  }

  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue / scale : null;
};

const parseCurrency = (value: unknown): number => {
  if (typeof value === 'number') return value;
  if (typeof value !== 'string') return 0;
  const numericString = value.replace(/[$,]/g, '').trim();
  const num = Number(numericString);
  return Number.isFinite(num) ? num : 0;
};

const extractTicker = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  const str = String(value).trim();
  // Ticker is the first word (e.g., "AAPL      Apple Inc." -> "AAPL")
  return str.split(/\s+/)[0];
};

export const parseExcel = async (file: File): Promise<StockData[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = 'All CCC';
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          reject(new Error(`Sheet "${sheetName}" not found`));
          return;
        }

        const json = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1 });
        // Data starts at row 6 (index 6)
        const rows = json.slice(6);
        
        const stocks: StockData[] = rows.map((row): StockData | null => {
          if (!row[0] || !row[1]) return null;
          
          return {
            name: String(row[0] || ''),
            symbol: String(row[1] || '').trim(),
            sector: String(row[2] || ''),
            industry: String(row[3] || ''),
            years: parseNumericCell(row[4]),
            price: parseNumericCell(row[8]),
            yield: parseNumericCell(row[9]), // It's already in percentage (e.g., 0.88 for 0.88%)
            dividend: parseNumericCell(row[10]),
            payout: parseNumericCell(row[25]),
            growth1yr: parseNumericCell(row[18]),
            growth3yr: parseNumericCell(row[19]),
            growth5yr: parseNumericCell(row[20]),
            growth10yr: parseNumericCell(row[21]),
            marketCap: parseNumericCell(row[37], 1000), // Column AK ($Mil)
            peRatio: parseNumericCell(row[26]),
            pbRatio: parseNumericCell(row[31]),
            roe: parseNumericCell(row[32]),
            past5yGrowth: parseNumericCell(row[35]),
            est5yGrowth: parseNumericCell(row[36]),
            debtEquity: parseNumericCell(row[39]), // Column AN (Debt/Equity)
            chowderRule: parseNumericCell(row[41]), // Column AP (Rule)
            roa: parseNumericCell(row[58]), // Column BG (ROA)
          };
        }).filter((stock): stock is StockData => stock !== null);

        resolve(stocks);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

export const parsePortfolioCSV = async (file: File): Promise<PortfolioItem[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet);

        const portfolio: PortfolioItem[] = json.map((row) => {
          const tickerKey = Object.keys(row).find(k => k.toLowerCase().trim() === 'ticker');
          const incomeKey = Object.keys(row).find(k => k.toLowerCase().trim() === 'annualincome');
          
          return {
            ticker: extractTicker(tickerKey ? row[tickerKey] : ''),
            annualIncome: parseCurrency(incomeKey ? row[incomeKey] : 0),
          };
        }).filter(item => item.ticker !== '');

        resolve(portfolio);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};
