export type NumericValue = number | null;

export interface PortfolioItem {
  ticker: string;
  annualIncome: number;
}

export interface StockData {
  name: string;
  symbol: string;
  sector: string;
  industry: string;
  years: NumericValue;
  price: NumericValue;
  yield: NumericValue;
  dividend: NumericValue;
  payout: NumericValue;
  growth1yr: NumericValue;
  growth3yr: NumericValue;
  growth5yr: NumericValue;
  growth10yr: NumericValue;
  marketCap: NumericValue;
  peRatio: NumericValue;
  pbRatio: NumericValue;
  roe: NumericValue;
  past5yGrowth: NumericValue;
  est5yGrowth: NumericValue;
  debtEquity: NumericValue;
  chowderRule: NumericValue;
  roa: NumericValue;
  portfolioAnnualIncome?: number;
  [key: string]: string | NumericValue | number | undefined;
}

export interface FilterConfig {
  metric: keyof StockData;
  label: string;
  active: boolean;
  cutoff: number;
  operator: 'gte' | 'lte';
}
