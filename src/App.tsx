import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ChangeEvent } from 'react';
import {
  Alert,
  AppBar,
  Box,
  Button,
  Card,
  CardContent,
  Chip,
  CircularProgress,
  Container,
  CssBaseline,
  Divider,
  IconButton,
  InputAdornment,
  Paper,
  Snackbar,
  Stack,
  Switch,
  TextField,
  ThemeProvider,
  Tooltip,
  Toolbar,
  Typography,
  createTheme,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';
import ErrorOutlineIcon from '@mui/icons-material/ErrorOutline';
import FilterListIcon from '@mui/icons-material/FilterList';
import InsightsIcon from '@mui/icons-material/Insights';
import UploadFileIcon from '@mui/icons-material/UploadFile';
import AccountBalanceWalletIcon from '@mui/icons-material/AccountBalanceWallet';
import VisibilityOffIcon from '@mui/icons-material/VisibilityOff';
import VisibilityIcon from '@mui/icons-material/Visibility';
import {
  DataGrid,
  GridToolbar,
  type GridColDef,
  type GridRenderCellParams,
  type GridSortModel,
  type GridRowClassNameParams,
} from '@mui/x-data-grid';
import type { FilterConfig, NumericValue, StockData, PortfolioItem } from './types';
import { parseExcel, parsePortfolioCSV } from './utils/parser';

const theme = createTheme({
  palette: {
    mode: 'light',
    background: {
      default: '#f6f8fb',
      paper: '#ffffff',
    },
    primary: {
      main: '#2563eb',
      dark: '#1d4ed8',
    },
    success: {
      main: '#16803c',
    },
    warning: {
      main: '#b45309',
    },
    text: {
      primary: '#1e293b',
      secondary: '#64748b',
    },
  },
  shape: {
    borderRadius: 8,
  },
  typography: {
    fontFamily:
      'Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    button: {
      textTransform: 'none',
      fontWeight: 700,
    },
  },
  components: {
    MuiCard: {
      styleOverrides: {
        root: {
          border: '1px solid #d9e2ec',
          boxShadow: '0 10px 30px rgba(15, 23, 42, 0.06)',
        },
      },
    },
  },
});

const DEFAULT_FILTERS: FilterConfig[] = [
  { metric: 'marketCap', label: 'Mkt Cap ($B)', active: true, cutoff: 2, operator: 'gte' },
  { metric: 'years', label: 'Record Years', active: true, cutoff: 20, operator: 'gte' },
  { metric: 'yield', label: 'Div Yield (%)', active: true, cutoff: 2, operator: 'gte' },
  { metric: 'growth1yr', label: '1Y DGR (%)', active: true, cutoff: 5, operator: 'gte' },
  { metric: 'growth3yr', label: '3Y DGR (%)', active: true, cutoff: 5, operator: 'gte' },
  { metric: 'growth5yr', label: '5Y DGR (%)', active: true, cutoff: 5, operator: 'gte' },
  { metric: 'growth10yr', label: '10Y DGR (%)', active: true, cutoff: 5, operator: 'gte' },
  { metric: 'payout', label: 'Payout ratio', active: true, cutoff: 60, operator: 'lte' },
  { metric: 'peRatio', label: 'P/E Ratio', active: true, cutoff: 20, operator: 'lte' },
  { metric: 'pbRatio', label: 'Price/Book', active: false, cutoff: 2, operator: 'lte' },
  { metric: 'roe', label: 'ROE (%)', active: false, cutoff: 10, operator: 'gte' },
  { metric: 'past5yGrowth', label: 'Past 5Y (%)', active: true, cutoff: 0, operator: 'gte' },
  { metric: 'est5yGrowth', label: 'Est 5Y (%)', active: true, cutoff: 0, operator: 'gte' },
  { metric: 'debtEquity', label: 'Debt/Equity', active: true, cutoff: 0.7, operator: 'lte' },
  { metric: 'chowderRule', label: 'Chowder Rule', active: false, cutoff: 11, operator: 'gte' },
];

const FILTERS_STORAGE_KEY = 'screener_filters';
const MISSING_VALUES_MATCH_STORAGE_KEY = 'screener_missing_values_match';
const STOCKS_STORAGE_KEY = 'screener_stocks';
const PORTFOLIO_STORAGE_KEY = 'screener_portfolio';
const SORT_MODEL_STORAGE_KEY = 'screener_sort_model';
const HIDDEN_SYMBOLS_STORAGE_KEY = 'screener_hidden_symbols';
const SHOW_HIDDEN_STORAGE_KEY = 'screener_show_hidden';
const DEFAULT_SORT_MODEL: GridSortModel = [{ field: 'name', sort: 'asc' }];

const loadSavedFilters = () => {
  try {
    const saved = localStorage.getItem(FILTERS_STORAGE_KEY);
    if (!saved) return DEFAULT_FILTERS;

    const parsed = JSON.parse(saved) as FilterConfig[];
    if (!Array.isArray(parsed)) return DEFAULT_FILTERS;

    return DEFAULT_FILTERS.map((defaultFilter) => {
      const savedFilter = parsed.find((filter) => filter.metric === defaultFilter.metric);
      return savedFilter ? { ...defaultFilter, ...savedFilter } : defaultFilter;
    });
  } catch {
    return DEFAULT_FILTERS;
  }
};

const loadSavedStocks = () => {
  try {
    const saved = localStorage.getItem(STOCKS_STORAGE_KEY);
    if (!saved) return [];

    const parsed = JSON.parse(saved) as StockData[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const loadSavedPortfolio = () => {
  try {
    const saved = localStorage.getItem(PORTFOLIO_STORAGE_KEY);
    if (!saved) return [];

    const parsed = JSON.parse(saved) as PortfolioItem[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const loadSavedSortModel = () => {
  try {
    const saved = localStorage.getItem(SORT_MODEL_STORAGE_KEY);
    if (!saved) return DEFAULT_SORT_MODEL;

    const parsed = JSON.parse(saved) as GridSortModel;
    return Array.isArray(parsed) ? parsed : DEFAULT_SORT_MODEL;
  } catch {
    return DEFAULT_SORT_MODEL;
  }
};

const loadSavedMissingValuesMatch = () => {
  try {
    return localStorage.getItem(MISSING_VALUES_MATCH_STORAGE_KEY) === 'true';
  } catch {
    return false;
  }
};

const loadSavedHiddenSymbols = (): Set<string> => {
  try {
    const saved = localStorage.getItem(HIDDEN_SYMBOLS_STORAGE_KEY);
    if (!saved) return new Set();
    const parsed = JSON.parse(saved);
    return Array.isArray(parsed) ? new Set(parsed) : new Set();
  } catch {
    return new Set();
  }
};

const loadSavedShowHidden = () => {
  try {
    return localStorage.getItem(SHOW_HIDDEN_STORAGE_KEY) === 'true';
  } catch {
    return false;
  }
};

const isValidNumber = (value: unknown): value is number =>
  typeof value === 'number' && Number.isFinite(value);
const formatMissingValue = (value: NumericValue | undefined, formatter: (numericValue: number) => string) =>
  isValidNumber(value) ? formatter(value) : 'N/A';
const formatPercent = (value: NumericValue | undefined) =>
  formatMissingValue(value, (numericValue) => `${numericValue.toFixed(1)}%`);
const formatRatio = (value: NumericValue | undefined) =>
  formatMissingValue(value, (numericValue) => numericValue.toFixed(1));
const formatDecimal = (value: NumericValue | undefined) =>
  formatMissingValue(value, (numericValue) => numericValue.toFixed(2));
const formatCurrency = (value: number | undefined) =>
  value !== undefined ? `$${value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '$0.00';
const getCutoffInputs = (filters: FilterConfig[]) =>
  Object.fromEntries(filters.map((filter) => [filter.metric, String(filter.cutoff)]));
const escapeSheetCell = (value: string) =>
  /[\t\r\n"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
const getRowSheetValues = (row: StockData) => [
  row.name,
  row.symbol,
  formatCurrency(row.portfolioAnnualIncome),
  row.sector,
  formatMissingValue(row.marketCap, (numericValue) => `${numericValue.toFixed(1)}B`),
  formatMissingValue(row.yield, (numericValue) => `${numericValue.toFixed(2)}%`),
  formatMissingValue(row.years, (numericValue) => String(numericValue)),
  formatPercent(row.payout),
  formatPercent(row.growth1yr),
  formatPercent(row.growth3yr),
  formatPercent(row.growth5yr),
  formatPercent(row.growth10yr),
  formatRatio(row.peRatio),
  formatRatio(row.pbRatio),
  formatPercent(row.roe),
  formatPercent(row.past5yGrowth),
  formatPercent(row.est5yGrowth),
  formatDecimal(row.debtEquity),
  formatRatio(row.chowderRule),
];
const formatRowForSheets = (row: StockData) => getRowSheetValues(row).map(escapeSheetCell).join('\t');

export default function App() {
  const [stocks, setStocks] = useState<StockData[]>(loadSavedStocks);
  const [portfolio, setPortfolio] = useState<PortfolioItem[]>(loadSavedPortfolio);
  const [filters, setFilters] = useState<FilterConfig[]>(loadSavedFilters);
  const [cutoffInputs, setCutoffInputs] = useState<Record<string, string>>(() =>
    getCutoffInputs(loadSavedFilters()),
  );
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [sortModel, setSortModel] = useState<GridSortModel>(loadSavedSortModel);
  const [missingValuesMatch, setMissingValuesMatch] = useState(loadSavedMissingValuesMatch);
  const [hiddenSymbols, setHiddenSymbols] = useState<Set<string>>(loadSavedHiddenSymbols);
  const [showHidden, setShowHidden] = useState(loadSavedShowHidden);
  const [copyMessage, setCopyMessage] = useState<string | null>(null);
  const [copyError, setCopyError] = useState<string | null>(null);

  useEffect(() => {
    localStorage.setItem(FILTERS_STORAGE_KEY, JSON.stringify(filters));
  }, [filters]);

  useEffect(() => {
    localStorage.setItem(SORT_MODEL_STORAGE_KEY, JSON.stringify(sortModel));
  }, [sortModel]);

  useEffect(() => {
    localStorage.setItem(MISSING_VALUES_MATCH_STORAGE_KEY, String(missingValuesMatch));
  }, [missingValuesMatch]);

  useEffect(() => {
    localStorage.setItem(HIDDEN_SYMBOLS_STORAGE_KEY, JSON.stringify(Array.from(hiddenSymbols)));
  }, [hiddenSymbols]);

  useEffect(() => {
    localStorage.setItem(SHOW_HIDDEN_STORAGE_KEY, String(showHidden));
  }, [showHidden]);

  useEffect(() => {
    if (stocks.length === 0) {
      localStorage.removeItem(STOCKS_STORAGE_KEY);
      return;
    }

    localStorage.setItem(STOCKS_STORAGE_KEY, JSON.stringify(stocks));
  }, [stocks]);

  useEffect(() => {
    if (portfolio.length === 0) {
      localStorage.removeItem(PORTFOLIO_STORAGE_KEY);
      return;
    }

    localStorage.setItem(PORTFOLIO_STORAGE_KEY, JSON.stringify(portfolio));
  }, [portfolio]);

  const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setStocks([]);
    setFilters(DEFAULT_FILTERS);
    setCutoffInputs(getCutoffInputs(DEFAULT_FILTERS));

    try {
      const data = await parseExcel(file);
      setStocks(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to parse file');
    } finally {
      setLoading(false);
      event.target.value = '';
    }
  };

  const handlePortfolioUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);

    try {
      const data = await parsePortfolioCSV(file);
      setPortfolio(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to parse portfolio file');
    } finally {
      setLoading(false);
      event.target.value = '';
    }
  };

  const handleStartNewFile = () => {
    localStorage.removeItem(STOCKS_STORAGE_KEY);
    localStorage.removeItem(FILTERS_STORAGE_KEY);
    localStorage.removeItem(SORT_MODEL_STORAGE_KEY);
    localStorage.removeItem(MISSING_VALUES_MATCH_STORAGE_KEY);
    localStorage.removeItem(PORTFOLIO_STORAGE_KEY);
    localStorage.removeItem(HIDDEN_SYMBOLS_STORAGE_KEY);
    localStorage.removeItem(SHOW_HIDDEN_STORAGE_KEY);
    setStocks([]);
    setPortfolio([]);
    setError(null);
    setFilters(DEFAULT_FILTERS);
    setCutoffInputs(getCutoffInputs(DEFAULT_FILTERS));
    setSortModel(DEFAULT_SORT_MODEL);
    setMissingValuesMatch(false);
    setHiddenSymbols(new Set());
    setShowHidden(false);
  };

  const joinedStocks = useMemo(() => {
    if (portfolio.length === 0) return stocks;

    const portfolioMap = new Map(
      portfolio.map((p) => [p.ticker.toUpperCase().trim(), p.annualIncome]),
    );

    return stocks.map((stock) => {
      const symbol = stock.symbol.toUpperCase().trim();
      return {
        ...stock,
        portfolioAnnualIncome: portfolioMap.get(symbol) || 0,
      };
    });
  }, [stocks, portfolio]);

  const filteredStocks = useMemo(() => {
    return joinedStocks.filter((stock) => {
      return filters.every((filter) => {
        if (!filter.active) return true;
        const value = stock[filter.metric];
        if (!isValidNumber(value)) return missingValuesMatch;
        return filter.operator === 'gte' ? value >= filter.cutoff : value <= filter.cutoff;
      });
    });
  }, [joinedStocks, filters, missingValuesMatch]);

  const displayedStocks = useMemo(() => {
    return filteredStocks.filter((stock) => showHidden || !hiddenSymbols.has(stock.symbol));
  }, [filteredStocks, hiddenSymbols, showHidden]);

  const toggleHide = useCallback((symbol: string) => {
    setHiddenSymbols((current) => {
      const next = new Set(current);
      if (next.has(symbol)) {
        next.delete(symbol);
      } else {
        next.add(symbol);
      }
      return next;
    });
  }, []);

  const toggleFilter = (index: number) => {
    setFilters((currentFilters) =>
      currentFilters.map((filter, currentIndex) =>
        currentIndex === index ? { ...filter, active: !filter.active } : filter,
      ),
    );
  };

  const updateCutoff = (index: number, rawValue: string) => {
    const metric = filters[index].metric;
    setCutoffInputs((currentInputs) => ({ ...currentInputs, [metric]: rawValue }));

    if (rawValue.trim() === '') return;

    const value = Number(rawValue);
    if (!Number.isFinite(value)) return;

    setFilters((currentFilters) =>
      currentFilters.map((filter, currentIndex) =>
        currentIndex === index ? { ...filter, cutoff: value } : filter,
      ),
    );
  };

  const resetCutoffInput = (index: number) => {
    const filter = filters[index];
    setCutoffInputs((currentInputs) => ({
      ...currentInputs,
      [filter.metric]: String(filter.cutoff),
    }));
  };

  const handleCopyRow = useCallback(async (row: StockData) => {
    try {
      await navigator.clipboard.writeText(formatRowForSheets(row));
      setCopyError(null);
      setCopyMessage(`Copied ${row.symbol} row`);
    } catch {
      setCopyMessage(null);
      setCopyError('Could not copy row. Please allow clipboard access and try again.');
    }
  }, []);

  const columns = useMemo<GridColDef<StockData>[]>(
    () => [
      {
        field: 'actions',
        headerName: '',
        width: 100,
        sortable: false,
        filterable: false,
        disableColumnMenu: true,
        align: 'center',
        renderCell: (params: GridRenderCellParams<StockData>) => {
          const isHidden = hiddenSymbols.has(params.row.symbol);
          return (
            <Stack direction="row" spacing={0.5}>
              <Tooltip title="Copy row for Google Sheets">
                <IconButton
                  aria-label={`Copy ${params.row.symbol} row for Google Sheets`}
                  size="small"
                  onClick={(event) => {
                    event.stopPropagation();
                    void handleCopyRow(params.row);
                  }}
                >
                  <ContentCopyIcon fontSize="inherit" />
                </IconButton>
              </Tooltip>
              <Tooltip title={isHidden ? 'Un-hide stock' : 'Hide stock'}>
                <IconButton
                  aria-label={`${isHidden ? 'Un-hide' : 'Hide'} ${params.row.symbol}`}
                  size="small"
                  onClick={(event) => {
                    event.stopPropagation();
                    toggleHide(params.row.symbol);
                  }}
                  color={isHidden ? 'warning' : 'default'}
                >
                  {isHidden ? (
                    <VisibilityIcon fontSize="inherit" />
                  ) : (
                    <VisibilityOffIcon fontSize="inherit" />
                  )}
                </IconButton>
              </Tooltip>
            </Stack>
          );
        },
      },
      {
        field: 'name',
        headerName: 'Company',
        minWidth: 220,
        flex: 1.2,
        renderCell: (params: GridRenderCellParams<StockData, string>) => (
          <Typography variant="body2" fontWeight={800} color="text.primary" noWrap>
            {params.value}
          </Typography>
        ),
      },
      {
        field: 'symbol',
        headerName: 'Ticker',
        width: 110,
        renderCell: (params: GridRenderCellParams<StockData, string>) => (
          <Chip label={params.value} size="small" variant="outlined" />
        ),
      },
      {
        field: 'portfolioAnnualIncome',
        headerName: 'Annual Income',
        width: 140,
        type: 'number',
        valueFormatter: (value: number | undefined) => formatCurrency(value),
        renderCell: (params: GridRenderCellParams<StockData, number>) => (
          <Typography
            variant="body2"
            fontWeight={800}
            color={params.value && params.value > 0 ? 'success.main' : 'text.secondary'}
          >
            {formatCurrency(params.value)}
          </Typography>
        ),
      },
      {
        field: 'sector',
        headerName: 'Sector',
        minWidth: 180,
        flex: 0.9,
        renderCell: (params: GridRenderCellParams<StockData, string>) => (
          <Typography variant="body2" color="text.secondary" fontWeight={700} noWrap>
            {params.value}
          </Typography>
        ),
      },
      {
        field: 'marketCap',
        headerName: 'Mkt Cap ($B)',
        width: 130,
        type: 'number',
        valueFormatter: (value: number | null | undefined) =>
            formatMissingValue(value, (numericValue) => `${numericValue.toFixed(1)}B`),
      },

            {
        field: 'years',
        headerName: 'Yrs',
        width: 90,
        type: 'number',
        valueFormatter: (value: number | null | undefined) =>
          formatMissingValue(value, (numericValue) => String(numericValue)),
      },


      {
        field: 'yield',
        headerName: 'Yield',
        width: 110,
        type: 'number',
        valueFormatter: (value: number | null | undefined) =>
            formatMissingValue(value, (numericValue) => `${numericValue.toFixed(2)}%`),
      },
      {
        field: 'payout',
        headerName: 'Payout',
        width: 110,
        type: 'number',
        valueFormatter: formatPercent,
      },
      { field: 'growth1yr', headerName: '1Y DGR', width: 110, type: 'number', valueFormatter: formatPercent },
      { field: 'growth3yr', headerName: '3Y DGR', width: 110, type: 'number', valueFormatter: formatPercent },
      {
        field: 'growth5yr',
        headerName: '5Y DGR',
        width: 110,
        type: 'number',
        valueFormatter: formatPercent,
      },
      { field: 'growth10yr', headerName: '10Y DGR', width: 120, type: 'number', valueFormatter: formatPercent },
      { field: 'peRatio', headerName: 'Price/Earning', width: 90, type: 'number', valueFormatter: formatRatio },
      { field: 'pbRatio', headerName: 'P/B', width: 90, type: 'number', valueFormatter: formatRatio },
      { field: 'roe', headerName: 'ROE', width: 100, type: 'number', valueFormatter: formatPercent },
      { field: 'past5yGrowth', headerName: 'Past 5Y', width: 110, type: 'number', valueFormatter: formatPercent },
      { field: 'est5yGrowth', headerName: 'Est. 5Y', width: 110, type: 'number', valueFormatter: formatPercent },
      { field: 'debtEquity', headerName: 'Debt/Equity', width: 110, type: 'number', valueFormatter: formatDecimal },
      {
        field: 'chowderRule',
        headerName: 'Chowder',
        width: 120,
        type: 'number',
        valueFormatter: formatRatio,
      },
    ],
    [handleCopyRow, hiddenSymbols, toggleHide],
  );

  const uploadControl = (label: string, variant: 'text' | 'contained' = 'contained') => (
    <Button
      component="label"
      variant={variant}
      startIcon={variant === 'contained' ? <UploadFileIcon /> : <CloudUploadIcon />}
      size={variant === 'contained' ? 'large' : 'medium'}
      fullWidth={variant === 'contained'}
    >
      {label}
      <input type="file" accept=".xlsx,.xlsm" onChange={handleFileUpload} hidden />
    </Button>
  );

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Box sx={{ minHeight: '100vh', bgcolor: 'background.default', display: 'flex', flexDirection: 'column' }}>
        <AppBar
          position="sticky"
          color="inherit"
          elevation={0}
          sx={{ borderBottom: '1px solid #d9e2ec', bgcolor: 'background.paper' }}
        >
          <Toolbar sx={{ gap: 2 }}>
            <Stack direction="row" spacing={1.25} alignItems="center" sx={{ minWidth: 0 }}>
              <InsightsIcon color="primary" />
              <Typography variant="h6" fontWeight={900} noWrap>
                Dividend Screener
              </Typography>
            </Stack>
            {stocks.length > 0 && (
              <Chip
                color="primary"
                variant="outlined"
                size="small"
                label={`${filteredStocks.length} stocks matched`}
              />
            )}
            <Box sx={{ flexGrow: 1 }} />
            <Stack direction="row" spacing={1} alignItems="center">
              {stocks.length > 0 && (
                <>
                  <Button
                    component="label"
                    variant="text"
                    startIcon={<AccountBalanceWalletIcon />}
                  >
                    {portfolio.length > 0 ? 'Update Portfolio' : 'Upload Portfolio'}
                    <input type="file" accept=".csv" onChange={handlePortfolioUpload} hidden />
                  </Button>
                  <Button variant="text" startIcon={<CloudUploadIcon />} onClick={handleStartNewFile}>
                    Upload a new file
                  </Button>
                </>
              )}
            </Stack>
          </Toolbar>
        </AppBar>

        <Container maxWidth={false} sx={{ maxWidth: 1600, py: 3, flex: 1 }}>
          {loading && (
            <Stack alignItems="center" justifyContent="center" spacing={2} sx={{ pt: 10 }}>
              <CircularProgress size={56} thickness={4} />
              <Typography variant="h6" fontWeight={800}>
                Processing File
              </Typography>
              <Typography color="text.secondary">Reading data...</Typography>
            </Stack>
          )}

          {!loading && stocks.length === 0 && (
            <Stack alignItems="center" sx={{ pt: 8 }}>
              <Card sx={{ maxWidth: 500, width: '100%', textAlign: 'center' }}>
                <CardContent sx={{ p: { xs: 3, sm: 6 } }}>
                  <CloudUploadIcon sx={{ fontSize: 64, color: 'grey.400', mb: 3 }} />
                  <Typography variant="h4" fontWeight={900} gutterBottom>
                    Start Your Analysis
                  </Typography>
                  <Typography color="text.secondary" sx={{ mb: 5 }}>
                    Upload the Dividend Champions Excel file to begin filtering. All processing
                    happens locally in your browser for maximum privacy.
                  </Typography>
                  {uploadControl('Select Excel File')}
                </CardContent>
              </Card>
            </Stack>
          )}

          {error && (
            <Alert severity="error" icon={<ErrorOutlineIcon />} sx={{ mb: 3 }}>
              <Typography fontWeight={800}>Error</Typography>
              {error}
            </Alert>
          )}

          {stocks.length > 0 && !loading && (
            <Stack direction={{ xs: 'column', lg: 'row' }} spacing={3} alignItems="flex-start">
              <Card
                sx={{
                  width: { xs: '100%', lg: 320 },
                  flexShrink: 0,
                  position: { lg: 'sticky' },
                  top: { lg: 88 },
                  maxHeight: { lg: 'calc(100vh - 112px)' },
                  overflowY: 'auto',
                }}
              >
                <Stack
                  direction="row"
                  alignItems="center"
                  spacing={1}
                  sx={{ p: 2.5, position: 'sticky', top: 0, bgcolor: 'background.paper', zIndex: 1 }}
                >
                  <FilterListIcon color="primary" />
                  <Typography variant="h6" fontWeight={900}>
                    Filters
                  </Typography>
                </Stack>
                <Divider />
                <Stack spacing={2.5} sx={{ p: 2.5 }}>
                  <Box>
                    <Stack direction="row" alignItems="center" justifyContent="space-between" spacing={2}>
                      <Box sx={{ minWidth: 0 }}>
                        <Typography variant="caption" fontWeight={900} color="text.secondary">
                          WHEN A FILTERED VALUE IS N/A
                        </Typography>
                        <Typography variant="body2" color="text.secondary">
                          {missingValuesMatch ? 'Include the stock in results' : 'Exclude the stock from results'}
                        </Typography>
                      </Box>
                      <Switch
                        checked={missingValuesMatch}
                        onChange={(event) => setMissingValuesMatch(event.target.checked)}
                        inputProps={{ 'aria-label': 'Treat missing values as matching filters' }}
                      />
                    </Stack>
                  </Box>
                  <Divider />
                  <Box>
                    <Stack direction="row" alignItems="center" justifyContent="space-between" spacing={2}>
                      <Box sx={{ minWidth: 0 }}>
                        <Typography variant="caption" fontWeight={900} color="text.secondary">
                          HIDDEN STOCKS
                        </Typography>
                        <Typography variant="body2" color="text.secondary">
                          {showHidden ? 'Showing all stocks (grayed out)' : 'Hiding researched stocks'}
                        </Typography>
                      </Box>
                      <Switch
                        checked={showHidden}
                        onChange={(event) => setShowHidden(event.target.checked)}
                        inputProps={{ 'aria-label': 'Show hidden stocks' }}
                      />
                    </Stack>
                    {hiddenSymbols.size > 0 && (
                      <Button
                        size="small"
                        onClick={() => setHiddenSymbols(new Set())}
                        sx={{ mt: 1 }}
                        color="warning"
                      >
                        Reset all hidden ({hiddenSymbols.size})
                      </Button>
                    )}
                  </Box>
                  <Divider />
                  {filters.map((filter, index) => (
                    <Box
                      key={filter.metric}
                      sx={{
                        opacity: filter.active ? 1 : 0.45,
                        transition: 'opacity 160ms ease',
                      }}
                    >
                      <Stack direction="row" alignItems="center" justifyContent="space-between" sx={{ mb: 1 }}>
                        <Typography variant="caption" fontWeight={900} color="text.secondary">
                          {filter.label.toUpperCase()}
                        </Typography>
                        <Switch
                          checked={filter.active}
                          onChange={() => toggleFilter(index)}
                          size="small"
                          inputProps={{ 'aria-label': `Toggle ${filter.label}` }}
                        />
                      </Stack>
                      <TextField
                        disabled={!filter.active}
                        type="number"
                        value={cutoffInputs[filter.metric] ?? String(filter.cutoff)}
                        onChange={(event) => updateCutoff(index, event.target.value)}
                        onBlur={() => resetCutoffInput(index)}
                        fullWidth
                        size="small"
                        inputProps={{ step: 0.1 }}
                        InputProps={{
                          startAdornment: (
                            <InputAdornment position="start">
                              {filter.operator === 'gte' ? '>=' : '<='}
                            </InputAdornment>
                          ),
                        }}
                      />
                    </Box>
                  ))}
                </Stack>
              </Card>

              <Paper
                variant="outlined"
                sx={{
                  flex: 1,
                  minWidth: 0,
                  width: '100%',
                  height: 640,
                  overflow: 'hidden',
                  bgcolor: 'background.paper',
                }}
              >
                <DataGrid
                  rows={displayedStocks}
                  columns={columns}
                  getRowId={(row) => row.symbol}
                  sortModel={sortModel}
                  onSortModelChange={setSortModel}
                  density="compact"
                  disableRowSelectionOnClick
                  slots={{ toolbar: GridToolbar }}
                  initialState={{
                    pagination: {
                      paginationModel: { pageSize: -1 },
                    },
                  }}
                  getRowClassName={(params: GridRowClassNameParams<StockData>) =>
                    hiddenSymbols.has(params.row.symbol) ? 'row--hidden' : ''
                  }
                  pageSizeOptions={[{ value: -1, label: 'All' }, 25, 50, 100]}
                  sx={{
                    border: 0,
                    '& .MuiDataGrid-columnHeaders': {
                      bgcolor: '#f8fafc',
                    },
                    '& .MuiDataGrid-columnHeaderTitle': {
                      fontWeight: 800,
                    },
                    '& .MuiDataGrid-cell': {
                      alignItems: 'center',
                    },
                    '& .MuiDataGrid-cell--textRight': {
                      justifyContent: 'flex-end',
                    },
                    '& .row--hidden': {
                      bgcolor: 'rgba(241, 245, 249, 0.6)',
                      color: 'text.disabled',
                      opacity: 0.6,
                      fontStyle: 'italic',
                      '&:hover': {
                        bgcolor: 'rgba(241, 245, 249, 0.8) !important',
                      },
                    },
                  }}
                />
              </Paper>
            </Stack>
          )}
        </Container>
        <Snackbar
          open={Boolean(copyMessage)}
          autoHideDuration={2200}
          onClose={() => setCopyMessage(null)}
          message={copyMessage}
        />
        <Snackbar
          open={Boolean(copyError)}
          autoHideDuration={4000}
          onClose={() => setCopyError(null)}
          anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
        >
          <Alert severity="error" variant="filled" onClose={() => setCopyError(null)}>
            {copyError}
          </Alert>
        </Snackbar>
      </Box>
    </ThemeProvider>
  );
}
