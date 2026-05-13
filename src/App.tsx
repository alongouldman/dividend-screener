import { useEffect, useMemo, useState } from 'react';
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
  InputAdornment,
  Paper,
  Stack,
  Switch,
  TextField,
  ThemeProvider,
  Toolbar,
  Typography,
  createTheme,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import ErrorOutlineIcon from '@mui/icons-material/ErrorOutline';
import FilterListIcon from '@mui/icons-material/FilterList';
import InsightsIcon from '@mui/icons-material/Insights';
import UploadFileIcon from '@mui/icons-material/UploadFile';
import {
  DataGrid,
  GridToolbar,
  type GridColDef,
  type GridRenderCellParams,
  type GridSortModel,
} from '@mui/x-data-grid';
import type { FilterConfig, NumericValue, StockData } from './types';
import { parseExcel } from './utils/parser';

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
  { metric: 'past5yGrowth', label: 'Past 5Y (%)', active: false, cutoff: 0, operator: 'gte' },
  { metric: 'est5yGrowth', label: 'Est 5Y (%)', active: false, cutoff: 0, operator: 'gte' },
  { metric: 'debtEquity', label: 'Debt/Equity', active: false, cutoff: 0.7, operator: 'lte' },
  { metric: 'chowderRule', label: 'Chowder Rule', active: true, cutoff: 11, operator: 'gte' },
  { metric: 'roa', label: 'ROA (%)', active: false, cutoff: 5, operator: 'gte' },
];

const FILTERS_STORAGE_KEY = 'screener_filters';
const MISSING_VALUES_MATCH_STORAGE_KEY = 'screener_missing_values_match';
const STOCKS_STORAGE_KEY = 'screener_stocks';
const SORT_MODEL_STORAGE_KEY = 'screener_sort_model';
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
const getCutoffInputs = (filters: FilterConfig[]) =>
  Object.fromEntries(filters.map((filter) => [filter.metric, String(filter.cutoff)]));

export default function App() {
  const [stocks, setStocks] = useState<StockData[]>(loadSavedStocks);
  const [filters, setFilters] = useState<FilterConfig[]>(loadSavedFilters);
  const [cutoffInputs, setCutoffInputs] = useState<Record<string, string>>(() =>
    getCutoffInputs(loadSavedFilters()),
  );
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [sortModel, setSortModel] = useState<GridSortModel>(loadSavedSortModel);
  const [missingValuesMatch, setMissingValuesMatch] = useState(loadSavedMissingValuesMatch);

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
    if (stocks.length === 0) {
      localStorage.removeItem(STOCKS_STORAGE_KEY);
      return;
    }

    localStorage.setItem(STOCKS_STORAGE_KEY, JSON.stringify(stocks));
  }, [stocks]);

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

  const handleStartNewFile = () => {
    localStorage.removeItem(STOCKS_STORAGE_KEY);
    localStorage.removeItem(FILTERS_STORAGE_KEY);
    localStorage.removeItem(SORT_MODEL_STORAGE_KEY);
    localStorage.removeItem(MISSING_VALUES_MATCH_STORAGE_KEY);
    setStocks([]);
    setError(null);
    setFilters(DEFAULT_FILTERS);
    setCutoffInputs(getCutoffInputs(DEFAULT_FILTERS));
    setSortModel(DEFAULT_SORT_MODEL);
    setMissingValuesMatch(false);
  };

  const filteredStocks = useMemo(() => {
    return stocks.filter((stock) => {
      return filters.every((filter) => {
        if (!filter.active) return true;
        const value = stock[filter.metric];
        if (!isValidNumber(value)) return missingValuesMatch;
        return filter.operator === 'gte' ? value >= filter.cutoff : value <= filter.cutoff;
      });
    });
  }, [stocks, filters, missingValuesMatch]);

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

  const columns = useMemo<GridColDef<StockData>[]>(
    () => [
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
        field: 'yield',
        headerName: 'Yield',
        width: 110,
        type: 'number',
        valueFormatter: (value: number | null | undefined) =>
          formatMissingValue(value, (numericValue) => `${numericValue.toFixed(2)}%`),
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
      { field: 'peRatio', headerName: 'P/E', width: 90, type: 'number', valueFormatter: formatRatio },
      { field: 'pbRatio', headerName: 'P/B', width: 90, type: 'number', valueFormatter: formatRatio },
      { field: 'roe', headerName: 'ROE', width: 100, type: 'number', valueFormatter: formatPercent },
      { field: 'past5yGrowth', headerName: 'Past 5Y', width: 110, type: 'number', valueFormatter: formatPercent },
      { field: 'est5yGrowth', headerName: 'Est. 5Y', width: 110, type: 'number', valueFormatter: formatPercent },
      { field: 'debtEquity', headerName: 'D/E', width: 90, type: 'number', valueFormatter: formatDecimal },
      {
        field: 'chowderRule',
        headerName: 'Chowder',
        width: 120,
        type: 'number',
        valueFormatter: formatRatio,
      },
    ],
    [],
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
                <Button variant="text" startIcon={<CloudUploadIcon />} onClick={handleStartNewFile}>
                  Upload a new file
                </Button>
              )}
            </Stack>
          </Toolbar>
        </AppBar>

        <Container maxWidth={false} sx={{ maxWidth: 1600, py: 3, flex: 1 }}>
          {loading && (
            <Stack alignItems="center" justifyContent="center" spacing={2} sx={{ pt: 10 }}>
              <CircularProgress size={56} thickness={4} />
              <Typography variant="h6" fontWeight={800}>
                Processing Excel File
              </Typography>
              <Typography color="text.secondary">Reading data from the All CCC sheet...</Typography>
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
              <Typography fontWeight={800}>Parsing Error</Typography>
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
                  rows={filteredStocks}
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
                  }}
                />
              </Paper>
            </Stack>
          )}
        </Container>
      </Box>
    </ThemeProvider>
  );
}
