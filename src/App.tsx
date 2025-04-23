import { ThemeProvider, createTheme, CssBaseline, Container, Box, Typography } from '@mui/material';
import ExcelProcessor from './components/ExcelProcessor';

// 创建Material UI主题
const theme = createTheme({
  palette: {
    primary: {
      main: '#4caf50', // 主色调为绿色
    },
    secondary: {
      main: '#2196f3', // 次要色调为蓝色
    },
    background: {
      default: '#f5f5f5', // 页面背景色
    },
  },
  typography: {
    fontFamily: '"Roboto", "Helvetica", "Arial", sans-serif',
    h1: {
      fontSize: '2rem',
      fontWeight: 500,
    },
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none', // 按钮文字不全大写
        },
      },
    },
  },
});

function App() {
  const currentYear = new Date().getFullYear();

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline /> {/* 重置CSS */}
      <Container maxWidth="lg" sx={{ py: 4 }}>
        <Box component="header" sx={{ mb: 4, textAlign: 'center' }}>
          <Typography variant="h4" component="h1" gutterBottom>
            Excel链接自动打开工具
          </Typography>
        </Box>
        
        <Box component="main">
          <ExcelProcessor />
        </Box>
        
        <Box component="footer" sx={{ mt: 6, textAlign: 'center', color: 'text.secondary', fontSize: '0.875rem', py: 2, borderTop: '1px solid #eee' }}>
          <Typography variant="body2">
            © {currentYear} Excel链接处理工具
          </Typography>
        </Box>
      </Container>
    </ThemeProvider>
  );
}

export default App;
