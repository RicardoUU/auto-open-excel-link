import { ChangeEvent, useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  Button, 
  Paper, 
  Typography, 
  Table, 
  TableBody, 
  TableCell, 
  TableContainer, 
  TableHead, 
  TableRow,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Checkbox,
  Card,
  CardHeader,
  CardContent,
  CardActions,
  Box,
  CircularProgress,
  Chip,
  Divider,
  Stack,
  Link,
  IconButton,
  SelectChangeEvent,
  ToggleButtonGroup,
  ToggleButton
} from '@mui/material';
import { 
  CloudUpload as CloudUploadIcon, 
  OpenInNew as OpenInNewIcon,
  ContentCopy as ContentCopyIcon,
  TableChart as TableChartIcon,
  ViewModule as ViewModuleIcon,
  CheckCircle as CheckCircleIcon,
  Link as LinkIcon
} from '@mui/icons-material';

interface ExcelLink {
  text: string;
  url: string;
  row: number;
  col: number;
}

const ExcelProcessor = () => {
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [links, setLinks] = useState<ExcelLink[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [openingLinks, setOpeningLinks] = useState(false);
  const [currentLinkIndex, setCurrentLinkIndex] = useState(-1);
  const [dragActive, setDragActive] = useState(false);
  const [selectedLinks, setSelectedLinks] = useState<Set<number>>(new Set());
  const [viewMode, setViewMode] = useState<'table' | 'cards'>('table');

  const handleDrag = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || 
          file.type === "application/vnd.ms-excel") {
        handleFile(file);
      } else {
        alert("请上传Excel文件 (.xlsx, .xls)");
      }
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (!selectedFile) {
      return;
    }
    handleFile(selectedFile);
  };

  const handleFile = (selectedFile: File) => {
    setFile(selectedFile);
    setIsProcessing(true);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheets = workbook.SheetNames;
        setSheetNames(sheets);
        
        if (sheets.length > 0) {
          setSelectedSheet(sheets[0]);
          extractLinksFromSheet(workbook, sheets[0]);
        }
        
        setIsProcessing(false);
      } catch (error) {
        console.error('Error reading Excel file:', error);
        setIsProcessing(false);
      }
    };

    reader.readAsArrayBuffer(selectedFile);
  };

  const handleSheetChange = (e: SelectChangeEvent) => {
    const sheetName = e.target.value;
    setSelectedSheet(sheetName);
    
    if (file) {
      setIsProcessing(true);
      const reader = new FileReader();
      
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          extractLinksFromSheet(workbook, sheetName);
          setIsProcessing(false);
        } catch (error) {
          console.error('Error reading Excel file:', error);
          setIsProcessing(false);
        }
      };
      
      reader.readAsArrayBuffer(file);
    }
  };

  const extractLinksFromSheet = (workbook: XLSX.WorkBook, sheetName: string) => {
    const worksheet = workbook.Sheets[sheetName];
    const extractedLinks: ExcelLink[] = [];

    // 查找包含超链接的单元格
    Object.keys(worksheet).forEach((cell) => {
      if (cell.startsWith('!')) return; // 跳过非数据单元格

      const cellData = worksheet[cell];
      
      // 检查单元格是否有超链接
      if (cellData.l && cellData.l.Target) {
        const cellAddress = XLSX.utils.decode_cell(cell);
        
        extractedLinks.push({
          text: cellData.v || '无文本',
          url: cellData.l.Target,
          row: cellAddress.r,
          col: cellAddress.c
        });
      }
    });

    setLinks(extractedLinks);
  };

  // 处理链接选择
  const toggleLinkSelection = (index: number) => {
    const newSelection = new Set(selectedLinks);
    if (newSelection.has(index)) {
      newSelection.delete(index);
    } else {
      newSelection.add(index);
    }
    setSelectedLinks(newSelection);
  };

  // 全选/取消全选
  const toggleSelectAll = () => {
    if (selectedLinks.size === links.length) {
      // 如果已全选，则取消所有选择
      setSelectedLinks(new Set());
    } else {
      // 如果未全选，则全选
      const allIndices = new Set(links.map((_, index) => index));
      setSelectedLinks(allIndices);
    }
  };

  const openAllLinks = () => {
    if (links.length === 0) return;
    
    // 如果已经在打开链接状态，则退出该状态
    if (openingLinks) {
      setOpeningLinks(false);
      setCurrentLinkIndex(-1);
      return;
    }
    
    // 如果浏览器支持特殊打开功能，则使用
    if ('clipboard' in navigator && 'writeText' in navigator.clipboard) {
      // 提示用户如何操作
      alert(`
为打开${links.length}个链接，我们有两种方式:

1. 点击"确定"后，系统将逐个打开链接。浏览器可能会询问您是否允许。

2. 也可以使用Ctrl(Win)或Command(Mac)键+点击表格中链接的方式手动打开多个链接。
      `.trim());
      
      // 尝试自动打开链接
      links.forEach((link, i) => {
        setTimeout(() => {
          try {
            window.open(link.url, '_blank');
          } catch (e) {
            console.error('打开链接失败:', e);
          }
        }, i * 300); // 每300毫秒打开一个链接
      });
    } else {
      // 进入特殊打开状态
      setOpeningLinks(true);
      setCurrentLinkIndex(0);
      
      alert(`
请点击确定，然后使用下方表格中的"打开"按钮逐个打开链接。

提示：
- 在表格中按住Ctrl(Windows)或Command(Mac)键点击链接，可以在新标签页打开多个链接。
- 点击"结束批量打开"按钮可退出打开模式。
      `.trim());
    }
  };
  
  // 打开下一个链接
  const openNextLink = () => {
    if (currentLinkIndex >= 0 && currentLinkIndex < links.length) {
      window.open(links[currentLinkIndex].url, '_blank');
      setCurrentLinkIndex(currentLinkIndex + 1);
    } else {
      // 如果已经打开完所有链接，退出打开状态
      setOpeningLinks(false);
      setCurrentLinkIndex(-1);
    }
  };

  // 一键拷贝所有链接
  const copyAllLinks = () => {
    const linkTexts = links.map(link => link.url).join('\n');
    navigator.clipboard.writeText(linkTexts).then(() => {
      alert('所有链接已复制到剪贴板，您可以在新标签页中粘贴并一次性打开所有链接。');
    }).catch(err => {
      console.error('复制失败:', err);
      alert('链接复制失败，请检查浏览器权限。');
    });
  };

  // 只打开选定的链接
  const openSelectedLinks = () => {
    if (selectedLinks.size === 0) {
      alert('请先选择要打开的链接');
      return;
    }

    alert(`将打开 ${selectedLinks.size} 个选定的链接`);

    // 将选定的链接转换为数组
    const linksToOpen = Array.from(selectedLinks).map(index => links[index]);
    
    // 打开选定的链接
    linksToOpen.forEach((link, i) => {
      setTimeout(() => {
        try {
          window.open(link.url, '_blank');
        } catch (e) {
          console.error('打开链接失败:', e);
        }
      }, i * 300);
    });
  };

  const handleViewModeChange = (_: React.MouseEvent<HTMLElement>, newMode: 'table' | 'cards' | null) => {
    if (newMode !== null) {
      setViewMode(newMode);
    }
  };

  return (
    <Paper elevation={3} sx={{ p: 3, borderRadius: 2, maxWidth: 1200, mx: 'auto' }}>
      <Typography variant="h4" component="h2" gutterBottom align="center" sx={{ mb: 3 }}>
        Excel链接解析器
      </Typography>
      
      <Box 
        sx={{
          border: dragActive ? '2px dashed #4caf50' : '2px dashed #e0e0e0',
          borderRadius: 2,
          p: 3,
          backgroundColor: dragActive ? 'rgba(76, 175, 80, 0.05)' : '#f9f9f9',
          textAlign: 'center',
          transition: 'all 0.3s',
          mb: 3
        }}
        onDragEnter={handleDrag}
        onDragOver={handleDrag}
        onDragLeave={handleDrag}
        onDrop={handleDrop}
      >
        <LinkIcon color="primary" sx={{ fontSize: 48, mb: 2, opacity: 0.7 }} />
        <Typography variant="h6" gutterBottom>
          选择或拖放Excel文件
        </Typography>
        <input
          id="file-upload"
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileChange}
          style={{ display: 'none' }}
        />
        <label htmlFor="file-upload">
          <Button
            variant="contained"
            component="span"
            startIcon={<CloudUploadIcon />}
            sx={{ mb: 1 }}
          >
            选择文件
          </Button>
        </label>
        <Typography variant="body2" color="textSecondary" sx={{ mt: 1 }}>
          支持的格式: .xlsx, .xls
        </Typography>
        {file && (
          <Chip 
            label={file.name} 
            color="primary" 
            variant="outlined" 
            sx={{ mt: 2 }} 
          />
        )}
      </Box>

      {isProcessing && (
        <Box sx={{ display: 'flex', justifyContent: 'center', alignItems: 'center', flexDirection: 'column', my: 3 }}>
          <CircularProgress size={40} />
          <Typography variant="body1" sx={{ mt: 2 }}>
            正在处理文件，请稍候...
          </Typography>
        </Box>
      )}

      {sheetNames.length > 0 && (
        <Box sx={{ my: 3, backgroundColor: '#f5f5f5', p: 2, borderRadius: 1 }}>
          <FormControl fullWidth>
            <InputLabel id="sheet-select-label">选择工作表</InputLabel>
            <Select
              labelId="sheet-select-label"
              id="sheet-select"
              value={selectedSheet}
              label="选择工作表"
              onChange={handleSheetChange}
            >
              {sheetNames.map((name) => (
                <MenuItem key={name} value={name}>{name}</MenuItem>
              ))}
            </Select>
          </FormControl>
        </Box>
      )}

      {links.length > 0 ? (
        <Box sx={{ mt: 3 }}>
          <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
            <Typography variant="h6" component="h3">
              找到 {links.length} 个链接
            </Typography>
            <ToggleButtonGroup
              value={viewMode}
              exclusive
              onChange={handleViewModeChange}
              size="small"
            >
              <ToggleButton value="table">
                <TableChartIcon fontSize="small" />
              </ToggleButton>
              <ToggleButton value="cards">
                <ViewModuleIcon fontSize="small" />
              </ToggleButton>
            </ToggleButtonGroup>
          </Box>

          <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} sx={{ mb: 3 }}>
            <Button 
              variant="contained" 
              color="primary" 
              startIcon={<OpenInNewIcon />}
              onClick={openAllLinks}
              fullWidth
            >
              {openingLinks ? '结束批量打开' : '打开所有链接'}
            </Button>
            
            {selectedLinks.size > 0 && (
              <Button 
                variant="contained" 
                color="warning" 
                startIcon={<CheckCircleIcon />}
                onClick={openSelectedLinks}
                fullWidth
              >
                打开已选择的链接 ({selectedLinks.size})
              </Button>
            )}
            
            {navigator.clipboard && (
              <Button 
                variant="outlined" 
                color="primary" 
                startIcon={<ContentCopyIcon />}
                onClick={copyAllLinks}
                fullWidth
              >
                复制所有链接
              </Button>
            )}
          </Stack>
          
          {openingLinks && currentLinkIndex < links.length && (
            <Paper elevation={1} sx={{ p: 2, mb: 3, display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderLeft: '4px solid #4caf50' }}>
              <Typography variant="body1" sx={{ fontWeight: 'medium' }}>
                正在准备打开第 {currentLinkIndex + 1}/{links.length} 个链接
              </Typography>
              <Button 
                variant="contained" 
                color="warning" 
                onClick={openNextLink}
                size="small"
              >
                打开下一个链接
              </Button>
            </Paper>
          )}
          
          {viewMode === 'table' ? (
            <TableContainer component={Paper} elevation={1}>
              <Table size="small">
                <TableHead>
                  <TableRow>
                    <TableCell padding="checkbox">
                      <Checkbox
                        checked={selectedLinks.size === links.length}
                        onChange={toggleSelectAll}
                        indeterminate={selectedLinks.size > 0 && selectedLinks.size < links.length}
                      />
                    </TableCell>
                    <TableCell>文本</TableCell>
                    <TableCell>URL</TableCell>
                    <TableCell>位置</TableCell>
                    {openingLinks && <TableCell>操作</TableCell>}
                  </TableRow>
                </TableHead>
                <TableBody>
                  {links.map((link, index) => (
                    <TableRow 
                      key={index}
                      sx={{ 
                        backgroundColor: currentLinkIndex === index ? 'rgba(76, 175, 80, 0.08)' : 'inherit',
                        borderLeft: currentLinkIndex === index ? '3px solid #4caf50' : 'none',
                        '&.Mui-selected, &.Mui-selected:hover': {
                          backgroundColor: 'rgba(33, 150, 243, 0.08)',
                        }
                      }}
                      selected={selectedLinks.has(index)}
                    >
                      <TableCell padding="checkbox">
                        <Checkbox
                          checked={selectedLinks.has(index)}
                          onChange={() => toggleLinkSelection(index)}
                        />
                      </TableCell>
                      <TableCell>{link.text}</TableCell>
                      <TableCell>
                        <Link
                          href={link.url}
                          target="_blank"
                          rel="noopener noreferrer"
                          sx={{ 
                            fontWeight: openingLinks ? 'bold' : 'regular',
                            color: openingLinks ? '#4caf50' : 'primary.main'
                          }}
                        >
                          {link.url}
                        </Link>
                      </TableCell>
                      <TableCell>行 {link.row + 1}, 列 {link.col + 1}</TableCell>
                      {openingLinks && (
                        <TableCell>
                          <IconButton
                            color="primary"
                            size="small"
                            onClick={() => window.open(link.url, '_blank')}
                          >
                            <OpenInNewIcon fontSize="small" />
                          </IconButton>
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </TableContainer>
          ) : (
            <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2 }}>
              {links.map((link, index) => (
                <Box 
                  key={index} 
                  sx={{ 
                    width: { xs: '100%', sm: 'calc(50% - 16px)', md: 'calc(33.333% - 16px)' },
                    mb: 2 
                  }}
                >
                  <Card 
                    variant="outlined" 
                    sx={{ 
                      height: '100%',
                      display: 'flex',
                      flexDirection: 'column',
                      border: selectedLinks.has(index) ? '1px solid #2196F3' : '1px solid #eee',
                      boxShadow: selectedLinks.has(index) ? '0 2px 8px rgba(33, 150, 243, 0.15)' : 'none',
                      transition: 'all 0.3s',
                      '&:hover': {
                        boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
                        transform: 'translateY(-2px)'
                      }
                    }}
                  >
                    <CardHeader
                      avatar={
                        <Checkbox
                          checked={selectedLinks.has(index)}
                          onChange={() => toggleLinkSelection(index)}
                        />
                      }
                      title={`链接 #${index + 1}`}
                      sx={{ backgroundColor: '#f5f7fa', borderBottom: '1px solid #eee' }}
                    />
                    <CardContent sx={{ flexGrow: 1 }}>
                      <Typography variant="subtitle1" gutterBottom sx={{ fontWeight: 'medium' }}>
                        {link.text}
                      </Typography>
                      <Link
                        href={link.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        sx={{ 
                          display: 'block', 
                          mb: 2,
                          wordBreak: 'break-all'
                        }}
                      >
                        {link.url}
                      </Link>
                      <Chip 
                        label={`行 ${link.row + 1}, 列 ${link.col + 1}`} 
                        size="small" 
                        variant="outlined"
                      />
                    </CardContent>
                    <Divider />
                    <CardActions>
                      <Button
                        startIcon={<OpenInNewIcon />}
                        fullWidth
                        onClick={() => window.open(link.url, '_blank')}
                        color="primary"
                      >
                        打开链接
                      </Button>
                    </CardActions>
                  </Card>
                </Box>
              ))}
            </Box>
          )}
        </Box>
      ) : (
        selectedSheet && !isProcessing && (
          <Typography variant="body1" align="center" sx={{ my: 3 }}>
            未在选定的工作表中找到链接
          </Typography>
        )
      )}
    </Paper>
  );
};

export default ExcelProcessor; 