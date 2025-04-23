import { ChangeEvent, useState } from 'react';
import * as XLSX from 'xlsx';
import './ExcelProcessor.css';

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

  const handleSheetChange = (e: ChangeEvent<HTMLSelectElement>) => {
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

  return (
    <div className="excel-processor">
      <h2>Excel链接解析器</h2>
      
      <div 
        className={`file-input ${dragActive ? 'drag-active' : ''}`}
        onDragEnter={handleDrag}
        onDragOver={handleDrag}
        onDragLeave={handleDrag}
        onDrop={handleDrop}
      >
        <div className="file-icon"></div>
        <p className="file-title">选择或拖放Excel文件</p>
        <input 
          id="file-upload"
          type="file" 
          accept=".xlsx, .xls" 
          onChange={handleFileChange} 
          className="file-upload"
        />
        <label htmlFor="file-upload" className="file-upload-label">
          选择文件
        </label>
        <p className="file-help">支持的格式: .xlsx, .xls</p>
        {file && <p className="selected-file">已选择: {file.name}</p>}
      </div>

      {isProcessing && (
        <div className="loading-container">
          <div className="loading-spinner"></div>
          <p>正在处理文件，请稍候...</p>
        </div>
      )}

      {sheetNames.length > 0 && (
        <div className="sheet-selector">
          <label htmlFor="sheet-select">选择工作表:</label>
          <select 
            id="sheet-select" 
            value={selectedSheet} 
            onChange={handleSheetChange}
          >
            {sheetNames.map((name) => (
              <option key={name} value={name}>{name}</option>
            ))}
          </select>
        </div>
      )}

      {links.length > 0 ? (
        <div className="links-section">
          <h3>找到 {links.length} 个链接</h3>
          <div className="buttons-group">
            <button onClick={openAllLinks} className="open-links-btn">
              {openingLinks ? '结束批量打开' : '打开所有链接'}
            </button>
            {selectedLinks.size > 0 && (
              <button onClick={openSelectedLinks} className="open-selected-btn">
                打开已选择的链接 ({selectedLinks.size})
              </button>
            )}
            {navigator.clipboard && (
              <button onClick={copyAllLinks} className="copy-links-btn">
                复制所有链接
              </button>
            )}
          </div>

          {/* 视图切换 */}
          <div className="view-toggle">
            <button 
              className={`view-btn ${viewMode === 'table' ? 'active' : ''}`}
              onClick={() => setViewMode('table')}
            >
              表格视图
            </button>
            <button 
              className={`view-btn ${viewMode === 'cards' ? 'active' : ''}`}
              onClick={() => setViewMode('cards')}
            >
              卡片视图
            </button>
          </div>
          
          {openingLinks && currentLinkIndex < links.length && (
            <div className="opening-status">
              <p>正在准备打开第 {currentLinkIndex + 1}/{links.length} 个链接</p>
              <button onClick={openNextLink} className="open-next-btn">
                打开下一个链接
              </button>
            </div>
          )}
          
          <div className="links-list">
            {viewMode === 'table' ? (
              <table>
                <thead>
                  <tr>
                    <th className="checkbox-cell">
                      <input 
                        type="checkbox" 
                        checked={selectedLinks.size === links.length}
                        onChange={toggleSelectAll}
                      />
                    </th>
                    <th>文本</th>
                    <th>URL</th>
                    <th>位置</th>
                    {openingLinks && <th>操作</th>}
                  </tr>
                </thead>
                <tbody>
                  {links.map((link, index) => (
                    <tr 
                      key={index} 
                      className={`
                        ${currentLinkIndex === index ? 'current-link' : ''}
                        ${selectedLinks.has(index) ? 'selected-link' : ''}
                      `}
                    >
                      <td className="checkbox-cell">
                        <input 
                          type="checkbox" 
                          checked={selectedLinks.has(index)}
                          onChange={() => toggleLinkSelection(index)}
                        />
                      </td>
                      <td>{link.text}</td>
                      <td>
                        <a 
                          href={link.url} 
                          target="_blank" 
                          rel="noopener noreferrer"
                          className={openingLinks ? 'highlight-link' : ''}
                        >
                          {link.url}
                        </a>
                      </td>
                      <td>行 {link.row + 1}, 列 {link.col + 1}</td>
                      {openingLinks && (
                        <td>
                          <button 
                            onClick={() => window.open(link.url, '_blank')}
                            className="open-this-link-btn"
                          >
                            打开
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div className="links-cards">
                {links.map((link, index) => (
                  <div 
                    key={index} 
                    className={`
                      link-card
                      ${currentLinkIndex === index ? 'current-link' : ''}
                      ${selectedLinks.has(index) ? 'selected-link' : ''}
                    `}
                  >
                    <div className="card-header">
                      <input 
                        type="checkbox" 
                        checked={selectedLinks.has(index)}
                        onChange={() => toggleLinkSelection(index)}
                      />
                      <span className="link-index">#{index + 1}</span>
                    </div>
                    <div className="card-body">
                      <div className="link-text">{link.text}</div>
                      <a 
                        href={link.url} 
                        target="_blank" 
                        rel="noopener noreferrer"
                        className="link-url"
                      >
                        {link.url}
                      </a>
                      <div className="link-position">
                        位置: 行 {link.row + 1}, 列 {link.col + 1}
                      </div>
                    </div>
                    <div className="card-footer">
                      <button 
                        onClick={() => window.open(link.url, '_blank')}
                        className="card-open-btn"
                      >
                        打开链接
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      ) : (
        selectedSheet && !isProcessing && <p>未在选定的工作表中找到链接</p>
      )}
    </div>
  );
};

export default ExcelProcessor; 