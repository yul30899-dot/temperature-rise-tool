import React, { useState, useRef, useMemo, useEffect } from 'react';
import { Settings, Sliders, FileSpreadsheet, Download, RefreshCw, UploadCloud, Trash2, CheckSquare, Square, LineChart as ChartIcon, FolderOpen, Save, FileJson, Eraser, HelpCircle } from 'lucide-react';
import { DEFAULT_CHANNELS, DEFAULT_CONFIG } from './constants';
import * as ExcelJS from 'exceljs';
import XlsxPopulate from 'xlsx-populate/browser/xlsx-populate.min.js';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { get, set } from 'idb-keyval';
import Papa from 'papaparse';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';

const CustomTooltip = ({ active, payload, label }) => {
  if (active && payload && payload.length) {
    const sortedPayload = [...payload].sort((a, b) => b.value - a.value);
    return (
      <div className="bg-white/95 backdrop-blur-md p-3 rounded-xl shadow-xl border border-slate-200/60 pointer-events-none">
        <p className="font-bold text-slate-700 mb-2 border-b border-slate-100 pb-1 text-center">{label}</p>
        <div className="grid grid-cols-4 md:grid-cols-5 lg:grid-cols-6 gap-x-4 gap-y-1">
          {sortedPayload.map((entry, index) => (
            <div key={`item-${index}`} className="flex items-center gap-1.5 text-xs whitespace-nowrap">
              <div className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: entry.color }} />
              <span className="text-slate-600 truncate max-w-[70px]" title={entry.name}>{entry.name}</span>
              <span className="font-semibold text-slate-800 ml-auto">{entry.value}</span>
            </div>
          ))}
        </div>
      </div>
    );
  }
  return null;
};

function App() {
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [channels, setChannels] = useState([]);
  const [rawData, setRawData] = useState([]);
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [chartSelectedIds, setChartSelectedIds] = useState(new Set());
  const [groupNotes, setGroupNotes] = useState({});
  const [activeTab, setActiveTab] = useState('config');
  const [focusedField, setFocusedField] = useState(null);
  const fileInputRef = useRef(null);
  const templateInputRef = useRef(null);
  const [isImporting, setIsImporting] = useState(false);
  const [autoExportTemplate, setAutoExportTemplate] = useState(true);
  const [importedFileName, setImportedFileName] = useState('');
  const [exportSuccessModal, setExportSuccessModal] = useState({ show: false, message: '', blob: null, fileName: '', savedPath: '' });
  const [toast, setToast] = useState({ show: false, message: '', type: 'success' });
  const [confirmModal, setConfirmModal] = useState({ show: false, message: '', onConfirm: null });
  const [isExporting, setIsExporting] = useState(false);
  const [showMaxTemp, setShowMaxTemp] = useState(false);
  const [fullRawData, setFullRawData] = useState([]);
  
  const [compareFullRawData, setCompareFullRawData] = useState([]);
  const [compareRawData, setCompareRawData] = useState([]);
  const [compareFileName, setCompareFileName] = useState('');
  const [isComparing, setIsComparing] = useState(false);
  const compareInputRef = useRef(null);

  const recalculateChannelsFromData = (data, compareData = null) => {
    if (!data || data.length === 0) return;
    const maxTemps = {};
    const stableTemps = {};
    data.forEach(row => {
      Object.keys(row).forEach(key => {
        if (key.startsWith('CH')) {
          const chId = parseInt(key.replace('CH', ''), 10);
          const val = row[key];
          if (!(chId in maxTemps) || val > maxTemps[chId]) {
            maxTemps[chId] = val;
          }
          stableTemps[chId] = val;
        }
      });
    });

    const compMaxTemps = {};
    const compStableTemps = {};
    if (compareData && compareData.length > 0) {
      compareData.forEach(row => {
        Object.keys(row).forEach(key => {
          if (key.startsWith('CH')) {
            const chId = parseInt(key.replace('CH', ''), 10);
            const val = row[key];
            if (!(chId in compMaxTemps) || val > compMaxTemps[chId]) {
              compMaxTemps[chId] = val;
            }
            compStableTemps[chId] = val;
          }
        });
      });
    }

    setChannels(prevChannels => prevChannels.map(ch => {
      const updates = {
        ...ch,
        temp: stableTemps[ch.id] !== undefined ? stableTemps[ch.id].toFixed(2) : ch.temp,
        maxTemp: maxTemps[ch.id] !== undefined ? maxTemps[ch.id].toFixed(2) : ch.maxTemp
      };
      if (compareData && compareData.length > 0) {
        updates.compareTemp = compStableTemps[ch.id] !== undefined ? compStableTemps[ch.id].toFixed(2) : ch.compareTemp;
        updates.compareMaxTemp = compMaxTemps[ch.id] !== undefined ? compMaxTemps[ch.id].toFixed(2) : ch.compareMaxTemp;
      }
      return updates;
    }));
  };

  const handleCropData = () => {
    if (!zoomRange || !fullRawData.length) return;
    const [start, end] = zoomRange;
    const cropped = fullRawData.slice(Math.max(0, Math.floor(start)), Math.min(fullRawData.length, Math.floor(end) + 1));
    setRawData(cropped);
    let compareCropped = [];
    if (isComparing && compareFullRawData.length > 0) {
      compareCropped = compareFullRawData.slice(Math.max(0, Math.floor(start)), Math.min(compareFullRawData.length, Math.floor(end) + 1));
      setCompareRawData(compareCropped);
    }
    recalculateChannelsFromData(cropped, isComparing ? compareCropped : null);
    showToast('已截取当前时间段并重新计算温度');
  };

  const handleResetData = () => {
    setRawData(fullRawData);
    if (isComparing) {
      setCompareRawData(compareFullRawData);
    }
    recalculateChannelsFromData(fullRawData, isComparing ? compareFullRawData : null);
    setZoomRange(fullRawData.length > 0 ? [0, fullRawData.length - 1] : null);
    if (fullRawData.length > 0) {
      setCropStartIndex(0);
      setCropEndIndex(fullRawData.length - 1);
    }
    setYDomain(['auto', 'auto']);
    showToast('已恢复全局数据');
  };

  const showToast = (message, type = 'success') => {
    setToast({ show: true, message, type });
    setTimeout(() => setToast({ show: false, message: '', type }), 3000);
  };

  const updateStyle = (field, updates) => {
    if (!field) return;
    setConfig(prev => ({
      ...prev,
      styles: {
        ...prev.styles,
        [field]: {
          ...(prev.styles?.[field] || DEFAULT_CONFIG.styles[field]),
          ...updates
        }
      }
    }));
  };

  const [exportDirHandle, setExportDirHandle] = useState(null);
  const [exportDirName, setExportDirName] = useState('');

  const isDragging = useRef(false);
  const lastClientX = useRef(0);

  // Load saved handle on mount
  useEffect(() => {
    async function loadHandle() {
      try {
        if (window.electronAPI) {
          const pathStr = await get('electron_export_dir');
          if (pathStr) {
            setExportDirName(pathStr);
          }
        } else {
          const handle = await get('export_dir_handle');
          if (handle) {
            setExportDirHandle(handle);
            setExportDirName(handle.name);
          }
        }
      } catch (e) {
        console.error(e);
      }
    }
    loadHandle();
  }, []);

  // Helper to chunk array into groups of 8
  const chunkedChannels = useMemo(() => {
    const chunks = [];
    for (let i = 0; i < channels.length; i += 8) {
      chunks.push(channels.slice(i, i + 8));
    }
    return chunks;
  }, [channels]);

  const [zoomRange, setZoomRange] = useState(null); // [startIdx, endIdx]
  const [yDomain, setYDomain] = useState(['auto', 'auto']);
  const [cropStartIndex, setCropStartIndex] = useState(0);
  const [cropEndIndex, setCropEndIndex] = useState(0);

  useEffect(() => {
    if (fullRawData && fullRawData.length > 0) {
      setCropStartIndex(0);
      setCropEndIndex(fullRawData.length - 1);
    }
  }, [fullRawData]);

  const handleCropExact = () => {
    if (cropStartIndex >= cropEndIndex) {
      showToast('起始时间不能晚于结束时间！', 'error');
      return;
    }
    const cropped = fullRawData.slice(cropStartIndex, cropEndIndex + 1);
    setRawData(cropped);
    
    let compareCropped = [];
    if (isComparing && compareFullRawData.length > 0) {
      compareCropped = compareFullRawData.slice(cropStartIndex, cropEndIndex + 1);
      setCompareRawData(compareCropped);
    }
    
    recalculateChannelsFromData(cropped, isComparing ? compareCropped : null);
    setZoomRange([0, cropped.length - 1]);
    showToast('已截取选定时间段记录数据');
  };

  useEffect(() => {
    if (rawData.length > 0) {
      setZoomRange([0, rawData.length - 1]);
      setYDomain(['auto', 'auto']);
    } else {
      setZoomRange(null);
      setYDomain(['auto', 'auto']);
    }
  }, [rawData]);

  const getVisibleYDomain = () => {
    if (yDomain[0] !== 'auto') return yDomain;
    let min = Infinity;
    let max = -Infinity;
    chartData.forEach(d => {
        chartSelectedIds.forEach(id => {
            const val = d[`CH${id}`];
            if (val !== undefined) {
                if (val < min) min = val;
                if (val > max) max = val;
            }
        });
    });
    if (min === Infinity) return [0, 100];
    const padding = (max - min) * 0.05 || 5;
    return [Math.floor(min - padding), Math.ceil(max + padding)];
  };

  // Sliced and downsampled data for rendering performance
  const chartData = useMemo(() => {
    if (rawData.length === 0 || !zoomRange) return [];
    let [start, end] = zoomRange;
    start = Math.max(0, Math.floor(start));
    end = Math.min(rawData.length - 1, Math.floor(end));
    if (start >= end) return [];
    
    const slice = rawData.slice(start, end + 1);
    const compareSlice = isComparing ? compareRawData.slice(start, end + 1) : [];
    
    // Use 80 points limit to drastically reduce React rendering load when 72+ channels are shown
    let step = 1;
    if (slice.length > 80) {
      step = Math.ceil(slice.length / 80);
    }
    
    const result = [];
    for (let i = 0; i < slice.length; i += step) {
      // make sure last point is included if it wasn't
      const isLast = (i + step >= slice.length);
      const actualIdx = isLast ? slice.length - 1 : i;
      
      const point = { ...slice[actualIdx] };
      if (isComparing && compareSlice[actualIdx]) {
        const cp = compareSlice[actualIdx];
        if (cp.time) {
          point.compare_time = cp.time;
        }
        Object.keys(cp).forEach(key => {
          if (key.startsWith('CH')) {
            point[`compare_${key}`] = cp[key];
          }
        });
      }
      result.push(point);
      if (isLast) break;
    }
    return result;
  }, [rawData, zoomRange, isComparing, compareRawData]);

  const chartDataRef = useRef(chartData);
  const chartSelectedIdsRef = useRef(chartSelectedIds);
  const yDomainRef = useRef(yDomain);

  useEffect(() => {
    chartDataRef.current = chartData;
    chartSelectedIdsRef.current = chartSelectedIds;
    yDomainRef.current = yDomain;
  }, [chartData, chartSelectedIds, yDomain]);

  const chartContainerRef = useRef(null);

  useEffect(() => {
    const container = chartContainerRef.current;
    if (!container || rawData.length === 0) return;

    let isWheelTicking = false;
    let wheelAccumulator = 0;
    const handleWheel = (e) => {
      e.preventDefault();
      wheelAccumulator += e.deltaY;
      
      if (!isWheelTicking) {
        isWheelTicking = true;
        setTimeout(() => {
          const deltaY = wheelAccumulator;
          wheelAccumulator = 0;
          isWheelTicking = false;

          const rect = container.getBoundingClientRect();
          const x = e.clientX - rect.left;
          const width = rect.width;
          const mouseRatio = Math.max(0, Math.min(1, x / width));

          setZoomRange(prev => {
            if (!prev) return prev;
            const [start, end] = prev;
            const currentLen = end - start;
            
            // 15% zoom per tick, scale by how fast they spin
            const ticks = Math.max(1, Math.abs(deltaY) / 100);
            const zoomAmount = Math.max(2, currentLen * 0.15 * ticks);
            
            let newStart = start;
            let newEnd = end;
            
            if (deltaY < 0) {
              newStart = start + (zoomAmount * mouseRatio);
              newEnd = end - (zoomAmount * (1 - mouseRatio));
            } else if (deltaY > 0) {
              newStart = start - (zoomAmount * mouseRatio);
              newEnd = end + (zoomAmount * (1 - mouseRatio));
            }
            newStart = Math.floor(Math.max(0, newStart));
            newEnd = Math.floor(Math.min(rawData.length - 1, newEnd));
            
            if (newEnd - newStart < 5 && deltaY < 0) return prev; 
            
            return [newStart, newEnd];
          });

          setYDomain(prev => {
             let min = prev[0];
             let max = prev[1];
             if (min === 'auto') {
                 min = Infinity; max = -Infinity;
                 chartDataRef.current.forEach(d => {
                     chartSelectedIdsRef.current.forEach(id => {
                         const val = d[`CH${id}`];
                         if (val !== undefined) {
                             if (val < min) min = val;
                             if (val > max) max = val;
                         }
                     });
                 });
                 if (min === Infinity) return prev;
                 const pad = (max - min) * 0.05 || 5;
                 min = min - pad;
                 max = max + pad;
             }
             
             const currentLen = max - min;
             const ticks = Math.max(1, Math.abs(deltaY) / 100);
             const zoomAmount = Math.max(0.5, currentLen * 0.15 * ticks);
             
             const y = e.clientY - rect.top;
             const mouseRatioY = Math.max(0, Math.min(1, y / rect.height));
             
             let newMin = min;
             let newMax = max;
             if (deltaY < 0) { 
                newMax = max - (zoomAmount * mouseRatioY);
                newMin = min + (zoomAmount * (1 - mouseRatioY));
             } else if (deltaY > 0) {
                newMax = max + (zoomAmount * mouseRatioY);
                newMin = min - (zoomAmount * (1 - mouseRatioY));
             }
             if (newMax - newMin < 1) return prev;
             return [newMin, newMax];
          });
        }, 40); // 25 fps throttle for smooth but much lighter CPU usage
      }
    };

    const handleMouseDown = (e) => {
      if (e.button !== 0) return; // Only left click
      isDragging.current = true;
      lastClientX.current = e.clientX;
      // Store lastClientY directly on isDragging ref as a workaround since we don't have it declared yet,
      // or actually I can just add lastClientY to the top level refs.
      if (!isDragging.lastClientY) isDragging.lastClientY = e.clientY;
      isDragging.lastClientY = e.clientY;
      container.style.cursor = 'grabbing';
    };

    let isMouseTicking = false;
    let accumulatedDeltaX = 0;
    let accumulatedDeltaY = 0;
    const handleMouseMove = (e) => {
      if (!isDragging.current) return;
      
      accumulatedDeltaX += (e.clientX - lastClientX.current);
      accumulatedDeltaY += (e.clientY - isDragging.lastClientY);
      lastClientX.current = e.clientX;
      isDragging.lastClientY = e.clientY;
      
      if (!isMouseTicking) {
        isMouseTicking = true;
        setTimeout(() => {
          const deltaX = accumulatedDeltaX;
          const deltaYDrag = accumulatedDeltaY;
          accumulatedDeltaX = 0;
          accumulatedDeltaY = 0;
          isMouseTicking = false;

          setZoomRange(prev => {
            if (!prev) return prev;
            const [start, end] = prev;
            const currentLen = end - start;
            
            if (currentLen >= rawData.length - 1) return prev; 

            const rect = container.getBoundingClientRect();
            const pointsPerPixel = currentLen / rect.width;
            
            const shift = Math.round(deltaX * pointsPerPixel);
            if (shift === 0) return prev;

            let newStart = start - shift;
            let newEnd = end - shift;

            if (newStart < 0) {
              newEnd -= newStart; 
              newStart = 0;
            }
            if (newEnd > rawData.length - 1) {
              newStart -= (newEnd - (rawData.length - 1));
              newEnd = rawData.length - 1;
            }

            newStart = Math.max(0, newStart);
            newEnd = Math.min(rawData.length - 1, newEnd);

            return [newStart, newEnd];
          });

          setYDomain(prev => {
             if (deltaYDrag === 0) return prev;
             let min = prev[0];
             let max = prev[1];
             if (min === 'auto') {
                 min = Infinity; max = -Infinity;
                 chartDataRef.current.forEach(d => {
                     chartSelectedIdsRef.current.forEach(id => {
                         const val = d[`CH${id}`];
                         if (val !== undefined) {
                             if (val < min) min = val;
                             if (val > max) max = val;
                         }
                     });
                 });
                 if (min === Infinity) return prev;
                 const pad = (max - min) * 0.05 || 5;
                 min = min - pad;
                 max = max + pad;
             }
             
             const currentLen = max - min;
             const rect = container.getBoundingClientRect();
             const pointsPerPixel = currentLen / rect.height;
             
             const shift = deltaYDrag * pointsPerPixel;
             return [min + shift, max + shift];
          });
        }, 40);
      }
    };

    const handleMouseUp = () => {
      if (isDragging.current) {
        isDragging.current = false;
        container.style.cursor = 'default';
      }
    };

    container.addEventListener('wheel', handleWheel, { passive: false });
    container.addEventListener('mousedown', handleMouseDown);
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    
    return () => {
      container.removeEventListener('wheel', handleWheel);
      container.removeEventListener('mousedown', handleMouseDown);
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [rawData.length, activeTab]);

  const resetZoom = () => {
    if (rawData.length > 0) {
      setZoomRange([0, rawData.length - 1]);
      setYDomain(['auto', 'auto']);
    }
  };

  const handleConfigChange = (e) => {
    setConfig({ ...config, [e.target.name]: e.target.value });
  };



  const handleChannelChange = (id, field, value) => {
    setChannels(channels.map(ch => ch.id === id ? { ...ch, [field]: value } : ch));
  };

  const toggleSelection = (id) => {
    const newSel = new Set(selectedIds);
    if (newSel.has(id)) newSel.delete(id);
    else newSel.add(id);
    setSelectedIds(newSel);
  };

  const toggleChartSelection = (id) => {
    const newSel = new Set(chartSelectedIds);
    if (newSel.has(id)) newSel.delete(id);
    else newSel.add(id);
    setChartSelectedIds(newSel);
  };

  const toggleSelectAll = () => {
    if (selectedIds.size === channels.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(channels.map(c => c.id)));
    }
  };

  const toggleAllChartSelection = () => {
    if (chartSelectedIds.size === channels.length) {
      setChartSelectedIds(new Set());
    } else {
      setChartSelectedIds(new Set(channels.map(c => c.id)));
    }
  };

  const deleteSelected = () => {
    if (selectedIds.size === 0) return;
    setConfirmModal({
      show: true,
      message: `确定要删除选中的 ${selectedIds.size} 个通道吗？`,
      onConfirm: () => {
        setChannels(channels.filter(ch => !selectedIds.has(ch.id)));
        setSelectedIds(new Set());
        setConfirmModal({ show: false, message: '', onConfirm: null });
      }
    });
  };

  const clearDeviceNames = () => {
    setConfirmModal({
      show: true,
      message: '确定要清除所有通道的器件名称吗？',
      onConfirm: () => {
        setChannels(prev => prev.map(ch => ({ ...ch, name: '' })));
        setConfirmModal({ show: false, message: '', onConfirm: null });
        showToast('器件名称已清除');
      }
    });
  };

  const calculateMaxTemps = (currentChannels = channels) => {
    const validChannels = currentChannels.filter(ch => ch.temp !== '' && !isNaN(parseFloat(ch.temp)));
    const sorted = validChannels.sort((a, b) => parseFloat(b.temp) - parseFloat(a.temp));
    if (sorted.length === 0) return '';
    const top3 = sorted.slice(0, 3);
    const text = top3.map(ch => `${ch.name ? `CH${ch.id}(${ch.name})` : `CH${ch.id}`} ${ch.temp}℃`).join('，');
    return `温升测试达到稳定状态后，温度最高为：${text}`;
  };

  useEffect(() => {
    setConfig(prevConfig => {
      if (!prevConfig.bottomNote || !prevConfig.bottomNote.startsWith('温升测试达到稳定状态后')) {
        return prevConfig;
      }
      const newNote = calculateMaxTemps(channels);
      if (newNote && prevConfig.bottomNote !== newNote) {
        return { ...prevConfig, bottomNote: newNote };
      }
      return prevConfig;
    });
  }, [channels]);

  const handleFileUpload = async (e, isCompare = false) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setIsImporting(true);
    if (isCompare) {
      setCompareFileName(file.name);
    } else {
      setImportedFileName(file.name);
    }
    try {
      const workbook = new ExcelJS.Workbook();
      let sheetToParse = null;
      let summarySheet = null;

      if (file.name.toLowerCase().endsWith('.csv')) {
        const text = await file.text();
        sheetToParse = workbook.addWorksheet('原始数据');
        
        await new Promise((resolve, reject) => {
          Papa.parse(text, {
            skipEmptyLines: true,
            complete: (results) => {
              results.data.forEach(row => {
                sheetToParse.addRow(row);
              });
              resolve();
            },
            error: (error) => {
              reject(error);
            }
          });
        });
      } else {
        const buffer = await file.arrayBuffer();
        await workbook.xlsx.load(buffer);
        
        let dataSheet = workbook.worksheets.find(s => s.name === '原始数据');
        let foundSummary = workbook.worksheets.find(s => s.name === '温升数据');

        if (!dataSheet) {
          dataSheet = workbook.worksheets.find(sheet => {
            let hasChannel = false;
            sheet.getRow(1).eachCell(cell => {
              if (cell.value === 'FullTime_Hidden' || (typeof cell.value === 'string' && cell.value.toUpperCase().startsWith('CH'))) {
                hasChannel = true;
              }
            });
            return hasChannel;
          });
        }

        if (!foundSummary) {
          foundSummary = workbook.worksheets.find(sheet => {
            const a2 = sheet.getCell('A2').text || '';
            const a5 = sheet.getCell('A5').text || '';
            const a7 = sheet.getCell('A7').text || '';
            return a2 === '测试条件' || a5.includes('备注') || a7 === 'GROUP 1';
          });
        }

        if (!dataSheet && workbook.worksheets.length >= 2) {
            foundSummary = foundSummary || workbook.worksheets[0];
            dataSheet = workbook.worksheets[1];
        }

        summarySheet = foundSummary;
        sheetToParse = dataSheet || workbook.worksheets[0];
      }
      
      if (!sheetToParse) throw new Error("No sheet found");

      const headerRow = sheetToParse.getRow(1);
      const colMap = {}; 
      const parsedChannelNames = {};
      let fullTimeCol = -1;
      
      headerRow.eachCell((cell, colNumber) => {
        if (cell.value === 'FullTime_Hidden') {
          fullTimeCol = colNumber;
        } else if (cell.value && typeof cell.value === 'string' && cell.value.toUpperCase().startsWith('CH')) {
          const match = cell.value.match(/CH(\d+)(?:\s+(.*))?/i);
          if (match) {
            const chId = parseInt(match[1], 10);
            if (!isNaN(chId)) {
              colMap[chId] = colNumber;
              if (match[2]) {
                parsedChannelNames[chId] = match[2].trim();
              }
            }
          }
        }
      });

      const maxTemps = {};
      const stableTemps = {};
      const newRawData = [];
      let dataPointCounter = 1;
      
      sheetToParse.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        
        let timeStr = '';
        let originalTimeStr = '';
        if (fullTimeCol !== -1) {
           const ftCell = row.getCell(fullTimeCol);
           timeStr = ftCell.type === ExcelJS.ValueType.Date ? ftCell.value.toLocaleTimeString() : ftCell.text;
           originalTimeStr = ftCell.type === ExcelJS.ValueType.Date ? ftCell.value.toLocaleString('zh-CN', { hour12: false }) : ftCell.text;
        }
        
        if (!timeStr) {
           const timeCell = row.getCell(1);
           originalTimeStr = timeCell.type === ExcelJS.ValueType.Date ? timeCell.value.toLocaleString('zh-CN', { hour12: false }) : timeCell.text;
           if (timeCell.type === ExcelJS.ValueType.Date) {
              timeStr = timeCell.value.toLocaleTimeString();
           } else {
              timeStr = timeCell.text;
              if (timeStr && timeStr.includes(' ')) {
                 timeStr = timeStr.split(' ').pop();
              }
           }
        }

        const dataPoint = { 
          time: timeStr || `P${dataPointCounter}`,
          originalTime: originalTimeStr || timeStr || `P${dataPointCounter}`
        };
        let hasValidData = false;
        
        for (const [chId, colNum] of Object.entries(colMap)) {
          const cell = row.getCell(colNum);
          const val = parseFloat(cell.value);
          // Ignore values <= 0 (unconnected/abnormal) and >= 1000 (like 1999)
          if (!isNaN(val) && val > 0 && val < 1000) {
            dataPoint[`CH${chId}`] = val;
            hasValidData = true;
            if (!(chId in maxTemps) || val > maxTemps[chId]) {
              maxTemps[chId] = val;
            }
            stableTemps[chId] = val;
          }
        }
        
        if (hasValidData) {
          newRawData.push(dataPoint);
          dataPointCounter++;
        }
      });

      // Post-process to interpolate missing times (for legacy exported files without FullTime_Hidden)
      if (fullTimeCol === -1) {
        const parseTime = (str) => {
          const parts = str.split(':');
          if (parts.length === 3) return parseInt(parts[0], 10)*3600 + parseInt(parts[1], 10)*60 + parseInt(parts[2], 10);
          if (parts.length === 2) return parseInt(parts[0], 10)*3600 + parseInt(parts[1], 10)*60;
          return NaN;
        };
        const formatTime = (s) => {
          const h = Math.floor(s/3600);
          const m = Math.floor((s%3600)/60);
          const sec = Math.floor(s%60);
          return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${sec.toString().padStart(2,'0')}`;
        };

        let lastValidIdx = -1;
        for (let i = 0; i < newRawData.length; i++) {
           const t = newRawData[i].time;
           if (t && !t.startsWith('P')) {
              lastValidIdx = i;
           } else if (t.startsWith('P')) {
              let nextValidIdx = -1;
              for (let j = i + 1; j < newRawData.length; j++) {
                 if (newRawData[j].time && !newRawData[j].time.startsWith('P')) {
                    nextValidIdx = j;
                    break;
                 }
              }
              if (lastValidIdx !== -1 && nextValidIdx !== -1) {
                 const t1 = parseTime(newRawData[lastValidIdx].time);
                 const t2 = parseTime(newRawData[nextValidIdx].time);
                 if (!isNaN(t1) && !isNaN(t2)) {
                    const fraction = (i - lastValidIdx) / (nextValidIdx - lastValidIdx);
                    const interpolated = t1 + (t2 - t1) * fraction;
                    newRawData[i].time = formatTime(interpolated);
                 } else {
                    newRawData[i].time = newRawData[lastValidIdx].time;
                 }
              } else if (lastValidIdx !== -1) {
                 newRawData[i].time = newRawData[lastValidIdx].time;
              }
           }
        }
      }

      // Post-process duplicate times (evenly space points recorded within the same minute)
      let timeIdx = 0;
      while (timeIdx < newRawData.length) {
          let j = timeIdx;
          while (j < newRawData.length && newRawData[j].time === newRawData[timeIdx].time) {
              j++;
          }
          const count = j - timeIdx;
          if (count > 1) {
              const baseTimeStr = newRawData[timeIdx].time;
              let h = 0, m = 0, s = 0;
              const parts = baseTimeStr.split(':');
              if (parts.length === 2) {
                  h = parseInt(parts[0], 10);
                  m = parseInt(parts[1], 10);
              } else if (parts.length === 3) {
                  h = parseInt(parts[0], 10);
                  m = parseInt(parts[1], 10);
                  s = parseInt(parts[2], 10);
              }
              
              if (!isNaN(h) && !isNaN(m)) {
                  const step = Math.floor(60 / count);
                  for (let k = 0; k < count; k++) {
                      const newS = s + k * step;
                      const finalS = newS % 60;
                      const finalM = m + Math.floor(newS / 60);
                      const finalH = h + Math.floor(finalM / 60);
                      
                      const formatNum = (num) => num.toString().padStart(2, '0');
                      newRawData[timeIdx + k].time = `${formatNum(finalH % 24)}:${formatNum(finalM % 60)}:${formatNum(finalS)}`;
                  }
              }
          }
          timeIdx = j;
      }


      // Recover Channels
      const existingMap = new Map(channels.map(c => [c.id, c]));

      let finalChannels = [];
      for (const [chIdStr, maxTemp] of Object.entries(maxTemps)) {
        const chId = parseInt(chIdStr, 10);
        const existing = existingMap.get(chId);
        let name = parsedChannelNames[chId] || '';
        if (!name && existing && existing.name) {
          name = existing.name;
        }
        finalChannels.push({ 
          id: chId, 
          name, 
          temp: stableTemps[chId] ? stableTemps[chId].toFixed(2) : '', 
          maxTemp: maxTemp.toFixed(2) 
        });
      }
      finalChannels.sort((a, b) => a.id - b.id);
      
      // Recover Config if exported file
      if (summarySheet) {
        let parsedGroupNotes = {};
        let parsedBottomNote = '';
        let parsedBottomNoteLabel = '底部总结备注';
        let row = 7;
        
        while (row <= 200) { // Limit to avoid infinite loop
          const groupLabel = summarySheet.getCell(`A${row}`).text || '';
          if (groupLabel.startsWith('GROUP')) {
            const gIndex = parseInt(groupLabel.replace('GROUP', '').trim(), 10) - 1;
            
            // Read channel names
            let col = 4; // Column D
            while (col <= 20) {
               const chIdText = summarySheet.getCell(row, col).text || '';
               if (chIdText.startsWith('CH')) {
                 const chId = parseInt(chIdText.replace('CH', '').trim(), 10);
                 const chName = summarySheet.getCell(row + 1, col).text || '';
                 if (chName) {
                   const channel = finalChannels.find(c => c.id === chId);
                   if (channel) channel.name = chName;
                 }
               } else {
                 break;
               }
               col++;
            }
            
            // Read group note by searching next few rows
            let noteRow = row + 1;
            while (noteRow <= row + 5) {
                if (summarySheet.getCell(`A${noteRow}`).text === '备注') {
                    const gNote = summarySheet.getCell(noteRow, 4).text || '';
                    if (gNote) {
                      parsedGroupNotes[gIndex] = gNote;
                    }
                    break;
                }
                noteRow++;
            }
          } else if (groupLabel && !groupLabel.startsWith('备注')) {
            // It might be the bottom note label
            parsedBottomNoteLabel = groupLabel;
            parsedBottomNote = summarySheet.getCell(`D${row}`).text || '';
            // Don't break, keep scanning just in case
          }
          row++;
        }

        setGroupNotes(parsedGroupNotes);

        setConfig(prev => ({
          ...prev,
          condition: summarySheet.getCell('A2').text || prev.condition,
          pv: summarySheet.getCell('D2').text || prev.pv,
          pvVoltage: summarySheet.getCell('F2').text || prev.pvVoltage,
          pvCurrent: summarySheet.getCell('H2').text || prev.pvCurrent,
          power: summarySheet.getCell('D3').text || prev.power,
          voltageVal: summarySheet.getCell('F3').text || prev.voltageVal,
          currentVal: summarySheet.getCell('H3').text || prev.currentVal,
          noteLabel: summarySheet.getCell('A5').text || prev.noteLabel || '备注',
          topNote: summarySheet.getCell('D5').text || prev.topNote,
          bottomNoteLabel: parsedBottomNoteLabel || prev.bottomNoteLabel,
          bottomNote: parsedBottomNote || calculateMaxTemps(finalChannels)
        }));
      } else {
        setConfig(prev => ({ ...prev, bottomNote: calculateMaxTemps(finalChannels) }));
      }

      if (isCompare) {
        setCompareRawData(newRawData);
        setCompareFullRawData(newRawData);
        setIsComparing(true);
        setChannels(prev => prev.map(ch => {
          const compCh = finalChannels.find(c => c.id === ch.id);
          if (compCh) {
            return {
              ...ch,
              compareTemp: compCh.temp,
              compareMaxTemp: compCh.maxTemp
            };
          }
          return ch;
        }));
        showToast(`对比数据成功导入`, 'success');
      } else {
        setChannels(finalChannels);
        setChartSelectedIds(new Set(finalChannels.map(c => c.id)));
        setRawData(newRawData);
        setFullRawData(newRawData);
        
        // Clear compare data when main data changes
        setCompareRawData([]);
        setCompareFullRawData([]);
        setIsComparing(false);
        setCompareFileName('');
        
        showToast(`数据导入成功！共识别到 ${Object.keys(maxTemps).length} 个有效通道。`, 'success');
      }
      console.error(err);
      showToast('导入失败，请确保文件格式正确。', 'error');
    } finally {
      setIsImporting(false);
      e.target.value = '';
    }
  };

  const clearData = () => {
    setConfirmModal({
      show: true,
      message: '确定要清除所有导入的数据吗？这将清空当前表格和图表数据。',
      onConfirm: () => {
        setChannels([]);
        setRawData([]);
        setFullRawData([]);
        setCompareRawData([]);
        setCompareFullRawData([]);
        setIsComparing(false);
        setCompareFileName('');
        setChartSelectedIds(new Set());
        setSelectedIds(new Set());
        setImportedFileName('');
        setGroupNotes({});
        setConfig(prev => ({ 
          ...prev, 
          power: '--',
          voltageVal: '--',
          currentVal: '--',
          topNote: '--',
          bottomNote: '' 
        }));
        setConfirmModal({ show: false, message: '', onConfirm: null });
      }
    });
  };

  const exportTemplate = (providedBaseName) => {
    let baseName = typeof providedBaseName === 'string' ? providedBaseName : '';
    if (!baseName) {
      if (importedFileName) {
        baseName = importedFileName.replace(/\.[^/.]+$/, "") + '_整理后';
      } else {
        const dateStr = new Date().toLocaleString('zh-CN', { hour12: false }).replace(/[\/\s:]/g, '');
        baseName = `温升数据报告_${dateStr}`;
      }
    }
    const templateData = {
      config,
      groupNotes,
      channels: channels.map(ch => ({ id: ch.id, name: ch.name }))
    };
    const blob = new Blob([JSON.stringify(templateData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${baseName}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('模板已导出');
  };

  const importTemplate = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = JSON.parse(event.target.result);
        if (data.config) setConfig(data.config);
        if (data.groupNotes) setGroupNotes(data.groupNotes);
        if (data.channels && Array.isArray(data.channels)) {
          const nameMap = {};
          data.channels.forEach(ch => { if (ch.name) nameMap[ch.id] = ch.name; });
          
          setChannels(prevChannels => prevChannels.map(ch => ({
            ...ch,
            name: nameMap[ch.id] || ch.name
          })));
          
          // Also save to localStorage cache so auto-fill works for future imports
          localStorage.setItem('temp_rise_channel_names', JSON.stringify(nameMap));
        }
        showToast('模板导入成功，已应用配置和名称', 'success');
      } catch (err) {
        console.error(err);
        showToast('模板解析失败，请确认文件格式', 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const changeExportPath = async () => {
    try {
      if (window.electronAPI) {
        const pathStr = await window.electronAPI.selectDirectory();
        if (pathStr) {
          await set('electron_export_dir', pathStr);
          setExportDirHandle(null);
          setExportDirName(pathStr);
        }
        return;
      }
      
      if (!window.showDirectoryPicker) {
        alert('您的浏览器不支持直接选择文件夹，请在导出时手动选择路径。');
        return;
      }
      const handle = await window.showDirectoryPicker();
      await set('export_dir_handle', handle);
      setExportDirHandle(handle);
      setExportDirName(handle.name);
    } catch (err) {
      if (err.name !== 'AbortError') {
        alert('无法选择文件夹: ' + err.message);
      }
    }
  };

  const exportExcel = async () => {
    try {
      setIsExporting(true);
      let exportBaseName = '';
      if (importedFileName) {
        exportBaseName = importedFileName.replace(/\.[^/.]+$/, "") + '_整理后';
      } else {
        const dateStr = new Date().toLocaleString('zh-CN', { hour12: false }).replace(/[\/\s:]/g, '');
        exportBaseName = `温升数据报告_${dateStr}`;
      }
      const dataLen = rawData ? rawData.length : 0;
      let templateName = isComparing ? './chart_template_compare_200.xlsx' : './chart_template_200.xlsx';
      if (dataLen > 3000) templateName = isComparing ? './chart_template_compare_5000.xlsx' : './chart_template_5000.xlsx';
      else if (dataLen > 2000) templateName = isComparing ? './chart_template_compare_3000.xlsx' : './chart_template_3000.xlsx';
      else if (dataLen > 1000) templateName = isComparing ? './chart_template_compare_2000.xlsx' : './chart_template_2000.xlsx';
      else if (dataLen > 500) templateName = isComparing ? './chart_template_compare_1000.xlsx' : './chart_template_1000.xlsx';
      else if (dataLen > 200) templateName = isComparing ? './chart_template_compare_500.xlsx' : './chart_template_500.xlsx';
      
      const response = await fetch(templateName);
      if (!response.ok) throw new Error('无法加载模板文件');
      let arrayBuffer = await response.arrayBuffer();

      const maxRowsMatch = templateName.match(/_(\d+)\.xlsx/);
      const maxRows = maxRowsMatch ? parseInt(maxRowsMatch[1]) : 200;

            // --- DYNAMICALLY REBUILD CHART SERIES USING JSZip ---
      try {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const chartFile = zip.file('xl/charts/chart1.xml');
        if (chartFile) {
          let xml = await chartFile.async('string');
          
          function getColName(colIndex) {
              let temp, letter = '';
              while (colIndex > 0) {
                  temp = (colIndex - 1) % 26;
                  letter = String.fromCharCode(temp + 65) + letter;
                  colIndex = (colIndex - temp - 1) / 26;
              }
              return letter;
          }

          const prefixMatch = xml.match(/<(c:)?ser>/);
          const p = prefixMatch && prefixMatch[1] ? 'c:' : '';

          let serNodes = '';
          channels.forEach((ch, idx) => {
              const colLetter = getColName(idx + 2);
              serNodes += `<${p}ser><${p}idx val="${idx}"/><${p}order val="${idx}"/><${p}tx><${p}strRef><${p}f>原始数据!$${colLetter}$1</${p}f><${p}strCache><${p}ptCount val="1"/><${p}pt idx="0"><${p}v>CH${ch.id}</${p}v></${p}pt></${p}strCache></${p}strRef></${p}tx><${p}spPr><a:ln w="28575"/></${p}spPr><${p}marker><${p}symbol val="none"/></${p}marker><${p}cat><${p}numRef><${p}f>原始数据!$A$2:$A$${maxRows}</${p}f><${p}numCache><${p}formatCode>General</${p}formatCode><${p}ptCount val="${maxRows-1}"/></${p}numCache></${p}numRef></${p}cat><${p}val><${p}numRef><${p}f>原始数据!$${colLetter}$2:$${colLetter}$${maxRows}</${p}f><${p}numCache><${p}formatCode>General</${p}formatCode><${p}ptCount val="${maxRows-1}"/></${p}numCache></${p}numRef></${p}val><${p}smooth val="1"/></${p}ser>`;
          });

          xml = xml.replace(new RegExp(`<${p}ser>[\s\S]*?<\/${p}ser>`, 'g'), '');
          xml = xml.replace(new RegExp(`(<${p}marker val="1"\/>|<${p}axId)`), (match) => serNodes + match);
          zip.file('xl/charts/chart1.xml', xml);
          
          if (isComparing) {
            const chartFile2 = zip.file('xl/charts/chart2.xml');
            if (chartFile2) {
              let xml2 = await chartFile2.async('string');
              const prefixMatch2 = xml2.match(/<(c:)?ser>/);
              const p2 = prefixMatch2 && prefixMatch2[1] ? 'c:' : '';
              let serNodes2 = '';
              channels.forEach((ch, idx) => {
                  const colLetter = getColName(idx + 203);
                  serNodes2 += `<${p2}ser><${p2}idx val="${idx}"/><${p2}order val="${idx}"/><${p2}tx><${p2}strRef><${p2}f>原始数据!$${colLetter}$1</${p2}f><${p2}strCache><${p2}ptCount val="1"/><${p2}pt idx="0"><${p2}v>CH${ch.id}</${p2}v></${p2}pt></${p2}strCache></${p2}strRef></${p2}tx><${p2}spPr><a:ln w="28575"/></${p2}spPr><${p2}marker><${p2}symbol val="none"/></${p2}marker><${p2}cat><${p2}numRef><${p2}f>原始数据!$A$2:$A$${maxRows}</${p2}f><${p2}numCache><${p2}formatCode>General</${p2}formatCode><${p2}ptCount val="${maxRows-1}"/></${p2}numCache></${p2}numRef></${p2}cat><${p2}val><${p2}numRef><${p2}f>原始数据!$${colLetter}$2:$${colLetter}$${maxRows}</${p2}f><${p2}numCache><${p2}formatCode>General</${p2}formatCode><${p2}ptCount val="${maxRows-1}"/></${p2}numCache></${p2}numRef></${p2}val><${p2}smooth val="1"/></${p2}ser>`;
              });
              xml2 = xml2.replace(new RegExp(`<${p2}ser>[\s\S]*?<\/${p2}ser>`, 'g'), '');
              xml2 = xml2.replace(new RegExp(`(<${p2}marker val="1"\/>|<${p2}axId)`), (match) => serNodes2 + match);
              zip.file('xl/charts/chart2.xml', xml2);
            }
          }

          arrayBuffer = await zip.generateAsync({ type: 'arraybuffer' });
        }
      } catch (err) {
        console.error("Failed to dynamically rebuild chart xml:", err);
      }
      // ----------------------------------------------------

      const workbook = await XlsxPopulate.fromDataAsync(arrayBuffer);
      
      // --- 1. Populate Raw Data Sheet (For the Chart) ---
      const dataSheet = workbook.sheet('原始数据');
      
      // Clear out old headers that might be pre-filled in the template
      for (let i = 1; i <= 200; i++) {
         dataSheet.cell(1, i + 1).value(undefined);
      }

      // Update chart legend with actual channel names contiguously
      channels.forEach((ch, idx) => {
        const name = ch.name ? `CH${ch.id} ${ch.name}` : `CH${ch.id}`;
        dataSheet.cell(1, idx + 2).value(name);
        if (isComparing) {
            dataSheet.cell(1, idx + 203).value(name);
        }
      });

      if (rawData && rawData.length > 0) {
        rawData.forEach((data, index) => {
          const row = index + 2;
          
          // Write the unmutated, original timestamp straight from the imported file
          dataSheet.cell(row, 1).value(data.originalTime || data.time);

          // Write actual channel values contiguously matching the rebuilt XML
          channels.forEach((ch, idx) => {
            const val = data[`CH${ch.id}`];
            if (val !== undefined) {
              dataSheet.cell(row, idx + 2).value(val);
            }
            if (isComparing) {
               const cpVal = data[`compare_CH${ch.id}`];
               if (cpVal !== undefined) {
                   dataSheet.cell(row, idx + 203).value(cpVal);
               }
            }
          });
        });
        
        // Hide all remaining unused rows so the chart doesn't render empty gaps at the end
        const maxRowsMatch = templateName.match(/_(\d+)\.xlsx/);
        const maxRows = maxRowsMatch ? parseInt(maxRowsMatch[1]) : 0;
        if (maxRows > 0) {
          for (let row = dataLen + 2; row <= maxRows; row++) {
            dataSheet.row(row).hidden(true);
          }
        }
      }

      // --- 2. Populate Summary Sheet ---
      const sheet = workbook.sheet('温升数据');

      const setBorderStyle = (range) => {
        range.style('border', true);
      };

      for(let i = 1; i <= 200; i++) {
        sheet.row(i).height(26);
      }
      
      for(let i = 1; i <= 20; i++) {
        sheet.column(i).style('fontFamily', 'Microsoft YaHei').style('fontSize', 11);
      }

      sheet.column('A').width(12);
      sheet.column('B').width(12);
      sheet.column('C').width(18);
      for(let i = 4; i <= 20; i++) {
        sheet.column(i).width(16);
      }

      sheet.range('D1:I1').merged(true);
      const titleCell = sheet.cell('D1');
      titleCell.value('原始数据');
      titleCell.style({ horizontalAlignment: 'center', verticalAlignment: 'center', fontSize: 16, bold: true });
      setBorderStyle(sheet.range('D1:I1'));

      const applyStyle = (cellRange, fieldName) => {
        const style = config.styles?.[fieldName] || DEFAULT_CONFIG.styles[fieldName] || { fontSize: 11, bold: false };
        cellRange.style('fontSize', style.fontSize);
        cellRange.style('bold', style.bold);
      };

      sheet.range('A2:C4').merged(true);
      const condCell = sheet.cell('A2');
      condCell.value(config.condition);
      condCell.style({ horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true });
      applyStyle(condCell, 'condition');
      setBorderStyle(sheet.range('A2:C4'));

      sheet.range('D2:E2').merged(true).value(config.pv);
      applyStyle(sheet.range('D2:E2'), 'pv');
      
      sheet.range('F2:G2').merged(true).value(config.pvVoltage);
      applyStyle(sheet.range('F2:G2'), 'pvVoltage');
      
      sheet.range('H2:I2').merged(true).value(config.pvCurrent);
      applyStyle(sheet.range('H2:I2'), 'pvCurrent');

      sheet.range('D3:E4').merged(true).value(config.power);
      applyStyle(sheet.range('D3:E4'), 'power');
      
      sheet.range('F3:G4').merged(true).value(config.voltageVal);
      applyStyle(sheet.range('F3:G4'), 'voltageVal');
      
      sheet.range('H3:I4').merged(true).value(config.currentVal);
      applyStyle(sheet.range('H3:I4'), 'currentVal');

      const pvRange = sheet.range('D2:I4');
      setBorderStyle(pvRange);
      pvRange.style({ horizontalAlignment: 'center', verticalAlignment: 'center' });
      sheet.range('D2:I2').style('fill', 'F2F2F2');

      sheet.range('A5:C5').merged(true);
      const noteLabelCell = sheet.cell('A5');
      noteLabelCell.value(config.noteLabel || '备注').style({ horizontalAlignment: 'center', verticalAlignment: 'center' });
      applyStyle(noteLabelCell, 'noteLabel');
      setBorderStyle(sheet.range('A5:C5'));
      
      sheet.range('D5:K5').merged(true);
      const topNoteCell = sheet.cell('D5');
      topNoteCell.value(config.topNote).style({ horizontalAlignment: 'left', verticalAlignment: 'center' });
      applyStyle(topNoteCell, 'topNote');
      setBorderStyle(sheet.range('D5:K5'));

      const getTempColorHex = (val, min, max) => {
        if (isNaN(val)) return 'FFFFFF';
        const green = [99, 190, 123];
        const yellow = [255, 235, 132];
        const red = [248, 105, 107];
        if (max === min) return 'FFEB84'; 
        const mid = (min + max) / 2;
        let r, g, b;
        if (val <= mid) {
          const ratio = Math.max(0, (val - min) / (mid - min));
          r = Math.round(green[0] + ratio * (yellow[0] - green[0]));
          g = Math.round(green[1] + ratio * (yellow[1] - green[1]));
          b = Math.round(green[2] + ratio * (yellow[2] - green[2]));
        } else {
          const ratio = Math.min(1, (val - mid) / (max - mid));
          r = Math.round(yellow[0] + ratio * (red[0] - yellow[0]));
          g = Math.round(yellow[1] + ratio * (red[1] - yellow[1]));
          b = Math.round(yellow[2] + ratio * (red[2] - yellow[2]));
        }
        const toHex = (n) => n.toString(16).padStart(2, '0').toUpperCase();
        return `${toHex(r)}${toHex(g)}${toHex(b)}`;
      };

      const validTemps = channels.flatMap(ch => [parseFloat(ch.temp), parseFloat(ch.maxTemp), parseFloat(ch.compareTemp), parseFloat(ch.compareMaxTemp)]).filter(val => !isNaN(val));
      const globalMin = validTemps.length > 0 ? Math.min(...validTemps) : 0;
      const globalMax = validTemps.length > 0 ? Math.max(...validTemps) : 100;

            let startRow = 7;
      chunkedChannels.forEach((groupChannels, gIndex) => {
        let rowsCount = showMaxTemp ? 4 : 3;
        if (isComparing) {
            rowsCount += showMaxTemp ? 2 : 1;
        }
        sheet.range(`A${startRow}:A${startRow + rowsCount - 1}`).merged(true);
        sheet.cell(`A${startRow}`).value(`GROUP ${gIndex + 1}`).style({ horizontalAlignment: 'center', verticalAlignment: 'center' });
        
        sheet.cell(`B${startRow}`).value('室温');
        sheet.cell(`C${startRow}`).value('采集器通道');
        
        groupChannels.forEach((ch, idx) => {
          sheet.cell(startRow, 4 + idx).value(`CH${ch.id}`);
        });

        sheet.cell(`C${startRow + 1}`).value('器件');
        groupChannels.forEach((ch, idx) => {
          sheet.cell(startRow + 1, 4 + idx).value(ch.name);
        });

        let currentRow = startRow + 2;
        sheet.cell(`C${currentRow}`).value('稳定温度/℃');
        groupChannels.forEach((ch, idx) => {
          const val = ch.temp ? parseFloat(ch.temp) : '';
          const cell = sheet.cell(currentRow, 4 + idx);
          cell.value(val);
          if (val !== '') {
              const hexColor = getTempColorHex(val, globalMin, globalMax);
              cell.style('fill', hexColor);
          }
        });
        currentRow++;

        if (isComparing) {
          sheet.cell(`C${currentRow}`).value('对比稳定温度/℃');
          groupChannels.forEach((ch, idx) => {
            const val = ch.compareTemp ? parseFloat(ch.compareTemp) : '';
            const cell = sheet.cell(currentRow, 4 + idx);
            cell.value(val);
            if (val !== '') {
                const hexColor = getTempColorHex(val, globalMin, globalMax);
                cell.style('fill', hexColor);
            }
          });
          currentRow++;
        }

        if (showMaxTemp) {
          sheet.cell(`C${currentRow}`).value('最高温度/℃');
          groupChannels.forEach((ch, idx) => {
            const val = ch.maxTemp ? parseFloat(ch.maxTemp) : '';
            const cell = sheet.cell(currentRow, 4 + idx);
            cell.value(val);
            if (val !== '') {
                const hexColor = getTempColorHex(val, globalMin, globalMax);
                cell.style('fill', hexColor);
            }
          });
          currentRow++;

          if (isComparing) {
            sheet.cell(`C${currentRow}`).value('对比最高温度/℃');
            groupChannels.forEach((ch, idx) => {
              const val = ch.compareMaxTemp ? parseFloat(ch.compareMaxTemp) : '';
              const cell = sheet.cell(currentRow, 4 + idx);
              cell.value(val);
              if (val !== '') {
                  const hexColor = getTempColorHex(val, globalMin, globalMax);
                  cell.style('fill', hexColor);
              }
            });
            currentRow++;
          }
        }


        const noteRow = startRow + rowsCount;
        sheet.range(`A${noteRow}:C${noteRow}`).merged(true);
        sheet.cell(`A${noteRow}`).value('备注').style({ horizontalAlignment: 'center', verticalAlignment: 'center' });
        
        const colsToMerge = Math.max(8, groupChannels.length);
        sheet.range(noteRow, 4, noteRow, 3 + colsToMerge).merged(true);
        sheet.cell(noteRow, 4).value(groupNotes[gIndex] || '').style({ horizontalAlignment: 'left', verticalAlignment: 'center' });

        const groupRange = sheet.range(startRow, 1, noteRow, 3 + colsToMerge);
        setBorderStyle(groupRange);
        groupRange.style({ horizontalAlignment: 'center', verticalAlignment: 'center' });
        sheet.cell(noteRow, 4).style('horizontalAlignment', 'left');

        sheet.range(startRow, 4, startRow, 3 + colsToMerge).style('fill', 'F2F2F2');
        sheet.range(startRow, 3, startRow + rowsCount - 1, 3).style('fill', 'F2F2F2');
        sheet.cell(startRow, 2).style('fill', 'F2F2F2');
        
        startRow = noteRow + 2;
      });

      sheet.range(`A${startRow}:C${startRow}`).merged(true);
      const bottomNoteLabelCell = sheet.cell(`A${startRow}`);
      bottomNoteLabelCell.value(config.bottomNoteLabel || '底部总结备注').style({ horizontalAlignment: 'center', verticalAlignment: 'center' });
      applyStyle(bottomNoteLabelCell, 'bottomNoteLabel');
      setBorderStyle(sheet.range(`A${startRow}:C${startRow}`));

      sheet.range(`D${startRow}:K${startRow}`).merged(true);
      const noteText = sheet.cell(`D${startRow}`);
      noteText.value(config.bottomNote || '').style({ horizontalAlignment: 'left', verticalAlignment: 'center', fontColor: 'FF0000' });
      applyStyle(noteText, 'bottomNote');
      setBorderStyle(sheet.range(`D${startRow}:K${startRow}`));

      const blob = await workbook.outputAsync();
      
      try {
        if (window.electronAPI && exportDirName) {
           let buffer;
           if (blob instanceof Blob) {
               buffer = await blob.arrayBuffer();
           } else {
               buffer = blob;
           }
           const fileName = `${exportBaseName}.xlsx`;
           const savedPath = await window.electronAPI.saveFileToPath(exportDirName, fileName, new Uint8Array(buffer));
           
           let extraMsg = '';
           if (autoExportTemplate) {
                const hasAnyName = channels.some(ch => ch.name && ch.name.trim() !== '');
                if (hasAnyName) {
                    try {
                        const templateFileName = `${exportBaseName}.json`;
                        const templateData = { config, groupNotes, channels: channels.map(ch => ({ id: ch.id, name: ch.name })) };
                        const tBuffer = new TextEncoder().encode(JSON.stringify(templateData, null, 2));
                        await window.electronAPI.saveFileToPath(exportDirName, templateFileName, tBuffer);
                        extraMsg = `\n(已同步导出配置模板: ${templateFileName})`;
                    } catch (e) {
                        console.error('静默导出模板失败', e);
                    }
                }
           }
           
           setExportSuccessModal({
             show: true,
             message: `已成功保存到文件夹:\n${exportDirName}/${fileName}${extraMsg}`,
             blob: blob,
             fileName: fileName,
             savedPath: savedPath
           });
           return;
        } else if (exportDirHandle) {
          // Verify permission
          const options = { mode: 'readwrite' };
          let permission = await exportDirHandle.queryPermission(options);
          if (permission !== 'granted') {
            permission = await exportDirHandle.requestPermission(options);
          }
          
          if (permission === 'granted') {
            const fileName = `${exportBaseName}.xlsx`;
            const fileHandle = await exportDirHandle.getFileHandle(fileName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(blob);
            await writable.close();
            
            let extraMsg = '';
            if (autoExportTemplate) {
                const hasAnyName = channels.some(ch => ch.name && ch.name.trim() !== '');
                if (hasAnyName) {
                    try {
                        const templateFileName = `${exportBaseName}.json`;
                        const templateData = { config, groupNotes, channels: channels.map(ch => ({ id: ch.id, name: ch.name })) };
                        const tFileHandle = await exportDirHandle.getFileHandle(templateFileName, { create: true });
                        const tWritable = await tFileHandle.createWritable();
                        await tWritable.write(new Blob([JSON.stringify(templateData, null, 2)], { type: 'application/json' }));
                        await tWritable.close();
                        extraMsg = `\n(已同步导出配置模板: ${templateFileName})`;
                    } catch (e) {
                        console.error('静默导出模板失败', e);
                    }
                }
            }

            setExportSuccessModal({
              show: true,
              message: `已成功静默保存到文件夹:\n${exportDirName}/${fileName}${extraMsg}`,
              blob: blob,
              fileName: fileName
            });
            return;
          }
        }
        
        if (window.electronAPI) {
          let buffer;
          if (blob instanceof Blob) {
            buffer = await blob.arrayBuffer();
          } else {
            buffer = blob;
          }
          const savedPath = await window.electronAPI.saveExcelFile(new Uint8Array(buffer), `${exportBaseName}.xlsx`);
          if (savedPath) {
            if (autoExportTemplate) {
                const hasAnyName = channels.some(ch => ch.name && ch.name.trim() !== '');
                if (hasAnyName) exportTemplate(exportBaseName);
            }
            setExportSuccessModal({
              show: true,
              message: `已成功保存到:\n${savedPath}`,
              blob: blob,
              fileName: `${exportBaseName}.xlsx`,
              savedPath: savedPath
            });
          }
          return;
        }
        
        if (window.showSaveFilePicker) {
          const handle = await window.showSaveFilePicker({
            suggestedName: `${exportBaseName}.xlsx`,
            types: [{
              description: 'Excel 文件',
              accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']},
            }],
          });
          const writable = await handle.createWritable();
          await writable.write(blob);
          await writable.close();
          
          if (autoExportTemplate) {
              const hasAnyName = channels.some(ch => ch.name && ch.name.trim() !== '');
              if (hasAnyName) exportTemplate(exportBaseName);
          }
          setExportSuccessModal({
            show: true,
            message: `已成功导出文件：\n${exportBaseName}.xlsx`,
            blob: blob,
            fileName: `${exportBaseName}.xlsx`,
            savedPath: ''
          });
        } else {
          saveAs(blob, `${exportBaseName}.xlsx`);
          if (autoExportTemplate) {
              const hasAnyName = channels.some(ch => ch.name && ch.name.trim() !== '');
              if (hasAnyName) exportTemplate(exportBaseName);
          }
          setExportSuccessModal({
            show: true,
            message: `已成功导出文件：\n${exportBaseName}.xlsx`,
            blob: blob,
            fileName: `${exportBaseName}.xlsx`,
            savedPath: ''
          });
        }
      } catch (err) {
        // If user cancelled, do nothing. Otherwise fallback
        if (err.name !== 'AbortError') {
          console.error('保存文件失败，使用默认下载:', err);
          saveAs(blob, `${exportBaseName}.xlsx`);
          if (autoExportTemplate) {
              const hasAnyName = channels.some(ch => ch.name && ch.name.trim() !== '');
              if (hasAnyName) exportTemplate(exportBaseName);
          }
          setExportSuccessModal({
            show: true,
            message: `已成功导出文件：\n${exportBaseName}.xlsx`,
            blob: blob,
            fileName: `${exportBaseName}.xlsx`,
            savedPath: ''
          });
        }
      }
    } catch (err) {
      console.error(err);
      alert('导出失败，请检查是否缺少图表模板。');
    } finally {
      setIsExporting(false);
    }
  };

  const renderSelectionBar = () => (
    <div className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-xl p-3 mb-6">
      <div className="flex items-center gap-4">
        <button 
          onClick={toggleSelectAll} 
          className="flex items-center gap-2 text-sm text-slate-600 hover:text-blue-600 transition-colors font-medium"
        >
          {selectedIds.size === channels.length && channels.length > 0 ? <CheckSquare size={16} className="text-blue-500" /> : <Square size={16} />}
          全选 ({selectedIds.size}/{channels.length})
        </button>
      </div>
      <div className="flex items-center gap-3">
        <button 
          onClick={clearDeviceNames}
          disabled={channels.length === 0}
          className="flex items-center gap-2 text-sm font-medium bg-white border border-orange-200 text-orange-600 px-4 py-1.5 rounded-lg hover:bg-orange-50 disabled:opacity-40 disabled:hover:bg-white transition-all shadow-sm"
        >
          <Eraser size={16} /> 清除器件名称
        </button>
        <button 
          onClick={deleteSelected}
          disabled={selectedIds.size === 0}
          className="flex items-center gap-2 text-sm font-medium bg-white border border-red-200 text-red-600 px-4 py-1.5 rounded-lg hover:bg-red-50 disabled:opacity-40 disabled:hover:bg-white transition-all shadow-sm"
        >
          <Trash2 size={16} /> 删除选中
        </button>
      </div>
    </div>
  );

  const validTemps = channels.flatMap(ch => [parseFloat(ch.temp), parseFloat(ch.maxTemp), parseFloat(ch.compareTemp), parseFloat(ch.compareMaxTemp)]).filter(val => !isNaN(val));
  const globalMin = validTemps.length > 0 ? Math.min(...validTemps) : 0;
  const globalMax = validTemps.length > 0 ? Math.max(...validTemps) : 100;
  
  const getTempColorHex = (val) => {
    if (isNaN(val)) return '#ffffff';
    const min = globalMin;
    const max = globalMax;
    const green = [99, 190, 123];
    const yellow = [255, 235, 132];
    const red = [248, 105, 107];
    if (max === min) return '#FFEB84'; 
    const mid = (min + max) / 2;
    let r, g, b;
    if (val <= mid) {
      const ratio = Math.max(0, (val - min) / (mid - min));
      r = Math.round(green[0] + ratio * (yellow[0] - green[0]));
      g = Math.round(green[1] + ratio * (yellow[1] - green[1]));
      b = Math.round(green[2] + ratio * (yellow[2] - green[2]));
    } else {
      const ratio = Math.min(1, (val - mid) / (max - mid));
      r = Math.round(yellow[0] + ratio * (red[0] - yellow[0]));
      g = Math.round(yellow[1] + ratio * (red[1] - yellow[1]));
      b = Math.round(yellow[2] + ratio * (red[2] - yellow[2]));
    }
    const toHex = (n) => n.toString(16).padStart(2, '0').toUpperCase();
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
  };

  return (
    <div className="min-h-screen p-4 md:p-8 text-slate-800">
      <div className="max-w-[1600px] mx-auto space-y-6">
        <header className="flex flex-wrap items-center justify-between gap-y-6 gap-x-4">
          <div className="flex flex-wrap items-center gap-4">
            <div>
              <h1 className="text-3xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-blue-600 to-teal-500 whitespace-nowrap">
                温升数据整理工具
              </h1>
              <p className="text-slate-500 mt-1 whitespace-nowrap">Temperature Rise Data Organizer</p>
            </div>
            {importedFileName && (
              <div className="bg-blue-50 border border-blue-200 text-blue-700 px-4 py-1.5 rounded-full flex items-center gap-2 shadow-sm animate-in fade-in zoom-in max-w-full">
                <FileSpreadsheet size={16} className="text-blue-500 shrink-0" />
                <span className="text-sm font-semibold max-w-[200px] sm:max-w-[300px] truncate" title={importedFileName}>
                  {importedFileName}
                </span>
              </div>
            )}
          </div>
          
          <div className="flex flex-wrap items-center justify-end gap-3 ml-auto">
              <button 
                onClick={clearData}
                disabled={channels.length === 0}
                className="flex items-center gap-2 bg-white border border-red-200 text-red-600 px-5 py-2.5 rounded-full font-medium hover:bg-red-50 transition-colors shadow-sm disabled:opacity-50 active:scale-95"
              >
                <Trash2 size={18} /> 清除数据
              </button>
              <input type="file" accept=".xlsx, .xls, .csv" ref={fileInputRef} className="hidden" onChange={(e) => handleFileUpload(e, false)} />
              <button 
                onClick={() => fileInputRef.current?.click()}
                disabled={isImporting}
                className="flex items-center gap-2 bg-white border border-blue-200 text-blue-600 px-5 py-2.5 rounded-full font-medium hover:bg-blue-50 transition-colors shadow-sm disabled:opacity-50 active:scale-95"
              >
                <UploadCloud size={18} />
                {isImporting ? '导入中...' : '导入数据'}
              </button>
              
              {channels.length > 0 && (
                <>
                  <input type="file" accept=".xlsx, .xls, .csv" ref={compareInputRef} className="hidden" onChange={(e) => handleFileUpload(e, true)} />
                  <button 
                    onClick={() => compareInputRef.current?.click()}
                    disabled={isImporting}
                    className="flex items-center gap-2 bg-white border border-amber-200 text-amber-600 px-5 py-2.5 rounded-full font-medium hover:bg-amber-50 transition-colors shadow-sm disabled:opacity-50 active:scale-95"
                    title={compareFileName ? `当前对比: ${compareFileName}` : ''}
                  >
                    <UploadCloud size={18} />
                    {isImporting ? '导入中...' : (isComparing ? '重新导入对比' : '导入对比数据')}
                  </button>
                  {isComparing && (
                    <button
                      onClick={() => {
                        setCompareRawData([]);
                        setCompareFullRawData([]);
                        setIsComparing(false);
                        setCompareFileName('');
                        setChannels(prev => prev.map(ch => {
                          const newCh = { ...ch };
                          delete newCh.compareTemp;
                          delete newCh.compareMaxTemp;
                          return newCh;
                        }));
                      }}
                      className="flex items-center gap-2 bg-white border border-red-200 text-red-500 px-5 py-2.5 rounded-full font-medium hover:bg-red-50 transition-colors shadow-sm active:scale-95"
                    >
                      <Trash2 size={18} /> 清除对比
                    </button>
                  )}
                </>
              )}
              
              {exportDirName && (
                <div className="text-sm text-slate-500 flex items-center gap-1.5 bg-white border border-slate-200 px-3 py-1.5 rounded-full shadow-sm" title={exportDirName}>
                  <FolderOpen size={14} className="text-teal-500" />
                  自动导出至: <span className="font-semibold text-slate-700 truncate max-w-[180px]">{exportDirName}</span>
                </div>
              )}
              {(window.showDirectoryPicker || window.electronAPI) && (
                <button 
                  onClick={changeExportPath}
                  className="flex items-center gap-2 bg-white border border-slate-300 text-slate-700 px-5 py-2.5 rounded-full font-medium hover:bg-blue-50 hover:text-blue-600 hover:border-blue-300 transition-all shadow-sm active:scale-95"
                >
                  <FolderOpen size={18} className={exportDirName ? "text-slate-400" : "text-blue-500"} />
                  {exportDirName ? '更改目录' : '设置默认目录'}
                </button>
              )}
              <button 
                onClick={exportExcel}
                disabled={channels.length === 0 || isExporting}
                className={`flex items-center gap-2 px-6 py-2.5 rounded-full font-medium transition-all ${
                  channels.length === 0 || isExporting
                    ? 'bg-slate-100 text-slate-400 cursor-not-allowed border border-slate-200' 
                    : 'bg-gradient-to-r from-blue-600 to-teal-500 text-white hover:shadow-lg hover:shadow-blue-500/30 active:scale-95'
                }`}
              >
                {(!window.electronAPI && isExporting) ? (
                  <svg className="animate-spin -ml-1 mr-2 h-5 w-5 text-slate-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                  </svg>
                ) : (
                  <Download size={18} />
                )}
                {(!window.electronAPI && isExporting) ? '导出中...' : '导出 Excel 报告'}
              </button>
            </div>
        </header>

        <main className="bg-white/70 backdrop-blur-xl border border-white/50 shadow-xl shadow-slate-200/50 rounded-3xl overflow-hidden">
          <div className="flex border-b border-slate-200/60 bg-slate-50/80">
            <button 
              onClick={() => setActiveTab('config')}
              className={`flex-1 py-4 text-sm flex items-center justify-center gap-2 transition-all border-r border-slate-200/60 ${activeTab === 'config' ? 'text-blue-700 font-bold border-b-4 border-blue-600 bg-blue-100/50 shadow-sm' : 'text-slate-500 font-medium hover:bg-slate-100'}`}
            >
              <Settings size={18} className={activeTab === 'config' ? 'text-blue-600' : ''} /> 基础测试配置
            </button>
            <button 
              onClick={() => setActiveTab('data')}
              className={`flex-1 py-4 text-sm flex items-center justify-center gap-2 transition-all border-r border-slate-200/60 ${activeTab === 'data' ? 'text-blue-700 font-bold border-b-4 border-blue-600 bg-blue-100/50 shadow-sm' : 'text-slate-500 font-medium hover:bg-slate-100'}`}
            >
              <FileSpreadsheet size={18} className={activeTab === 'data' ? 'text-blue-600' : ''} /> 温度与通道数据
            </button>
            <button 
              onClick={() => setActiveTab('chart')}
              className={`flex-1 py-4 text-sm flex items-center justify-center gap-2 transition-all ${activeTab === 'chart' ? 'text-blue-700 font-bold border-b-4 border-blue-600 bg-blue-100/50 shadow-sm' : 'text-slate-500 font-medium hover:bg-slate-100'}`}
            >
              <ChartIcon size={18} className={activeTab === 'chart' ? 'text-blue-600' : ''} /> 温升曲线图
            </button>
          </div>

          <div className="p-8">
            {activeTab === 'config' && (
              <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <div className="w-1.5 h-4 bg-blue-500 rounded-full"></div>
                    <h3 className="text-base font-semibold text-slate-700">Excel 报告所见即所得编辑器</h3>
                  </div>
                  <div className="flex items-center gap-3">
                    {focusedField && (
                      <div className="flex items-center bg-slate-800 text-white px-1.5 py-1 rounded-lg shadow-lg animate-in fade-in zoom-in duration-200">
                        <span className="text-xs font-medium px-2 border-r border-slate-600 text-slate-300">
                          字号: {config.styles?.[focusedField]?.fontSize || 11}
                        </span>
                        <button 
                          onClick={() => updateStyle(focusedField, { fontSize: Math.max(6, (config.styles?.[focusedField]?.fontSize || 11) - 1) })}
                          className="px-2.5 py-1 hover:bg-slate-700 rounded transition-colors text-sm font-bold"
                          title="缩小字号"
                        >A-</button>
                        <button 
                          onClick={() => updateStyle(focusedField, { fontSize: Math.min(36, (config.styles?.[focusedField]?.fontSize || 11) + 1) })}
                          className="px-2.5 py-1 hover:bg-slate-700 rounded transition-colors text-sm font-bold"
                          title="放大字号"
                        >A+</button>
                        <div className="w-px h-4 bg-slate-600 mx-1"></div>
                        <button 
                          onClick={() => updateStyle(focusedField, { bold: !(config.styles?.[focusedField]?.bold) })}
                          className={`px-2.5 py-1 rounded transition-colors text-sm font-bold ${config.styles?.[focusedField]?.bold ? 'bg-blue-500 text-white' : 'hover:bg-slate-700'}`}
                          title="加粗"
                        >B</button>
                      </div>
                    )}

                    <span className="text-xs font-medium text-blue-600 bg-blue-50 px-3 py-1 rounded-full border border-blue-100 shadow-sm animate-pulse">
                      ✨ 点击任意文字即可修改内容和格式
                    </span>
                  </div>
                </div>
                
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm bg-white hover:border-blue-300 transition-colors focus-within:border-blue-400 focus-within:ring-4 focus-within:ring-blue-500/10">
                  <table className="w-full text-sm text-center border-collapse">
                    <tbody>
                      <tr>
                        <td rowSpan={2} className="border border-slate-200 w-[25%] p-0 align-middle">
                          <input 
                            name="condition" 
                            value={config.condition || ''} 
                            onChange={handleConfigChange} 
                            onFocus={() => setFocusedField('condition')}
                            style={{ fontWeight: config.styles?.condition?.bold ? 'bold' : 'normal', fontSize: config.styles?.condition?.fontSize ? `${config.styles.condition.fontSize}pt` : undefined }}
                            className="w-full h-full min-h-[96px] p-4 bg-white/50 outline-none text-center text-slate-700 placeholder-slate-300 hover:bg-slate-50 focus:bg-white transition-all block" 
                            placeholder="填写测试条件..." 
                          />
                        </td>
                        <td className="border border-slate-200 p-0 bg-slate-50 w-[25%] relative group">
                          <input name="pv" value={config.pv || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('pv')} style={{ fontWeight: config.styles?.pv?.bold ? 'bold' : 'normal', fontSize: config.styles?.pv?.fontSize ? `${config.styles.pv.fontSize}pt` : undefined }} placeholder="如: PV" className="w-full p-3 bg-transparent text-slate-600 outline-none text-center hover:bg-slate-100 focus:bg-white transition-all placeholder-slate-400 block" />
                        </td>
                        <td className="border border-slate-200 p-0 bg-slate-50 w-[25%] relative group">
                          <input name="pvVoltage" value={config.pvVoltage || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('pvVoltage')} style={{ fontWeight: config.styles?.pvVoltage?.bold ? 'bold' : 'normal', fontSize: config.styles?.pvVoltage?.fontSize ? `${config.styles.pvVoltage.fontSize}pt` : undefined }} placeholder="如: PV电压" className="w-full p-3 bg-transparent text-slate-600 outline-none text-center hover:bg-slate-100 focus:bg-white transition-all placeholder-slate-400 block" />
                        </td>
                        <td className="border border-slate-200 p-0 bg-slate-50 w-[25%] relative group">
                          <input name="pvCurrent" value={config.pvCurrent || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('pvCurrent')} style={{ fontWeight: config.styles?.pvCurrent?.bold ? 'bold' : 'normal', fontSize: config.styles?.pvCurrent?.fontSize ? `${config.styles.pvCurrent.fontSize}pt` : undefined }} placeholder="如: PV电流" className="w-full p-3 bg-transparent text-slate-600 outline-none text-center hover:bg-slate-100 focus:bg-white transition-all placeholder-slate-400 block" />
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-slate-200 p-0">
                          <input name="power" value={config.power || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('power')} style={{ fontWeight: config.styles?.power?.bold ? 'bold' : 'normal', fontSize: config.styles?.power?.fontSize ? `${config.styles.power.fontSize}pt` : undefined }} className="w-full p-4 text-slate-800 outline-none text-center hover:bg-slate-50 focus:bg-white transition-all placeholder-slate-300 block" placeholder="填写功率..." />
                        </td>
                        <td className="border border-slate-200 p-0">
                          <input name="voltageVal" value={config.voltageVal || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('voltageVal')} style={{ fontWeight: config.styles?.voltageVal?.bold ? 'bold' : 'normal', fontSize: config.styles?.voltageVal?.fontSize ? `${config.styles.voltageVal.fontSize}pt` : undefined }} className="w-full p-4 text-slate-800 outline-none text-center hover:bg-slate-50 focus:bg-white transition-all placeholder-slate-300 block" placeholder="填写电压值..." />
                        </td>
                        <td className="border border-slate-200 p-0">
                          <input name="currentVal" value={config.currentVal || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('currentVal')} style={{ fontWeight: config.styles?.currentVal?.bold ? 'bold' : 'normal', fontSize: config.styles?.currentVal?.fontSize ? `${config.styles.currentVal.fontSize}pt` : undefined }} className="w-full p-4 text-slate-800 outline-none text-center hover:bg-slate-50 focus:bg-white transition-all placeholder-slate-300 block" placeholder="填写电流值..." />
                        </td>
                      </tr>
                      <tr>
                        <td className="border border-slate-200 p-0 bg-slate-50">
                          <input name="noteLabel" value={config.noteLabel || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('noteLabel')} style={{ fontWeight: config.styles?.noteLabel?.bold ? 'bold' : 'normal', fontSize: config.styles?.noteLabel?.fontSize ? `${config.styles.noteLabel.fontSize}pt` : undefined }} className="w-full p-3 bg-transparent text-slate-600 outline-none text-center hover:bg-slate-100 focus:bg-white transition-all placeholder-slate-400 block" placeholder="如: 备注" />
                        </td>
                        <td colSpan={3} className="border border-slate-200 p-0">
                          <input name="topNote" value={config.topNote || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('topNote')} style={{ fontWeight: config.styles?.topNote?.bold ? 'bold' : 'normal', fontSize: config.styles?.topNote?.fontSize ? `${config.styles.topNote.fontSize}pt` : undefined }} className="w-full p-3 text-left text-slate-700 outline-none hover:bg-slate-50 focus:bg-white transition-all px-4 placeholder-slate-300 block" placeholder="在此填写顶部总结备注..." />
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
                
                <div className="mt-4 border border-slate-200 rounded-xl overflow-hidden shadow-sm bg-white hover:border-red-300 transition-colors focus-within:border-red-400 focus-within:ring-4 focus-within:ring-red-500/10">
                  <table className="w-full text-sm border-collapse">
                    <tbody>
                      <tr>
                        <td className="border border-slate-200 p-0 bg-slate-50 w-[25%] align-middle text-center">
                           <input name="bottomNoteLabel" value={config.bottomNoteLabel || ''} onChange={handleConfigChange} onFocus={() => setFocusedField('bottomNoteLabel')} style={{ fontWeight: config.styles?.bottomNoteLabel?.bold ? 'bold' : 'normal', fontSize: config.styles?.bottomNoteLabel?.fontSize ? `${config.styles.bottomNoteLabel.fontSize}pt` : undefined }} className="w-full p-3 bg-transparent text-slate-600 outline-none text-center hover:bg-slate-100 focus:bg-white transition-all placeholder-slate-400 block" placeholder="如: 底部总结备注" />
                        </td>
                        <td className="border border-slate-200 p-0 w-[75%]">
                           <textarea 
                             name="bottomNote" 
                             value={config.bottomNote || ''} 
                             onChange={(e) => setConfig({...config, bottomNote: e.target.value})} 
                             onFocus={() => setFocusedField('bottomNote')}
                             style={{ fontWeight: config.styles?.bottomNote?.bold ? 'bold' : 'normal', fontSize: config.styles?.bottomNote?.fontSize ? `${config.styles.bottomNote.fontSize}pt` : undefined }}
                             className="w-full h-16 p-3 text-left text-red-600 outline-none resize-none hover:bg-red-50/30 focus:bg-white transition-all px-4 placeholder-slate-300 block" 
                             placeholder="此处可填写底部补充备注，或者在导入数据后自动生成..." 
                           />
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 导入数据预览区 */}
                <div className="mt-6 border border-slate-200 rounded-xl overflow-hidden shadow-sm bg-white hover:border-blue-200 transition-colors">
                  <div className="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between">
                    <h4 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                      <div className="w-1.5 h-4 bg-teal-500 rounded-full"></div>
                      导入数据预览 (只读)
                    </h4>
                    <div className="flex items-center gap-4">
                      <div className="flex items-center gap-2 text-sm text-slate-600 bg-white px-3 py-1 rounded border border-slate-200 shadow-sm">
                        <input 
                          type="checkbox" 
                          id="showMaxTemp" 
                          checked={showMaxTemp}
                          onChange={(e) => setShowMaxTemp(e.target.checked)}
                          className="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer"
                        />
                        <label htmlFor="showMaxTemp" className="cursor-pointer font-medium select-none text-xs">
                          显示最高温度
                        </label>
                      </div>
                      <span className="text-xs text-slate-400">实时同步“温度与通道数据”</span>
                    </div>
                  </div>
                  {channels.length > 0 ? (
                    <div className="overflow-x-auto p-4 bg-slate-50/50">
                      {chunkedChannels.map((group, gIndex) => {
                        const cols = Math.max(8, group.length);
                        const totalRows = 4 + (isComparing ? 1 : 0) + (showMaxTemp ? (isComparing ? 2 : 1) : 0);
                        return (
                        <table key={gIndex} className="w-full text-xs text-center border-collapse border-2 border-slate-400 mb-6 bg-white last:mb-0 min-w-max">
                          <tbody>
                            <tr>
                              <td rowSpan={totalRows} className="border border-slate-400 w-16 align-middle font-medium text-slate-700 bg-white">GROUP {gIndex + 1}</td>
                              <td className="border border-slate-400 w-12 bg-white">室温</td>
                              <td className="border border-slate-400 w-24 bg-white">采集器通道</td>
                              {[...Array(cols)].map((_, i) => (
                                <td key={`ch-${i}`} className="border border-slate-400 h-8 bg-white text-slate-700 min-w-[60px]">
                                  {group[i] ? `CH${group[i].id}` : ''}
                                </td>
                              ))}
                            </tr>
                            <tr>
                              <td rowSpan={totalRows - 1} className="border border-slate-400 bg-white"></td>
                              <td className="border border-slate-400 bg-white">器件</td>
                              {[...Array(cols)].map((_, i) => (
                                <td key={`dev-${i}`} className="border border-slate-400 h-8 bg-white text-slate-700">
                                  {group[i]?.name || ''}
                                </td>
                              ))}
                            </tr>
                            <tr>
                              <td className="border border-slate-400 bg-white">稳定温度/℃</td>
                              {[...Array(cols)].map((_, i) => {
                                const tempStr = group[i]?.temp;
                                const temp = (tempStr !== undefined && tempStr !== '') ? parseFloat(tempStr) : NaN;
                                const bg = !isNaN(temp) ? getTempColorHex(temp) : '#ffffff';
                                return (
                                  <td key={`tmp-${i}`} className="border border-slate-400 h-8 text-slate-800" style={{ backgroundColor: bg }}>
                                    {!isNaN(temp) ? temp : ''}
                                  </td>
                                )
                              })}
                            </tr>
                            {isComparing && (
                              <tr>
                                <td className="border border-slate-400 bg-white border-t-0 text-slate-500 text-[10px]">对比稳定温度/℃</td>
                                {[...Array(cols)].map((_, i) => {
                                  const tempStr = group[i]?.compareTemp;
                                  const temp = (tempStr !== undefined && tempStr !== '') ? parseFloat(tempStr) : NaN;
                                  const bg = !isNaN(temp) ? getTempColorHex(temp) : '#ffffff';
                                  return (
                                    <td key={`comp-tmp-${i}`} className="border border-slate-400 border-t-0 border-dashed h-6 text-slate-500 text-[10px]" style={{ backgroundColor: bg }}>
                                      {!isNaN(temp) ? temp : ''}
                                    </td>
                                  )
                                })}
                              </tr>
                            )}
                            {showMaxTemp && (
                              <>
                                <tr>
                                  <td className="border border-slate-400 bg-white font-medium">最高温度/℃</td>
                                  {[...Array(cols)].map((_, i) => {
                                    const tempStr = group[i]?.maxTemp;
                                    const temp = (tempStr !== undefined && tempStr !== '') ? parseFloat(tempStr) : NaN;
                                    const bg = !isNaN(temp) ? getTempColorHex(temp) : '#ffffff';
                                    return (
                                      <td key={`max-${i}`} className="border border-slate-400 h-8 text-slate-800 font-medium" style={{ backgroundColor: bg }}>
                                        {!isNaN(temp) ? temp : ''}
                                      </td>
                                    )
                                  })}
                                </tr>
                                {isComparing && (
                                  <tr>
                                    <td className="border border-slate-400 bg-white border-t-0 text-slate-500 text-[10px] font-medium">对比最高温度/℃</td>
                                    {[...Array(cols)].map((_, i) => {
                                      const tempStr = group[i]?.compareMaxTemp;
                                      const temp = (tempStr !== undefined && tempStr !== '') ? parseFloat(tempStr) : NaN;
                                      const bg = !isNaN(temp) ? getTempColorHex(temp) : '#ffffff';
                                      return (
                                        <td key={`comp-max-${i}`} className="border border-slate-400 border-t-0 border-dashed h-6 text-slate-500 text-[10px] font-medium" style={{ backgroundColor: bg }}>
                                          {!isNaN(temp) ? temp : ''}
                                        </td>
                                      )
                                    })}
                                  </tr>
                                )}
                              </>
                            )}
                            <tr>
                              <td colSpan={3} className="border border-slate-400 bg-white h-8">备注</td>
                              <td colSpan={cols} className="border border-slate-400 bg-white text-left px-3 text-slate-600">
                                {groupNotes[gIndex] || ''}
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      )})}
                    </div>
                  ) : (
                    <div className="p-8 text-center text-slate-400 text-sm bg-slate-50/50">
                      暂无数据，请在“温度与通道数据”面板导入
                    </div>
                  )}
                </div>

              </div>
            )}


            {activeTab === 'data' && (
              <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="flex items-center justify-between bg-blue-50/50 border border-blue-100 rounded-2xl p-4">
                  <div className="flex gap-6 items-center">
                    <div>
                      <h3 className="text-sm font-semibold text-blue-800 flex items-center gap-2">
                        一键导入数据
                        {importedFileName && (
                          <span className="text-[10px] font-medium bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full">
                            {importedFileName}
                          </span>
                        )}
                      </h3>
                      <p className="text-xs text-blue-600/80 mt-1">自动解析 Excel / CSV 原始测试数据文件</p>
                    </div>
                    
                    <div className="flex items-center gap-4">
                      <div className="group relative flex items-center">
                        <HelpCircle size={16} className="text-slate-400 cursor-help hover:text-blue-500 transition-colors" />
                        <div className="absolute top-full mt-2 left-1/2 -translate-x-1/2 w-64 bg-slate-800 text-slate-100 text-xs rounded-xl p-3 shadow-xl opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all z-50 pointer-events-none leading-relaxed text-left font-normal">
                          <p className="font-semibold mb-1 text-white">💡 模板使用说明</p>
                          <p className="mb-1">1. 勾选“同步导出配置模板”后，每次导出Excel会同时生成一个.json配置文件。</p>
                          <p>2. 下次测试时，点击“加载模板”即可一键恢复之前的测试条件、器件名称、备注等，免去重新手敲的烦恼。</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2 mr-2 -ml-2 text-sm text-slate-600 bg-white/60 px-3 py-1.5 rounded-lg border border-slate-200">
                        <input 
                          type="checkbox" 
                          id="autoExport" 
                          checked={autoExportTemplate}
                          onChange={(e) => setAutoExportTemplate(e.target.checked)}
                          className="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 focus:ring-blue-500"
                        />
                        <label htmlFor="autoExport" className="cursor-pointer font-medium select-none text-xs">
                          同步导出配置模板
                        </label>
                      </div>

                      <button 
                        onClick={exportTemplate}
                        disabled={channels.length === 0 || autoExportTemplate || !channels.some(ch => ch.name && ch.name.trim() !== '')}
                        className={`flex items-center gap-2 border px-4 py-2 rounded-xl text-sm font-medium transition-colors shadow-sm ${
                          (channels.length === 0 || autoExportTemplate || !channels.some(ch => ch.name && ch.name.trim() !== ''))
                            ? 'bg-slate-50 text-slate-400 border-slate-200 cursor-not-allowed'
                            : 'bg-white border-indigo-200 text-indigo-600 hover:bg-indigo-50'
                        }`}
                      >
                        <Save size={16} /> 导出模板
                      </button>

                      <input type="file" accept=".json" ref={templateInputRef} className="hidden" onChange={importTemplate} />
                      <button 
                        onClick={() => templateInputRef.current?.click()}
                        className="flex items-center gap-2 bg-white border border-teal-200 text-teal-600 px-4 py-2 rounded-xl text-sm font-medium hover:bg-teal-50 transition-colors shadow-sm"
                      >
                        <FileJson size={16} />
                        加载模板
                      </button>
                      <button 
                        onClick={() => fileInputRef.current?.click()}
                        disabled={isImporting}
                        className="flex items-center gap-2 bg-white border border-blue-200 text-blue-600 px-4 py-2 rounded-xl text-sm font-medium hover:bg-blue-50 transition-colors shadow-sm disabled:opacity-50"
                      >
                        <UploadCloud size={16} />
                        {isImporting ? '解析中...' : '导入数据'}
                      </button>
                      
                      {channels.length > 0 && (
                        <>
                          <button 
                            onClick={() => compareInputRef.current?.click()}
                            disabled={isImporting}
                            className="flex items-center gap-2 bg-white border border-amber-200 text-amber-600 px-4 py-2 rounded-xl text-sm font-medium hover:bg-amber-50 transition-colors shadow-sm disabled:opacity-50"
                            title={compareFileName ? `当前对比: ${compareFileName}` : ''}
                          >
                            <UploadCloud size={16} />
                            {isImporting ? '解析中...' : (isComparing ? '重新导入对比' : '导入对比数据')}
                          </button>
                          {isComparing && (
                            <button
                              onClick={() => {
                                setCompareRawData([]);
                                setCompareFullRawData([]);
                                setIsComparing(false);
                                setCompareFileName('');
                                setChannels(prev => prev.map(ch => {
                                  const newCh = { ...ch };
                                  delete newCh.compareTemp;
                                  delete newCh.compareMaxTemp;
                                  return newCh;
                                }));
                              }}
                              className="flex items-center gap-2 bg-white border border-red-200 text-red-500 px-4 py-2 rounded-xl text-sm font-medium hover:bg-red-50 transition-colors shadow-sm"
                            >
                              <Trash2 size={16} /> 清除对比
                            </button>
                          )}
                        </>
                      )}
                    </div>
                  </div>
                </div>

                {channels.length > 0 && renderSelectionBar()}

                {channels.length === 0 && (
                  <div className="text-center py-16 bg-white/50 rounded-2xl border border-dashed border-slate-200">
                    <div className="w-16 h-16 bg-blue-50 text-blue-400 rounded-full flex items-center justify-center mx-auto mb-4">
                      <FileSpreadsheet size={28} />
                    </div>
                    <h3 className="text-lg font-semibold text-slate-700 mb-2">还未导入任何数据</h3>
                    <p className="text-slate-500 mb-6">点击上方的“一键导入数据”按钮，选择您的 Excel 测试文件。</p>
                    <div className="flex items-center justify-center gap-4">
                      <button 
                        onClick={() => fileInputRef.current?.click()}
                        className="bg-blue-600 text-white px-6 py-2.5 rounded-full text-sm font-medium hover:bg-blue-700 transition-colors shadow-md shadow-blue-500/20"
                      >
                        立即导入数据
                      </button>
                    </div>
                  </div>
                )}

                {chunkedChannels.map((group, gIndex) => (
                  <div key={gIndex} className="bg-white rounded-2xl p-6 border border-slate-100 shadow-sm">
                    <h4 className="font-medium text-slate-600 mb-4 flex items-center gap-2">
                      <div className="w-2 h-2 rounded-full bg-teal-500"></div> GROUP {gIndex + 1}
                    </h4>
                    <div className="grid grid-cols-8 gap-3">
                      {group.map(ch => (
                        <div key={ch.id} className={`space-y-2.5 relative p-3 rounded-xl border transition-all ${selectedIds.has(ch.id) ? 'bg-blue-50/30 border-blue-200 shadow-sm' : 'bg-slate-50/50 border-slate-100 hover:border-slate-200'}`}>
                          <div className="flex items-center gap-2">
                            <button onClick={() => toggleSelection(ch.id)} className="text-slate-400 hover:text-blue-500 transition-colors flex-shrink-0">
                              {selectedIds.has(ch.id) ? <CheckSquare size={16} className="text-blue-500" /> : <Square size={16} />}
                            </button>
                            <span className="text-sm font-bold text-slate-700">CH{ch.id}</span>
                          </div>
                          
                          <div className="space-y-2">
                            <input 
                              value={ch.name} 
                              onChange={(e) => handleChannelChange(ch.id, 'name', e.target.value)} 
                              className="w-full bg-white border border-slate-200 rounded-lg text-xs md:text-sm px-2.5 py-1.5 focus:ring-1 focus:ring-blue-500/30 focus:border-blue-500 outline-none transition-all placeholder-slate-300" 
                              placeholder="器件名称"
                            />
                            <div className="relative">
                              <input 
                                value={ch.temp} 
                                readOnly
                                style={{ backgroundColor: ch.temp ? getTempColorHex(parseFloat(ch.temp)) : '#ffffff' }}
                                className="w-full border border-slate-200/80 rounded-lg text-xs md:text-sm font-semibold pl-2.5 pr-6 py-1.5 outline-none placeholder-slate-300 text-slate-800 shadow-inner" 
                                placeholder="温度"
                                type="number"
                                step="0.01"
                              />
                              <span className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-600/60 font-medium text-xs pointer-events-none">℃</span>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                    
                    <div className="mt-4 pt-4 border-t border-slate-100">
                      <div className="flex items-center gap-3">
                        <label className="text-sm font-medium text-slate-500 whitespace-nowrap">本组备注</label>
                        <input 
                          value={groupNotes[gIndex] || ''} 
                          onChange={(e) => setGroupNotes({...groupNotes, [gIndex]: e.target.value})} 
                          className="w-full bg-slate-50 border border-slate-200 rounded-lg px-4 py-2 text-sm focus:bg-white focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 outline-none transition-all placeholder-slate-400" 
                          placeholder="填写该组备注（选填），没填写就留空..."
                        />
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {activeTab === 'chart' && (
              <div className="space-y-4 animate-in fade-in slide-in-from-bottom-4 duration-500">
                {rawData.length > 0 && (
                  <div className="flex justify-between items-center px-1 mb-2">
                    <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-slate-200 shadow-sm">
                      <span className="text-xs font-semibold text-slate-600">截取指定时间段:</span>
                      <select 
                        className="text-xs border border-slate-300 rounded px-2 py-1 outline-none focus:border-blue-500 bg-slate-50 min-w-[100px] cursor-pointer"
                        value={cropStartIndex}
                        onChange={(e) => setCropStartIndex(Number(e.target.value))}
                      >
                        {fullRawData.map((d, i) => (
                          <option key={`start-${i}`} value={i}>{d.time}</option>
                        ))}
                      </select>
                      <span className="text-xs text-slate-400 font-medium">至</span>
                      <select 
                        className="text-xs border border-slate-300 rounded px-2 py-1 outline-none focus:border-blue-500 bg-slate-50 min-w-[100px] cursor-pointer"
                        value={cropEndIndex}
                        onChange={(e) => setCropEndIndex(Number(e.target.value))}
                      >
                        {fullRawData.map((d, i) => (
                          <option key={`end-${i}`} value={i}>{d.time}</option>
                        ))}
                      </select>
                      <button onClick={handleCropExact} className="ml-2 px-3 py-1 bg-blue-600 text-white rounded shadow-sm text-xs font-medium hover:bg-blue-700 transition active:scale-95">
                        确认截取
                      </button>
                      <button onClick={handleResetData} className="px-3 py-1 bg-slate-100 text-slate-600 rounded shadow-sm text-xs font-medium hover:bg-slate-200 transition active:scale-95 border border-slate-200">
                        恢复全局
                      </button>
                    </div>
                    <div className="text-xs font-medium text-slate-500 bg-slate-50 px-3 py-1.5 rounded-full border border-slate-200 flex items-center gap-2 shadow-sm">
                      <ChartIcon size={14} className="text-blue-500" /> 滚轮缩放(上下左右)，左键拖拽平移，双击还原
                    </div>
                  </div>
                )}
                <div 
                  className={`bg-white rounded-2xl p-6 shadow-sm border border-slate-100 ${isComparing ? 'h-[800px]' : 'h-[600px]'} relative select-none`}
                  ref={chartContainerRef}
                  onDoubleClick={resetZoom}
                >
                  {rawData.length > 0 ? (
                    <>
                      {isComparing ? (
                        <div className="w-full h-full flex flex-col gap-4">
                          <div className="flex-1 min-h-0 relative">
                            <div className="absolute top-2 left-6 text-xs font-bold text-slate-500 z-10 bg-white/80 px-2 py-1 rounded shadow-sm border border-slate-100">当前测试数据</div>
                            <ResponsiveContainer width="100%" height="100%">
                              <LineChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 0 }} syncId="tempChart">
                                <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                                <XAxis dataKey="time" minTickGap={30} tick={{ fontSize: 12, fill: '#64748b' }} />
                                <YAxis tick={{ fontSize: 12, fill: '#64748b' }} domain={getVisibleYDomain()} tickFormatter={(val) => Number.isInteger(val) ? val : parseFloat(val).toFixed(1)} />
                                <Tooltip content={<CustomTooltip />} />
                                <Legend wrapperStyle={{ paddingTop: '10px', paddingBottom: '10px' }} />
                                {Array.from(chartSelectedIds).map(id => {
                                    const ch = channels.find(c => c.id === id);
                                    if (!ch) return null;
                                    const name = ch.name ? `${ch.name} (CH${id})` : `CH${id}`;
                                    const hue = (id * 137.5) % 360;
                                    return <Line key={id} isAnimationActive={false} type="monotone" dataKey={`CH${id}`} name={name} stroke={`hsl(${hue}, 70%, 50%)`} dot={false} strokeWidth={2} />;
                                })}
                              </LineChart>
                            </ResponsiveContainer>
                          </div>
                          <div className="flex-1 min-h-0 relative">
                            <div className="absolute top-2 left-6 text-xs font-bold text-slate-500 z-10 bg-white/80 px-2 py-1 rounded shadow-sm border border-slate-100">历史对比数据</div>
                            <ResponsiveContainer width="100%" height="100%">
                              <LineChart data={chartData} margin={{ top: 0, right: 30, left: 20, bottom: 20 }} syncId="tempChart">
                                <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                                <XAxis dataKey="compare_time" minTickGap={30} tick={{ fontSize: 12, fill: '#64748b' }} />
                                <YAxis tick={{ fontSize: 12, fill: '#64748b' }} domain={getVisibleYDomain()} tickFormatter={(val) => Number.isInteger(val) ? val : parseFloat(val).toFixed(1)} />
                                <Tooltip content={<CustomTooltip />} />
                                {Array.from(chartSelectedIds).map(id => {
                                    const ch = channels.find(c => c.id === id);
                                    if (!ch) return null;
                                    const name = ch.name ? `${ch.name} (CH${id})` : `CH${id}`;
                                    const hue = (id * 137.5) % 360;
                                    return <Line key={id} isAnimationActive={false} type="monotone" dataKey={`compare_CH${id}`} name={`${name} (对比)`} stroke={`hsl(${hue}, 70%, 50%)`} dot={false} strokeWidth={2} />;
                                })}
                              </LineChart>
                            </ResponsiveContainer>
                          </div>
                        </div>
                      ) : (
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 10 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                            <XAxis dataKey="time" minTickGap={30} tick={{ fontSize: 12, fill: '#64748b' }} />
                            <YAxis tick={{ fontSize: 12, fill: '#64748b' }} domain={getVisibleYDomain()} tickFormatter={(val) => Number.isInteger(val) ? val : parseFloat(val).toFixed(1)} />
                            <Tooltip content={<CustomTooltip />} />
                            <Legend wrapperStyle={{ paddingTop: '20px' }} />
                            {Array.from(chartSelectedIds).map(id => {
                                const ch = channels.find(c => c.id === id);
                                if (!ch) return null;
                                const name = ch.name ? `${ch.name} (CH${id})` : `CH${id}`;
                                const hue = (id * 137.5) % 360;
                                return <Line key={id} isAnimationActive={false} type="monotone" dataKey={`CH${id}`} name={name} stroke={`hsl(${hue}, 70%, 50%)`} dot={false} strokeWidth={2} />;
                            })}
                          </LineChart>
                        </ResponsiveContainer>
                      )}
                    </>
                  ) : (
                    <div className="w-full h-full flex flex-col items-center justify-center text-slate-400">
                      <ChartIcon size={48} className="mb-4 opacity-20" />
                      <p>请先导入数据</p>
                    </div>
                  )}
                </div>
                {rawData.length > 0 && (
                  <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100">
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-lg font-medium text-slate-800">图表显示通道</h3>
                      <button
                        onClick={toggleAllChartSelection}
                        className="flex items-center gap-2 text-sm text-slate-600 hover:text-blue-600 transition-colors font-medium"
                      >
                        {chartSelectedIds.size === channels.length && channels.length > 0 ? <CheckSquare size={16} className="text-blue-500" /> : <Square size={16} />}
                        全选 ({chartSelectedIds.size}/{channels.length})
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {channels.map((ch) => {
                        const isSelected = chartSelectedIds.has(ch.id);
                        const hue = (ch.id * 137.5) % 360;
                        const name = ch.name ? `${ch.name} (CH${ch.id})` : `CH${ch.id}`;
                        return (
                          <button
                            key={ch.id}
                            onClick={() => toggleChartSelection(ch.id)}
                            className={`px-3 py-1.5 rounded-full text-xs font-medium border transition-all ${isSelected ? 'text-white' : 'bg-white text-slate-600 border-slate-200 hover:border-slate-300'}`}
                            style={isSelected ? { backgroundColor: `hsl(${hue}, 70%, 50%)`, borderColor: `hsl(${hue}, 70%, 50%)` } : {}}
                          >
                            {name}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </main>
      </div>

      {exportSuccessModal.show && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center animate-in fade-in">
          <div className="bg-white rounded-2xl p-6 shadow-2xl max-w-md w-full mx-4 border border-slate-100 flex flex-col items-center text-center">
            <div className="w-16 h-16 bg-green-100 text-green-500 rounded-full flex items-center justify-center mb-4">
              <CheckSquare size={32} />
            </div>
            <h3 className="text-xl font-bold text-slate-800 mb-2">导出成功</h3>
            <p className="text-sm text-slate-600 mb-6 whitespace-pre-wrap break-all">
              {exportSuccessModal.message}
            </p>
            <div className="flex w-full gap-3">
              <button 
                onClick={() => setExportSuccessModal({ show: false, message: '', blob: null, fileName: '' })}
                className="flex-1 bg-slate-100 text-slate-700 py-2.5 rounded-xl font-medium hover:bg-slate-200 transition-colors"
              >
                关闭
              </button>
              <button 
                onClick={() => {
                  if (window.electronAPI && exportSuccessModal.savedPath) {
                    window.electronAPI.openFile(exportSuccessModal.savedPath);
                    setExportSuccessModal({ show: false, message: '', blob: null, fileName: '' });
                  } else {
                    const textToCopy = exportSuccessModal.message.split('\n')[1] || exportSuccessModal.fileName;
                    navigator.clipboard.writeText(textToCopy).then(() => {
                      showToast('复制路径成功', 'success');
                      setExportSuccessModal({ show: false, message: '', blob: null, fileName: '' });
                    }).catch(err => {
                      console.error('复制失败', err);
                      showToast('复制路径失败', 'error');
                      setExportSuccessModal({ show: false, message: '', blob: null, fileName: '' });
                    });
                  }
                }}
                className="flex-1 bg-blue-600 text-white py-2.5 rounded-xl font-medium hover:bg-blue-700 transition-colors shadow-md shadow-blue-500/20"
              >
                {window.electronAPI ? '打开文件' : '复制路径'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirm Modal */}
      {confirmModal.show && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center animate-in fade-in">
          <div className="bg-white rounded-2xl p-6 shadow-2xl max-w-md w-full mx-4 border border-slate-100 flex flex-col items-center text-center">
            <div className="w-16 h-16 bg-amber-100 text-amber-500 rounded-full flex items-center justify-center mb-4">
              <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg>
            </div>
            <h3 className="text-xl font-bold text-slate-800 mb-2">请确认</h3>
            <p className="text-sm text-slate-600 mb-6 whitespace-pre-wrap break-all">
              {confirmModal.message}
            </p>
            <div className="flex w-full gap-3">
              <button 
                onClick={() => setConfirmModal({ show: false, message: '', onConfirm: null })}
                className="flex-1 bg-slate-100 text-slate-700 py-2.5 rounded-xl font-medium hover:bg-slate-200 transition-colors"
              >
                取消
              </button>
              <button 
                onClick={confirmModal.onConfirm}
                className="flex-1 bg-red-600 text-white py-2.5 rounded-xl font-medium hover:bg-red-700 transition-colors shadow-md shadow-red-500/20"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toast Notification */}
      {toast.show && (
        <div className="fixed bottom-20 left-1/2 -translate-x-1/2 z-[100] animate-in slide-in-from-bottom-4 fade-in duration-300">
          <div className={`flex items-center gap-3 px-6 py-3 rounded-full shadow-lg border ${
            toast.type === 'success' ? 'bg-green-50 border-green-200 text-green-700' : 'bg-red-50 border-red-200 text-red-700'
          }`}>
            {toast.type === 'success' ? <CheckSquare size={20} className="text-green-500" /> : <ChartIcon size={20} className="text-red-500" />}
            <span className="font-medium text-sm">{toast.message}</span>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
