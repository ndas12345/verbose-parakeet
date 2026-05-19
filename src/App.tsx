/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  Upload, 
  ArrowRight, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  ChevronRight,
  Database,
  Plus,
  Search,
  Trash2,
  X
} from 'lucide-react';
import { cn } from './lib/utils';
import { motion, AnimatePresence } from 'motion/react';
import Fuse from 'fuse.js';

type Step = 'upload-source' | 'upload-template' | 'mapping' | 'preview' | 'complete';

interface ColumnMapping {
  templateColumn: string;
  sourceColumn: string;
}

interface ExcelData {
  headers: string[];
  rows: any[][];
  fileName: string;
}

const MEMORY_KEY = 'postmaster_bridge_mappings_v2';

export default function App() {
  const [activeStep, setActiveStep] = useState<Step>('upload-source');
  const [sourceData, setSourceData] = useState<ExcelData | null>(null);
  const [templateData, setTemplateData] = useState<ExcelData | null>(null);
  const [mappings, setMappings] = useState<ColumnMapping[]>([]);
  const [selectedRowIndices, setSelectedRowIndices] = useState<Set<number>>(new Set());
  const [previewSearchQuery, setPreviewSearchQuery] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Load Mapping Memory
  const [memory, setMemory] = useState<Record<string, string>>(() => {
    try {
      const saved = localStorage.getItem(MEMORY_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  const handleFileUpload = useCallback((type: 'source' | 'template') => (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        if (jsonData.length === 0) {
          setError(`File ${file.name} is empty`);
          return;
        }

        const headers = jsonData[0].map(h => String(h).trim());
        const rows = jsonData.slice(1);

        const excelData = {
          headers,
          rows,
          fileName: file.name
        };

        if (type === 'source') {
          setSourceData(excelData);
          // Auto-select all rows by default
          setSelectedRowIndices(new Set(rows.map((_, i) => i)));
          setActiveStep('upload-template');
        } else {
          setTemplateData(excelData);
          // Auto-mapping logic: Priority to Memory, then Fuzzy Receiver Matching
          const initialMappings: ColumnMapping[] = [];
          const receiverKeywords = ['rec', 'to', 'addressee', 'name', 'addr', 'city', 'state', 'pin', 'zip', 'line'];
          const senderKeywords = ['sender', 'from', 'return', 'origin'];
          let memCount = 0;

          headers.forEach(tCol => {
            const lowTCol = tCol.toLowerCase();
            
            // 1. Check Memory first (Manual preference)
            const rememberedSource = memory[tCol];
            if (rememberedSource && sourceData?.headers.includes(rememberedSource)) {
              initialMappings.push({ templateColumn: tCol, sourceColumn: rememberedSource });
              memCount++;
              return;
            }

            // 2. Skip auto-mapping if it looks like a sender field
            const isSender = senderKeywords.some(key => lowTCol.includes(key));
            if (isSender) return;

            // 3. Fallback to Receiver-like fuzzy matching
            const isReceiver = receiverKeywords.some(key => lowTCol.includes(key));
            
            if (isReceiver) {
              const match = sourceData?.headers.find(sCol => 
                sCol.toLowerCase().includes(lowTCol) || 
                lowTCol.includes(sCol.toLowerCase())
              );
              if (match) {
                initialMappings.push({ templateColumn: tCol, sourceColumn: match });
              }
            }
          });
          setMappings(initialMappings);
          setActiveStep('mapping');
        }
        setError(null);
      } catch (err) {
        setError("Error reading Excel file. Please ensure it's a valid .xlsx or .xls file.");
      }
    };
    reader.readAsArrayBuffer(file);
  }, [sourceData, memory]);

  const addMapping = (templateCol: string, sourceCol: string) => {
    setMappings(prev => {
      const filtered = prev.filter(m => m.templateColumn !== templateCol);
      const newMappings = [...filtered, { templateColumn: templateCol, sourceColumn: sourceCol }];
      
      // Update Persistent Memory
      const newMemory = { ...memory, [templateCol]: sourceCol };
      setMemory(newMemory);
      localStorage.setItem(MEMORY_KEY, JSON.stringify(newMemory));
      
      return newMappings;
    });
  };

  const removeMapping = (templateCol: string) => {
    setMappings(prev => prev.filter(m => m.templateColumn !== templateCol));
    // Also remove from memory if it was a negative choice? 
    // Maybe not, just let them select something else.
  };

  const reset = () => {
    setSourceData(null);
    setTemplateData(null);
    setMappings([]);
    setSelectedRowIndices(new Set());
    setPreviewSearchQuery('');
    setActiveStep('upload-source');
    setError(null);
  };

  const previewDataRows = useMemo(() => {
    if (!sourceData || !templateData || mappings.length === 0) return [];
    
    const sourceColIndices = new Map(sourceData.headers.map((h, i) => [h, i]));
    const templateColIndices = new Map(templateData.headers.map((h, i) => [h, i]));

    return sourceData.rows.map((sRow: any[], sIdx: number) => {
      const mappedRow: any = { _sourceIndex: sIdx };
      
      mappings.forEach(m => {
        const sIdxInRow = sourceColIndices.get(m.sourceColumn);
        if (sIdxInRow !== undefined) {
          let value = sRow[sIdxInRow as number];
          if (value !== undefined && value !== null) {
            let strValue = String(value).trim();
            if (m.templateColumn.toLowerCase().includes('pin') || m.templateColumn.toLowerCase().includes('zip')) {
              strValue = strValue.replace(/\D/g, '');
            }
            if (m.templateColumn.toLowerCase().includes('name')) {
              strValue = strValue.toLowerCase().split(' ').map(s => s.charAt(0).toUpperCase() + s.substring(1)).join(' ');
            }
            mappedRow[m.templateColumn] = strValue;
          } else {
            mappedRow[m.templateColumn] = '';
          }
        }
      });
      return mappedRow;
    });
  }, [sourceData, templateData, mappings]);

  const filteredPreviewRows = useMemo(() => {
    if (!previewSearchQuery.trim()) return previewDataRows;
    
    const options = {
      keys: mappings.map(m => m.templateColumn),
      threshold: 0.4, // Allows for some misspelling (0 is exact, 1 is anything)
      distance: 100,
      minMatchCharLength: 2
    };

    const fuse = new Fuse(previewDataRows, options);
    const results = fuse.search(previewSearchQuery);
    
    return results.map(result => result.item);
  }, [previewDataRows, previewSearchQuery, mappings]);

  const toggleRowSelection = (idx: number) => {
    setSelectedRowIndices(prev => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  };

  const setAllVisibleSelection = (selected: boolean) => {
    const allVisibleIndices = filteredPreviewRows.map(r => r._sourceIndex);
    setSelectedRowIndices(prev => {
      const next = new Set(prev);
      if (selected) {
        allVisibleIndices.forEach(idx => next.add(idx));
      } else {
        allVisibleIndices.forEach(idx => next.delete(idx));
      }
      return next;
    });
  };

  const processAndDownload = () => {
    if (!sourceData || !templateData || selectedRowIndices.size === 0) return;
    setIsProcessing(true);

    try {
      const sourceColIndices = new Map(sourceData.headers.map((h, i) => [h, i]));
      const templateColIndices = new Map(templateData.headers.map((h, i) => [h, i]));

      const resultRows = sourceData.rows
        .filter((_, idx) => selectedRowIndices.has(idx))
        .map((sRow: any[]) => {
          const newRow = new Array(templateData.headers.length).fill('');
          
          mappings.forEach(m => {
            const tIdx = templateColIndices.get(m.templateColumn);
            const sIdx = sourceColIndices.get(m.sourceColumn);
            
            if (tIdx !== undefined && sIdx !== undefined) {
              const value = sRow[sIdx as number];
              if (value !== undefined && value !== null) {
                let strValue = String(value).trim();
                if (m.templateColumn.toLowerCase().includes('pin') || m.templateColumn.toLowerCase().includes('zip')) {
                  strValue = strValue.replace(/\D/g, '');
                }
                if (m.templateColumn.toLowerCase().includes('name')) {
                  strValue = strValue.toLowerCase().split(' ').map(s => s.charAt(0).toUpperCase() + s.substring(1)).join(' ');
                }
                newRow[tIdx as number] = strValue;
              } else {
                newRow[tIdx as number] = '';
              }
            }
          });
          return newRow;
        });

      // 3. Create workbook
      const ws = XLSX.utils.aoa_to_sheet([templateData.headers, ...resultRows]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Bulk Posting");

      // 4. Trigger download
      XLSX.writeFile(wb, `Merged_Post_Office_Bulk_${Date.now()}.xlsx`);
      
      setActiveStep('complete');
    } catch (err) {
      setError("Failed to generate file. Check column mappings.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="flex flex-col h-screen w-full bg-slate-100 font-sans text-slate-900 overflow-hidden">
      {/* Header */}
      <header className="h-14 bg-white border-b border-slate-200 flex items-center justify-between px-6 shrink-0 z-10">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-orange-500 rounded flex items-center justify-center">
            <Database className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-lg font-bold tracking-tight text-slate-800">
            PostMaster Bridge <span className="text-slate-400 font-normal ml-2">v1.4</span>
          </h1>
        </div>
        <div className="flex items-center gap-4">
          <div className={cn(
            "flex items-center gap-2 px-3 py-1 rounded-full border text-xs font-semibold transition-colors",
            sourceData && templateData ? "bg-green-50 text-green-700 border-green-200" : "bg-slate-50 text-slate-500 border-slate-200"
          )}>
            <span className={cn("w-2 h-2 rounded-full", sourceData && templateData ? "bg-green-500" : "bg-slate-300")} />
            {sourceData && templateData ? "Ready to Process" : "Awaiting Data"}
          </div>
          <button 
            onClick={reset}
            className="px-4 py-2 bg-slate-800 text-white rounded text-sm font-medium hover:bg-slate-700 transition-colors"
          >
            New Session
          </button>
        </div>
      </header>

      <main className="flex-1 flex gap-4 p-4 min-h-0">
        {/* Sidebar */}
        <div className="w-68 flex flex-col gap-4 shrink-0">
          {/* Source Config */}
          <div className="bg-white rounded-lg border border-slate-200 p-4 flex flex-col h-1/2 shadow-sm">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold mb-3 flex items-center justify-between">
              Source Configuration
              {sourceData && <CheckCircle2 className="w-3 h-3 text-green-500" />}
            </div>
            
            {sourceData ? (
              <div className="flex flex-col h-full min-h-0">
                <div className="p-3 border-2 border-dashed border-slate-200 rounded-md bg-slate-50 mb-3 flex flex-col items-center justify-center text-center">
                  <div className="text-orange-600 font-bold text-xs truncate w-full">{sourceData.fileName}</div>
                  <div className="text-[10px] text-slate-400">{sourceData.rows.length.toLocaleString()} rows • {sourceData.headers.length} columns</div>
                </div>
                <div className="space-y-2 flex-1 overflow-y-auto pr-1">
                  <div className="text-[11px] font-medium text-slate-400">DETECTED FIELDS:</div>
                  {sourceData.headers.map(h => (
                    <div key={h} className="px-2 py-1.5 bg-slate-100 rounded text-[10px] font-mono border border-slate-200 truncate">
                      {h}
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="flex-1 flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-md bg-slate-50">
                <div className="text-[10px] text-slate-400 uppercase font-mono">No Source File</div>
              </div>
            )}
          </div>

          {/* Target Config */}
          <div className="bg-white rounded-lg border border-slate-200 p-4 flex flex-col h-1/2 shadow-sm">
            <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold mb-3 flex items-center justify-between">
              Target Template
              {templateData && <CheckCircle2 className="w-3 h-3 text-green-500" />}
            </div>

            {templateData ? (
              <div className="flex-1 flex flex-col min-h-0">
                <div className="p-3 border-2 border-dashed border-orange-200 rounded-md bg-orange-50 mb-3 flex flex-col items-center justify-center text-center">
                  <div className="text-orange-600 font-bold text-xs truncate w-full">{templateData.fileName}</div>
                  <div className="text-[10px] text-orange-400">Template Required Fields: {templateData.headers.length}</div>
                </div>
                <div className="space-y-2 flex-1 overflow-y-auto pr-1">
                  <div className="text-[11px] font-medium text-slate-400">TEMPLATE HEADERS:</div>
                  {templateData.headers.map(h => (
                    <div key={h} className="px-2 py-1.5 bg-white border border-slate-200 rounded text-[10px] font-mono italic truncate">
                      {h}
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="flex-1 flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-md bg-slate-50">
                <div className="text-[10px] text-slate-400 uppercase font-mono">No Template File</div>
              </div>
            )}
          </div>
        </div>

        {/* Main Workspace */}
        <div className="flex-1 bg-white rounded-lg border border-slate-200 flex flex-col shadow-sm min-w-0">
          <AnimatePresence mode="wait">
            {!sourceData || !templateData ? (
              <motion.div 
                key="upload-zone"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="flex-1 flex flex-col items-center justify-center p-12 text-center"
              >
                {!sourceData ? (
                  <div className="max-w-md">
                    <div className="w-16 h-16 bg-slate-100 rounded-2xl flex items-center justify-center mx-auto mb-6">
                      <Database className="w-8 h-8 text-slate-400" />
                    </div>
                    <h2 className="text-2xl font-bold text-slate-800 mb-2">Step 1: Upload Source</h2>
                    <p className="text-slate-500 mb-8">Drop your Excel address database here to begin mapping.</p>
                    <label className="inline-flex items-center px-8 py-3 bg-slate-800 text-white rounded font-bold cursor-pointer hover:bg-slate-700 transition-all shadow-lg shadow-slate-200">
                      <Upload className="w-4 h-4 mr-2" />
                      Browse Source File
                      <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload('source')} />
                    </label>
                  </div>
                ) : (
                  <div className="max-w-md">
                    <div className="w-16 h-16 bg-orange-50 rounded-2xl flex items-center justify-center mx-auto mb-6">
                      <FileSpreadsheet className="w-8 h-8 text-orange-500" />
                    </div>
                    <h2 className="text-2xl font-bold text-slate-800 mb-2">Step 2: Upload Template</h2>
                    <p className="text-slate-500 mb-8">Select the Post Office bulk posting template to map your data into.</p>
                    <label className="inline-flex items-center px-8 py-3 bg-orange-600 text-white rounded font-bold cursor-pointer hover:bg-orange-500 transition-all shadow-lg shadow-orange-100">
                      <Upload className="w-4 h-4 mr-2" />
                      Browse Template File
                      <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleFileUpload('template')} />
                    </label>
                  </div>
                )}
              </motion.div>
            ) : activeStep === 'mapping' ? (
              <motion.div 
                key="mapping-view"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0, x: -20 }}
                className="flex-1 flex flex-col min-h-0"
              >
                <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50 flex-none">
                  <div className="text-sm font-bold text-slate-700">Active Field Mapping</div>
                  <div className="flex gap-2">
                    <span className="text-[10px] bg-blue-100 text-blue-700 px-2 py-0.5 rounded font-bold tracking-tight">AUTOMAPPED: {mappings.length}</span>
                    <span className="text-[10px] bg-amber-100 text-amber-700 px-2 py-0.5 rounded font-bold tracking-tight">TOTAL FIELDS: {templateData.headers.length}</span>
                  </div>
                </div>

                <div className="flex-1 overflow-auto bg-white">
                  <table className="w-full text-left border-collapse min-w-[600px]">
                    <thead className="bg-slate-50 text-[10px] uppercase font-bold text-slate-500 sticky top-0 shadow-sm z-10">
                      <tr>
                        <th className="p-3 pl-6 border-b border-slate-200">Source Field (Your File)</th>
                        <th className="p-3 border-b border-slate-200 text-center w-12"></th>
                        <th className="p-3 border-b border-slate-200">Target Field (Post Office)</th>
                        <th className="p-3 border-b border-slate-200">Transformation</th>
                        <th className="p-3 pr-6 border-b border-slate-200 text-right">Status</th>
                      </tr>
                    </thead>
                    <tbody className="text-[11px] font-mono text-slate-600">
                      {templateData.headers.map((tCol) => {
                        const currentMapping = mappings.find(m => m.templateColumn === tCol);
                        const isMapped = !!currentMapping;
                        const isSender = ['sender', 'from', 'return', 'origin'].some(key => tCol.toLowerCase().includes(key));
                        const isReceiver = !isSender && ['rec', 'to', 'addressee', 'name', 'addr', 'city', 'state', 'pin', 'zip', 'line'].some(key => tCol.toLowerCase().includes(key));
                        
                        return (
                          <tr key={tCol} className={cn(
                            "hover:bg-slate-50 border-b border-slate-100 group",
                            isSender && "bg-slate-50/50"
                          )}>
                            <td className="p-3 pl-6">
                              <select 
                                className={cn(
                                  "w-full border rounded px-2 py-1 focus:outline-none transition-all appearance-none cursor-pointer",
                                  isMapped ? "bg-white border-blue-200" : "bg-slate-50 border-slate-100 italic",
                                  isSender && "opacity-60"
                                )}
                                value={currentMapping?.sourceColumn || ''}
                                onChange={(e) => {
                                  if (e.target.value === '') removeMapping(tCol);
                                  else addMapping(tCol, e.target.value);
                                }}
                              >
                                <option value="" className="text-slate-400 bg-white">[ SELECT SOURCE ]</option>
                                {sourceData.headers.map(h => (
                                  <option key={h} value={h} className="bg-white">{h}</option>
                                ))}
                              </select>
                            </td>
                            <td className="p-3 text-center">
                              <ArrowRight className={cn("w-3 h-3 mx-auto transition-colors", isMapped ? "text-blue-500" : "text-slate-200")} />
                            </td>
                            <td className="p-3">
                              <div className="flex items-center gap-2">
                                <span className={cn(
                                  "text-[8px] font-bold px-1 rounded",
                                  isSender ? "bg-slate-200 text-slate-500" : 
                                  isReceiver ? "bg-orange-100 text-orange-600" : "bg-blue-100 text-blue-600"
                                )}>
                                  {isSender ? "SENDER" : "RECEIVER"}
                                </span>
                                <div className={cn("font-bold", isSender ? "text-slate-500" : "text-slate-800")}>{tCol}</div>
                              </div>
                            </td>
                            <td className="p-3">
                              <div className="text-slate-400">
                                {tCol.toLowerCase().includes('name') ? "Proper Case" : 
                                 (tCol.toLowerCase().includes('pin') || tCol.toLowerCase().includes('zip')) ? "Numeric Only" : 
                                 "None"}
                              </div>
                            </td>
                            <td className="p-3 pr-6 text-right font-bold h-12">
                              {isMapped ? (
                                <span className="text-green-500 underline underline-offset-4 decoration-green-200">Matched</span>
                              ) : (
                                <span className="text-slate-300">Unset</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>


              </motion.div>
            ) : activeStep === 'preview' ? (
              <motion.div 
                key="preview-view"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                className="flex-1 flex flex-col min-h-0"
              >
                <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between bg-white flex-none">
                  <div className="flex items-center gap-4 flex-1">
                    <div className="text-sm font-bold text-slate-700 whitespace-nowrap">Data Review</div>
                    <div className="relative flex-1 max-w-sm">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                      <input 
                        type="text"
                        placeholder="Search addresses..."
                        className="w-full bg-slate-50 border border-slate-200 rounded px-9 py-1.5 text-xs focus:outline-none focus:border-blue-400 transition-all font-mono"
                        value={previewSearchQuery}
                        onChange={(e) => setPreviewSearchQuery(e.target.value)}
                      />
                      {previewSearchQuery && (
                        <button 
                          onClick={() => setPreviewSearchQuery('')}
                          className="absolute right-3 top-1/2 -translate-y-1/2"
                        >
                          <X className="w-3 h-3 text-slate-400 hover:text-slate-600" />
                        </button>
                      )}
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => setAllVisibleSelection(true)}
                      className="text-[10px] font-bold px-3 py-1 border border-slate-200 rounded text-slate-600 hover:bg-slate-50 transition-colors uppercase whitespace-nowrap"
                    >
                      Select All
                    </button>
                    <button 
                      onClick={() => setAllVisibleSelection(false)}
                      className="text-[10px] font-bold px-3 py-1 border border-slate-200 rounded text-slate-600 hover:bg-slate-50 transition-colors uppercase whitespace-nowrap"
                    >
                      Deselect All
                    </button>
                    <span className="text-[10px] bg-slate-100 text-slate-700 px-2 py-1 rounded font-bold tracking-tight whitespace-nowrap">SELECTED: {selectedRowIndices.size} / {sourceData?.rows.length}</span>
                  </div>
                </div>

                <div className="flex-1 overflow-auto bg-slate-50">
                  <table className="w-full text-left border-collapse min-w-[800px]">
                    <thead className="bg-white text-[10px] uppercase font-bold text-slate-500 sticky top-0 shadow-sm z-10">
                      <tr>
                        <th className="p-3 pl-6 border-b border-slate-200 w-12 text-center select-none">ID</th>
                        {mappings.map(m => (
                          <th key={m.templateColumn} className="p-3 border-b border-slate-200 text-[9px] tracking-widest">{m.templateColumn}</th>
                        ))}
                        <th className="p-3 pr-6 border-b border-slate-200 text-right w-16">Action</th>
                      </tr>
                    </thead>
                    <tbody className="text-[10px] font-mono text-slate-600">
                      <AnimatePresence initial={false}>
                        {filteredPreviewRows.map((row) => {
                          const isSelected = selectedRowIndices.has(row._sourceIndex);
                          return (
                            <motion.tr 
                              key={row._sourceIndex}
                              layout
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              className={cn(
                                "border-b border-slate-100 transition-colors",
                                isSelected ? "bg-white hover:bg-slate-50" : "bg-red-50/30 opacity-60 line-through grayscale text-slate-400"
                              )}
                            >
                              <td className="p-3 pl-6 font-bold text-slate-400 border-r border-slate-50 text-center">
                                {row._sourceIndex + 1}
                              </td>
                              {mappings.map(m => (
                                <td key={m.templateColumn} className="p-3 truncate max-w-[150px]">
                                  {row[m.templateColumn] || '—'}
                                </td>
                              ))}
                              <td className="p-3 pr-6 text-right">
                                <button 
                                  onClick={() => toggleRowSelection(row._sourceIndex)}
                                  className={cn(
                                    "p-1.5 rounded transition-colors group",
                                    isSelected ? "text-slate-300 hover:bg-red-50 hover:text-red-500" : "text-emerald-500 hover:bg-emerald-50"
                                  )}
                                >
                                  {isSelected ? <Trash2 className="w-3.5 h-3.5" /> : <Plus className="w-3.5 h-3.5" />}
                                </button>
                              </td>
                            </motion.tr>
                          );
                        })}
                      </AnimatePresence>
                      {filteredPreviewRows.length === 0 && (
                        <tr>
                          <td colSpan={mappings.length + 2} className="p-20 text-center text-slate-400 uppercase tracking-widest">
                            No matching records found
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                <div className="p-4 bg-slate-900 text-slate-400 font-mono text-[9px] uppercase flex justify-between">
                  <div>Showing {filteredPreviewRows.length} of {sourceData?.rows.length} total records</div>
                  <div className="text-blue-400">Tip: Red/strikethrough rows will be skipped during export</div>
                </div>
              </motion.div>
            ) : (
              <motion.div
                key="complete-view"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex-1 flex flex-col items-center justify-center p-12 text-center"
              >
                <div className="w-20 h-20 bg-green-50 rounded-full flex items-center justify-center mb-6">
                  <CheckCircle2 className="w-10 h-10 text-green-500" />
                </div>
                <h2 className="text-3xl font-bold text-slate-800 mb-2">Process Complete</h2>
                <p className="text-slate-500 mb-8 max-w-sm">
                  Your Excel file has been generated with {selectedRowIndices.size} selected rows.
                </p>
                <button 
                  onClick={reset}
                  className="px-8 py-3 bg-slate-800 text-white rounded font-bold hover:bg-slate-700 transition-all shadow-lg"
                >
                  Start New Session
                </button>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </main>

      {/* Footer */}
      <footer className="h-16 bg-white border-t border-slate-200 flex items-center justify-between px-6 shrink-0">
        <div className="text-xs text-slate-500 flex items-center gap-2">
          {sourceData && (
            <>
              Merged Output Target: 
              <span className="font-bold font-mono px-2 py-0.5 bg-slate-100 rounded border border-slate-200">
                output_batch_{new Date().toISOString().split('T')[0]}.xlsx
              </span>
            </>
          )}
        </div>
        <div className="flex gap-3">
          {activeStep === 'complete' ? (
             <button 
              onClick={reset}
              className="px-5 py-2 border border-slate-300 rounded text-sm font-semibold hover:bg-slate-50 text-slate-700 transition-colors"
            >
              Start New Batch
            </button>
          ) : activeStep === 'mapping' ? (
            <button 
              onClick={() => setActiveStep('preview')}
              disabled={mappings.length === 0}
              className={cn(
                "px-8 py-2 rounded text-sm font-bold shadow-lg transition-all",
                mappings.length === 0 
                  ? "bg-slate-200 text-slate-400 cursor-not-allowed shadow-none" 
                  : "bg-slate-800 text-white hover:bg-slate-700 shadow-slate-200"
              )}
            >
              Preview Processed Data
            </button>
          ) : activeStep === 'preview' && (
            <button 
              onClick={() => setActiveStep('mapping')}
              className="px-5 py-2 border border-slate-300 rounded text-sm font-semibold hover:bg-slate-50 text-slate-700 transition-colors"
            >
              Back to Mapping
            </button>
          )}

          {activeStep === 'complete' ? (
            <button 
              onClick={processAndDownload}
              className="px-8 py-2 bg-slate-800 text-white rounded text-sm font-bold shadow-lg shadow-slate-200"
            >
              Re-download Report
            </button>
          ) : (
            <button 
              onClick={activeStep === 'preview' ? processAndDownload : () => {}}
              disabled={!sourceData || !templateData || mappings.length === 0 || isProcessing || (activeStep === 'preview' && selectedRowIndices.size === 0)}
              className={cn(
                "px-8 py-2 rounded text-sm font-bold shadow-lg transition-all",
                (!sourceData || !templateData || mappings.length === 0 || (activeStep === 'preview' && selectedRowIndices.size === 0)) 
                  ? "bg-slate-200 text-slate-400 cursor-not-allowed shadow-none" 
                  : "bg-orange-600 text-white hover:bg-orange-700 shadow-orange-100",
                activeStep === 'mapping' && "hidden" // Hide run button in mapping, use preview button instead
              )}
            >
              {isProcessing ? 'Processing...' : 'Run Processing Task'}
            </button>
          )}
        </div>
      </footer>
    </div>
  );
}
