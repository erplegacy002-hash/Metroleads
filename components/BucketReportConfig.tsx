import React, { useState, useEffect, useRef } from 'react';
import { 
  Check, 
  ChevronsUpDown, 
  Search, 
  X, 
  Filter, 
  RotateCcw, 
  Users, 
  Layers, 
  Eye, 
  Calendar,
  Sparkles,
  Info
} from 'lucide-react';
import { 
  BucketReportAnalysis, 
  computeBucketReportTable, 
  BucketTableSummary, 
  DEFAULT_BUCKET_LIST,
  formatToDDMMYYYY 
} from '../utils/bucketSiteVisitProcessor';

interface BucketReportConfigProps {
  analysis: BucketReportAnalysis;
  file: File;
  reportTitle: string;
  onReportTitleChange: (title: string) => void;
  selectedSalesColIdx: number;
  onSalesColIdxChange: (idx: number) => void;
  selectedBucketColIdx: number;
  onBucketColIdxChange: (idx: number) => void;
  selectedAgencyColIdx?: number;
  onAgencyColIdxChange?: (idx: number) => void;
  selectedSalesUsers: string[];
  onSelectedSalesUsersChange: (users: string[]) => void;
  selectedBuckets: string[];
  onSelectedBucketsChange: (buckets: string[]) => void;
  columnFilters: Record<number, string[]>;
  onColumnFiltersChange: (filters: Record<number, string[]>) => void;
  onResetToDefaultTitle: () => void;
  reportType?: 'sales' | 'presales';
}

const BucketReportConfig: React.FC<BucketReportConfigProps> = ({
  analysis,
  file,
  reportTitle,
  onReportTitleChange,
  selectedSalesColIdx,
  onSalesColIdxChange,
  selectedBucketColIdx,
  onBucketColIdxChange,
  selectedAgencyColIdx = -1,
  onAgencyColIdxChange,
  selectedSalesUsers,
  onSelectedSalesUsersChange,
  selectedBuckets,
  onSelectedBucketsChange,
  columnFilters,
  onColumnFiltersChange,
  onResetToDefaultTitle,
  reportType = 'sales',
}) => {
  const isPresales = reportType === 'presales';
  const [openDropdownColIdx, setOpenDropdownColIdx] = useState<number | null>(null);
  const [searchTerms, setSearchTerms] = useState<Record<number, string>>({});
  const [previewSummary, setPreviewSummary] = useState<BucketTableSummary | null>(null);
  const [isPreviewLoading, setIsPreviewLoading] = useState(false);

  const dropdownRef = useRef<HTMLDivElement | null>(null);

  // Close dropdown on outside click
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setOpenDropdownColIdx(null);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Update live preview whenever filters, buckets, sales users, agency, or title change
  useEffect(() => {
    let isMounted = true;
    setIsPreviewLoading(true);

    computeBucketReportTable(file, {
      salesColIdx: selectedSalesColIdx,
      bucketColIdx: selectedBucketColIdx,
      agencyColIdx: selectedAgencyColIdx,
      selectedSales: selectedSalesUsers,
      selectedBuckets: selectedBuckets,
      columnFilters,
      reportTitle: reportTitle || analysis.defaultTitle,
      headerIndex: analysis.headerIndex,
      reportType,
      dimensionLabel: isPresales ? 'Telecaller' : 'Sales'
    })
      .then(summary => {
        if (isMounted) {
          setPreviewSummary(summary);
          setIsPreviewLoading(false);
        }
      })
      .catch(err => {
        console.error("Error computing preview summary:", err);
        if (isMounted) setIsPreviewLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [
    file,
    selectedSalesColIdx,
    selectedBucketColIdx,
    selectedAgencyColIdx,
    selectedSalesUsers,
    selectedBuckets,
    columnFilters,
    reportTitle,
    analysis.headerIndex,
    analysis.defaultTitle,
    reportType,
    isPresales
  ]);

  const toggleValueForColumn = (colIdx: number, val: string) => {
    if (colIdx === selectedSalesColIdx) {
      if (selectedSalesUsers.includes(val)) {
        onSelectedSalesUsersChange(selectedSalesUsers.filter(u => u !== val));
      } else {
        onSelectedSalesUsersChange([...selectedSalesUsers, val]);
      }
    } else if (colIdx === selectedBucketColIdx) {
      if (selectedBuckets.includes(val)) {
        onSelectedBucketsChange(selectedBuckets.filter(b => b !== val));
      } else {
        onSelectedBucketsChange([...selectedBuckets, val]);
      }
    } else {
      const current = columnFilters[colIdx] || [];
      if (current.includes(val)) {
        onColumnFiltersChange({
          ...columnFilters,
          [colIdx]: current.filter(item => item !== val)
        });
      } else {
        onColumnFiltersChange({
          ...columnFilters,
          [colIdx]: [...current, val]
        });
      }
    }
  };

  const selectAllForColumn = (colIdx: number, allVals: string[]) => {
    if (colIdx === selectedSalesColIdx) {
      onSelectedSalesUsersChange(allVals);
    } else if (colIdx === selectedBucketColIdx) {
      onSelectedBucketsChange(allVals);
    } else {
      onColumnFiltersChange({
        ...columnFilters,
        [colIdx]: allVals
      });
    }
  };

  const deselectAllForColumn = (colIdx: number) => {
    if (colIdx === selectedSalesColIdx) {
      onSelectedSalesUsersChange([]);
    } else if (colIdx === selectedBucketColIdx) {
      onSelectedBucketsChange([]);
    } else {
      onColumnFiltersChange({
        ...columnFilters,
        [colIdx]: []
      });
    }
  };

  const getSelectedValuesForCol = (colIdx: number): string[] => {
    if (colIdx === selectedSalesColIdx) {
      return selectedSalesUsers;
    }
    if (colIdx === selectedBucketColIdx) {
      return selectedBuckets;
    }
    return columnFilters[colIdx] || [];
  };

  return (
    <div className="w-full max-w-5xl mx-auto space-y-6 mb-8 animate-in fade-in duration-300" ref={dropdownRef}>
      {/* 1. Report Title Customization Card */}
      <div className="bg-white border border-amber-200/80 rounded-xl p-5 sm:p-6 shadow-sm">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-3">
          <label className="block text-sm font-bold text-slate-800 font-serif flex items-center gap-2">
            <Sparkles className="w-4 h-4 text-[#d4af37]" />
            Report Title (Prefilled with Report Format)
          </label>
          <button
            type="button"
            onClick={onResetToDefaultTitle}
            className="text-xs font-semibold text-amber-700 hover:text-amber-800 flex items-center gap-1.5 self-start sm:self-auto py-1 px-2.5 rounded bg-amber-50 hover:bg-amber-100 transition-colors border border-amber-200"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Reset to Default Format
          </button>
        </div>
        <input
          type="text"
          value={reportTitle}
          onChange={(e) => onReportTitleChange(e.target.value)}
          placeholder={isPresales ? "e.g. Leads Report | Legacy Ekam | Week Report (07-09-2026 to 13-09-2026)" : "e.g. Site Visits Report | Legacy Ekam | Week Report (07-09-2026 to 13-09-2026)"}
          className="w-full px-4 py-2.5 text-sm sm:text-base font-medium border border-slate-300 rounded-lg focus:ring-2 focus:ring-[#d4af37] focus:border-[#d4af37] text-slate-900 bg-slate-50/50"
        />
        <div className="mt-2.5 flex items-start gap-2 text-xs text-slate-500">
          <Info className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
          <span>
            Default format: <strong>{isPresales ? 'Leads Report' : 'Site Visits Report'} | {analysis.detectedProject || 'Legacy Ekam'} | Week Report ({analysis.detectedStartDate} to {analysis.detectedEndDate})</strong>. You can customize this title above as needed.
          </span>
        </div>
      </div>

      {/* 2. Detected Columns Having Most Records as Same */}
      <div className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-sm">
        <div className="border-b border-slate-100 pb-4 mb-5">
          <h3 className="text-lg font-serif font-bold text-slate-900 flex items-center gap-2">
            <Layers className="w-5 h-5 text-[#d4af37]" />
            Detected Fields with Categorical Records
          </h3>
          <p className="text-xs sm:text-sm text-slate-500 mt-1">
            The system detected the following columns having repeated/matching records. Use the multiselect dropdowns below to select which records to include in the generated report.
          </p>
        </div>

        {/* Primary Role Mapping Controls */}
        <div className={`grid grid-cols-1 ${isPresales ? 'sm:grid-cols-3' : 'sm:grid-cols-2'} gap-4 mb-6 p-4 bg-amber-50/40 border border-amber-200/60 rounded-lg`}>
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
              <Users className="w-4 h-4 text-slate-600" />
              Row Dimension: {isPresales ? 'Telecaller' : 'Sales Column'}
            </label>
            <select
              value={selectedSalesColIdx}
              onChange={(e) => {
                const newIdx = parseInt(e.target.value, 10);
                onSalesColIdxChange(newIdx);
                const col = analysis.detectedColumns.find(c => c.colIndex === newIdx);
                if (col) onSelectedSalesUsersChange(col.uniqueValues);
              }}
              className="w-full text-xs sm:text-sm font-medium border border-slate-300 rounded-md px-3 py-2 bg-white text-slate-800 focus:ring-[#d4af37] focus:border-[#d4af37]"
            >
              {analysis.detectedColumns.map(col => (
                <option key={col.colIndex} value={col.colIndex}>
                  {col.headerName} ({col.uniqueValues.length} unique values)
                </option>
              ))}
            </select>
          </div>

          {isPresales && (
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                <Layers className="w-4 h-4 text-purple-600" />
                Sub-Row Dimension: Agency Name
              </label>
              <select
                value={selectedAgencyColIdx}
                onChange={(e) => {
                  const newIdx = parseInt(e.target.value, 10);
                  if (onAgencyColIdxChange) onAgencyColIdxChange(newIdx);
                }}
                className="w-full text-xs sm:text-sm font-medium border border-slate-300 rounded-md px-3 py-2 bg-white text-slate-800 focus:ring-[#d4af37] focus:border-[#d4af37]"
              >
                <option value={-1}>None (Single Row)</option>
                {analysis.detectedColumns.map(col => (
                  <option key={col.colIndex} value={col.colIndex}>
                    {col.headerName} ({col.uniqueValues.length} unique agencies)
                  </option>
                ))}
              </select>
              <p className="text-[11px] text-purple-800 mt-1 flex items-center gap-1 font-medium">
                <Check className="w-3 h-3 text-purple-600 shrink-0" />
                Blank agency records are categorized as "(blank)"
              </p>
            </div>
          )}

          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
              <Filter className="w-4 h-4 text-slate-600" />
              Column Dimension: {isPresales ? 'AI Lead Level / Buckets' : 'Buckets / Stages'}
            </label>
            <select
              value={selectedBucketColIdx}
              onChange={(e) => {
                const newIdx = parseInt(e.target.value, 10);
                onBucketColIdxChange(newIdx);
                const col = analysis.detectedColumns.find(c => c.colIndex === newIdx);
                if (col) onSelectedBucketsChange(col.uniqueValues);
              }}
              className="w-full text-xs sm:text-sm font-medium border border-slate-300 rounded-md px-3 py-2 bg-white text-slate-800 focus:ring-[#d4af37] focus:border-[#d4af37]"
            >
              {analysis.detectedColumns.map(col => (
                <option key={col.colIndex} value={col.colIndex}>
                  {col.headerName} ({col.uniqueValues.length} unique stages)
                </option>
              ))}
            </select>
            {analysis.detectedColumns.find(c => c.colIndex === selectedBucketColIdx && (c.headerName.toLowerCase().includes('enquiry level') || c.headerName.toLowerCase().includes('lead level') || c.role === 'bucket')) && (
              <p className="text-[11px] text-amber-800 mt-1 flex items-center gap-1 font-medium">
                <Check className="w-3 h-3 text-emerald-600 shrink-0" />
                Blank Enquiry/Lead Level records are counted as "Open"
              </p>
            )}
          </div>
        </div>

        {/* Detected Columns Grid with Multiselect Dropdowns */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {analysis.detectedColumns.map(col => {
            const isSales = col.colIndex === selectedSalesColIdx;
            const isAgency = col.colIndex === selectedAgencyColIdx;
            const isBucket = col.colIndex === selectedBucketColIdx;
            const selectedVals = getSelectedValuesForCol(col.colIndex);
            const isOpen = openDropdownColIdx === col.colIndex;
            const term = searchTerms[col.colIndex] || '';

            const filteredVals = col.uniqueValues.filter(v => 
              v.toLowerCase().includes(term.toLowerCase())
            );

            let roleBadge = 'Filter';
            let roleBadgeColor = 'bg-slate-100 text-slate-700 border-slate-200';
            if (isSales) {
              roleBadge = isPresales ? 'Telecaller Rows' : 'Sales Rows';
              roleBadgeColor = 'bg-blue-50 text-blue-800 border-blue-200 font-semibold';
            } else if (isAgency) {
              roleBadge = 'Agency Sub-Rows';
              roleBadgeColor = 'bg-purple-50 text-purple-800 border-purple-200 font-semibold';
            } else if (isBucket) {
              roleBadge = isPresales ? 'AI Lead Level' : 'Bucket Columns';
              roleBadgeColor = 'bg-amber-100 text-amber-900 border-amber-300 font-semibold';
            } else if (col.role === 'project') {
              roleBadge = 'Project Filter';
              roleBadgeColor = 'bg-emerald-50 text-emerald-800 border-emerald-200';
            }

            return (
              <div 
                key={col.colIndex}
                className={`relative border rounded-lg p-3.5 bg-white transition-all ${
                  isOpen ? 'ring-2 ring-[#d4af37] border-transparent shadow-md' : 'border-slate-200 hover:border-slate-300'
                }`}
              >
                {/* Field Header */}
                <div className="flex items-center justify-between gap-2 mb-2">
                  <div className="min-w-0">
                    <span className="text-xs font-bold text-slate-900 block truncate" title={col.headerName}>
                      {col.headerName}
                    </span>
                    <span className="text-[11px] text-slate-400">
                      {col.totalCount} records • {col.uniqueValues.length} distinct
                    </span>
                  </div>
                  <span className={`text-[10px] px-2 py-0.5 rounded-full border shrink-0 ${roleBadgeColor}`}>
                    {roleBadge}
                  </span>
                </div>

                {/* Multiselect Trigger Button */}
                <button
                  type="button"
                  onClick={() => setOpenDropdownColIdx(isOpen ? null : col.colIndex)}
                  className="w-full flex items-center justify-between px-3 py-2 text-xs bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-md text-slate-800 font-medium transition-colors"
                >
                  <span className="truncate">
                    {selectedVals.length === 0
                      ? 'None selected'
                      : selectedVals.length === col.uniqueValues.length
                        ? `All (${selectedVals.length}) selected`
                        : `${selectedVals.length} of ${col.uniqueValues.length} selected`}
                  </span>
                  <ChevronsUpDown className="w-3.5 h-3.5 text-slate-400 shrink-0 ml-1.5" />
                </button>

                {/* Multiselect Dropdown Panel */}
                {isOpen && (
                  <div className="absolute left-0 right-0 top-full mt-1.5 bg-white border border-slate-300 rounded-lg shadow-xl z-50 p-3 animate-in fade-in duration-150">
                    {/* Search inside dropdown */}
                    <div className="relative mb-2">
                      <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Filter records..."
                        value={term}
                        onChange={(e) => setSearchTerms({ ...searchTerms, [col.colIndex]: e.target.value })}
                        className="w-full pl-8 pr-2 py-1.5 text-xs border border-slate-200 rounded focus:ring-1 focus:ring-[#d4af37] focus:border-[#d4af37]"
                      />
                    </div>

                    {/* Action buttons */}
                    <div className="flex items-center justify-between border-b border-slate-100 pb-2 mb-2">
                      <span className="text-[11px] font-semibold text-slate-500">
                        {selectedVals.length} / {col.uniqueValues.length} selected
                      </span>
                      <div className="space-x-2">
                        <button
                          type="button"
                          onClick={() => selectAllForColumn(col.colIndex, col.uniqueValues)}
                          className="text-[11px] font-bold text-amber-700 hover:text-amber-800 hover:underline"
                        >
                          Select All
                        </button>
                        <span className="text-slate-300">|</span>
                        <button
                          type="button"
                          onClick={() => deselectAllForColumn(col.colIndex)}
                          className="text-[11px] font-bold text-slate-500 hover:text-slate-700 hover:underline"
                        >
                          Clear
                        </button>
                      </div>
                    </div>

                    {/* Scrollable Checkbox List */}
                    <div className="max-h-48 overflow-y-auto space-y-1 pr-1 scrollbar-thin text-xs">
                      {filteredVals.length === 0 ? (
                        <div className="py-3 text-center text-slate-400 text-xs italic">
                          No matching records
                        </div>
                      ) : (
                        filteredVals.map(val => {
                          const isChecked = selectedVals.includes(val);
                          const count = col.valueCounts[val] || 0;
                          return (
                            <label
                              key={val}
                              className={`flex items-center justify-between px-2 py-1.5 rounded cursor-pointer transition-colors ${
                                isChecked ? 'bg-amber-50/60 font-semibold text-slate-900' : 'hover:bg-slate-50 text-slate-600'
                              }`}
                            >
                              <div className="flex items-center space-x-2 truncate">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => toggleValueForColumn(col.colIndex, val)}
                                  className="h-3.5 w-3.5 rounded border-slate-300 text-[#d4af37] focus:ring-[#d4af37]"
                                />
                                <span className="truncate">{val}</span>
                              </div>
                              <span className="text-[10px] text-slate-400 px-1.5 py-0.5 bg-slate-100 rounded shrink-0 ml-2">
                                {count}
                              </span>
                            </label>
                          );
                        })
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* 3. Live Interactive Preview of Report Format */}
      <div className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-sm overflow-hidden">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-100 pb-3 mb-4">
          <div className="flex items-center gap-2">
            <Eye className="w-5 h-5 text-[#d4af37]" />
            <h4 className="text-base font-serif font-bold text-slate-900">
              Live Report Preview (Matches Report Output Format)
            </h4>
          </div>
          <span className="text-xs text-slate-500 bg-slate-100 px-2.5 py-1 rounded-full font-medium">
            {previewSummary ? `${previewSummary.rows.length} Sales Executives • ${previewSummary.buckets.length} Buckets` : 'Updating...'}
          </span>
        </div>

        {/* Styled Table matching image.png */}
        <div className="overflow-x-auto pb-2">
          {previewSummary && previewSummary.rows.length > 0 ? (
            <div className="inline-block min-w-full align-middle">
              <table className="border-collapse border-2 border-black font-sans text-black w-full bg-white text-xs sm:text-sm">
                <thead>
                  {/* Row 1: Merged Title matching image */}
                  <tr>
                    <th 
                      colSpan={2 + previewSummary.buckets.length} 
                      className="border border-black px-3 py-2 text-center font-bold text-black bg-white tracking-wide text-sm sm:text-base"
                    >
                      {previewSummary.reportTitle}
                    </th>
                  </tr>
                  {/* Row 2: Column Headers matching image */}
                  <tr>
                    <th className="border border-black px-3 py-1.5 text-left font-bold text-black bg-white min-w-[140px]">
                      {previewSummary.dimensionLabel || (isPresales ? 'Telecaller' : 'Sales')}
                    </th>
                    <th className="border border-black px-3 py-1.5 text-center font-bold text-black bg-white min-w-[90px] whitespace-nowrap">
                      Grand Total
                    </th>
                    {previewSummary.buckets.map(b => (
                      <th key={b} className="border border-black px-2.5 py-1.5 text-center font-bold text-black bg-white whitespace-nowrap">
                        {b}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewSummary.rows.map((row) => {
                    if (row.subRows && row.subRows.length > 0) {
                      return (
                        <React.Fragment key={row.salesUser}>
                          <tr className="bg-slate-50/80">
                            <td className="border border-black px-3 py-1.5 text-left font-bold text-black whitespace-nowrap">
                              {row.salesUser}
                            </td>
                            <td className="border border-black px-3 py-1.5 text-center font-bold text-black">
                              {row.grandTotal}
                            </td>
                            {previewSummary.buckets.map(b => (
                              <td key={b} className="border border-black px-2.5 py-1.5 text-center font-bold text-black">
                                {row.bucketCounts[b] || 0}
                              </td>
                            ))}
                          </tr>
                          {row.subRows.map(sub => (
                            <tr key={`${row.salesUser}-${sub.agency}`} className="hover:bg-slate-50/50">
                              <td className="border border-black px-3 py-1 text-left font-normal text-slate-800 whitespace-nowrap">
                                {sub.agency}
                              </td>
                              <td className="border border-black px-3 py-1 text-center font-normal text-slate-800">
                                {sub.grandTotal}
                              </td>
                              {previewSummary.buckets.map(b => (
                                <td key={b} className="border border-black px-2.5 py-1 text-center font-normal text-slate-800">
                                  {sub.bucketCounts[b] || 0}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </React.Fragment>
                      );
                    }

                    return (
                      <tr key={row.salesUser} className="hover:bg-slate-50/50">
                        <td className="border border-black px-3 py-1.5 text-left font-normal text-black whitespace-nowrap">
                          {row.salesUser}
                        </td>
                        <td className="border border-black px-3 py-1.5 text-center font-normal text-black">
                          {row.grandTotal}
                        </td>
                        {previewSummary.buckets.map(b => (
                          <td key={b} className="border border-black px-2.5 py-1.5 text-center font-normal text-black">
                            {row.bucketCounts[b] || 0}
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                  {/* Footer Row: Grand Total matching image */}
                  <tr className="font-bold bg-white">
                    <td className="border border-black px-3 py-1.5 text-left font-bold text-black">
                      Grand Total
                    </td>
                    <td className="border border-black px-3 py-1.5 text-center font-bold text-black">
                      {previewSummary.columnTotals.grandTotal}
                    </td>
                    {previewSummary.buckets.map(b => (
                      <td key={b} className="border border-black px-2.5 py-1.5 text-center font-bold text-black">
                        {previewSummary.columnTotals.bucketTotals[b] || 0}
                      </td>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>
          ) : (
            <div className="py-8 text-center text-slate-400 italic text-sm">
              {isPreviewLoading ? 'Calculating preview table...' : `No ${isPresales ? 'presales' : 'sales'} records match the current selections.`}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default BucketReportConfig;
