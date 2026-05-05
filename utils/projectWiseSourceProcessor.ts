import { read, write, utils } from 'xlsx';
import { toPng } from 'html-to-image';
import JSZip from 'jszip';
import { GeneratedImage, ProcessResponse } from '../types';
import { USER_PROJECT_MAPPING, USER_TEAM_MAPPING, DEFAULT_SITE } from './projectMapping';

// --- Helpers ---

function getCellValue(cell: any): string {
    if (cell === null || cell === undefined) return '';
    if (typeof cell === 'object' && cell.v !== undefined) return String(cell.v);
    return String(cell);
}

function findColumnIndex(headerRow: any[], possibleNames: string[]): number {
  return headerRow.findIndex(h => {
    if (!h) return false;
    const lowerH = getCellValue(h).toLowerCase().trim();
    return possibleNames.some(p => lowerH === p || lowerH.includes(p));
  });
}

const ALLOWED_SOURCES = [
  'WhatsApp',
  'Hoarding',
  'Incoming Call',
  'RRPL',
  'Website',
  'Website Bot',
  'Channel Partner',
  'Walk-In'
];

function determineSource(rawStrings: string[]): string | null {
  for (let str of rawStrings) {
    if (!str) continue;
    const s = str.toLowerCase();
    
    if (s.includes('channel partner') || s.includes('cp')) return 'Channel Partner';
    if (s.includes('walk-in') || s.includes('walk in') || s.includes('walkin')) return 'Walk-In';
    if (s.includes('whatsapp')) return 'WhatsApp';
    if (s.includes('hoarding') || s.includes('hoardings') || s.includes('site branding')) return 'Hoarding';
    if (s.includes('incoming call') || s.includes('incoming')) return 'Incoming Call';
    if (s.includes('rrpl')) return 'RRPL';
    if (s.includes('website bot') || s.includes('bot')) return 'Website Bot';
    if (s.includes('website')) return 'Website';
  }
  return null; // Ignore if it doesn't match the specific ones
}

async function generateSourceSummaryImage(
  siteName: string, 
  sourceStats: Record<string, number>,
  dateText: string
): Promise<string> {
  const container = document.createElement('div');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    width: '400px', 
    backgroundColor: '#ffffff', 
    padding: '15px', 
    fontFamily: "'Calibri', sans-serif",
    color: '#000000', 
    zIndex: '-9999',
    pointerEvents: 'none'
  });

  const totalCount = Object.values(sourceStats).reduce((a, b) => a + b, 0);

  // HTML Structure
  container.innerHTML = `
    <div style="background-color: #ffffff; width: 100%; border: 1px solid #000000; box-sizing: border-box;">
      <div style="padding: 12px 15px; background-color: #ffffff; text-align: center;">
        <div style="font-size: 16px; font-weight: 900; text-transform: uppercase;">PROJECT WISE SOURCE REPORT</div>
        <div style="font-size: 14px; font-weight: bold; margin-top: 4px;">SITE: ${siteName.toUpperCase()}</div>
        ${dateText ? `<div style="font-size: 12px; font-weight: bold; margin-top: 4px; color: #4b5563;">${dateText}</div>` : ''}
      </div>
      
      <div style="padding: 15px; background-color: #ffffff; border-top: 1px solid #000000;">
        <table style="width: 100%; border-collapse: collapse;">
          <thead>
            <tr>
              <th style="padding: 6px; text-align: left; border: 1px solid #000000; font-size: 13px; font-weight: 900; background-color: #f3f4f6; width: 70%;">Source</th>
              <th style="padding: 6px; text-align: center; border: 1px solid #000000; font-size: 13px; font-weight: 900; background-color: #f3f4f6;">Count</th>
            </tr>
          </thead>
          <tbody>
            ${ALLOWED_SOURCES.map(source => {
                const count = sourceStats[source] || 0;
                if (count === 0) return '';
                return `
                <tr>
                    <td style="padding: 6px; border: 1px solid #000000; font-size: 13px; text-align: left; font-weight: bold;">${source}</td>
                    <td style="padding: 6px; border: 1px solid #000000; font-size: 13px; text-align: center; font-weight: bold;">${count}</td>
                </tr>
                `;
            }).join('')}
            <tr>
              <td style="padding: 6px; border: 1px solid #000000; font-size: 13px; text-align: right; font-weight: 900; background-color: #f9fafb;">Total</td>
              <td style="padding: 6px; border: 1px solid #000000; font-size: 13px; text-align: center; font-weight: 900; background-color: #f9fafb;">${totalCount}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `;

  document.body.appendChild(container);
  try {
    await document.fonts.ready;
    await new Promise(resolve => setTimeout(resolve, 50)); 
    const dataUrl = await toPng(container, { quality: 1.0, pixelRatio: 3 });
    return dataUrl;
  } finally {
    document.body.removeChild(container);
  }
}

export async function processProjectWiseSourceFile(
  file: File,
  manualStartDate?: string,
  manualEndDate?: string
): Promise<ProcessResponse> {
  try {
    const data = await file.arrayBuffer();
    const workbook = read(data, { type: 'array', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Parse raw rows without skipping empty lines first to find header
    const rawData = utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: null }) as any[][];
    if (rawData.length === 0) throw new Error("Excel file is empty");

    // Preserve hyperlinks
    for(let R = 0; R < rawData.length; R++) {
        const row = rawData[R];
        if (!row || !Array.isArray(row)) continue;
        for(let C = 0; C < row.length; C++) {
             const cell_ref = utils.encode_cell({ c: C, r: R });
             const originalCell = worksheet[cell_ref];
             if (originalCell && originalCell.l && originalCell.l.Target) {
                 row[C] = { v: row[C] !== null ? row[C] : '', t: originalCell.t || 's', l: originalCell.l };
             }
        }
    }

    const maxColLength = Math.max(...rawData.map(r => (r && Array.isArray(r)) ? r.length : 0));

    let headerIndex = -1;
    let nameIdx = -1, phoneIdx = -1, projectIdx = -1, assignedToIdx = -1;
    let sourceIdx = -1, subSourceIdx = -1, subSubSourceIdx = -1, additionalSourceIdx = -1;
    let adminRemarkIdx = -1;

    for (let i = 0; i < Math.min(20, rawData.length); i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        nameIdx = findColumnIndex(row, ['name', 'lead name', 'client name']);
        phoneIdx = findColumnIndex(row, ['phone', 'mobile', 'contact', 'number']);
        projectIdx = findColumnIndex(row, ['project', 'site name', 'project name']);
        assignedToIdx = findColumnIndex(row, ['assigned to', 'user', 'owner']);
        
        sourceIdx = findColumnIndex(row, ['source', 'lead source']);
        subSourceIdx = findColumnIndex(row, ['sub source']);
        subSubSourceIdx = findColumnIndex(row, ['sub sub source']);
        additionalSourceIdx = findColumnIndex(row, ['additional source', 'additional source (info)']);
        adminRemarkIdx = findColumnIndex(row, ['admin remark', 'admin remarks', 'remark', 'admin comment']);
        
        if (assignedToIdx !== -1) {
            headerIndex = i;
            break;
        }
    }

    if (headerIndex === -1 || assignedToIdx === -1) {
       throw new Error("Could not find required 'Assigned To' header in the first 20 rows.");
    }

    const normalizedMapping: Record<string, string> = {};
    Object.keys(USER_PROJECT_MAPPING).forEach(k => {
      normalizedMapping[k.toLowerCase().trim()] = USER_PROJECT_MAPPING[k];
    });

    const sites: Record<string, { rawRows: any[][], stats: Record<string, number> }> = {};

    for (let i = headerIndex + 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        // Filter out 'test'
        const isExcluded = row.some(cell => {
             if (!cell) return false;
             const s = getCellValue(cell).toLowerCase();
             return s.includes('metro') || s.includes('test') || s.includes('ramesh bodke');
        });
        if (isExcluded) continue;

        const rawAssigned = row[assignedToIdx];
        const assignedStr = rawAssigned ? getCellValue(rawAssigned).trim() : "Unassigned";
        const assignedLower = assignedStr.toLowerCase();

        // 1. First, check the Project Column
        let siteName = DEFAULT_SITE;
        const rawProject = projectIdx !== -1 ? getCellValue(row[projectIdx]).toLowerCase().trim() : '';
        if (rawProject.includes('kairos')) siteName = 'Kairos';
        else if (rawProject.includes('aqua') || rawProject.includes('aqualife')) siteName = 'Aqua Life';
        else if (rawProject.includes('milestone')) siteName = 'Milestone';
        else if (rawProject.includes('statement')) siteName = 'Statement';
        else if (rawProject.includes('ekam')) siteName = 'Legacy Ekam';

        // 2. If Project Column didn't help (or was missing), try mapping
        let matchedUserKey = Object.keys(normalizedMapping).find(k => {
            return assignedLower === k || assignedLower.includes(k) || k.includes(assignedLower);
        });
        
        if (siteName === DEFAULT_SITE) {
            if (matchedUserKey) {
            siteName = normalizedMapping[matchedUserKey];
            }
        }

        const sourceVals = [
            sourceIdx !== -1 ? getCellValue(row[sourceIdx]) : '',
            subSourceIdx !== -1 ? getCellValue(row[subSourceIdx]) : '',
            subSubSourceIdx !== -1 ? getCellValue(row[subSubSourceIdx]) : '',
            additionalSourceIdx !== -1 ? getCellValue(row[additionalSourceIdx]) : ''
        ];

        const determinedSource = determineSource(sourceVals);
        
        if (!sites[siteName]) {
            const headerRow = [...rawData[headerIndex]];
            while (headerRow.length < maxColLength) headerRow.push('');
            headerRow[maxColLength] = 'Determined Source';
            sites[siteName] = { 
                rawRows: [headerRow], // Push header
                stats: {} 
            };
        }
        
        const newRow = [...row];
        while (newRow.length < maxColLength) newRow.push('');
        newRow[maxColLength] = determinedSource || '';
        
        sites[siteName].rawRows.push(newRow);
        
        if (determinedSource) {
            if (!sites[siteName].stats[determinedSource]) {
                sites[siteName].stats[determinedSource] = 0;
            }
            sites[siteName].stats[determinedSource]++;
        }
    }

    const images: GeneratedImage[] = [];
    const zip = new JSZip();

    let dateText = "Total Source Summary";
    if (manualStartDate && manualEndDate) {
      const sDate = new Date(manualStartDate);
      const eDate = new Date(manualEndDate);
      const sFormat = `${sDate.getDate().toString().padStart(2, '0')}-${(sDate.getMonth()+1).toString().padStart(2, '0')}-${sDate.getFullYear()}`;
      const eFormat = `${eDate.getDate().toString().padStart(2, '0')}-${(eDate.getMonth()+1).toString().padStart(2, '0')}-${eDate.getFullYear()}`;
      dateText = `From: ${sFormat}  To: ${eFormat}`;
    } else if (manualStartDate) {
       const sDate = new Date(manualStartDate);
       const sFormat = `${sDate.getDate().toString().padStart(2, '0')}-${(sDate.getMonth()+1).toString().padStart(2, '0')}-${sDate.getFullYear()}`;
       dateText = `Date: ${sFormat}`;
    } else if (manualEndDate) {
       const eDate = new Date(manualEndDate);
       const eFormat = `${eDate.getDate().toString().padStart(2, '0')}-${(eDate.getMonth()+1).toString().padStart(2, '0')}-${eDate.getFullYear()}`;
       dateText = `Date: ${eFormat}`;
    }

    for (const siteName of Object.keys(sites)) {
        if (siteName === DEFAULT_SITE) continue;
        
        const siteData = sites[siteName];
        if (siteData.rawRows.length <= 1) continue; // Only header
        
        // Ensure all allowed sources are in stats even if 0
        const totalSourcesMatched = Object.values(siteData.stats).reduce((a, b) => a + b, 0);
        
        if (totalSourcesMatched > 0) {
            const dataUrl = await generateSourceSummaryImage(siteName, siteData.stats, dateText);
            const title = `${siteName} Source Summary`;
            images.push({ project_name: title, image_url: dataUrl, filename: `${siteName}_source_summary.png` });
        }
        
        const wb = utils.book_new();
        const ws = utils.aoa_to_sheet(siteData.rawRows);
        utils.book_append_sheet(wb, ws, "Source Data");
        const buf = write(wb, { bookType: 'xlsx', type: 'array' });
        zip.file(`${siteName} Source Data.xlsx`, buf);
    }
    
    // Add leftover General Project
    if (sites[DEFAULT_SITE] && sites[DEFAULT_SITE].rawRows.length > 1) {
       const wb = utils.book_new();
       const ws = utils.aoa_to_sheet(sites[DEFAULT_SITE].rawRows);
       utils.book_append_sheet(wb, ws, "Source Data");
       const buf = write(wb, { bookType: 'xlsx', type: 'array' });
       zip.file(`Uncategorized General Source Data.xlsx`, buf);
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    return { images, zip_url: URL.createObjectURL(zipBlob), message: "Success" };
  } catch (error: any) {
    throw new Error(`Project Wise Source Report Error: ${error.message}`);
  }
}
