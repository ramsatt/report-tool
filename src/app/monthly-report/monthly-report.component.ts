import { Component, OnInit, AfterViewInit, ChangeDetectorRef, HostListener, NgZone } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import jspdf from 'jspdf';
import html2canvas from 'html2canvas';
import * as htmlToImage from 'html-to-image';
import PptxGenJS from 'pptxgenjs';
import * as XLSX from 'xlsx-js-style';

declare var lucide: any;

@Component({
  selector: 'app-monthly-report',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './monthly-report.component.html',
  styleUrls: ['./monthly-report.component.css']
})
export class MonthlyReportComponent implements OnInit, AfterViewInit {
  activeTab: 'input' | 'preview' = 'input';
  currentMonth = 'May';
  currentYear = '2026';

  // Core Highlights
  coreHighlights: string[] = [];

  // App Highlights
  appHighlights: string[] = [];

  // Core Delivery Data
  coreDeliveryData = {
    sprintMonth: '',
    committed: 0,
    delivered: 0,
    achieved: '',
    features: [],
    deploymentStatus: '',
    bugs: 0,
    comments: ''
  };

  coreDeliveryDataRows: any[] = [];

  // App Delivery Data
  appDeliveryDataRows: any[] = [];

  // Migration Status Data (Block 3)
  migrationData: any[] = [
    { module: 'Home Dashboard', start: '19-Mar-25', end: '01-Apr-25', pct: '100%', status: 'Completed', comments: 'Demo done in DEV env' },
    { module: 'User Specific Screen', start: '19-Mar-25', end: '13-May-25', pct: '100%', status: 'Completed', comments: 'Demo done in DEV env' },
    { module: 'Custom Field', start: '30-Apr-25', end: '27-May-25', pct: '100%', status: 'Completed', comments: 'Demo done in DEV env' },
    { module: 'Template', start: '14-May-25', end: '22-Jul-25', pct: '100%', status: 'Completed', comments: 'Demo done in INT env' },
    { module: 'Fleet List', start: '13-Oct-25', end: '28-Oct-25', pct: '100%', status: 'Completed', comments: 'Demo done in INT env' },
    { module: 'Digital Factory', start: '19-Aug-25', end: '16-Sep-25', pct: '100%', status: 'Completed', comments: 'Demo done in INT env' },
    { module: 'Filtered Data Report', start: '17-Sep-25', end: '30-Sep-25', pct: '100%', status: 'Completed', comments: 'Demo done in INT env' },
    { module: 'Chart', start: '29-Oct-25', end: '11-Nov-25', pct: '100%', status: 'Completed', comments: 'Demo done in INT env' },
    { module: 'Report Group', start: '12-Nov-25', end: '25-Nov-25', pct: '100%', status: 'Completed', comments: 'Demo Pending' }
  ];

  // Feedback Data
  feedbackData: any[] = [];

  // 6 Month Metrics Data
  sixMonthMetrics: any[] = [];

  // Delivery Strategy Data (Slide 7)
  deliveryStrategyData: any[] = [];

  // Defect Analysis Data (Slide 8)
  defectAnalysisData: any[] = [];

  // Defect Backlog Data (Slide 9)
  defectBacklogData: any[] = [];

  // Defect Metrics Data (Slide 10)
  defectMetricsData: any[] = [
      { title: 'Not Triage', backlog: 10, sprints: {}, type: 'critical' },
      { title: 'Old/Historic Bug', backlog: null, sprints: {9: 1, 10: 2, 15: 1, 22: 2, 23: 2}, type: 'info' },
      { title: 'Data Error', backlog: null, sprints: {15: 8, 16: 1, 19: 1}, type: 'info' },
      { title: 'Sprint 18(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 19(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 20(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 21(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 22(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 23(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 24(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 25(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 26(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 27(2024)', backlog: null, sprints: {}, type: 'info' },
      { title: 'Sprint 1(2025)', backlog: null, sprints: {9: 1}, type: 'info' }
  ];

  // Leave Plan Data
  leavePlanData: any[] = [];

  // Team Actions Data
  teamActionsData: any[] = [];
  
  // Excel Processing
  onFileChange(evt: any) {
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      this.ngZone.run(() => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      // 1. General Info
      const wsGeneral = wb.Sheets['General Info'];
      if(wsGeneral) {
          const data = XLSX.utils.sheet_to_json(wsGeneral, {header: 1}) as any[][];
          this.coreHighlights = [];
          this.appHighlights = [];
          
          let currentSection = '';
          data.forEach(row => {
             const key = (row[0] || '').toString().trim();
             const val = row[1];

             // Section detection
             if(key.toLowerCase().includes('core highlights')) { currentSection = 'core'; return; }
             if(key.toLowerCase().includes('app highlights')) { currentSection = 'app'; return; }
             
             // Update Month/Year if present (key-value pairs)
             if(key === 'Month' && val) this.currentMonth = val;
             if(key === 'Year' && val) this.currentYear = val;
             
             // Value extraction (Relaxed: Take any non-empty value in Col B if we are in a section)
             if (val && typeof val === 'string' && val.trim().length > 0) {
                 // Exclude known metadata keys if they happen to share the section (unlikely but safe)
                 if (key !== 'Month' && key !== 'Year') {
                     if(currentSection === 'core') this.coreHighlights.push(val);
                     else if(currentSection === 'app') this.appHighlights.push(val);
                 }
             }
          });
      }

      // 2. Delivery Metrics
      const wsDelivery = wb.Sheets['Delivery Metrics'];
      if(wsDelivery) {
        const data = XLSX.utils.sheet_to_json(wsDelivery) as any[];
        // Filter into Core and App
        this.coreDeliveryDataRows = [];
        this.appDeliveryDataRows = [];
        
        data.forEach(row => {
            // Reliable Percentage Calculation
            const committed = row['Committed'] || 0;
            const delivered = row['Delivered'] || 0;
            let achievedVal = 0;
            
            if (committed > 0) {
                achievedVal = Math.round((delivered / committed) * 100);
            } else if (row['Delivery %']) {
                // Fallback: If < 1 (e.g. 0.85), assume fraction -> 85. If > 1 (e.g. 85), assume whole number.
                const raw = row['Delivery %'];
                if (typeof raw === 'number') {
                    achievedVal = raw <= 1 ? Math.round(raw * 100) : Math.round(raw);
                } else {
                    achievedVal = parseInt(raw.toString().replace('%', '').trim()) || 0;
                }
            }

            const item = {
                sprintMonth: `${row['Sprint'] || ''}\n(${this.formatExcelDate(row['Month'])})`,
                committed: committed,
                delivered: delivered,
                achieved: achievedVal + '%', 
                features: (row['Features Delivered'] || '').split(/,|\n/).map((s: string) => s.trim()).filter((s: string) => s),
                deploymentStatus: row['Deployment Status'] || 'No Deployment',
                bugs: row['Bugs'] || 0,
                comments: row['Comments'] || ''
            };
            
            const stream = (row['Stream'] || '').toString().toLowerCase();
            if(stream.includes('core')) {
                this.coreDeliveryDataRows.push(item);
            } else {
                this.appDeliveryDataRows.push(item);
            }
        });
      }
      
      // 3a. Defect Analysis (Slide 8)
      const wsDefectAnalysis = wb.Sheets['Defect Analysis'];
      if(wsDefectAnalysis) {
          const data = XLSX.utils.sheet_to_json(wsDefectAnalysis) as any[];
          this.defectAnalysisData = data.map((row: any) => ({
             id: row['S.No'] || row['ID'] || '',
             desc: row['Description'],
             priority: row['Priority'],
             created: this.formatExcelDate(row['Created On'] || row['Created Date']),
             assigned: row['Assigned'] || row['Assigned To'],
             intro: row['When Introduced?'] || row['When Introduced'],
             eta: row['ETA'] || row['Fix Sprint'],
             status: row['Status'] 
          }));
      }

      // 3b. Backlog Items (Slide 9)
      const wsBacklog = wb.Sheets['Backlog Items']; 

      if(wsBacklog) {
          const data = XLSX.utils.sheet_to_json(wsBacklog) as any[];
          this.defectBacklogData = data.map((row: any) => ({
             id: row['S.No'] || row['ID'] || '',
             desc: row['Description'],
             priority: row['Priority'],
             created: this.formatExcelDate(row['Created On'] || row['Created Date']),
             assigned: row['Assigned'] || row['Assigned To'],
             intro: row['When Introduced?'] || row['When Introduced'],
             eta: row['ETA'] || row['Fix Sprint'],
             status: row['Status'],
             statusColor: (row['Status'] === 'IN-PROGRESS' ? '#fff8e1' : '#e3f2fd'),
             statusText: (row['Status'] === 'IN-PROGRESS' ? '#ff8f00' : '#1565c0')
         }));
      }

      // 4. Feedback
      const wsFeedback = wb.Sheets['Feedback'];
      if(wsFeedback) {
          const data = XLSX.utils.sheet_to_json(wsFeedback) as any[];
          this.feedbackData = data
              .filter((row: any) => {
                  // Filter out rows that have NO date AND NO owner AND NO status (likely garbage or wrap-text artifacts)
                  const hasDate = !!row['Action Date'];
                  const hasOwner = !!row['Owner'];
                  const hasStatus = !!row['Status'];
                  // Keep if at least 2 of 3 key fields are present, OR if it has a valid Item description + at least 1 other field
                  return (hasDate || hasOwner || hasStatus); 
              })
              .map((row: any) => {
                   // Fuzzy comment search
                   const commentKey = Object.keys(row).find(k => /comment|note|remark/i.test(k));
                   return {
                      date: row['Action Date'] || '',
                      item: row['Action Item'] || '',
                      owner: row['Owner'] || '',
                      status: row['Status'] || 'PENDING',
                      comments: commentKey ? row[commentKey] : ''
                   };
              });
      }

      // 5. Velocity Strategy
      const wsVelocity = wb.Sheets['Velocity Strategy'];
      if(wsVelocity) {
          const data = XLSX.utils.sheet_to_json(wsVelocity) as any[];
          this.deliveryStrategyData = data.map((row: any) => {
              // 1. Handle Month (Excel Date Serial vs String)
              let monthVal = row['Month'];
              if (typeof monthVal === 'number' && monthVal > 20000) {
                   const date = new Date((monthVal - 25569) * 86400 * 1000);
                   monthVal = date.toLocaleString('default', { month: 'short', year: 'numeric' });
              }

              // 2. Reliable Percentage Calculation
              const committed = row['Committed'] || 0;
              const delivered = row['Delivered'] || 0;
              let pct = '0%';
              if (committed > 0) {
                  pct = Math.round((delivered / committed) * 100) + '%';
              } else if (row['Delivery %']) {
                  // Fallback to Excel value if valid
                  const val = row['Delivery %'];
                  pct = typeof val === 'number' ? Math.round(val > 1 ? val : val * 100) + '%' : val;
              }

              // Find comment field fuzzily
              const commentKey = Object.keys(row).find(k => /comment|note|remark/i.test(k));
              const comment = commentKey ? row[commentKey] : '';

              return {
                  month: monthVal,
                  sprint: row['Sprint'],
                  committed: committed,
                  delivered: delivered,
                  pct: pct,
                  planned: '-', actual: '-', // Defaults
                  comment: comment
              };
          });
      }

      // 5b. App Migration (Block 3)
      const wsMigration = wb.Sheets['App Migration'];
      if(wsMigration) {
          const data = XLSX.utils.sheet_to_json(wsMigration) as any[];
          this.migrationData = data.map((row: any) => ({
              module: row['Module'],
              start: this.formatExcelDate(row['Start']),
              end: this.formatExcelDate(row['End']),
              pct: row['%'] || row['Completed %'] || '0%',
              status: row['Status'],
              comments: row['Comments']
          }));
      }

      // 6. Defect Metrics (Matrix)
      const wsDefectMatrix = wb.Sheets['Defect Metrics'];
      if(wsDefectMatrix) {
          const data = XLSX.utils.sheet_to_json(wsDefectMatrix) as any[];
          // Start fresh
           this.defectMetricsData = [];
           data.forEach(row => {
               const title = row['Category'];
               const backlog = row['Backlog'] || null;
               // Sprints 1-26
               const sprints: any = {};
               for(let i=1; i<=26; i++) {
                   if(row[i.toString()]) sprints[i] = row[i.toString()];
               }
               
               this.defectMetricsData.push({
                   title, 
                   backlog, 
                   sprints, 
                   type: title === 'Not Triage' ? 'critical' : 'info'
               });
           });
      }

      // 7. Leave Plan
      const wsLeave = wb.Sheets['Leave Plan'];
      if(wsLeave) {
          const data = XLSX.utils.sheet_to_json(wsLeave) as any[];
          this.leavePlanData = data.map((row: any) => ({
              date: row['Date'],
              event: row['Event'],
              member: row['Team Member']
          }));
      }

      // 8. Team Actions
      const wsTeamActions = wb.Sheets['Team Actions'];
      if(wsTeamActions) {
          const data = XLSX.utils.sheet_to_json(wsTeamActions) as any[];
          this.teamActionsData = data.map((row: any) => ({
              item: row['Action Item'],
              duration: row['Duration'],
              comments: row['Comments']
          }));
      }

      this.updatePreviews();
      this.cd.detectChanges(); // Force update view
      this.updatePreviews();
      this.cd.detectChanges(); // Force update view
      
      this.importStats = {
          coreHighlights: this.coreHighlights.length,
          appHighlights: this.appHighlights.length,
          sprints: this.coreDeliveryDataRows.length + this.appDeliveryDataRows.length,
          backlog: this.defectBacklogData.length,
          velocity: this.deliveryStrategyData.length
      };
      this.importModalVisible = true;
      
      });
    };
    reader.readAsBinaryString(target.files[0]);
    (evt.target as HTMLInputElement).value = ''; // Allow re-uploading same file
  }

  // Import Modal State
  importModalVisible = false;
  importStats = { coreHighlights: 0, appHighlights: 0, sprints: 0, backlog: 0, velocity: 0 };
  
  closeImportModal() {
      this.importModalVisible = false;
  }

  previewHtmlSafe: SafeHtml | null = null;
  slides: string[] = [];
  
  // Presentation State
  isPresenting = false;
  currentSlideIndex = 0;

  @HostListener('window:keydown', ['$event'])
  handleKeyboardEvent(event: KeyboardEvent) {
    if (!this.isPresenting) return;
    
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown' || event.key === 'Space') {
      this.nextSlide();
    } else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
      this.prevSlide();
    } else if (event.key === 'Escape') {
      this.endPresentation();
    }
  }

  constructor(public sanitizer: DomSanitizer, private cd: ChangeDetectorRef, private ngZone: NgZone) {}

  ngOnInit() {
    this.updatePreviews();
  }

  ngAfterViewInit() {
    this.initLucide();
  }

  initLucide() {
    if (typeof lucide !== 'undefined') {
      lucide.createIcons();
    }
  }

  switchTab(tab: 'input' | 'preview') {
    this.activeTab = tab;
    if (tab === 'preview') {
        this.updatePreviews();
    }
  }

  addCoreHighlight() {
    this.coreHighlights.push('New Highlight');
  }

  removeCoreHighlight(index: number) {
    this.coreHighlights.splice(index, 1);
  }

  addAppHighlight() {
      this.appHighlights.push('New Highlight');
  }

  removeAppHighlight(index: number) {
      this.appHighlights.splice(index, 1);
  }
  
  // Helper to format Excel serial dates
  formatExcelDate(val: any): string {
      if (!val) return '';
      if (typeof val === 'number' && val > 20000) {
          // Excel serial date to JS Date (approximate for modern dates)
          const date = new Date((val - 25569) * 86400 * 1000);
          return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
      }
      return val.toString();
  }

  // Helper to calculate totals for tables
  getTotals(rows: any[]) {
      const committed = rows.reduce((acc, row) => acc + row.committed, 0);
      const delivered = rows.reduce((acc, row) => acc + row.delivered, 0);
      const achieved = committed > 0 ? Math.round((delivered / committed) * 100) + '%' : '0%';
      return { committed, delivered, achieved };
  }

  generateSlides(): string[] {
      const coreTotals = this.getTotals(this.coreDeliveryDataRows);
      const appTotals = this.getTotals(this.appDeliveryDataRows);
      
      const slides = [];

      // Global Print/Preview Styles for dynamic tooltips
      // Global Print/Preview Styles for dynamic tooltips
      const tooltipStyles = `
        <style>
            .custom-tooltip-wrapper {
                position: relative;
                cursor: pointer;
            }
            .custom-tooltip-wrapper:hover .custom-tooltip-content {
                visibility: visible;
                opacity: 1;
                transform: translate(-50%, 10px);
            }
            .custom-tooltip-content {
                visibility: hidden;
                opacity: 0;
                width: 280px;
                background-color: #000048;
                color: #fff;
                text-align: left;
                border-radius: 6px;
                padding: 10px 14px;
                position: absolute;
                z-index: 999;
                top: 100%;
                left: 50%;
                transform: translate(-50%, 0);
                transition: opacity 0.3s, transform 0.3s;
                box-shadow: 0 5px 20px rgba(0,0,0,0.3);
                font-size: 11px;
                font-weight: 400;
                line-height: 1.5;
                white-space: normal; 
                pointer-events: none;
            }
            .custom-tooltip-content::after {
                content: "";
                position: absolute;
                bottom: 100%;
                left: 50%;
                margin-left: -6px;
                border-width: 6px;
                border-style: solid;
                border-color: transparent transparent #000048 transparent;
            }

            /* Flip Top Variant for bottom rows */
            .custom-tooltip-wrapper.flip-top .custom-tooltip-content {
                 top: auto;
                 bottom: 100%;
            }
            .custom-tooltip-wrapper.flip-top:hover .custom-tooltip-content {
                 transform: translate(-50%, -10px);
            }
            .custom-tooltip-wrapper.flip-top .custom-tooltip-content::after {
                 top: 100%;
                 bottom: auto;
                 border-color: #000048 transparent transparent transparent;
            }
        </style>
      `;

      // Slide 1: Title
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #000048; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            <!-- explicit background for export -->
            <div style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; background-color: #000048;"></div>
            
            <!-- Background Elements -->
            <div style="position: absolute; top: 0; right: 0; width: 55%; height: 100%; background: linear-gradient(135deg, #00155c 0%, #000048 100%); clip-path: polygon(25% 0, 100% 0, 100% 100%, 0% 100%);"></div>
            <!-- Accent Footer Line -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 12px; background: linear-gradient(to right, #2f78c4, #26C6DA);"></div>
            
            <!-- Header Logo Area -->
            <!-- Header Logo Area (Empty now) -->
            <div style="position: absolute; top: 50px; left: 60px;">
                <div style="width: 50px; height: 5px; background-color: #26C6DA; margin-top: 10px;"></div>
            </div>

            <!-- Logo Area Top Right -->
            <div style="position: absolute; top: 40px; right: 60px; display: flex; align-items: center; gap: 30px;">
                <img src="data:image/png;base64,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" style="height: 60px;" alt="Cognizant">
                <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAlgAAAJYCAYAAAC+ZpjcAAAAAXNSR0IArs4c6QAAIABJREFUeF7snQeUVUX29Q85Z9ABFQMMAoqAjGLAUcAAgooiAkYUA0pSsuQcJIOiiIggCoIMAiKS5W/+GEUUERzEhDDa3eTUTYdv7XIasW26X6i691bdXWuxmJF7T1X9TvXr/apOnZMnIyMjQ9hIgARIgARIgARIgAS0EchDgaWNJQ2RAAmQAAmQAAmQgCJAgcWFQAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQZKcyRAAiRAAiRAAiRAgcU1QAIkQAIkQAIkQAKaCVBgaQaanbljx47JkSNHZO/evZKUlCS//vqr/Pe//z35N/47/hw4cEAOHjyonsU7J06ckNTUVPUns+XNm1fwJ3/+/FKoUCH1p3jx4lKqVCkpXbq0lC1bVs444wz529/+JmeeeaZUrFhR/V2+fHn170WLFpV8+fJ5MGt2QQIkQAIkQALhJUCBpdH3ycnJsn//ftm9e7f8/PPPsnPnTvXnxx9/lD179ihxBREF8ZSSkiJpaWmSkZGhcQR/mMqTJ48UKFBAChcuLMWKFZMyZcoooXXeeefJBRdcIFWqVFF/n3PO2VKmTFkpUqSIkXHQKAmQAAmQAAmEkQAFVoxehzjat2+fEk/btm2TLVu2yNatW5Wgwg4VhBREVFAbdrFKlCihdrogti666CKpW7eu+rty5cpqR4yNBEiABEiABEggNgIUWBFyS09PV8d4//nPf+Tzzz+X//f//p8SVRBY2LWC4LK94dixXLly8ve//13q1asnV111lVx66aVyzjnnqKNINhIgARIgARIggcgIUGDlwAlHfhBQGzdulA8++ECJKuxQQVCFoUFwnXXWWfKPf/xDGjZsKA0aNJBq1arxODEMzuccSYAESIAE4iJAgZUFH3aqvv/+e/m///s/effdd+XTTz+VX3755U+B5nERt/RlxHQheL5+/frSvHlzufbaa9XRIgPmLXUoh00CJEACPhPABS5sWCCsJiEhQf3v48ePq1EhfhhhLJmXthBHjC/9NjUKrP95Czf3/v3vf8uiRYtkxYoVSmS5cOxnYjHiFiOC5Rs3biwtW7aUK664gjFbJkBrsvntt9+q9cymh8D555+vdnL9bocOHZLPPvtMsNPORgImCFx44YXqs15ng6hCqA1Ohd5//3356quv1CWww4cPq7WcefELX+oLFiyoRBZuw9eoUUP++c9/ytVXXy0YF/4t6C30AgvB6OvXr5fZs2fLhg0bVOA6W+QESpYsKZdffrncfffd0qRJE/WDwBYcAviS8Pjjj8urr74anEFZPpJ27drJs88+6/vu7ddffy0333yz/Pbbb5YT5fCDSmDUqFHy5JNPahkeLn19/PHH6rNo9erV6qZ91lv0EFX4Ao+G06Ts/h2/Y66//nq57777lNgK8g340Aos7FitXbtWXnzxRSWwjh49qmURhdUItm5xC/Hee++V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" style="height: 60px; " alt="CAT">
            </div>

            <!-- Main Title Content -->
            <div style="position: absolute; top: 180px; left: 60px; z-index: 10;">
                <div style="display: inline-block; padding: 6px 16px; background-color: rgba(255,255,255,0.08); border-left: 4px solid #26C6DA; color: #7BE4F1; text-transform: uppercase; font-size: 14px; font-weight: 700; letter-spacing: 1.5px; margin-bottom: 24px;">
                    Project Status Report
                </div>
                <h1 style="font-size: 64px; font-weight: 300; color: white; margin: 0; line-height: 1.1;">
                    CAT Technology<br>
                    <span style="font-weight: 800;">DQME – Core/App</span>
                </h1>
                
                <div style="margin-top: 50px; display: flex; align-items: center; gap: 20px;">
                     <!-- Date Tag -->
                     <div style="background-color: #2f78c4; color: white; padding: 12px 30px; font-weight: bold; font-size: 20px; box-shadow: 0 10px 20px rgba(0,0,0,0.2);">
                         ${this.currentMonth.toUpperCase()} <span style="font-weight: 300; opacity: 0.8; margin-left: 5px;">${this.currentYear}</span>
                     </div>
                     <div style="display: flex; flex-direction: column; justify-content: center; border-left: 1px solid rgba(255,255,255,0.3); padding-left: 20px; height: 50px;">
                         <span style="color: rgba(255,255,255,0.6); font-size: 12px; text-transform: uppercase; letter-spacing: 1px;">Generated On</span>
                         <span style="color: white; font-size: 16px; font-weight: 500;">${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' })}</span>
                     </div>
                </div>
            </div>

            <!-- Right Side Abstract Visuals -->
            <div style="position: absolute; bottom: 80px; right: -50px; opacity: 0.15; pointer-events: none;">
                <svg width="400" height="400" viewBox="0 0 400 400" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <circle cx="200" cy="200" r="150" stroke="white" stroke-width="40"/>
                    <circle cx="200" cy="200" r="100" stroke="white" stroke-width="20"/>
                </svg>
            </div>
            
            <!-- Footer Confidentiality -->
            <div style="position: absolute; bottom: 30px; left: 60px; color: rgba(255,255,255,0.4); font-size: 11px; text-transform: uppercase; letter-spacing: 1px;">
                &copy; ${new Date().getFullYear()} Cognizant. All rights reserved. &nbsp;|&nbsp; <span style="color: #26C6DA;">Caterpillar: Confidential Green</span>
            </div>
        </div>
      `);

      // Slide 2: Core Highlights
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Deliverable Highlights</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Core Platform</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <!-- Content Area -->
            <div style="padding: 40px 60px; position: relative;">
                
                <!-- Background Decoration -->
                <div style="position: absolute; top: 20px; right: 40px; opacity: 0.05;">
                     <svg width="200" height="200" viewBox="0 0 200 200" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <path d="M100 0L200 100L100 200L0 100L100 0Z" fill="#000048"/>
                     </svg>
                </div>

                <h3 style="color: #000048; font-size: 20px; font-weight: 700; margin-bottom: 25px; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px; display: inline-block;">
                    Key Achievements
                </h3>
                
                <ul style="list-style: none; padding: 0;">
                    ${this.coreHighlights.map(h => `
                    <li style="margin-bottom: 18px; display: flex; align-items: flex-start; font-size: 18px; color: #333; line-height: 1.5;">
                        <span style="min-width: 8px; height: 8px; background-color: #26C6DA; border-radius: 50%; margin-top: 9px; margin-right: 20px; box-shadow: 0 0 0 3px rgba(38, 198, 218, 0.2);"></span>
                        <span style="color: #000048; font-weight: 500;">${h}</span>
                    </li>
                    `).join('')}
                </ul>
            </div>
            
            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);

      // Slide 3: Core Delivery Table (Redesigned & Compacted)
        const coreChunks = [];
        for (let i = 0; i < this.coreDeliveryDataRows.length; i += 2) {
             coreChunks.push(this.coreDeliveryDataRows.slice(i, i + 2));
        }

        coreChunks.forEach(chunk => {
             slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 15px 40px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 10; position: relative;">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 24px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Delivery Metrics</h2>
                         <span style="color: #26C6DA; font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Core Platform</span>
                     </div>
                </div>
                <div style="display: flex; flex-direction: column; align-items: flex-end;">
                    <div style="color: rgba(255,255,255,0.8); font-size: 13px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
                </div>
            </div>

            <!-- Content Area: Sprint Cards (Compacted) -->
            <div style="padding: 15px 40px; display: flex; flex-direction: column; gap: 12px; height: 435px; justify-content: flex-start;">
                
                ${chunk.map((row) => `
                <!-- Sprint Card -->
                <div style="background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); overflow: hidden; border-left: 5px solid #000048; display: flex; flex-direction: column;">
                    
                    <!-- Card Header -->
                    <div style="padding: 8px 21px; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; justify-content: space-between; background: linear-gradient(to right, #f8f9fa, #fff);">
                        <!-- Sprint Title -->
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="width: 32px; height: 32px; background-color: #e8eaf6; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #000048;">
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
                            </div>
                            <div>
                                <h4 style="margin: 0; color: #000048; font-size: 15px; font-weight: 700;">${row.sprintMonth.split('\n')[0]}</h4>
                                <span style="font-size: 11px; color: #666;">${row.sprintMonth.split('\n')[1] || ''}</span>
                            </div>
                        </div>

                        <!-- Metrics Grid -->
                        <div style="display: flex; gap: 30px;">
                            <div style="text-align: center; min-width: 70px;">
                                <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600; margin-bottom: 0px;">Committed</div>
                                <div style="font-size: 16px; font-weight: 700; color: #333;">${row.committed}</div>
                            </div>
                            <div style="text-align: center; min-width: 70px;">
                                <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600; margin-bottom: 0px;">Delivered</div>
                                <div style="font-size: 16px; font-weight: 700; color: #333;">${row.delivered}</div>
                            </div>
                            <div style="text-align: center; min-width: 70px;">
                                <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600; margin-bottom: 0px;">Achieved</div>
                                <div style="font-size: 16px; font-weight: 700; color: ${parseInt(row.achieved) >= 100 ? '#2e7d32' : '#ef6c00'};">${row.achieved}</div>
                            </div>
                        </div>
                    </div>

                    <!-- Card Body: Features -->
                    <div style="padding: 8px 21px; background-color: #fff; flex-grow: 1;">
                       <!-- Header Row: Title + Status Metrics -->
                       <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; border-bottom: 1px solid #f0f0f0; padding-bottom: 4px;">
                           <div style="font-size: 10px; font-weight: 700; color: #000048; text-transform: uppercase;">Features Delivered</div>
                           
                           <div style="display: flex; align-items: center; gap: 15px;">
                                 <!-- Deployment -->
                                 <div style="display: flex; align-items: center; gap: 5px;">
                                    <span style="color: #666; font-size: 9px; font-weight: 600;">Deployment:</span>
                                    <span style="padding: 1px 6px; border-radius: 4px; background-color: ${row.deploymentStatus === 'Success' ? '#e8f5e9' : '#f5f5f5'}; color: ${row.deploymentStatus === 'Success' ? '#2e7d32' : '#666'}; font-weight: 700; font-size: 9px;">${row.deploymentStatus}</span>
                                 </div>
                                 <!-- Bugs -->
                                 <div style="display: flex; align-items: center; gap: 5px;">
                                    <span style="color: #666; font-size: 9px; font-weight: 600;">Bugs:</span>
                                    <span style="color: ${row.bugs > 0 ? '#d32f2f' : '#666'}; font-weight: 700; font-size: 9px;">${row.bugs}</span>
                                 </div>
                                 <!-- Comments -->
                                 ${row.comments ? `
                                 <div style="display: flex; align-items: center; gap: 5px; max-width: 150px;">
                                     <span style="color: #666; font-size: 9px; font-weight: 600;">Note:</span>
                                     <span style="color: #666; font-size: 9px; font-style: italic; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.comments}</span>
                                 </div>
                                 ` : ''}
                           </div>
                       </div>
                       
                       <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 4px 15px;">
                            ${row.features.map((f: string) => `
                            <div style="display: flex; align-items: flex-start; gap: 6px;">
                                <span style="min-width: 4px; height: 4px; margin-top: 5px; background-color: #26C6DA; border-radius: 50%;"></span>
                                <span style="font-size: 10px; color: #444; line-height: 1.3; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 380px;">${f.length > 60 ? f.substring(0, 60) + '...' : f}</span>
                            </div>
                            `).join('')}
                       </div>
                    </div>
                </div>
                `).join('')}

                <!-- Grand Total Block (Ultra Compact) -->
                <div style="background: linear-gradient(90deg, #00155c 0%, #0033A0 100%); border-radius: 6px; padding: 10px 21px; display: flex; align-items: center; justify-content: space-between; color: white; margin-top: auto;">
                    <div style="font-size: 12px; font-weight: 700; text-transform: uppercase;">Grand Total</div>
                    <div style="display: flex; gap: 30px;">
                        <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
                            <span style="font-size: 9px; opacity: 0.7; text-transform: uppercase;">Committed</span>
                            <span style="font-size: 14px; font-weight: 700;">${coreTotals.committed}</span>
                        </div>
                        <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
                             <span style="font-size: 9px; opacity: 0.7; text-transform: uppercase;">Delivered</span>
                             <span style="font-size: 14px; font-weight: 700;">${coreTotals.delivered}</span>
                        </div>
                        <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
                             <span style="font-size: 9px; text-transform: uppercase; color: #26C6DA; font-weight: 700;">Achieved</span>
                             <span style="font-size: 14px; font-weight: 700; color: #26C6DA;">${coreTotals.achieved}</span>
                        </div>
                    </div>
                </div>

            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);
      });

      // Slide 4: App Highlights
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Deliverable Highlights</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">App Platform</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <!-- Content Area -->
            <div style="padding: 40px 60px; position: relative;">
                
                <!-- Background Decoration -->
                <div style="position: absolute; top: 20px; right: 40px; opacity: 0.05;">
                     <svg width="200" height="200" viewBox="0 0 200 200" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <rect x="0" y="0" width="100" height="100" stroke="#000048" stroke-width="20"/>
                          <rect x="50" y="50" width="100" height="100" stroke="#000048" stroke-width="10"/>
                     </svg>
                </div>

                <h3 style="color: #000048; font-size: 20px; font-weight: 700; margin-bottom: 25px; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px; display: inline-block;">
                    Key Achievements
                </h3>
                
                <ul style="list-style: none; padding: 0;">
                    ${this.appHighlights.map(h => `
                    <li style="margin-bottom: 18px; display: flex; align-items: flex-start; font-size: 18px; color: #333; line-height: 1.5;">
                        <span style="min-width: 8px; height: 8px; background-color: #26C6DA; border-radius: 50%; margin-top: 9px; margin-right: 20px; box-shadow: 0 0 0 3px rgba(38, 198, 218, 0.2);"></span>
                        <span style="color: #000048; font-weight: 500;">${h}</span>
                    </li>
                    `).join('')}
                </ul>
            </div>
            
            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);

      // Slide 5: App Delivery Table (Redesigned & Compacted)
        const appChunks = [];
        for (let i = 0; i < this.appDeliveryDataRows.length; i += 2) {
             appChunks.push(this.appDeliveryDataRows.slice(i, i + 2));
        }

        appChunks.forEach(chunk => {
        slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 15px 40px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 10; position: relative;">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 24px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Delivery Metrics</h2>
                         <span style="color: #26C6DA; font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">App Platform</span>
                     </div>
                </div>
                <div style="display: flex; flex-direction: column; align-items: flex-end;">
                    <div style="color: rgba(255,255,255,0.8); font-size: 13px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
                </div>
            </div>

            <!-- Content Area: Sprint Cards (Compacted) -->
            <div style="padding: 15px 40px; display: flex; flex-direction: column; gap: 12px; height: 435px; justify-content: flex-start;">
                
                ${chunk.map((row) => `
                <!-- Sprint Card -->
                <div style="background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); overflow: hidden; border-left: 5px solid #000048; display: flex; flex-direction: column;">
                    
                    <!-- Card Header -->
                    <div style="padding: 8px 21px; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; justify-content: space-between; background: linear-gradient(to right, #f8f9fa, #fff);">
                        <!-- Sprint Title -->
                        <div style="display: flex; align-items: center; gap: 12px;">
                            <div style="width: 32px; height: 32px; background-color: #e8eaf6; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #000048;">
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
                            </div>
                            <div>
                                <h4 style="margin: 0; color: #000048; font-size: 15px; font-weight: 700;">${row.sprintMonth.split('\n')[0]}</h4>
                                <span style="font-size: 11px; color: #666;">${row.sprintMonth.split('\n')[1] || ''}</span>
                            </div>
                        </div>

                        <!-- Metrics Grid -->
                        <div style="display: flex; gap: 30px;">
                            <div style="text-align: center; min-width: 70px;">
                                <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600; margin-bottom: 0px;">Committed</div>
                                <div style="font-size: 16px; font-weight: 700; color: #333;">${row.committed}</div>
                            </div>
                            <div style="text-align: center; min-width: 70px;">
                                <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600; margin-bottom: 0px;">Delivered</div>
                                <div style="font-size: 16px; font-weight: 700; color: #333;">${row.delivered}</div>
                            </div>
                            <div style="text-align: center; min-width: 70px;">
                                <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600; margin-bottom: 0px;">Achieved</div>
                                <div style="font-size: 16px; font-weight: 700; color: ${parseInt(row.achieved) >= 100 ? '#2e7d32' : '#ef6c00'};">${row.achieved}</div>
                            </div>
                        </div>
                    </div>

                    <!-- Card Body: Features -->
                    <div style="padding: 8px 21px; background-color: #fff; flex-grow: 1;">
                       <!-- Header Row: Title + Status Metrics -->
                       <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; border-bottom: 1px solid #f0f0f0; padding-bottom: 4px;">
                           <div style="font-size: 10px; font-weight: 700; color: #000048; text-transform: uppercase;">Features Delivered</div>
                           
                           <div style="display: flex; align-items: center; gap: 15px;">
                                 <!-- Deployment -->
                                 <div style="display: flex; align-items: center; gap: 5px;">
                                    <span style="color: #666; font-size: 9px; font-weight: 600;">Deployment:</span>
                                    <span style="padding: 1px 6px; border-radius: 4px; background-color: ${row.deploymentStatus === 'Success' ? '#e8f5e9' : '#f5f5f5'}; color: ${row.deploymentStatus === 'Success' ? '#2e7d32' : '#666'}; font-weight: 700; font-size: 9px;">${row.deploymentStatus}</span>
                                 </div>
                                 <!-- Bugs -->
                                 <div style="display: flex; align-items: center; gap: 5px;">
                                    <span style="color: #666; font-size: 9px; font-weight: 600;">Bugs:</span>
                                    <span style="color: ${row.bugs > 0 ? '#d32f2f' : '#666'}; font-weight: 700; font-size: 9px;">${row.bugs}</span>
                                 </div>
                                 <!-- Comments -->
                                 ${row.comments ? `
                                 <div style="display: flex; align-items: center; gap: 5px; max-width: 150px;">
                                     <span style="color: #666; font-size: 9px; font-weight: 600;">Note:</span>
                                     <span style="color: #666; font-size: 9px; font-style: italic; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.comments}</span>
                                 </div>
                                 ` : ''}
                           </div>
                       </div>
                       
                       <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 4px 15px;">
                            ${row.features.map((f: string) => `
                            <div style="display: flex; align-items: flex-start; gap: 6px;">
                                <span style="min-width: 4px; height: 4px; margin-top: 5px; background-color: #26C6DA; border-radius: 50%;"></span>
                                <span style="font-size: 10px; color: #444; line-height: 1.3; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 380px;">${f.length > 60 ? f.substring(0, 60) + '...' : f}</span>
                            </div>
                            `).join('')}
                       </div>
                    </div>
                </div>
                `).join('')}

                <!-- Grand Total Block (Ultra Compact) -->
                <div style="background: linear-gradient(90deg, #00155c 0%, #0033A0 100%); border-radius: 6px; padding: 10px 21px; display: flex; align-items: center; justify-content: space-between; color: white; margin-top: auto;">
                    <div style="font-size: 12px; font-weight: 700; text-transform: uppercase;">Grand Total</div>
                    <div style="display: flex; gap: 30px;">
                        <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
                            <span style="font-size: 9px; opacity: 0.7; text-transform: uppercase;">Committed</span>
                            <span style="font-size: 14px; font-weight: 700;">${appTotals.committed}</span>
                        </div>
                        <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
                             <span style="font-size: 9px; opacity: 0.7; text-transform: uppercase;">Delivered</span>
                             <span style="font-size: 14px; font-weight: 700;">${appTotals.delivered}</span>
                        </div>
                        <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
                             <span style="font-size: 9px; text-transform: uppercase; color: #26C6DA; font-weight: 700;">Achieved</span>
                             <span style="font-size: 14px; font-weight: 700; color: #26C6DA;">${appTotals.achieved}</span>
                        </div>
                    </div>
                </div>

            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);
      });

      // Slide 5b: Migration Status (Block 3)
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
             <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">App Migration Status</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Block 3 Overview</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>
            
            <!-- Content -->
            <div style="padding: 20px 50px;">
                <div style="background: white; border-radius: 8px; box-shadow: 0 2px 15px rgba(0,0,0,0.05); overflow: hidden; border: 1px solid #eee;">
                    <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 11px;">
                        <thead>
                            <tr style="background-color: #000048; color: white;">
                                <th style="padding: 8px 15px; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1);">Module name</th>
                                <th style="padding: 8px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1);">Start date</th>
                                <th style="padding: 8px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1);">End date</th>
                                <th style="padding: 8px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1);">Completed%</th>
                                <th style="padding: 8px; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1);">Status</th>
                                <th style="padding: 8px 15px; font-weight: 600; text-align: left;">Comments</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${this.migrationData.map((row, i) => `
                            <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">
                                <td style="padding: 6px 15px; color: #000048; font-weight: 600; border-right: 1px solid #eee;">${row.module}</td>
                                <td style="padding: 6px; color: #444; text-align: center; border-right: 1px solid #eee;">${row.start}</td>
                                <td style="padding: 6px; color: #444; text-align: center; border-right: 1px solid #eee;">${row.end}</td>
                                <td style="padding: 6px; color: #333; font-weight: 700; text-align: center; border-right: 1px solid #eee;">${row.pct}</td>
                                <td style="padding: 6px; border-right: 1px solid #eee;">
                                    <span style="padding: 2px 8px; border-radius: 12px; font-size: 10px; font-weight: 700; background-color: #e8f5e9; color: #2e7d32;">
                                        ${row.status}
                                    </span>
                                </td>
                                <td style="padding: 6px 15px; color: #666; font-style: italic;">${row.comments}</td>
                            </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);



      // Slide 6: Feedback/Action Items
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Feedback / Action Items</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Core & App Platform</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <!-- Content Area -->
            <div style="padding: 30px 50px;">
                <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 12px; box-shadow: 0 2px 15px rgba(0,0,0,0.05); border-radius: 8px; overflow: hidden; background: white;">
                    <thead>
                        <tr style="background-color: #000048; color: white;">
                            <th style="padding: 15px; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1);">Action Date</th>
                            <th style="padding: 15px; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1); width: 35%;">Action Item</th>
                            <th style="padding: 15px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1);">Owner</th>
                            <th style="padding: 15px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1);">Status</th>
                            <th style="padding: 15px; font-weight: 600; text-align: left;">Comments</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${this.feedbackData.map((row, i) => `
                        <tr style="border-bottom: 1px solid #eee;">
                            <td style="padding: 15px; color: #000048; font-weight: 600; vertical-align: middle; border-right: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">${row.date}</td>
                            <td style="padding: 15px; color: #444; vertical-align: middle; border-right: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'}; line-height: 1.4;">${row.item}</td>
                            <td style="padding: 15px; color: #000048; text-align: center; font-weight: 500; vertical-align: middle; border-right: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">${row.owner}</td>
                            <td style="padding: 15px; vertical-align: middle; border-right: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'}; text-align: center;">
                                <span style="
                                    display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: 700; 
                                    background-color: ${(row.status || '').toLowerCase() === 'completed' ? '#e0f2f1' : '#e3f2fd'}; 
                                    color: ${(row.status || '').toLowerCase() === 'completed' ? '#00695c' : '#1565c0'};">
                                    ${(row.status || 'PENDING').toUpperCase()}
                                </span>
                            </td>
                            <td style="padding: 15px; color: #666; font-style: italic; vertical-align: middle; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">${row.comments}</td>
                        </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);

      // Slide 7: Last 6 Month Delivery Metrics
      // deliveryStrategyData is now a class property

      // Helper to calculate bar height (max approx 50)
      const maxVal = 50;
      const getBarHeight = (val: number) => Math.max(4, (val / maxVal) * 100);

      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 20px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Delivery Metrics Strategy</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Last 6 Months - App Platform</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <!-- Main Content: Split View -->
            <div style="padding: 20px 40px; display: grid; grid-template-rows: 170px 1fr; gap: 20px; height: 430px;">
                
                <!-- Top Section: Trend Chart -->
                <div style="background: white; border-radius: 8px; padding: 15px 25px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); display: flex; flex-direction: column;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                        <span style="font-size: 12px; font-weight: 700; color: #000048; text-transform: uppercase;">Delivery Velocity Trend</span>
                        <div style="display: flex; gap: 15px; font-size: 10px;">
                            <div style="display: flex; align-items: center; gap: 4px;"><span style="width: 8px; height: 8px; background-color: #e0e0e0; border-radius: 2px;"></span> Committed</div>
                            <div style="display: flex; align-items: center; gap: 4px;"><span style="width: 8px; height: 8px; background-color: #000048; border-radius: 2px;"></span> Delivered</div>
                        </div>
                    </div>
                    
                    <!-- Chart Area -->
                    <div style="flex-grow: 1; display: flex; align-items: flex-end; justify-content: space-between; padding-bottom: 5px; padding-top: 10px;">
                        ${[...this.deliveryStrategyData].reverse().map(d => `
                            <!-- Bar Group -->
                            <div style="display: flex; flex-direction: column; align-items: center; gap: 6px; width: 10%;">
                                <div style="display: flex; gap: 4px; align-items: flex-end; height: 80px;">
                                    <!-- Committed Bar -->
                                    <div style="width: 12px; height: ${getBarHeight(d.committed)}%; background-color: #e0e0e0; border-radius: 2px 2px 0 0; position: relative; group: hover;">
                                       ${d.committed > 0 ? `<span style="position: absolute; top: -14px; left: 50%; transform: translateX(-50%); font-size: 8px; color: #888;">${d.committed}</span>` : ''}
                                    </div>
                                    <!-- Delivered Bar -->
                                    <div style="width: 12px; height: ${getBarHeight(d.delivered)}%; background-color: ${parseFloat(d.pct) >= 100 ? '#2e7d32' : parseFloat(d.pct) < 60 ? '#ef6c00' : '#000048'}; border-radius: 2px 2px 0 0; position: relative;">
                                        <span style="position: absolute; top: -14px; left: 50%; transform: translateX(-50%); font-size: 9px; font-weight: 700; color: #333;">${d.delivered}</span>
                                    </div>
                                </div>
                                <div style="font-size: 9px; color: #666; font-weight: 600;">${d.sprint.replace('Sprint ', 'S')}</div>
                            </div>
                        `).join('')}
                    </div>
                </div>

                <!-- Bottom Section: Roadmap Timeline -->
                <div style="background: white; border-radius: 8px; padding: 0; box-shadow: 0 2px 10px rgba(0,0,0,0.05); overflow: hidden; display: flex; flex-direction: column;">
                     <div style="padding: 10px 20px; border-bottom: 1px solid #f0f0f0; background-color: #fafafa; font-size: 11px; font-weight: 700; color: #666; text-transform: uppercase;">Sprint Roadmap</div>
                     <div style="overflow-y: auto; padding: 10px 20px;">
                        <div style="display: flex; flex-direction: column; gap: 0;">
                            ${(() => {
                                let html = '';
                                let lastMonth = '';
                                const reversedData = [...this.deliveryStrategyData].reverse();
                                reversedData.forEach((row, idx) => {
                                    const isNewMonth = row.month && row.month !== lastMonth;
                                    if (isNewMonth) lastMonth = row.month;
                                    
                                    html += `
                                    <div style="display: flex; border-bottom: 1px solid #f9f9f9; padding: 8px 0; align-items: center;">
                                        <!-- Month Column -->
                                        <div style="width: 90px; flex-shrink: 0;">
                                            ${isNewMonth ? `<div style="font-weight: 700; color: #000048; font-size: 11px;">${row.month}</div>` : ''}
                                        </div>
                                        
                                        <!-- Timeline Connector -->
                                        <div style="width: 30px; display: flex; flex-direction: column; align-items: center; position: relative;">
                                            <div style="width: 10px; height: 10px; border-radius: 50%; background-color: ${parseFloat(row.pct) >= 100 ? '#26C6DA' : '#ff9800'}; border: 2px solid white; box-shadow: 0 0 0 1px #eee; z-index: 2;"></div>
                                            ${idx < reversedData.length - 1 ? `<div style="width: 1px; background-color: #eee; height: 100%; position: absolute; top: 10px; z-index: 1;"></div>` : ''}
                                        </div>

                                        <!-- Sprint Card Content -->
                                        <div style="flex-grow: 1; display: flex; align-items: center; justify-content: space-between; padding-left: 10px;">
                                            <div style="display: flex; flex-direction: column; gap: 2px;">
                                                <div style="display: flex; align-items: center; gap: 8px;">
                                                    <span style="font-weight: 700; color: #333; font-size: 11px;">${row.sprint}</span>
                                                    ${row.planned !== '-' ? `<span style="font-size: 9px; padding: 1px 4px; background-color: #f0f4c3; color: #558b2f; border-radius: 4px;">Deployed: ${row.actual}</span>` : ''}
                                                </div>
                                                <div style="font-size: 10px; color: #666; font-style: italic;">${row.comment || ''}</div>
                                            </div>
                                            
                                            <!-- Stats Badge -->
                                            <div style="display: flex; align-items: center; gap: 10px;">
                                                 <div style="text-align: right;">
                                                    <span style="font-size: 9px; color: #999; display: block; line-height: 1;">Delivery</span>
                                                    <span style="font-size: 12px; font-weight: 700; color: ${parseFloat(row.pct) >= 100 ? '#2e7d32' : '#e65100'};">${row.pct}</span>
                                                 </div>
                                            </div>
                                        </div>
                                    </div>
                                    `;
                                });
                                return html;
                            })()}
                        </div>
                     </div>
                </div>
            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);

      // Slide 8: Defect Analysis Report
      // Chunking Logic: Max 8 items per slide
      const defectChunks = [];
      const defectChunkSize = 8;
      for (let i = 0; i < this.defectAnalysisData.length; i += defectChunkSize) {
           defectChunks.push(this.defectAnalysisData.slice(i, i + defectChunkSize));
      }

      const daFontSize = '11px';
      const daPadding = '10px 12px';
      const daHeaderPadding = '12px';

      defectChunks.forEach((chunk, chunkIndex) => {
          slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 15px 40px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 24px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Defect Analysis Report</h2>
                         <span style="color: #26C6DA; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Last 6 Months - DQME App</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 12px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <!-- Content Area -->
            <div style="padding: 15px 40px; height: 435px; display: flex; flex-direction: column;">
                 <div style="overflow: visible; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); background: white;">
                    <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: ${daFontSize}; table-layout: fixed;">
                        <thead>
                            <tr style="background-color: #000048; color: white;">
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 4%;">S.No</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1); width: 35%;">Description</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 5%;">Priority</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 10%;">Created On</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 12%;">Assigned</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 12%;">When Introduced?</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 12%;">When Fix?</th>
                                <th style="padding: ${daHeaderPadding}; font-weight: 600; text-align: center; width: 10%;">Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${chunk.map((row: any, i: number) => {
                                const statusUpper = (row.status || 'PENDING').toUpperCase();
                                let badgeBg = '#e3f2fd'; // Blue default
                                let badgeColor = '#1565c0';
                                
                                if (statusUpper === 'RESOLVED' || statusUpper === 'COMPLETED') { 
                                    badgeBg = '#e0f2f1'; badgeColor = '#00695c'; 
                                } else if (statusUpper === 'CLOSED') {
                                    badgeBg = '#e8f5e9'; badgeColor = '#2e7d32';
                                } else if (statusUpper === 'IN-PROGRESS' || statusUpper === 'ONGOING') {
                                    badgeBg = '#fff8e1'; badgeColor = '#ff8f00';
                                }

                                // Flip tooltip for last 3 items in the current page/chunk
                                const isBottomRow = i >= chunk.length - 3;

                                return `
                                <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">
                                    <td style="padding: ${daPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.id}</td>
                                    <td style="padding: ${daPadding}; color: #000048; font-weight: 500; vertical-align: middle; border-right: 1px solid #eee; line-height: 1.3;">
                                        <div class="custom-tooltip-wrapper ${isBottomRow ? 'flip-top' : ''}">
                                            <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.desc}</div>
                                            <div class="custom-tooltip-content">${row.desc}</div>
                                        </div>
                                    </td>
                                    <td style="padding: ${daPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.priority}</td>
                                    <td style="padding: ${daPadding}; color: #666; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.created}</td>
                                    <td style="padding: ${daPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">
                                        <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.assigned}</div>
                                    </td>
                                    <td style="padding: ${daPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.intro}</td>
                                    <td style="padding: ${daPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">
                                        <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.eta}</div>
                                    </td>
                                    <td style="padding: ${daPadding}; vertical-align: middle; text-align: center;">
                                        <span style="display: inline-block; padding: 2px 6px; border-radius: 10px; font-size: 10px; font-weight: 700; background-color: ${badgeBg}; color: ${badgeColor}; white-space: nowrap;">
                                            ${statusUpper}
                                        </span>
                                    </td>
                                </tr>`;
                            }).join('')}
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);
      });
      // Slide 9: Defect Report Backlog
      // Chunking Logic: max 8 items per slide
      const backlogChunks = [];
      const backlogChunkSize = 8;
      for (let i = 0; i < this.defectBacklogData.length; i += backlogChunkSize) {
           backlogChunks.push(this.defectBacklogData.slice(i, i + backlogChunkSize));
      }
      
      const blFontSize = '10px';
      const blPadding = '7px 10px';
      const blHeaderPadding = '10px';

      backlogChunks.forEach((chunk, chunkIndex) => {
          slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Defect Report</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Backlog Items - DQME App</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <!-- Content Area -->
            <div style="padding: 20px 40px; height: 410px; display: flex; flex-direction: column;">
                <div style="overflow: visible; border-radius: 8px; box-shadow: 0 2px 15px rgba(0,0,0,0.05); background: white;">
                    <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: ${blFontSize}; table-layout: fixed;">
                        <thead>
                            <tr style="background-color: #000048; color: white;">
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 4%;">S.No</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1); width: 28%;">Description</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 6%;">Priority</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 12%;">Created On</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 14%;">Assigned</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 14%;">When Introduced?</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 8%;">ETA</th>
                                <th style="padding: ${blHeaderPadding}; font-weight: 600; text-align: center; width: 14%;">Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${chunk.map((row, i) => `
                            ${(() => {
                                const isBottomRow = i >= chunk.length - 3;
                                return '';
                            })()}
                            <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">
                                <td style="padding: ${blPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.id}</td>
                                <td style="padding: ${blPadding}; color: #000048; font-weight: 500; vertical-align: middle; border-right: 1px solid #eee; line-height: 1.3;">
                                    <div class="custom-tooltip-wrapper ${i >= chunk.length - 3 ? 'flip-top' : ''}">
                                        <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.desc}</div>
                                        <div class="custom-tooltip-content">${row.desc}</div>
                                    </div>
                                </td>
                                <td style="padding: ${blPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.priority}</td>
                                <td style="padding: ${blPadding}; color: #666; font-weight: 500; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.created}</td>
                                <td style="padding: ${blPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">
                                     <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.assigned}</div>
                                </td>
                                <td style="padding: ${blPadding}; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">
                                     <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.intro}</div>
                                </td>
                                <td style="padding: ${blPadding}; color: #666; font-style: italic; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.eta}</td>
                                <td style="padding: ${blPadding}; vertical-align: middle; text-align: center;">
                                    <span style="display: inline-block; padding: 2px 6px; border-radius: 10px; font-size: 10px; font-weight: 700; background-color: ${row.statusColor}; color: ${row.statusText}; white-space: nowrap;">${row.status}</span>
                                </td>
                            </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);
      });

      // Slide 10: DQME App - Defect Analysis metrics for last 6 months
      // Slide 10: Defect Analysis Metrics
      // Define the fixed sprints 13 to 26
      const defectGridSprints = Array.from({length: 14}, (_, i) => i + 13);
      
      // Define the fixed categories we MUST show
      const fixedCategories = [
          'Not Triage',
          'Old/Historic Bug',
          'Data Error',
          'Sprint 18(2024)',
          'Sprint 19(2024)',
          'Sprint 20(2024)',
          'Sprint 21(2024)',
          'Sprint 22(2024)',
          'Sprint 23(2024)',
          'Sprint 24(2024)',
          'Sprint 25(2024)',
          'Sprint 26(2024)',
          'Sprint 27(2024)',
          'Sprint 1(2025)'
      ];

      // Helper to find data for a category
      const getData = (cat: string) => this.defectMetricsData.find(d => d.title === cat) || { backlog: null, sprints: {} };

      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
             <!-- Top Header Bar (Standard Style) -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 20px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 24px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Defect Analysis</h2>
                         <span style="color: #26C6DA; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Last 6 Months Metrics - DQME App</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>
            
            <!-- Content Area -->
            <div style="padding: 15px 40px; display: flex; justify-content: center;">
                 
                 <div style="background: white; padding: 8px; border-radius: 8px; box-shadow: 0 2px 15px rgba(0,0,0,0.05); width: 100%;">
                     <table style="border-collapse: collapse; width: 100%; font-size: 10px; font-family: 'Segoe UI', sans-serif; color: #333;">
                        <thead>
                            <!-- Row 1: Top Headers -->
                            <tr>
                                <th rowspan="3" style="border: 1px solid #4472C4; color: #0033A0; font-weight: 800; padding: 3px; width: 140px; background-color: #f8f9fa; vertical-align: middle; font-size: 11px;">
                                    Bug<br>Introduced
                                </th>
                                <th rowspan="3" style="border: 1px solid #4472C4; color: #4472C4; font-weight: 700; padding: 3px; width: 50px; background-color: #f8f9fa; vertical-align: bottom;">
                                    Backlog
                                </th>
                                <th colspan="14" style="border: 1px solid #4472C4; color: #0033A0; font-weight: 800; padding: 3px; background-color: #f8f9fa; font-size: 11px;">
                                    Bug Fixed and Deployed in PROD
                                </th>
                            </tr>
                            <!-- Row 2: Sprints Year Header -->
                            <tr>
                                <th colspan="14" style="border: 1px solid #4472C4; color: #0033A0; font-weight: 700; padding: 3px; background-color: #f8f9fa; font-size: 10px;">
                                    ${this.currentYear} Sprints
                                </th>
                            </tr>
                            <!-- Row 3: Sprint Numbers -->
                            <tr style="height: 20px;">
                                ${defectGridSprints.map(s => `
                                    <th style="border: 1px solid #4472C4; background-color: #002060; color: white; padding: 2px; font-weight: 700; font-size: 9px;">${s}</th>
                                `).join('')}
                            </tr>
                        </thead>
                        <tbody>
                            <!-- Data Rows (Iterate Fixed Categories) -->
                             ${fixedCategories.map((cat, i) => {
                                const row = getData(cat);
                                const bg = i % 2 === 0 ? '#fff' : '#fcfcfc';
                                return `
                                <tr style="height: 19px;">
                                    <!-- Category Title -->
                                    <td style="border: 1px solid #4472C4; color: #4472C4; font-weight: 700; padding: 2px 8px; text-align: left; background-color: ${bg};">
                                        ${cat}
                                    </td>
                                    
                                    <!-- Backlog Count -->
                                    <td style="border: 1px solid #4472C4; color: #0033A0; font-weight: 700; padding: 2px; text-align: center; background-color: ${bg};">
                                        ${row.backlog !== null && row.backlog !== undefined ? row.backlog : ''}
                                    </td>

                                    <!-- Sprints Data -->
                                    ${defectGridSprints.map(s => `
                                        <td style="border: 1px solid #4472C4; color: #0033A0; font-weight: 700; padding: 2px; text-align: center; background-color: ${bg}; font-size: 9px;">
                                            ${row.sprints[s] ? row.sprints[s] : ''}
                                        </td>
                                    `).join('')}
                                </tr>
                                `;
                             }).join('')}
                             
                             <tr style="height: 19px;">
                                 <td style="border: 1px solid #4472C4; background-color: #fff;">&nbsp;</td>
                                 <td style="border: 1px solid #4472C4; background-color: #fff;">&nbsp;</td>
                                 ${defectGridSprints.map(() => `<td style="border: 1px solid #4472C4; background-color: #fff;">&nbsp;</td>`).join('')}
                             </tr>

                        </tbody>
                     </table>
                 </div>

            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);



      // Slide 12: People Update
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Top Header Bar -->
            <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
                <div style="display: flex; align-items: center; gap: 15px;">
                     <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
                     <div>
                         <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">People Update</h2>
                         <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Holiday Plan & Action Items</span>
                     </div>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
            </div>

            <div style="padding: 25px 50px; display: flex; flex-direction: column; gap: 30px;">
                
                <!-- Section 1: Holiday / Leave Plan -->
                <div>
                    <h3 style="color: #000048; font-size: 18px; font-weight: 700; margin-bottom: 15px; display: flex; align-items: center; gap: 10px;">
                        <span style="display: inline-block; width: 4px; height: 18px; background-color: #26C6DA;"></span>
                        App team – Holiday / Leave Plan
                    </h3>
                    <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); border-radius: 8px; overflow: hidden; background: white;">
                        <thead>
                            <tr style="background-color: #000048; color: white;">
                                <th style="padding: 12px; font-weight: 600; text-align: left; width: 25%;">Date</th>
                                <th style="padding: 12px; font-weight: 600; text-align: left; width: 45%;">Event</th>
                                <th style="padding: 12px; font-weight: 600; text-align: left; width: 30%;">Team Member</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${this.leavePlanData.map(l => `
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 12px; color: #000048; vertical-align: middle;">${l.date}</td>
                                <td style="padding: 12px; color: #444; vertical-align: middle;">${l.event}</td>
                                <td style="padding: 12px; color: #666; vertical-align: middle;">${l.member}</td>
                            </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>

                <!-- Section 2: Action Items -->
                <div>
                     <h3 style="color: #000048; font-size: 18px; font-weight: 700; margin-bottom: 15px; display: flex; align-items: center; gap: 10px;">
                        <span style="display: inline-block; width: 4px; height: 18px; background-color: #26C6DA;"></span>
                        App team – Action items
                    </h3>
                     <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); border-radius: 8px; overflow: hidden; background: white;">
                        <thead>
                            <tr style="background-color: #000048; color: white;">
                                <th style="padding: 12px; font-weight: 600; text-align: left; width: 40%;">Action items</th>
                                <th style="padding: 12px; font-weight: 600; text-align: left; width: 20%;">Duration</th>
                                <th style="padding: 12px; font-weight: 600; text-align: left; width: 40%;">Comments</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${this.teamActionsData.map(a => `
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 15px 12px; color: #444; vertical-align: middle;">${a.item || ''}</td>
                                <td style="padding: 15px 12px; color: #666; vertical-align: middle;">${a.duration || ''}</td>
                                <td style="padding: 15px 12px; color: #666; vertical-align: middle; font-style: italic;">${a.comments || ''}</td>
                            </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>

            </div>

            <!-- Footer -->
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
                <div style="color: #666; font-size: 11px; font-weight: 600;">
                    <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="color: #ccc; font-size: 12px;">Page</div>
                    <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
                </div>
            </div>
            <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
        </div>
      `);

      // Slide 12: Thank You
      slides.push(`
        <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: white; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
            
            <!-- Logo Top Right -->
            <div style="position: absolute; top: 40px; right: 50px;">
                <h2 style="color: #000048; font-size: 28px; font-weight: 800; margin: 0; letter-spacing: -0.5px;">Cognizant</h2>
            </div>
            
            <!-- Center Content -->
            <div style="position: absolute; top: 45%; left: 80px; transform: translateY(-50%);">
                <h1 style="color: #000048; font-size: 56px; font-weight: 600; margin: 0; line-height: 1.2;">Thank you</h1>
                <div style="width: 120px; height: 4px; background-color: #26C6DA; margin-top: 15px;"></div>
            </div>
            
            <!-- Abstract Graphic Bottom Right -->
             <div style="
                position: absolute; 
                bottom: -150px; 
                right: -100px; 
                width: 700px; 
                height: 400px; 
                background: linear-gradient(135deg, #000048 0%, #0033A0 40%, #2979ff 70%, #26C6DA 100%); 
                transform: rotate(-15deg);
                opacity: 0.9;
                border-top-left-radius: 200px;
                z-index: 1;">
            </div>
            <!-- Decorative accent lines in the swoosh -->
             <div style="
                position: absolute; 
                bottom: -120px; 
                right: -80px; 
                width: 700px; 
                height: 400px; 
                border-top: 2px solid rgba(255,255,255,0.2);
                transform: rotate(-15deg);
                border-radius: 200px;
                pointer-events: none;
                z-index: 2;">
            </div>
             <div style="
                position: absolute; 
                bottom: -100px; 
                right: -60px; 
                width: 700px; 
                height: 400px; 
                border-top: 1px solid rgba(255,255,255,0.1);
                transform: rotate(-15deg);
                border-radius: 200px;
                pointer-events: none;
                z-index: 2;">
            </div>

             <!-- Footer Minimal -->
            <div style="position: absolute; bottom: 20px; left: 50px; z-index: 3;">
                <p style="color: #666; font-size: 11px; margin: 0;">Caterpillar: Confidential Green</p>
            </div>
        </div>
      `);
      
      let pageCounter = 1;
      return slides.map(s => (tooltipStyles + s).replace('<!--PAGE-->', () => (pageCounter++).toString()));
  }
  
  updatePreviews() {
    this.slides = this.generateSlides();
    this.previewHtmlSafe = this.sanitizer.bypassSecurityTrustHtml(
        ` <div style="font-family: Arial, sans-serif; bg-white">
            ${this.slides.join('')}
          </div>`
    );
  }

  // Presentation State
  presentationScale = 1;

  @HostListener('window:resize')
  onResize() {
     if (this.isPresenting) {
         this.calculatePresentationScale();
     }
  }

  calculatePresentationScale() {
      // Slide dimensions: 1000 x 562.5
      const slideWidth = 1000;
      const slideHeight = 562.5;
      
      const windowWidth = window.innerWidth;
      const windowHeight = window.innerHeight;
      
      const widthScale = windowWidth / slideWidth;
      const heightScale = windowHeight / slideHeight;
      
      // Use the smaller scale to ensure it fits, multiply by 0.95 for some padding
      this.presentationScale = Math.min(widthScale, heightScale) * 0.95;
  }

  // Presentation Logic
  startPresentation() {
      this.slides = this.generateSlides(); // Ensure latest
      this.currentSlideIndex = 0;
      this.isPresenting = true;
      
      // Calculate initial scale
      this.calculatePresentationScale();

      // Request Fullscreen
      const elem = document.documentElement as any;
      if (elem.requestFullscreen) {
        elem.requestFullscreen();
      } else if (elem.webkitRequestFullscreen) { /* Safari */
        elem.webkitRequestFullscreen();
      } else if (elem.msRequestFullscreen) { /* IE11 */
        elem.msRequestFullscreen();
      }
  }

  endPresentation() {
      this.isPresenting = false;
      
      // Exit Fullscreen
      if (document.fullscreenElement) {
         const doc = document as any;
         if (doc.exitFullscreen) {
            doc.exitFullscreen();
         } else if (doc.webkitExitFullscreen) { /* Safari */
            doc.webkitExitFullscreen();
         } else if (doc.msExitFullscreen) { /* IE11 */
            doc.msExitFullscreen();
         }
      }
  }

  nextSlide() {
      if (this.currentSlideIndex < this.slides.length - 1) {
          this.currentSlideIndex++;
      }
  }

  prevSlide() {
      if (this.currentSlideIndex > 0) {
          this.currentSlideIndex--;
      }
  }



  downloadExcelTemplate() {
    const wb = XLSX.utils.book_new();

    // Helper: Apply Styles
    const applyTableStyles = (ws: any) => {
         if (!ws['!ref']) return;
         const range = XLSX.utils.decode_range(ws['!ref']);
         const cols = [];
         
         for(let R = range.s.r; R <= range.e.r; ++R) {
            for(let C = range.s.c; C <= range.e.c; ++C) {
                const cell_address = XLSX.utils.encode_cell({r:R, c:C});
                if(!ws[cell_address]) continue;
                
                // Base Style
                const style: any = {
                    font: { name: 'Segoe UI', sz: 11 },
                    border: {
                        top: { style: 'thin', color: { rgb: "CCCCCC" } },
                        bottom: { style: 'thin', color: { rgb: "CCCCCC" } },
                        left: { style: 'thin', color: { rgb: "CCCCCC" } },
                        right: { style: 'thin', color: { rgb: "CCCCCC" } }
                    }
                };

                // Header Style (Row 0)
                if(R === 0) {
                    style.font = { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: "FFFFFF" } };
                    style.fill = { fgColor: { rgb: "000048" } };
                    style.alignment = { horizontal: "center", vertical: "center" };
                }

                // Preserve existing Props if any (like formula type) but overwrite style
                ws[cell_address].s = style;
            }
         }

         // Set Default Width
         for(let C = range.s.c; C <= range.e.c; ++C) {
             cols.push({ wch: 25 });
         }
         ws['!cols'] = cols;
    };

    // 1. General Info Sheet
    const generalData = [
      ['Section', 'Value'],
      ['Month', this.currentMonth],
      ['Year', this.currentYear],
      ['Project Name', 'DQME App'],
      ['Client', 'Caterpillar'],
      [],
      ['Core Highlights (Add more rows as needed)', ''],
      ...this.coreHighlights.map(h => ['Highlight', h]),
      [],
      ['App Highlights (Add more rows as needed)', ''],
      ...this.appHighlights.map(h => ['Highlight', h])
    ];
    const wsGeneral = XLSX.utils.aoa_to_sheet(generalData);
    applyTableStyles(wsGeneral);
    XLSX.utils.book_append_sheet(wb, wsGeneral, 'General Info');

    // 2. Delivery Metrics Sheet
    const deliveryHeaders = [['Stream', 'Sprint', 'Month', 'Committed', 'Delivered', 'Delivery %', 'Features Delivered', 'Deployed Date', 'Deployment Status', 'Bugs', 'Comments']];
    const deliverySample = [
        ['Core Platform', 'Sprint 05', 'Mar 2026', 21, 23, null, 'Feature A, Feature B', '12-Mar-2026', 'Success', 0, 'Completed'],
        ['App Platform', 'Sprint 11', 'June 2026', 38, 36, null, 'Feature X, Feature Y', '', '', 0, 'Spillover due to scope']
    ];
    const wsDelivery = XLSX.utils.aoa_to_sheet([...deliveryHeaders, ...deliverySample]);
    
    // Formulas
    wsDelivery['F2'] = { t: 'n', f: 'E2/D2', F: '0%' };
    wsDelivery['F3'] = { t: 'n', f: 'E3/D3', F: '0%' };
    
    applyTableStyles(wsDelivery);
    XLSX.utils.book_append_sheet(wb, wsDelivery, 'Delivery Metrics');

    // 3. Defect Analysis Sheet
    const defectHeaders = [['ID', 'Description', 'Priority', 'Created Date', 'Assigned To', 'When Introduced', 'Fix Sprint', 'Status']];
    const defectSample = [
        ['BUG2122267', 'DQME portal allowing Tableau URL from other sites', 3, '07 Nov 2025', 'Sprint 23', 'Integration Error', 'Sprint 23', 'RESOLVED']
    ];
    const wsDefects = XLSX.utils.aoa_to_sheet([...defectHeaders, ...defectSample]);
    applyTableStyles(wsDefects);
    XLSX.utils.book_append_sheet(wb, wsDefects, 'Defect Analysis');

    // 4. Feedback / Action Items Sheet
    const feedbackHeaders = [['Action Date', 'Action Item', 'Owner', 'Status', 'Comments']];
    const feedbackSample = [
        ['11-19-2024', 'End to end demo of sample automation test script', 'Manikandan', 'ONGOING', 'Part of the migration activity']
    ];
    const wsFeedback = XLSX.utils.aoa_to_sheet([...feedbackHeaders, ...feedbackSample]);
    applyTableStyles(wsFeedback);
    XLSX.utils.book_append_sheet(wb, wsFeedback, 'Feedback');

    // 5. Velocity & Roadmap Sheet
    const velocityHeaders = [['Month', 'Sprint', 'Committed', 'Delivered', 'Delivery %', 'Comments']];
    const velocitySample = [
        ['June 2025', 'Sprint 11', 38, 36, null, 'One bug spill over'],
        ['June 2025', 'Sprint 12', 40, 40, null, 'Completed']
    ];
    const wsVelocity = XLSX.utils.aoa_to_sheet([...velocityHeaders, ...velocitySample]);
    // Formulas
    wsVelocity['E2'] = { t: 'n', f: 'D2/C2', F: '0%' };
    wsVelocity['E3'] = { t: 'n', f: 'D3/C3', F: '0%' };
    
    applyTableStyles(wsVelocity);
    XLSX.utils.book_append_sheet(wb, wsVelocity, 'Velocity Strategy');

    // 6. Defect Backlog Sheet
    const backlogHeaders = [['S.No', 'Description', 'Priority', 'Created On', 'Assigned', 'When Introduced?', 'ETA', 'Status']];
    const backlogSample = [
        [1, 'BUG2382601 - Filter Report', 4, '19 Dec 2024', 'Carthipadmin', 'Base Error', 'TBD', 'IN-PROGRESS']
    ];
    const wsBacklog = XLSX.utils.aoa_to_sheet([...backlogHeaders, ...backlogSample]);
    applyTableStyles(wsBacklog);
    XLSX.utils.book_append_sheet(wb, wsBacklog, 'Backlog Items');

    // 7. Defect Metrics (Matrix) Sheet
    // Sprints 13-26 to match slide 10 structure
    const sprintNums = Array.from({length: 14}, (_, i) => (i + 13).toString());
    const defectMatrixHeaders = [['Category', 'Backlog', ...sprintNums]];
    
    const createRow = (category: string): (string | number)[] => [category, '', ...sprintNums.map(() => '')];
    
    const categories = [
        'Not Triage',
        'Old/Historic Bug', 
        'Data Error',
        'Sprint 18(2024)',
        'Sprint 19(2024)',
        'Sprint 20(2024)',
        'Sprint 21(2024)',
        'Sprint 22(2024)',
        'Sprint 23(2024)',
        'Sprint 24(2024)',
        'Sprint 25(2024)',
        'Sprint 26(2024)',
        'Sprint 27(2024)',
        'Sprint 1(2025)'
    ];

    const defectMatrixSample = categories.map(cat => {
        const row = createRow(cat);
        // Add sample data for visualization consistency with default state
        if (cat === 'Not Triage') row[1] = 10; // Backlog
        if (cat === 'Old/Historic Bug') {
             // Sprint 15 is index 2+2=4 (Category, Backlog, 13, 14, 15) -> array index 4
             // array index = 2 + (sprint - 13)
             row[2 + (15-13)] = 1; 
             row[2 + (22-13)] = 2; 
             row[2 + (23-13)] = 2;
        }
        if (cat === 'Data Error') {
             row[2 + (15-13)] = 8;
             row[2 + (16-13)] = 1;
             row[2 + (19-13)] = 1;
        }
        return row;
    });

    const wsDefectMatrix = XLSX.utils.aoa_to_sheet([...defectMatrixHeaders, ...defectMatrixSample]);
    applyTableStyles(wsDefectMatrix);
    XLSX.utils.book_append_sheet(wb, wsDefectMatrix, 'Defect Metrics');

    // 8. Leave Plan Sheet
    const leaveHeaders = [['Date', 'Event', 'Team Member']];
    const leaveSample = [['15th Jan', 'Pongal', 'App Team'], ['26th Jan', 'Republic Day', 'App Team']];
    const wsLeave = XLSX.utils.aoa_to_sheet([...leaveHeaders, ...leaveSample]);
    applyTableStyles(wsLeave);
    XLSX.utils.book_append_sheet(wb, wsLeave, 'Leave Plan');

    // 9. Team Actions Sheet
    const teamActionHeaders = [['Action Item', 'Duration', 'Comments']];
    const teamActionSample = [['NA', '', '']];
    const wsTeamActions = XLSX.utils.aoa_to_sheet([...teamActionHeaders, ...teamActionSample]);
    applyTableStyles(wsTeamActions);
    XLSX.utils.book_append_sheet(wb, wsTeamActions, 'Team Actions');

    // 10. App Migration Sheet
    const migrationHeaders = [['Module', 'Start', 'End', '%', 'Status', 'Comments']];
    const migrationSample = [
        ['Home Dashboard', '19-Mar-25', '01-Apr-25', '100%', 'Completed', 'Demo done in DEV env'],
        ['User Specific Screen', '19-Mar-25', '13-May-25', '100%', 'Completed', 'Demo done in DEV env']
    ];
    const wsMigration = XLSX.utils.aoa_to_sheet([...migrationHeaders, ...migrationSample]);
    applyTableStyles(wsMigration);
    XLSX.utils.book_append_sheet(wb, wsMigration, 'App Migration');

    // Write File
    XLSX.writeFile(wb, 'Monthly_Report_Input_Template.xlsx');
  }

  // Export State
  isExporting = false;
  exportProgress = '';

  // Export Logic
  async exportReport(format: 'pdf' | 'pptx' | 'image' | 'pptx-content') {
    this.isExporting = true;
    this.exportProgress = 'Initializing...';
    
    // Create a loading overlay
    const overlay = document.createElement('div');
    overlay.style.position = 'fixed';
    overlay.style.inset = '0';
    overlay.style.backgroundColor = 'rgba(0,0,0,0.9)';
    overlay.style.zIndex = '10000';
    overlay.style.display = 'flex';
    overlay.style.flexDirection = 'column';
    overlay.style.alignItems = 'center';
    overlay.style.justifyContent = 'center';
    overlay.style.color = 'white';
    overlay.innerHTML = '<div style="font-size: 24px; margin-bottom: 20px;">Processing Report...</div><div id="export-status">Preparing content...</div>';
    document.body.appendChild(overlay);

    const statusEl = document.getElementById('export-status');
    const updateStatus = (msg: string) => {
        if(statusEl) statusEl.innerText = msg;
        this.exportProgress = msg;
    };

    try {
        // For PPTX Content export, we don't need DOM rendering
        if (format === 'pptx-content') {
            // Export PPTX with actual editable content
            updateStatus('Creating PowerPoint with editable content...');
            const pptx = new PptxGenJS();
            pptx.layout = 'LAYOUT_16x9';
            pptx.defineLayout({ name: 'CUSTOM', width: 10, height: 5.625 });
            pptx.layout = 'CUSTOM';

            // Helper function to add footer to slides
            const addFooter = (slide: any, pageNum: number) => {
                slide.addText('Cognizant | Caterpillar: Confidential Green', { 
                    x: 0.5, y: 5.2, w: 6, h: 0.3, fontSize: 9, color: '666666' 
                });
                slide.addText(`${pageNum}`, { 
                    x: 9, y: 5.2, w: 0.8, h: 0.3, fontSize: 10, bold: true, 
                    color: 'FFFFFF', fill: { color: '000048' }, align: 'center', valign: 'middle'
                });
            };

            let pageNum = 1;

            // ===== Slide 1: Title Slide =====
            updateStatus('Creating title slide...');
            const slide1 = pptx.addSlide();
            slide1.background = { color: '000048' };
            slide1.addText('Cognizant', { x: 8.2, y: 0.4, w: 1.5, h: 0.4, fontSize: 14, bold: true, color: 'FFFFFF' });
            slide1.addText('CAT', { x: 8.2, y: 0.8, w: 1.5, h: 0.4, fontSize: 14, bold: true, color: 'FFFFFF' });
            slide1.addText('CAT Technology', { x: 0.6, y: 1.8, w: 8, h: 0.6, fontSize: 48, bold: false, color: 'FFFFFF' });
            slide1.addText('DQME – Core/App', { x: 0.6, y: 2.5, w: 8, h: 0.6, fontSize: 48, bold: true, color: 'FFFFFF' });
            slide1.addText(`${this.currentMonth.toUpperCase()} ${this.currentYear}`, { 
                x: 0.6, y: 3.5, w: 2, h: 0.5, fontSize: 20, bold: true, color: 'FFFFFF', 
                fill: { color: '2f78c4' }, align: 'center', valign: 'middle'
            });
            slide1.addText(`© ${new Date().getFullYear()} Cognizant. All rights reserved. | Caterpillar: Confidential Green`, {
                x: 0.6, y: 5.0, w: 8, h: 0.3, fontSize: 9, color: 'FFFFFF'
            });

            // ===== Slide 2: Core Highlights =====
            updateStatus('Creating Core Highlights...');
            pageNum = 2;
            const slide2 = pptx.addSlide();
            slide2.background = { color: 'F8F9FA' };
            slide2.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: '000048' } });
            slide2.addText('Deliverable Highlights', { x: 0.5, y: 0.15, w: 8, h: 0.4, fontSize: 22, bold: true, color: 'FFFFFF' });
            slide2.addText('Core Platform', { x: 0.5, y: 0.45, w: 8, h: 0.2, fontSize: 12, bold: true, color: '26C6DA' });
            slide2.addText(`${this.currentMonth} ${this.currentYear}`, { x: 8.5, y: 0.25, w: 1.3, h: 0.3, fontSize: 12, color: 'FFFFFF' });
            
            let yPos = 1.0;
            slide2.addText('Key Achievements', { x: 0.6, y: yPos, w: 8, h: 0.3, fontSize: 16, bold: true, color: '000048' });
            yPos += 0.4;
            this.coreHighlights.forEach((highlight) => {
                if (yPos < 4.5) {
                    slide2.addText('• ' + highlight, { x: 0.8, y: yPos, w: 8.5, h: 0.25, fontSize: 14, color: '333333' });
                    yPos += 0.3;
                }
            });
            addFooter(slide2, pageNum);

            // ===== Slides 3+: Core Delivery Metrics =====
            updateStatus('Creating Core Delivery Metrics...');
            const coreChunks = [];
            for (let i = 0; i < this.coreDeliveryDataRows.length; i += 2) {
                coreChunks.push(this.coreDeliveryDataRows.slice(i, i + 2));
            }
            
            coreChunks.forEach((chunk) => {
                pageNum++;
                const slide = pptx.addSlide();
                slide.background = { color: 'F8F9FA' };
                slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: '000048' } });
                slide.addText('Delivery Metrics', { x: 0.5, y: 0.1, w: 8, h: 0.3, fontSize: 20, bold: true, color: 'FFFFFF' });
                slide.addText('Core Platform', { x: 0.5, y: 0.35, w: 8, h: 0.2, fontSize: 11, bold: true, color: '26C6DA' });
                slide.addText(`${this.currentMonth} ${this.currentYear}`, { x: 8.5, y: 0.2, w: 1.3, h: 0.25, fontSize: 11, color: 'FFFFFF' });
                
                chunk.forEach((row, idx) => {
                    const cardY = idx === 0 ? 0.8 : 2.4;
                    slide.addShape(pptx.ShapeType.rect, { x: 0.4, y: cardY, w: 9.2, h: 1.4, fill: { color: 'FFFFFF' }, 
                        line: { color: '000048', width: 2 }
                    });
                    slide.addText(row.sprintMonth.split('\n')[0], { x: 0.6, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 13, bold: true, color: '000048' });
                    slide.addText(`C: ${row.committed}  D: ${row.delivered}  A: ${row.achieved}`, { 
                        x: 3, y: cardY + 0.1, w: 3, h: 0.25, fontSize: 10, bold: true, 
                        color: parseInt(row.achieved) >= 100 ? '2e7d32' : 'ef6c00' 
                    });
                    slide.addText('Features:', { x: 0.6, y: cardY + 0.45, w: 8.8, h: 0.2, fontSize: 9, bold: true, color: '000048' });
                    const features = row.features.slice(0, 4).map((f: string) => '• ' + (f.length > 80 ? f.substring(0, 80) + '...' : f)).join('\n');
                    slide.addText(features, { x: 0.6, y: cardY + 0.65, w: 8.8, h: 0.6, fontSize: 8, color: '444444' });
                });
                
                // Grand Total
                const totals = this.getTotals(this.coreDeliveryDataRows);
                slide.addShape(pptx.ShapeType.rect, { x: 0.4, y: 4.2, w: 9.2, h: 0.4, fill: { color: '00155c' } });
                slide.addText(`Grand Total: C: ${totals.committed}  D: ${totals.delivered}  A: ${totals.achieved}`, {
                    x: 0.6, y: 4.3, w: 8.8, h: 0.2, fontSize: 11, bold: true, color: 'FFFFFF'
                });
                addFooter(slide, pageNum);
            });

            // ===== App Highlights =====
            updateStatus('Creating App Highlights...');
            pageNum++;
            const slide3 = pptx.addSlide();
            slide3.background = { color: 'F8F9FA' };
            slide3.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: '000048' } });
            slide3.addText('Deliverable Highlights', { x: 0.5, y: 0.15, w: 8, h: 0.4, fontSize: 22, bold: true, color: 'FFFFFF' });
            slide3.addText('App Platform', { x: 0.5, y: 0.45, w: 8, h: 0.2, fontSize: 12, bold: true, color: '26C6DA' });
            
            yPos = 1.0;
            slide3.addText('Key Achievements', { x: 0.6, y: yPos, w: 8, h: 0.3, fontSize: 16, bold: true, color: '000048' });
            yPos += 0.4;
            this.appHighlights.forEach((highlight) => {
                if (yPos < 4.5) {
                    slide3.addText('• ' + highlight, { x: 0.8, y: yPos, w: 8.5, h: 0.25, fontSize: 14, color: '333333' });
                    yPos += 0.3;
                }
            });
            addFooter(slide3, pageNum);

            // ===== App Delivery Metrics =====
            updateStatus('Creating App Delivery Metrics...');
            const appChunks = [];
            for (let i = 0; i < this.appDeliveryDataRows.length; i += 2) {
                appChunks.push(this.appDeliveryDataRows.slice(i, i + 2));
            }
            
            appChunks.forEach((chunk) => {
                pageNum++;
                const slide = pptx.addSlide();
                slide.background = { color: 'F8F9FA' };
                slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: '000048' } });
                slide.addText('Delivery Metrics', { x: 0.5, y: 0.1, w: 8, h: 0.3, fontSize: 20, bold: true, color: 'FFFFFF' });
                slide.addText('App Platform', { x: 0.5, y: 0.35, w: 8, h: 0.2, fontSize: 11, bold: true, color: '26C6DA' });
                
                chunk.forEach((row, idx) => {
                    const cardY = idx === 0 ? 0.8 : 2.4;
                    slide.addShape(pptx.ShapeType.rect, { x: 0.4, y: cardY, w: 9.2, h: 1.4, fill: { color: 'FFFFFF' }, 
                        line: { color: '000048', width: 2 }
                    });
                    slide.addText(row.sprintMonth.split('\n')[0], { x: 0.6, y: cardY + 0.1, w: 2, h: 0.3, fontSize: 13, bold: true, color: '000048' });
                    slide.addText(`C: ${row.committed}  D: ${row.delivered}  A: ${row.achieved}`, { 
                        x: 3, y: cardY + 0.1, w: 3, h: 0.25, fontSize: 10, bold: true, 
                        color: parseInt(row.achieved) >= 100 ? '2e7d32' : 'ef6c00' 
                    });
                    const features = row.features.slice(0, 4).map((f: string) => '• ' + (f.length > 80 ? f.substring(0, 80) + '...' : f)).join('\n');
                    slide.addText(features, { x: 0.6, y: cardY + 0.65, w: 8.8, h: 0.6, fontSize: 8, color: '444444' });
                });
                addFooter(slide, pageNum);
            });

            // ===== Migration Status =====
            if (this.migrationData.length > 0) {
                updateStatus('Creating Migration Status...');
                pageNum++;
                const migSlide = pptx.addSlide();
                migSlide.background = { color: 'F8F9FA' };
                migSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: '000048' } });
                migSlide.addText('Migration Status', { x: 0.5, y: 0.2, w: 8, h: 0.4, fontSize: 22, bold: true, color: 'FFFFFF' });
                
                const migRows = this.migrationData.slice(0, 8).map(row => [
                    row.module, row.start, row.end, row.pct, row.status
                ]);
                migSlide.addTable(migRows, {
                    x: 0.5, y: 1.0, w: 9, h: 3.8,
                    colW: [3, 1.2, 1.2, 0.8, 1.5, 1.3],
                    border: { pt: 1, color: '000048' },
                    fill: { color: 'FFFFFF' },
                    fontSize: 9,
                    color: '333333'
                });
                addFooter(migSlide, pageNum);
            }

            // ===== Feedback =====
            if (this.feedbackData.length > 0) {
                updateStatus('Creating Feedback slide...');
                pageNum++;
                const fbSlide = pptx.addSlide();
                fbSlide.background = { color: 'F8F9FA' };
                fbSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: '000048' } });
                fbSlide.addText('Feedback & Actions', { x: 0.5, y: 0.2, w: 8, h: 0.4, fontSize: 22, bold: true, color: 'FFFFFF' });
                
                yPos = 1.0;
                this.feedbackData.slice(0, 10).forEach((fb) => {
                    fbSlide.addText(`• ${fb.item || 'N/A'}`, { x: 0.6, y: yPos, w: 8.5, h: 0.25, fontSize: 11, color: '000048', bold: true });
                    fbSlide.addText(`Owner: ${fb.owner || 'N/A'}  |  Status: ${fb.status || 'PENDING'}`, { 
                        x: 0.8, y: yPos + 0.25, w: 8.5, h: 0.15, fontSize: 9, color: '666666' 
                    });
                    yPos += 0.45;
                });
                addFooter(fbSlide, pageNum);
            }

            // ===== Defect Backlog =====
            if (this.defectBacklogData.length > 0) {
                updateStatus('Creating Defect Backlog...');
                pageNum++;
                const defectSlide = pptx.addSlide();
                defectSlide.background = { color: 'F8F9FA' };
                defectSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: '000048' } });
                defectSlide.addText('Defect Analysis', { x: 0.5, y: 0.2, w: 8, h: 0.4, fontSize: 22, bold: true, color: 'FFFFFF' });
                
                // Create table with defect information
                const defectRows: any[] = [['ID', 'Description', 'Priority', 'Status']];
                this.defectBacklogData.slice(0, 8).forEach(defect => {
                    defectRows.push([
                        defect.id || 'N/A',
                        (defect.desc || '').substring(0, 80) + ((defect.desc || '').length > 80 ? '...' : ''),
                        String(defect.priority || 'N/A'),
                        defect.status || 'NEW'
                    ]);
                });
                
                defectSlide.addTable(defectRows, {
                    x: 0.5, y: 1.0, w: 9, h: 4.0,
                    colW: [1.2, 5, 0.8, 2],
                    fontSize: 9,
                    color: '333333',
                    border: { pt: 1, color: '000048' },
                    fill: { color: 'FFFFFF' },
                    align: 'left',
                    valign: 'middle'
                });
                addFooter(defectSlide, pageNum);
            }

            // ===== People Update =====
            if (this.leavePlanData.length > 0 || this.teamActionsData.length > 0) {
                updateStatus('Creating People Update...'); 
                pageNum++;
                const peopleSlide = pptx.addSlide();
                peopleSlide.background = { color: 'F8F9FA' };
                peopleSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.7, fill: { color: '000048' } });
                peopleSlide.addText('People Update', { x: 0.5, y: 0.2, w: 8, h: 0.4, fontSize: 22, bold: true, color: 'FFFFFF' });
                
                // Leave Plan
                peopleSlide.addText('Holiday / Leave Plan', { x: 0.5, y: 1.0, w: 9, h: 0.25, fontSize: 14, bold: true, color: '000048' });
                if (this.leavePlanData.length > 0) {
                    const leaveRows = this.leavePlanData.slice(0, 5).map(l => [l.date, l.event, l.member]);
                    peopleSlide.addTable(leaveRows, {
                        x: 0.5, y: 1.3, w: 9, h: 1.2,
                        colW: [1.5, 4.5, 3],
                        fontSize: 10,
                        border: { pt: 1, color: 'CCCCCC' }
                    });
                }
                
                // Action Items
                peopleSlide.addText('Action Items', { x: 0.5, y: 2.7, w: 9, h: 0.25, fontSize: 14, bold: true, color: '000048' });
                if (this.teamActionsData.length > 0) {
                    const actionRows = this.teamActionsData.slice(0, 5).map(a => [a.item || 'N/A', a.duration || '', a.comments || '']);
                    peopleSlide.addTable(actionRows, {
                        x: 0.5, y: 3.0, w: 9, h: 1.8,
                        colW: [4, 1.5, 3.5],
                        fontSize: 9,
                        border: { pt: 1, color: 'CCCCCC' }
                    });
                }
                addFooter(peopleSlide, pageNum);
            }

            // ===== Thank You Slide =====
            updateStatus('Creating Thank You slide...');
            const thankYouSlide = pptx.addSlide();
            thankYouSlide.background = { color: 'FFFFFF' };
            thankYouSlide.addText('Thank you', { x: 0.8, y: 2.3, w: 8, h: 0.8, fontSize: 48, bold: true, color: '000048' });
            thankYouSlide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 3.2, w: 1.2, h: 0.05, fill: { color: '26C6DA' } });
            thankYouSlide.addText('Cognizant', { x: 8.5, y: 0.4, w: 1.2, h: 0.4, fontSize: 24, bold: true, color: '000048' });

            updateStatus('Saving PPTX with content...');
            await pptx.writeFile({ fileName: `Monthly_Report_Content_${this.currentMonth}_${this.currentYear}.pptx` });
            
            // Clean up and exit early
            if(document.body.contains(overlay)) document.body.removeChild(overlay);
            this.isExporting = false;
            this.exportProgress = '';
            return;
        }

        // For other formats, create DOM container
        const container = document.createElement('div');
        container.style.position = 'absolute';
        container.style.top = '0';
        container.style.left = '0';
        container.style.width = '1000px';
        container.style.zIndex = '9999'; 
        container.style.backgroundColor = '#f0f0f0';
        document.body.appendChild(container);

        const slides = this.generateSlides();
        
        // Append all slides to the DOM
        slides.forEach((slideHtml, index) => {
            const wrapper = document.createElement('div');
            wrapper.style.width = '1000px';
            wrapper.style.height = '562.5px';
            wrapper.style.marginBottom = '20px'; // Space between slides
            wrapper.style.position = 'relative';
            wrapper.style.background = 'white';
            wrapper.style.overflow = 'hidden';
            wrapper.innerHTML = slideHtml;
            container.appendChild(wrapper);
        });

        // Wait for DOM to settle and images/fonts to load
        updateStatus('Rendering content...');
        await new Promise(resolve => setTimeout(resolve, 2000));

        const slideElements = Array.from(container.children) as HTMLElement[];
        
        // Helper to prepare element for html-to-image
        const prepareSlide = (el: HTMLElement) => {
            // Ensure no transforms that might confuse capture
            el.style.transform = 'none';
            el.style.margin = '0';
        };

        slideElements.forEach(prepareSlide);
        
        if (format === 'pdf') {
            const pdf = new jspdf({
                orientation: 'landscape',
                unit: 'px',
                format: [1000, 562.5]
            });

            for (let i = 0; i < slideElements.length; i++) {
                updateStatus(`Capturing Slide ${i + 1} of ${slideElements.length}...`);
                const slideElement = slideElements[i];
                await new Promise(resolve => setTimeout(resolve, 100)); // Breathing room for UI

                try {
                    const imgData = await htmlToImage.toPng(slideElement, {
                        pixelRatio: 2,
                        width: 1000,
                        height: 562.5,
                        backgroundColor: '#ffffff',
                        skipAutoScale: true
                    });
                     // html-to-image returns the data URL directly

                    if (i > 0) pdf.addPage([1000, 562.5]);
                    pdf.addImage(imgData, 'PNG', 0, 0, 1000, 562.5);
                } catch (e) {
                    console.error(`Error capturing slide ${i+1}:`, e);
                }
            }
            updateStatus('Saving PDF...');
            pdf.save(`Monthly_Report_${this.currentMonth}_${this.currentYear}.pdf`);

        } else if (format === 'pptx') {
            const pptx = new PptxGenJS();
            pptx.layout = 'LAYOUT_16x9';

            for (let i = 0; i < slideElements.length; i++) {
                updateStatus(`Capturing Slide ${i + 1} of ${slideElements.length}...`);
                const slideElement = slideElements[i];
                await new Promise(resolve => setTimeout(resolve, 100));

                try {
                     const imgData = await htmlToImage.toPng(slideElement, {
                         pixelRatio: 2,
                         width: 1000,
                         height: 562.5,
                         backgroundColor: '#ffffff',
                         skipAutoScale: true
                     });
                     // html-to-image returns the data URL directly

                    const slide = pptx.addSlide();
                    slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });
                } catch (e) {
                     console.error(`Error capturing slide ${i+1}:`, e);
                }
            }
            updateStatus('Saving PPTX...');
            pptx.writeFile({ fileName: `Monthly_Report_${this.currentMonth}_${this.currentYear}.pptx` });

        } else if (format === 'image') {
             // Export all slides as separate images
             for (let i = 0; i < slideElements.length; i++) {
                 updateStatus(`Capturing Slide ${i + 1} of ${slideElements.length}...`);
                 const slideElement = slideElements[i];
                 await new Promise(resolve => setTimeout(resolve, 100));
                 
                 try {
                     const imgData = await htmlToImage.toPng(slideElement, {
                         pixelRatio: 2,
                         width: 1000,
                         height: 562.5,
                         backgroundColor: '#ffffff',
                         skipAutoScale: true
                     });
                     
                     // Download each slide as a separate image
                     const link = document.createElement('a');
                     link.href = imgData;
                     link.download = `Slide_${String(i + 1).padStart(2, '0')}_${this.currentMonth}_${this.currentYear}.png`;
                     link.click();
                     
                     // Small delay between downloads to prevent browser blocking
                     await new Promise(resolve => setTimeout(resolve, 200));
                 } catch (e) {
                     console.error(`Error capturing slide ${i+1}:`, e);
                 }
             }
         }

    } catch (error) {
        console.error('Export Error:', error);
        alert('An error occurred during export.');
    } finally {
        // Clean up DOM elements (container only exists for non-pptx-content exports)
        const container = document.querySelector('div[style*="position: absolute"][style*="z-index: 9999"]') as HTMLElement;
        if(container && document.body.contains(container)) document.body.removeChild(container);
        if(document.body.contains(overlay)) document.body.removeChild(overlay);
        this.isExporting = false;
        this.exportProgress = '';
    }
  }
}


