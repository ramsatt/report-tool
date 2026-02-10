import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx-js-style';
import { BehaviorSubject, Observable } from 'rxjs';
import { MonthlyData, DefectItem } from './monthly-connect.models';

@Injectable({
  providedIn: 'root'
})
export class MonthlyConnectService {
  private readonly STORAGE_KEY = 'monthly_connect_data';
  
  private monthlyDataSubject = new BehaviorSubject<MonthlyData | null>(null);
  public monthlyData$ = this.monthlyDataSubject.asObservable();

  constructor() {
    this.loadData();
  }

  // --- Data Loading ---

  loadData() {
    const stored = localStorage.getItem(this.STORAGE_KEY);
    if (stored) {
      try {
        const data = JSON.parse(stored);
        this.monthlyDataSubject.next(data);
        console.log('Data loaded from local storage');
      } catch (e) {
        console.error('Failed to parse stored data', e);
        this.resetData();
      }
    }
  }

  saveData(data: MonthlyData) {
    this.monthlyDataSubject.next(data);
    localStorage.setItem(this.STORAGE_KEY, JSON.stringify(data));
    console.log('Data saved to local storage:', data);
  }

  resetData() {
    localStorage.removeItem(this.STORAGE_KEY);
    this.monthlyDataSubject.next(null);
    console.log('Data reset');
  }

  // --- Excel Processing ---

  processExcelFile(file: File): Promise<void> {
    return new Promise((resolve, reject) => {
      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {
        try {
          const bstr: string = e.target.result;
          const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
          const data = this.parseWorkbook(wb);
          this.saveData(data);
          resolve();
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsBinaryString(file);
    });
  }

  private parseWorkbook(wb: XLSX.WorkBook): MonthlyData {
    const data: MonthlyData = {
      generalInfo: { month: '', year: new Date().getFullYear(), projectName: '' },
      deliveryMetrics: [],
      velocityStrategy: [],
      defectAnalysis: [],
      backlogItems: [],
      defectMetrics: [],
      roadmap: [],
      feedback: [],
      leavePlan: [],
      teamActions: []
    };

    // 1. General Info
    const wsGeneral = wb.Sheets['General Info'];
    if (wsGeneral) {
      data.generalInfo.coreHighlights = [];
      data.generalInfo.appHighlights = [];
      const rows = XLSX.utils.sheet_to_json(wsGeneral, { header: 1 }) as any[][];
      
      let currentSection = ''; // 'CORE', 'APP', or ''

      rows.forEach(row => {
        const key = (row[0] || '').toString().trim();
        const val = row[1];
        
        // General Metadata
        if (key === 'Month') data.generalInfo.month = val;
        if (key === 'Year') data.generalInfo.year = val;
        if (key === 'Project Name') data.generalInfo.projectName = val;
        
        // Detect Section Headers
        if (key.startsWith('Core Highlights')) {
          currentSection = 'CORE';
        } else if (key.startsWith('App Highlights')) {
          currentSection = 'APP';
        }

        // Collect Highlights based on current section
        if (key === 'Highlight' && val) {
          if (currentSection === 'CORE') {
             data.generalInfo.coreHighlights?.push(val);
          } else if (currentSection === 'APP') {
             data.generalInfo.appHighlights?.push(val);
          }
        }
      });
    }

    // 2. Delivery Metrics
    const wsDelivery = wb.Sheets['Delivery Metrics'];
    if (wsDelivery) {
      const rows = XLSX.utils.sheet_to_json(wsDelivery) as any[];
      data.deliveryMetrics = rows.map(row => ({
        stream: row['Stream'],
        sprint: row['Sprint'],
        month: this.excelDateToJSDate(row['Month']),
        committed: row['Committed'] || 0,
        delivered: row['Delivered'] || 0,
        deliveryPct: row['Delivery %'] || 0,
        features: row['Features Delivered'],
        deployedDate: this.excelDateToJSDate(row['Deployed Date']),
        deploymentStatus: row['Deployment Status'],
        bugs: row['Bugs'] || 0,
        comments: row['Comments']
      }));
    }

    // 3. Velocity Strategy
    const wsVelocity = wb.Sheets['Velocity Strategy'];
    if (wsVelocity) {
      const rows = XLSX.utils.sheet_to_json(wsVelocity) as any[];
      data.velocityStrategy = rows.map(row => ({
        month: this.excelDateToJSDate(row['Month']),
        sprint: row['Sprint'],
        committed: row['Committed'] || 0,
        delivered: row['Delivered'] || 0,
        deliveryPct: row['Delivery %'] || 0,
        comments: row['Comment'] || row['Comment '] || row['Comments'] || '' 
      }));
    }

    // 4. Defect Analysis
    const wsDefects = wb.Sheets['Defect Analysis'];
    if (wsDefects) {
      const rows = XLSX.utils.sheet_to_json(wsDefects) as any[];
      data.defectAnalysis = rows.map(row => this.mapDefectItem(row));
    }

    // 5. Backlog Items
    const wsBacklog = wb.Sheets['Backlog Items'];
    if (wsBacklog) {
      const rows = XLSX.utils.sheet_to_json(wsBacklog) as any[];
      data.backlogItems = rows.map(row => this.mapDefectItem(row));
    }

    // 6. Defect Metrics
    const wsDefectMetrics = wb.Sheets['Defect Metrics'];
    if (wsDefectMetrics) {
      const rows = XLSX.utils.sheet_to_json(wsDefectMetrics) as any[];
      data.defectMetrics = rows.map(row => this.mapDefectItem(row));
    }

    // 7. Roadmap
    const wsRoadmap = wb.Sheets['Forward Looking Roadmap'];
    if (wsRoadmap) {
      const rows = XLSX.utils.sheet_to_json(wsRoadmap) as any[];
      data.roadmap = rows.map(row => ({
        sno: row[this.scrubKey(row, 'S. No')],
        workItem: row[this.scrubKey(row, 'Work Item')],
        details: row[this.scrubKey(row, 'Work Item Details')],
        type: row[this.scrubKey(row, 'Type')],
        priority: row[this.scrubKey(row, 'Priority')],
        status: row[this.scrubKey(row, 'Discussion Status')]
      }));
    }

    // 8. Feedback
    const wsFeedback = wb.Sheets['Feedback'];
    if (wsFeedback) {
      const rows = XLSX.utils.sheet_to_json(wsFeedback) as any[];
      data.feedback = rows.map(row => ({
        date: this.scrubVal(row[this.scrubKey(row, 'Action Date')]) || '', 
        actionItem: this.scrubVal(row[this.scrubKey(row, 'Action Item')]),
        owner: this.scrubVal(row[this.scrubKey(row, 'Owner')]),
        status: this.scrubVal(row[this.scrubKey(row, 'Status')]),
        comments: this.scrubVal(row[this.scrubKey(row, 'Comments')])
      }));
    }

    // 9. Leave Plan
    const wsLeave = wb.Sheets['Leave Plan'];
    if (wsLeave) {
      const rows = XLSX.utils.sheet_to_json(wsLeave) as any[];
      data.leavePlan = rows.map(row => ({
        date: row['Date'],
        event: row['Event'],
        member: row['Team Member']
      }));
    }

    // 10. Team Actions
    const wsActions = wb.Sheets['Team Actions'];
    if (wsActions) {
      const rows = XLSX.utils.sheet_to_json(wsActions) as any[];
      data.teamActions = rows.map(row => ({
        actionItem: row['Action Item'],
        duration: row['Duration'],
        comments: row['Comments']
      }));
    }

    return data;
  }

  private mapDefectItem(row: any): DefectItem {
    return {
      id: row[this.scrubKey(row, 'ID')],
      type: row[this.scrubKey(row, 'Work Item Type')],
      title: row[this.scrubKey(row, 'Title')],
      priority: row[this.scrubKey(row, 'Priority')],
      createdDate: this.excelDateToJSDate(row[this.scrubKey(row, 'Created Date')]),
      resolvedDate: this.excelDateToJSDate(row[this.scrubKey(row, 'Resolved Date')]),
      assignedTo: row[this.scrubKey(row, 'Assigned To')],
      iteration: row[this.scrubKey(row, 'Iteration Path')],
      state: row[this.scrubKey(row, 'State')],
      tags: row[this.scrubKey(row, 'Tags')],
      defectType: row[this.scrubKey(row, 'Defect Type')],
      resolvedReason: row[this.scrubKey(row, 'Resolved Reason')],
      rootCause: row[this.scrubKey(row, 'Defect Root Cause')]
    };
  }

  private scrubKey(row: any, target: string): string {
    const keys = Object.keys(row);
    const found = keys.find(k => k.trim().replace(/\u200B/g, '') === target);
    return found || target;
  }

  private scrubVal(val: any): any {
    if (typeof val === 'string') {
      return val.trim().replace(/\u200B/g, '');
    }
    return val;
  }

  private excelDateToJSDate(serial: any): string {
    if (!serial) return '';
    if (typeof serial === 'string') return this.scrubVal(serial);
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400; 
    const date_info = new Date(utc_value * 1000);
    
    // Return YYYY-MM-DD format
    const year = date_info.getFullYear();
    const month = String(date_info.getMonth() + 1).padStart(2, '0');
    const day = String(date_info.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
}
