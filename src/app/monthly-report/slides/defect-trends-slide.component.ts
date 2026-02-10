import { Component, Input } from '@angular/core';
import { DefectTrendData } from '../models/report-data.model';

@Component({
  selector: 'app-defect-trends-slide',
  standalone: true,
  template: ''
})
export class DefectTrendsSlideComponent {
  @Input() defectTrendsData: DefectTrendData[] = [];
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  generateSlide(): string {
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const today = new Date();
    const last6Months: string[] = [];
    
    for (let i = 0; i < 6; i++) {
      const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
      last6Months.push(`${monthNames[d.getMonth()]} ${d.getFullYear()}`);
    }

    const defectsByMonth = this.groupDefectsByMonth(last6Months);
    const metrics = this.calculateMetrics(defectsByMonth, last6Months);

    return `
      <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #020617; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', 'Plus Jakarta Sans', sans-serif; color: #f8fafc;">
        ${this.generateStyles()}
        ${this.generateHeader()}
        ${this.generateSummaryStrip(metrics)}
        ${this.generateMasonryContent(last6Months, defectsByMonth)}
        ${this.generateFooter()}
      </div>
    `;
  }

  private groupDefectsByMonth(last6Months: string[]): { [key: string]: DefectTrendData[] } {
    const defectsByMonth: { [key: string]: DefectTrendData[] } = {};
    last6Months.forEach(m => defectsByMonth[m] = []);

    this.defectTrendsData.forEach(d => {
      const targetMonth = last6Months.find(m => 
        d.month && d.month.includes(m.split(' ')[0]) && d.month.includes(m.split(' ')[1])
      );
      if (targetMonth) {
        defectsByMonth[targetMonth].push(d);
      }
    });

    return defectsByMonth;
  }

  private calculateMetrics(defectsByMonth: { [key: string]: DefectTrendData[] }, last6Months: string[]) {
    let totalVolume = 0;
    let zeroDefectMonths = 0;
    let peakCycleMonth = "-";
    let maxDefects = -1;
    const totalTypeCounts: { [key: string]: number } = {};

    last6Months.forEach(month => {
      const count = defectsByMonth[month].length;
      totalVolume += count;
      if (count === 0) zeroDefectMonths++;
      if (count > 0 && count > maxDefects) {
        maxDefects = count;
        peakCycleMonth = month;
      }

      defectsByMonth[month].forEach(d => {
        let t = d.issueType || 'Unknown';
        if (t.toLowerCase().includes('coding')) t = 'Coding';
        else if (t.toLowerCase().includes('data')) t = 'Data';
        else if (t.toLowerCase().includes('config')) t = 'Config';
        else if (t.toLowerCase().includes('int')) t = 'Integration';
        else if (t.toLowerCase().includes('user')) t = 'User';
        totalTypeCounts[t] = (totalTypeCounts[t] || 0) + 1;
      });
    });

    if (maxDefects === -1) peakCycleMonth = "None";
    const stabilityIndex = Math.round((zeroDefectMonths / 6) * 100);

    let primaryRootCause = "None";
    let primaryRootCausePct = 0;
    if (totalVolume > 0) {
      const sortedTypes = Object.entries(totalTypeCounts).sort((a, b) => b[1] - a[1]);
      if (sortedTypes.length > 0) {
        primaryRootCause = sortedTypes[0][0];
        primaryRootCausePct = Math.round((sortedTypes[0][1] / totalVolume) * 100);
      }
    }

    return {
      totalVolume,
      stabilityIndex,
      peakCycleMonth,
      maxDefects: maxDefects === -1 ? 0 : maxDefects,
      primaryRootCause,
      primaryRootCausePct
    };
  }

  private generateStyles(): string {
    return `
      <style>
        .slide-11-container {
          --bg-deep: #020617;
          --surface-glass: rgba(15, 23, 42, 0.6);
          --surface-accent: #1e293b;
          --border: rgba(51, 65, 85, 0.4);
          --accent-primary: #38bdf8;
          --text-main: #f8fafc;
          --text-dim: #94a3b8;
        }
        .s-metric { display: flex; flex-direction: column; gap: 0px; }
        .s-label { font-size: 8px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 2px; }
        .s-val { font-size: 15px; font-weight: 700; color: #f8fafc; }
        .s-val span { font-size: 10px; color: #38bdf8; font-weight: 500; margin-left: 4px; }
        .month-card {
          break-inside: avoid;
          margin-bottom: 12px;
          background: rgba(15, 23, 42, 0.6);
          border: 1px solid rgba(51, 65, 85, 0.4);
          border-radius: 8px;
          box-sizing: border-box;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .pill {
          background: #1e293b;
          padding: 1px 7px;
          border-radius: 20px;
          font-size: 9px;
          font-weight: 700;
          border: 1px solid rgba(51, 65, 85, 0.4);
          color: #fff;
        }
        .pill.active {
          background: rgba(56, 189, 248, 0.1);
          color: #38bdf8;
          border-color: rgba(56, 189, 248, 0.3);
        }
        .data-table th {
          font-size: 7px;
          text-transform: uppercase;
          text-align: left;
          padding: 4px 8px;
          color: #94a3b8;
          background: rgba(0,0,0,0.2);
          font-weight: 600;
        }
        .data-table td {
          padding: 4px 8px;
          font-size: 9px;
          border-bottom: 1px solid rgba(255,255,255,0.03);
          vertical-align: middle;
          color: #f8fafc;
          line-height: 1.1;
        }
        .type-tag {
          display: inline-block;
          padding: 1px 5px;
          border-radius: 3px;
          font-size: 7px;
          font-weight: 600;
          background: rgba(255,255,255,0.05);
          border: 1px solid rgba(255,255,255,0.1);
        }
        .tooltip-wrapper {
          position: relative;
          cursor: help;
          text-decoration: underline dotted #38bdf8;
          text-decoration-thickness: 1px;
        }
        .tooltip-text {
          visibility: hidden;
          width: 200px;
          background-color: #0f172a;
          color: #f8fafc;
          text-align: left;
          border-radius: 6px;
          padding: 8px 10px;
          position: absolute;
          z-index: 9999;
          left: 0;
          top: 100%;
          margin-top: 4px;
          opacity: 0;
          transition: opacity 0.2s, visibility 0.2s;
          border: 1px solid #38bdf8;
          box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.8);
          font-family: 'Segoe UI', sans-serif;
          font-size: 10px;
          font-weight: 400;
          line-height: 1.4;
          white-space: normal;
        }
        .tooltip-wrapper:hover .tooltip-text {
          visibility: visible;
          opacity: 1;
        }
      </style>
    `;
  }

  private generateHeader(): string {
    return `
      <div style="padding: 10px 30px; background: rgba(15, 23, 42, 0.3); border-bottom: 1px solid rgba(51, 65, 85, 0.4); display: flex; justify-content: space-between; align-items: center; height: 48px; box-sizing: border-box;">
        <div style="display: flex; align-items: center;">
          <div style="width: 4px; height: 22px; background: #38bdf8; border-radius: 2px; margin-right: 12px;"></div>
          <h1 style="font-size: 19px; font-weight: 700; margin: 0; display: flex; align-items: center; color: #f8fafc;">Defect Analysis</h1>
          <p style="color: #38bdf8; font-size: 9px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px; margin: 0 0 0 16px; padding-top: 4px;">Last 6 Months Quality Metrics</p>
        </div>
        <div style="text-align: right; color: #94a3b8; font-size: 10px; font-family: Consolas, monospace;">
          <div>${this.currentMonth.toUpperCase()} ${this.currentYear}</div>
        </div>
      </div>
    `;
  }

  private generateSummaryStrip(metrics: any): string {
    return `
      <div style="display: flex; gap: 40px; padding: 8px 30px; background: rgba(2, 6, 23, 0.5); border-bottom: 1px solid rgba(51, 65, 85, 0.4); height: 44px; box-sizing: border-box; align-items: center;">
        <div class="s-metric">
          <span class="s-label">Total Volume</span>
          <span class="s-val">${metrics.totalVolume} <span>Defects</span></span>
        </div>
        <div class="s-metric">
          <span class="s-label">Stability Index</span>
          <span class="s-val">${metrics.stabilityIndex}% <span>Zero-Months</span></span>
        </div>
        <div class="s-metric">
          <span class="s-label">Peak Cycle</span>
          <span class="s-val">${metrics.peakCycleMonth} <span>(${metrics.maxDefects})</span></span>
        </div>
        <div class="s-metric">
          <span class="s-label">Primary Root Cause</span>
          <span class="s-val" style="color: ${metrics.primaryRootCause === 'Coding' ? '#818cf8' : '#fbbf24'};">${metrics.primaryRootCause} <span>(${metrics.primaryRootCausePct}%)</span></span>
        </div>
      </div>
    `;
  }

  private generateMasonryContent(last6Months: string[], defectsByMonth: { [key: string]: DefectTrendData[] }): string {
    return `
      <div class="slide-11-container" style="padding: 15px 30px; height: 430px; overflow: hidden; column-count: 3; column-gap: 15px; width: 100%; box-sizing: border-box; background: radial-gradient(circle at 50% 50%, rgba(56, 189, 248, 0.03) 0%, transparent 70%);">
        ${last6Months.map(month => this.generateMonthCard(month, defectsByMonth[month])).join('')}
      </div>
    `;
  }

  private generateMonthCard(month: string, defects: DefectTrendData[]): string {
    const defectsBySprint: { [key: string]: DefectTrendData[] } = {};
    defects.forEach(d => {
      const sp = d.closedSprint || 'Other';
      if (!defectsBySprint[sp]) defectsBySprint[sp] = [];
      defectsBySprint[sp].push(d);
    });

    const sortedSprints = Object.keys(defectsBySprint).sort((a, b) => {
      if (a === 'Other') return 1;
      if (b === 'Other') return -1;
      const numA = parseInt(a.replace(/\D/g, '') || '0');
      const numB = parseInt(b.replace(/\D/g, '') || '0');
      return numA - numB;
    });

    const borderStyle = defects.length > 0 ? 'rgba(56, 189, 248, 0.4)' : 'rgba(51, 65, 85, 0.4)';
    const bgStyle = defects.length > 0 ? 'rgba(15, 23, 42, 0.8)' : 'rgba(15, 23, 42, 0.6)';

    return `
      <div class="month-card" style="border-color: ${borderStyle}; background: ${bgStyle};">
        <div style="padding: 6px 10px; display: flex; justify-content: space-between; align-items: center; border-bottom: ${defects.length > 0 ? '1px solid rgba(51, 65, 85, 0.4)' : 'none'}; background: rgba(255, 255, 255, 0.02); min-height: 26px;">
          <span style="font-size: 12px; font-weight: 700; color: #f8fafc;">${month}</span>
          <span class="pill ${defects.length > 0 ? 'active' : ''}">${defects.length} Total</span>
        </div>
        ${defects.length > 0 ? `
          <div style="width: 100%;">
            <table class="data-table" style="width: 100%; border-collapse: collapse; table-layout: fixed;">
              <thead>
                <tr>
                  <th style="width: 38%;">ID</th>
                  <th style="width: 18%; text-align: center;">Rel.</th>
                  <th style="width: 44%; text-align: right;">Type</th>
                </tr>
              </thead>
              <tbody>
                ${sortedSprints.map(sprint => `
                  <tr style="background: rgba(56, 189, 248, 0.05);">
                    <td colspan="3" style="font-size: 7px; font-weight: 700; color: #38bdf8; text-transform: uppercase; letter-spacing: 0.05em; padding: 2px 8px;">${sprint}</td>
                  </tr>
                  ${defectsBySprint[sprint].map(d => this.generateDefectRow(d)).join('')}
                `).join('')}
              </tbody>
            </table>
          </div>
        ` : ''}
      </div>
    `;
  }

  private generateDefectRow(d: DefectTrendData): string {
    let typeColor = '#94a3b8';
    let typeLabel = d.issueType || 'Unk';
    const t = typeLabel.toLowerCase();
    
    if (t.includes('data')) typeColor = '#fbbf24';
    else if (t.includes('code') || t.includes('coding')) typeColor = '#818cf8';
    else if (t.includes('int')) typeColor = '#2dd4bf';
    else if (t.includes('config')) typeColor = '#f472b6';

    let displayType = typeLabel;
    if (displayType.length > 9) displayType = displayType.substring(0, 9) + '..';

    return `
      <tr>
        <td style="font-family: Consolas, monospace; opacity: 0.9;">
          <div class="tooltip-wrapper">
            ${d.id}
            <div class="tooltip-text">${d.desc ? d.desc.replace(/"/g, '&quot;') : 'No description available'}</div>
          </div>
        </td>
        <td style="text-align: center;">${d.createdSprint ? d.createdSprint.replace('Sprint ', 'S') : '-'}</td>
        <td style="text-align: right;">
          <span class="type-tag" style="color: ${typeColor}; border-color: ${typeColor}40;">${displayType}</span>
        </td>
      </tr>
    `;
  }

  private generateFooter(): string {
    return `
      <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 40px; background: #000; display: flex; align-items: center; justify-content: space-between; padding: 0 30px; border-top: 1px solid rgba(51, 65, 85, 0.4);">
        <div style="color: #f8fafc; font-size: 9px; font-weight: 700;">
          Cognizant <span style="font-weight: 400; color: #94a3b8;">&nbsp;|&nbsp; Caterpillar: Confidential</span>
        </div>
        <div style="display: flex; align-items: center; gap: 10px;">
          <span style="color: #94a3b8; font-size: 9px;">Reporting Engine v2.4</span>
          <div style="background: #1e293b; padding: 2px 6px; border-radius: 4px; border: 1px solid rgba(51, 65, 85, 0.4); color: #f8fafc; font-size: 9px;">Slide 11</div>
        </div>
      </div>
    `;
  }
}
