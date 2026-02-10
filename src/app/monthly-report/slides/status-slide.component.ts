import { Component, Input } from '@angular/core';
import { MigrationData, AutomationData } from '../models/report-data.model';
import { SlideGeneratorService } from '../services/slide-generator.service';

@Component({
  selector: 'app-status-slide',
  standalone: true,
  template: ''
})
export class StatusSlideComponent {
  @Input() statusData: MigrationData[] | AutomationData[] = [];
  @Input() type: 'migration' | 'automation' = 'migration';
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  constructor(private slideGenerator: SlideGeneratorService) {}

  generateSlide(): string {
    const title = this.type === 'migration' ? 'Migration Status' : 'Automation Status';
    const subtitle = this.type === 'migration' ? 'Platform Migration Progress' : 'Test Automation Coverage';
    
    const header = this.slideGenerator.generateHeader(title, subtitle, this.currentMonth, this.currentYear);
    const footer = this.slideGenerator.generateFooter();

    const content = `
      <div style="padding: 30px 50px;">
        <div style="background: white; border-radius: 8px; box-shadow: 0 2px 15px rgba(0,0,0,0.05); overflow: hidden;">
          <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 12px;">
            <thead>
              <tr style="background-color: #000048; color: white;">
                ${this.generateTableHeaders()}
              </tr>
            </thead>
            <tbody>
              ${this.generateTableRows()}
            </tbody>
          </table>
        </div>
      </div>
    `;

    return this.slideGenerator.wrapSlideContent(header + content + footer);
  }

  private generateTableHeaders(): string {
    if (this.type === 'migration') {
      return `
        <th style="padding: 12px; font-weight: 600; text-align: left; width: 40%;">Module</th>
        <th style="padding: 12px; font-weight: 600; text-align: center; width: 20%;">Status</th>
        <th style="padding: 12px; font-weight: 600; text-align: center; width: 20%;">Progress</th>
        <th style="padding: 12px; font-weight: 600; text-align: left; width: 20%;">ETA</th>
      `;
    } else {
      return `
        <th style="padding: 12px; font-weight: 600; text-align: left; width: 40%;">Feature</th>
        <th style="padding: 12px; font-weight: 600; text-align: center; width: 20%;">Status</th>
        <th style="padding: 12px; font-weight: 600; text-align: center; width: 20%;">Coverage</th>
        <th style="padding: 12px; font-weight: 600; text-align: left; width: 20%;">Notes</th>
      `;
    }
  }

  private generateTableRows(): string {
    return this.statusData.map((row, i) => {
      if (this.type === 'migration') {
        const migRow = row as MigrationData;
        return `
          <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">
            <td style="padding: 15px 12px; color: #000048; font-weight: 600; vertical-align: middle;">${migRow.module}</td>
            <td style="padding: 15px 12px; text-align: center; vertical-align: middle;">
              <span style="padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 700; background: ${this.getStatusColor(migRow.status).bg}; color: ${this.getStatusColor(migRow.status).text};">
                ${migRow.status}
              </span>
            </td>
            <td style="padding: 15px 12px; text-align: center; vertical-align: middle;">
              <div style="width: 100%; background: #e0e0e0; border-radius: 10px; height: 20px; position: relative; overflow: hidden;">
                <div style="width: ${migRow.progress}%; background: linear-gradient(90deg, #2f78c4, #26C6DA); height: 100%; border-radius: 10px; display: flex; align-items: center; justify-content: center; color: white; font-size: 10px; font-weight: 700;">
                  ${migRow.progress}%
                </div>
              </div>
            </td>
            <td style="padding: 15px 12px; color: #666; vertical-align: middle;">${migRow.eta || 'TBD'}</td>
          </tr>
        `;
      } else {
        const autoRow = row as AutomationData;
        return `
          <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">
            <td style="padding: 15px 12px; color: #000048; font-weight: 600; vertical-align: middle;">${autoRow.feature}</td>
            <td style="padding: 15px 12px; text-align: center; vertical-align: middle;">
              <span style="padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 700; background: ${this.getStatusColor(autoRow.status).bg}; color: ${this.getStatusColor(autoRow.status).text};">
                ${autoRow.status}
              </span>
            </td>
            <td style="padding: 15px 12px; text-align: center; vertical-align: middle;">
              <div style="width: 100%; background: #e0e0e0; border-radius: 10px; height: 20px; position: relative; overflow: hidden;">
                <div style="width: ${autoRow.coverage}%; background: linear-gradient(90deg, #2f78c4, #26C6DA); height: 100%; border-radius: 10px; display: flex; align-items: center; justify-content: center; color: white; font-size: 10px; font-weight: 700;">
                  ${autoRow.coverage}%
                </div>
              </div>
            </td>
            <td style="padding: 15px 12px; color: #666; font-style: italic; vertical-align: middle;">${autoRow.notes || '-'}</td>
          </tr>
        `;
      }
    }).join('');
  }

  private getStatusColor(status: string): { bg: string; text: string } {
    const s = status.toLowerCase();
    if (s.includes('complete') || s.includes('done') || s.includes('success')) {
      return { bg: '#e8f5e9', text: '#2e7d32' };
    } else if (s.includes('progress') || s.includes('ongoing')) {
      return { bg: '#fff3e0', text: '#ef6c00' };
    } else if (s.includes('pending') || s.includes('planned')) {
      return { bg: '#e3f2fd', text: '#1976d2' };
    }
    return { bg: '#f5f5f5', text: '#666' };
  }
}
