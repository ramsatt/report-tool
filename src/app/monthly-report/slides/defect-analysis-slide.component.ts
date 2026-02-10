import { Component, Input } from '@angular/core';
import { DefectData } from '../models/report-data.model';
import { SlideGeneratorService } from '../services/slide-generator.service';

@Component({
  selector: 'app-defect-analysis-slide',
  standalone: true,
  template: ''
})
export class DefectAnalysisSlideComponent {
  @Input() defectData: DefectData[] = [];
  @Input() type: 'overview' | 'backlog' = 'overview';
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  constructor(private slideGenerator: SlideGeneratorService) {}

  generateSlides(): string[] {
    const slides: string[] = [];
    const chunkSize = this.type === 'overview' ? 10 : 8;
    const chunks = this.slideGenerator.chunkArray(this.defectData, chunkSize);

    chunks.forEach((chunk, index) => {
      if (this.type === 'overview') {
        slides.push(this.generateOverviewSlide(chunk, index, chunks.length));
      } else {
        slides.push(this.generateBacklogSlide(chunk, index, chunks.length));
      }
    });

    return slides;
  }

  private generateOverviewSlide(chunk: DefectData[], index: number, total: number): string {
    const tooltipStyles = this.slideGenerator.getTooltipStyles();
    const header = `
      <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 15px 40px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <div style="display: flex; align-items: center; gap: 15px;">
          <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
          <div>
            <h2 style="color: white; font-size: 24px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Defect Analysis Report</h2>
            <span style="color: #26C6DA; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Overview ${total > 1 ? `(${index + 1}/${total})` : ''}</span>
          </div>
        </div>
        <div style="color: rgba(255,255,255,0.8); font-size: 12px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
      </div>
    `;

    const content = `
      <div style="padding: 20px 40px; height: 435px; overflow-y: auto;">
        <div style="background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); overflow: hidden;">
          <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 11px;">
            <thead>
              <tr style="background-color: #000048; color: white;">
                <th style="padding: 10px; text-align: left;">S.No</th>
                <th style="padding: 10px; text-align: left;">Description</th>
                <th style="padding: 10px; text-align: center;">Priority</th>
                <th style="padding: 10px; text-align: center;">Status</th>
                <th style="padding: 10px; text-align: left;">Remarks</th>
              </tr>
            </thead>
            <tbody>
              ${chunk.length > 0 ? chunk.map((d, i) => `
                <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9f9f9'};">
                  <td style="padding: 8px; color: #666;">${d.id}</td>
                  <td style="padding: 8px; color: #333; font-weight: 600;">${d.desc}</td>
                  <td style="padding: 8px; text-align: center;">${d.priority}</td>
                  <td style="padding: 8px; text-align: center;">
                    <span style="padding: 2px 8px; border-radius: 10px; font-size: 10px; background: ${d.status === 'RESOLVED' || d.status === 'Closed' ? '#e8f5e9' : '#fff3e0'}; color: ${d.status === 'RESOLVED' || d.status === 'Closed' ? '#2e7d32' : '#ef6c00'}; font-weight: 700;">
                      ${d.status}
                    </span>
                  </td>
                  <td style="padding: 8px; color: #666; font-style: italic;">${d.intro || ''}</td>
                </tr>
              `).join('') : '<tr><td colspan="5" style="padding: 20px; text-align: center;">No defects found for this selection.</td></tr>'}
            </tbody>
          </table>
        </div>
      </div>
    `;

    const footer = this.slideGenerator.generateFooter();
    return tooltipStyles + this.slideGenerator.wrapSlideContent(header + content + footer);
  }

  private generateBacklogSlide(chunk: DefectData[], index: number, total: number): string {
    const tooltipStyles = this.slideGenerator.getTooltipStyles();
    const header = this.slideGenerator.generateHeader(
      'Defect Report',
      'Backlog Items - DQME App',
      this.currentMonth,
      this.currentYear
    );

    const content = `
      <div style="padding: 20px 50px; height: 410px; display: flex; flex-direction: column;">
        <div style="overflow: visible; border-radius: 8px; box-shadow: 0 2px 15px rgba(0,0,0,0.05); background: white;">
          <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-family: 'Segoe UI', sans-serif; font-size: 10px; table-layout: fixed;">
            <thead>
              <tr style="background-color: #000048; color: white;">
                <th style="padding: 10px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 4%;">S.No</th>
                <th style="padding: 10px; font-weight: 600; text-align: left; border-right: 1px solid rgba(255,255,255,0.1); width: 28%;">Description</th>
                <th style="padding: 10px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 6%;">Priority</th>
                <th style="padding: 10px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 12%;">Created On</th>
                <th style="padding: 10px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 14%;">Assigned</th>
                <th style="padding: 10px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 14%;">When Introduced?</th>
                <th style="padding: 10px; font-weight: 600; text-align: center; border-right: 1px solid rgba(255,255,255,0.1); width: 8%;">ETA</th>
                <th style="padding: 10px; font-weight: 600; text-align: center; width: 14%;">Status</th>
              </tr>
            </thead>
            <tbody>
              ${chunk.map((row, i) => `
                <tr style="border-bottom: 1px solid #eee; background-color: ${i % 2 === 0 ? '#fff' : '#f9fbfd'};">
                  <td style="padding: 7px 10px; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.id}</td>
                  <td style="padding: 7px 10px; color: #000048; font-weight: 500; vertical-align: middle; border-right: 1px solid #eee; line-height: 1.3;">
                    <div class="custom-tooltip-wrapper ${i >= chunk.length - 3 ? 'flip-top' : ''}">
                      <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.desc}</div>
                      <div class="custom-tooltip-content">${row.desc}</div>
                    </div>
                  </td>
                  <td style="padding: 7px 10px; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.priority}</td>
                  <td style="padding: 7px 10px; color: #666; font-weight: 500; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.created || '-'}</td>
                  <td style="padding: 7px 10px; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">
                    <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.assigned || '-'}</div>
                  </td>
                  <td style="padding: 7px 10px; color: #444; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">
                    <div style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${row.intro || '-'}</div>
                  </td>
                  <td style="padding: 7px 10px; color: #666; font-style: italic; text-align: center; vertical-align: middle; border-right: 1px solid #eee;">${row.eta || '-'}</td>
                  <td style="padding: 7px 10px; vertical-align: middle; text-align: center;">
                    <span style="display: inline-block; padding: 2px 6px; border-radius: 10px; font-size: 10px; font-weight: 700; background-color: ${row.statusColor || '#f5f5f5'}; color: ${row.statusText || '#666'}; white-space: nowrap;">${row.status}</span>
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;

    const footer = this.slideGenerator.generateFooter();
    return tooltipStyles + this.slideGenerator.wrapSlideContent(header + content + footer);
  }
}
