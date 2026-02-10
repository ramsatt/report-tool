import { Component, Input } from '@angular/core';
import { DeliveryData, DeliveryTotals } from '../models/report-data.model';
import { SlideGeneratorService } from '../services/slide-generator.service';

@Component({
  selector: 'app-delivery-metrics-slide',
  standalone: true,
  template: ''
})
export class DeliveryMetricsSlideComponent {
  @Input() deliveryData: DeliveryData[] = [];
  @Input() type: 'core' | 'app' = 'core';
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  constructor(private slideGenerator: SlideGeneratorService) {}

  generateSlides(): string[] {
    const slides: string[] = [];
    const chunkSize = this.type === 'core' ? 2 : 2;
    const chunks = this.slideGenerator.chunkArray(this.deliveryData, chunkSize);
    const totals = this.slideGenerator.calculateTotals(this.deliveryData);

    chunks.forEach((chunk) => {
      if (this.type === 'core') {
        slides.push(this.generateCoreSlide(chunk, totals));
      } else {
        slides.push(this.generateAppSlide(chunk, totals));
      }
    });

    return slides;
  }

  private generateCoreSlide(chunk: DeliveryData[], totals: DeliveryTotals): string {
    const header = `
      <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 15px 40px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 10; position: relative;">
        <div style="display: flex; align-items: center; gap: 15px;">
          <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
          <div>
            <h2 style="color: white; font-size: 24px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">Delivery Metrics</h2>
            <span style="color: #26C6DA; font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">Core Platform</span>
          </div>
        </div>
        <div style="color: rgba(255,255,255,0.8); font-size: 13px; font-weight: 400;">${this.currentMonth} ${this.currentYear}</div>
      </div>
    `;

    const content = `
      <div style="padding: 15px 40px; display: flex; flex-direction: column; gap: 12px; height: 435px; justify-content: flex-start;">
        ${chunk.map((row) => this.generateCoreSprintCard(row)).join('')}
        ${this.generateGrandTotal(totals)}
      </div>
    `;

    const footer = this.slideGenerator.generateFooter();
    return this.slideGenerator.wrapSlideContent(header + content + footer);
  }

  private generateCoreSprintCard(row: DeliveryData): string {
    return `
      <div style="background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); overflow: hidden; border-left: 5px solid #000048; display: flex; flex-direction: column;">
        <div style="padding: 8px 21px; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; justify-content: space-between; background: linear-gradient(to right, #f8f9fa, #fff);">
          <div style="display: flex; align-items: center; gap: 12px;">
            <div style="width: 32px; height: 32px; background-color: #e8eaf6; border-radius: 6px; display: flex; align-items: center; justify-content: center; color: #000048;">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
            </div>
            <div>
              <h4 style="margin: 0; color: #000048; font-size: 15px; font-weight: 700;">${row.sprintMonth.split('\n')[0]}</h4>
              <span style="font-size: 11px; color: #666;">${row.sprintMonth.split('\n')[1] || ''}</span>
            </div>
          </div>
          <div style="display: flex; gap: 30px;">
            <div style="text-align: center; min-width: 70px;">
              <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600;">Committed</div>
              <div style="font-size: 16px; font-weight: 700; color: #333;">${row.committed}</div>
            </div>
            <div style="text-align: center; min-width: 70px;">
              <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600;">Delivered</div>
              <div style="font-size: 16px; font-weight: 700; color: #333;">${row.delivered}</div>
            </div>
            <div style="text-align: center; min-width: 70px;">
              <div style="font-size: 9px; color: #888; text-transform: uppercase; font-weight: 600;">Achieved</div>
              <div style="font-size: 16px; font-weight: 700; color: ${parseInt(row.achieved) >= 100 ? '#2e7d32' : '#ef6c00'};">${row.achieved}</div>
            </div>
          </div>
        </div>
        <div style="padding: 8px 21px; background-color: #fff; flex-grow: 1;">
          <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; border-bottom: 1px solid #f0f0f0; padding-bottom: 4px;">
            <div style="font-size: 10px; font-weight: 700; color: #000048; text-transform: uppercase;">Features Delivered</div>
            <div style="display: flex; align-items: center; gap: 15px;">
              <div style="display: flex; align-items: center; gap: 5px;">
                <span style="color: #666; font-size: 9px; font-weight: 600;">Deployment:</span>
                <span style="padding: 1px 6px; border-radius: 4px; background-color: ${row.deploymentStatus === 'Success' ? '#e8f5e9' : '#f5f5f5'}; color: ${row.deploymentStatus === 'Success' ? '#2e7d32' : '#666'}; font-weight: 700; font-size: 9px;">${row.deploymentStatus}</span>
              </div>
              <div style="display: flex; align-items: center; gap: 5px;">
                <span style="color: #666; font-size: 9px; font-weight: 600;">Bugs:</span>
                <span style="color: ${row.bugs > 0 ? '#d32f2f' : '#666'}; font-weight: 700; font-size: 9px;">${row.bugs}</span>
              </div>
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
    `;
  }

  private generateAppSlide(chunk: DeliveryData[], totals: DeliveryTotals): string {
    const header = `
      <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 20px 40px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border-bottom: 1px solid #334155;">
        <div style="display: flex; align-items: center; gap: 15px;">
          <div style="width: 5px; height: 35px; background-color: #26C6DA; border-radius: 2px;"></div>
          <div>
            <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: -0.5px;">Delivery Metrics</h2>
            <span style="color: #26C6DA; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 2px;">App Platform — Release Streams</span>
          </div>
        </div>
        <div style="text-align: right; color: rgba(255,255,255,0.7); font-family: Consolas, monospace; font-size: 12px;">
          <div style="font-weight: 700; color: white; letter-spacing: 1px;">${this.currentMonth.toUpperCase()} ${this.currentYear}</div>
          <div style="font-size: 10px;">Cycle: Q1-${this.currentYear.substring(2)}</div>
        </div>
      </div>
    `;

    const content = `
      <div style="padding: 20px 40px; display: flex; flex-direction: column; gap: 20px; height: 410px;">
        ${chunk.map((row) => this.generateAppSprintCard(row)).join('')}
      </div>
      ${this.generateAppGrandTotal(totals)}
    `;

    const footer = `
      <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
        <div style="color: #666; font-size: 11px; font-weight: 600;">
          <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
        </div>
        <div style="display: flex; align-items: center; gap: 10px;">
          <div style="color: #ccc; font-size: 12px;">Reporting Engine v2.4</div>
          <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;"><!--PAGE--></div>
        </div>
      </div>
      <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
    `;

    return this.slideGenerator.wrapSlideContent(header + content + footer);
  }

  private generateAppSprintCard(row: DeliveryData): string {
    return `
      <div style="background: white; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); overflow: hidden; display: flex; flex-direction: column; flex: 1;">
        <div style="padding: 15px 20px; border-bottom: 1px solid #f1f5f9; display: flex; align-items: center; justify-content: space-between; background: #fff;">
          <div style="display: flex; align-items: center; gap: 15px;">
            <div style="width: 40px; height: 40px; background-color: #f1f5f9; border-radius: 8px; display: flex; align-items: center; justify-content: center; color: #0f172a;">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
            </div>
            <div>
              <div style="font-size: 16px; font-weight: 700; color: #0f172a;">${row.sprintMonth.split('\n')[0]}</div>
              <div style="font-size: 11px; color: #64748b;">${row.sprintMonth.split('\n')[1] || ''}</div>
            </div>
          </div>
          <div style="display: flex; gap: 30px; align-items: center;">
            <div style="text-align: center;">
              <span style="display: block; font-size: 9px; text-transform: uppercase; color: #64748b; font-weight: 700; margin-bottom: 2px;">Committed</span>
              <span style="font-size: 18px; font-weight: 700; color: #0f172a;">${row.committed}</span>
            </div>
            <div style="text-align: center;">
              <span style="display: block; font-size: 9px; text-transform: uppercase; color: #64748b; font-weight: 700; margin-bottom: 2px;">Delivered</span>
              <span style="font-size: 18px; font-weight: 700; color: #0f172a;">${row.delivered}</span>
            </div>
            <div style="width: 40px; height: 40px; border-radius: 50%; border: 3px solid #e2e8f0; border-top-color: #26C6DA; display: flex; align-items: center; justify-content: center; position: relative;">
              <span style="font-size: 10px; font-weight: 800; color: #0d9488;">${row.achieved}</span>
            </div>
          </div>
        </div>
        <div style="padding: 8px 20px; background-color: #f8fafc; border-bottom: 1px solid #f1f5f9; display: flex; gap: 20px; font-size: 10px; align-items: center;">
          <div><span style="color: #64748b; font-weight: 600;">Deployment:</span> <span style="padding: 2px 8px; border-radius: 4px; font-weight: 700; background: ${row.deploymentStatus==='Success'?'#dcfce7':'#fee2e2'}; color: ${row.deploymentStatus==='Success'?'#166534':'#991b1b'}; border: 1px solid ${row.deploymentStatus==='Success'?'#bbf7d0':'#fecaca'};">${row.deploymentStatus || '-'}</span></div>
          <div><span style="color: #64748b; font-weight: 600;">Bugs Found:</span> <span style="padding: 2px 8px; border-radius: 4px; font-weight: 700; background: ${row.bugs > 0 ? '#fee2e2' : '#f1f5f9'}; color: ${row.bugs > 0 ? '#991b1b' : '#64748b'};">${row.bugs}</span></div>
          <div style="flex: 1; text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;"><span style="color: #64748b; font-weight: 600;">Notes:</span> <span style="color: #64748b; font-style: italic;">${row.comments || 'No significant notes'}</span></div>
        </div>
        <div style="padding: 12px 20px; flex-grow: 1; background: #fff;">
          <ul style="list-style: none; padding: 0; margin: 0; display: grid; grid-template-columns: 1fr 1fr; gap: 6px 30px;">
            ${row.features.map((f: string) => `
              <li style="font-size: 10px; color: #334155; position: relative; padding-left: 14px; line-height: 1.4;">
                <span style="position: absolute; left: 0; color: #26C6DA; font-weight: 700;">→</span>
                ${f.length > 55 ? f.substring(0, 55) + '...' : f}
              </li>
            `).join('')}
          </ul>
        </div>
      </div>
    `;
  }

  private generateGrandTotal(totals: DeliveryTotals): string {
    return `
      <div style="background: linear-gradient(90deg, #00155c 0%, #0033A0 100%); border-radius: 6px; padding: 10px 21px; display: flex; align-items: center; justify-content: space-between; color: white; margin-top: auto;">
        <div style="font-size: 12px; font-weight: 700; text-transform: uppercase;">Grand Total</div>
        <div style="display: flex; gap: 30px;">
          <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
            <span style="font-size: 9px; opacity: 0.7; text-transform: uppercase;">Committed</span>
            <span style="font-size: 14px; font-weight: 700;">${totals.committed}</span>
          </div>
          <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
            <span style="font-size: 9px; opacity: 0.7; text-transform: uppercase;">Delivered</span>
            <span style="font-size: 14px; font-weight: 700;">${totals.delivered}</span>
          </div>
          <div style="display: flex; flex-direction: row; align-items: center; gap: 8px; min-width: 70px; justify-content: center;">
            <span style="font-size: 9px; text-transform: uppercase; color: #26C6DA; font-weight: 700;">Achieved</span>
            <span style="font-size: 14px; font-weight: 700; color: #26C6DA;">${totals.achieved}</span>
          </div>
        </div>
      </div>
    `;
  }

  private generateAppGrandTotal(totals: DeliveryTotals): string {
    return `
      <div style="position: absolute; bottom: 50px; left: 0; width: 100%; background: linear-gradient(90deg, #0f172a 0%, #1e293b 100%); padding: 12px 50px; display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #334155;">
        <div style="font-size: 13px; font-weight: 700; text-transform: uppercase; color: #26C6DA; letter-spacing: 1px;">Grand Total Performance</div>
        <div style="display: flex; gap: 40px; color: white;">
          <div style="text-align: center;">
            <span style="display: block; font-size: 8px; text-transform: uppercase; opacity: 0.7; margin-bottom: 2px;">Total Committed</span>
            <span style="font-size: 16px; font-weight: 700;">${totals.committed}</span>
          </div>
          <div style="text-align: center;">
            <span style="display: block; font-size: 8px; text-transform: uppercase; opacity: 0.7; margin-bottom: 2px;">Total Delivered</span>
            <span style="font-size: 16px; font-weight: 700;">${totals.delivered}</span>
          </div>
          <div style="text-align: center;">
            <span style="display: block; font-size: 8px; text-transform: uppercase; opacity: 0.7; margin-bottom: 2px;">Overall Achievement</span>
            <span style="font-size: 16px; font-weight: 700; color: #26C6DA;">${totals.achieved}</span>
          </div>
        </div>
      </div>
    `;
  }
}
