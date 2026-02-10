import { Injectable } from '@angular/core';

/**
 * Shared service for common slide generation utilities
 */
@Injectable({
  providedIn: 'root'
})
export class SlideGeneratorService {

  /**
   * Generate common tooltip styles for all slides
   */
  getTooltipStyles(): string {
    return `
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
        .custom-tooltip-wrapper.flip-top .custom-tooltip-content {
          top: auto;
          bottom: 100%;
        }
        .custom-tooltip-wrapper.flip-top .custom-tooltip-content::after {
          top: 100%;
          bottom: auto;
          border-color: #000048 transparent transparent transparent;
        }
        .custom-tooltip-wrapper.left-align .custom-tooltip-content {
          left: -5px;
          transform: translate(0, 10px);
        }
        .custom-tooltip-wrapper.left-align .custom-tooltip-content::after {
          left: 15px;
          margin-left: 0;
        }
        .custom-tooltip-wrapper.left-align.flip-top .custom-tooltip-content {
          transform: translate(0, -10px);
        }
      </style>
    `;
  }

  /**
   * Generate common header for slides
   */
  generateHeader(title: string, subtitle: string, currentMonth: string, currentYear: string): string {
    return `
      <div style="background: linear-gradient(90deg, #000048 0%, #00155c 100%); padding: 25px 50px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <div style="display: flex; align-items: center; gap: 15px;">
          <div style="width: 6px; height: 35px; background-color: #26C6DA;"></div>
          <div>
            <h2 style="color: white; font-size: 26px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">${title}</h2>
            <span style="color: #26C6DA; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 1.5px;">${subtitle}</span>
          </div>
        </div>
        <div style="color: rgba(255,255,255,0.8); font-size: 14px; font-weight: 400;">${currentMonth} ${currentYear}</div>
      </div>
    `;
  }

  /**
   * Generate common footer for slides
   */
  generateFooter(pageNumber?: number): string {
    const pageDisplay = pageNumber ? pageNumber.toString() : '<!--PAGE-->';
    return `
      <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 50px; background-color: white; display: flex; align-items: center; justify-content: space-between; padding: 0 50px; border-top: 1px solid #eee;">
        <div style="color: #666; font-size: 11px; font-weight: 600;">
          <span style="color: #000048; font-weight: 800;">Cognizant</span> &nbsp;|&nbsp; Caterpillar: Confidential Green
        </div>
        <div style="display: flex; align-items: center; gap: 10px;">
          <div style="color: #ccc; font-size: 12px;">Page</div>
          <div style="background-color: #000048; color: white; width: 24px; height: 24px; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 12px; border-radius: 4px;">${pageDisplay}</div>
        </div>
      </div>
      <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 4px; background: linear-gradient(to right, #000048 0%, #2979ff 100%);"></div>
    `;
  }

  /**
   * Wrap slide content in standard container
   */
  wrapSlideContent(content: string): string {
    return `
      <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #F8F9FA; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
        ${content}
      </div>
    `;
  }

  /**
   * Calculate totals for delivery data
   */
  calculateTotals(rows: any[]): { committed: number; delivered: number; achieved: string } {
    const totals = { committed: 0, delivered: 0, achieved: '0%' };
    rows.forEach(r => {
      totals.committed += (r.committed || 0);
      totals.delivered += (r.delivered || 0);
    });
    if (totals.committed > 0) {
      totals.achieved = Math.round((totals.delivered / totals.committed) * 100) + '%';
    }
    return totals;
  }

  /**
   * Chunk array into smaller arrays for pagination
   */
  chunkArray<T>(array: T[], chunkSize: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }
}
