import { Component, Input } from '@angular/core';

@Component({
  selector: 'app-title-slide',
  standalone: true,
  template: ''
})
export class TitleSlideComponent {
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  generateTitleSlide(): string {
    return `
      <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: #000048; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
        <div style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; background-color: #000048;"></div>
        <div style="position: absolute; top: 0; right: 0; width: 55%; height: 100%; background: linear-gradient(135deg, #00155c 0%, #000048 100%); clip-path: polygon(25% 0, 100% 0, 100% 100%, 0% 100%);"></div>
        <div style="position: absolute; bottom: 0; left: 0; width: 100%; height: 12px; background: linear-gradient(to right, #2f78c4, #26C6DA);"></div>
        
        <div style="position: absolute; top: 50px; left: 60px;">
          <div style="width: 50px; height: 5px; background-color: #26C6DA; margin-top: 10px;"></div>
        </div>

        <div style="position: absolute; top: 40px; right: 60px; display: flex; align-items: center; gap: 30px;">
          <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAlgAAAJYCAYAAAC+ZpjcAA..." style="height: 60px;" alt="Cognizant">
          <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAlgAAAJYCAYAAAC+ZpjcAA..." style="height: 60px;" alt="CAT">
        </div>

        <div style="position: absolute; top: 180px; left: 60px; z-index: 10;">
          <div style="display: inline-block; padding: 6px 16px; background-color: rgba(255,255,255,0.08); border-left: 4px solid #26C6DA; color: #7BE4F1; text-transform: uppercase; font-size: 14px; font-weight: 700; letter-spacing: 1.5px; margin-bottom: 24px;">
            Project Status Report
          </div>
          <h1 style="font-size: 64px; font-weight: 300; color: white; margin: 0; line-height: 1.1;">
            CAT Technology<br>
            <span style="font-weight: 800;">DQME – Core/App</span>
          </h1>
          
          <div style="margin-top: 50px; display: flex; align-items: center; gap: 20px;">
            <div style="background-color: #2f78c4; color: white; padding: 12px 30px; font-weight: bold; font-size: 20px; box-shadow: 0 10px 20px rgba(0,0,0,0.2);">
              ${this.currentMonth.toUpperCase()} <span style="font-weight: 300; opacity: 0.8; margin-left: 5px;">${this.currentYear}</span>
            </div>
            <div style="display: flex; flex-direction: column; justify-content: center; border-left: 1px solid rgba(255,255,255,0.3); padding-left: 20px; height: 50px;">
              <span style="color: rgba(255,255,255,0.6); font-size: 12px; text-transform: uppercase; letter-spacing: 1px;">Generated On</span>
              <span style="color: white; font-size: 16px; font-weight: 500;">${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' })}</span>
            </div>
          </div>
        </div>

        <div style="position: absolute; bottom: 30px; left: 60px; color: rgba(255,255,255,0.4); font-size: 11px; text-transform: uppercase; letter-spacing: 1px;">
          &copy; ${new Date().getFullYear()} Cognizant. All rights reserved. &nbsp;|&nbsp; <span style="color: #26C6DA;">Caterpillar: Confidential Green</span>
        </div>
      </div>
    `;
  }

  generateThankYouSlide(): string {
    return `
      <div class="slide" style="width: 1000px; height: 562.5px; position: relative; background-color: white; overflow: hidden; page-break-after: always; flex-shrink: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
        <div style="position: absolute; top: 40px; right: 50px;">
          <h2 style="color: #000048; font-size: 28px; font-weight: 800; margin: 0; letter-spacing: -0.5px;">Cognizant</h2>
        </div>
        
        <div style="position: absolute; top: 45%; left: 80px; transform: translateY(-50%);">
          <h1 style="color: #000048; font-size: 56px; font-weight: 600; margin: 0; line-height: 1.2;">Thank you</h1>
          <div style="width: 120px; height: 4px; background-color: #26C6DA; margin-top: 15px;"></div>
        </div>
        
        <div style="position: absolute; bottom: -150px; right: -100px; width: 700px; height: 400px; background: linear-gradient(135deg, #000048 0%, #0033A0 40%, #2979ff 70%, #26C6DA 100%); transform: rotate(-15deg); opacity: 0.9; border-top-left-radius: 200px; z-index: 1;"></div>
        
        <div style="position: absolute; bottom: 20px; left: 50px; z-index: 3;">
          <p style="color: #666; font-size: 11px; margin: 0;">Caterpillar: Confidential Green</p>
        </div>
      </div>
    `;
  }
}
