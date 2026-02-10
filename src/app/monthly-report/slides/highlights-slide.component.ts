import { Component, Input } from '@angular/core';
import { SlideGeneratorService } from '../services/slide-generator.service';

@Component({
  selector: 'app-highlights-slide',
  standalone: true,
  template: ''
})
export class HighlightsSlideComponent {
  @Input() highlights: string[] = [];
  @Input() type: 'core' | 'app' = 'core';
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  constructor(private slideGenerator: SlideGeneratorService) {}

  generateSlide(): string {
    const title = 'Deliverable Highlights';
    const subtitle = this.type === 'core' ? 'Core Platform' : 'App Platform';

    const header = this.slideGenerator.generateHeader(title, subtitle, this.currentMonth, this.currentYear);
    const footer = this.slideGenerator.generateFooter();

    const content = `
      <div style="padding: 40px 60px; position: relative;">
        <div style="position: absolute; top: 20px; right: 40px; opacity: 0.05;">
          <svg width="200" height="200" viewBox="0 0 200 200" fill="none" xmlns="http://www.w3.org/2000/svg">
            ${this.type === 'core' 
              ? '<path d="M100 0L200 100L100 200L0 100L100 0Z" fill="#000048"/>'
              : '<rect x="0" y="0" width="100" height="100" stroke="#000048" stroke-width="20"/><rect x="50" y="50" width="100" height="100" stroke="#000048" stroke-width="10"/>'
            }
          </svg>
        </div>

        <h3 style="color: #000048; font-size: 20px; font-weight: 700; margin-bottom: 25px; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px; display: inline-block;">
          Key Achievements
        </h3>
        
        <ul style="list-style: none; padding: 0;">
          ${this.highlights.map(h => `
            <li style="margin-bottom: 18px; display: flex; align-items: flex-start; font-size: 18px; color: #333; line-height: 1.5;">
              <span style="min-width: 8px; height: 8px; background-color: #26C6DA; border-radius: 50%; margin-top: 9px; margin-right: 20px; box-shadow: 0 0 0 3px rgba(38, 198, 218, 0.2);"></span>
              <span style="color: #000048; font-weight: 500;">${h}</span>
            </li>
          `).join('')}
        </ul>
      </div>
    `;

    return this.slideGenerator.wrapSlideContent(header + content + footer);
  }
}
