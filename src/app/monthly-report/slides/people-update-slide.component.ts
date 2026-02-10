import { Component, Input } from '@angular/core';
import { LeavePlanData, TeamActionData } from '../models/report-data.model';
import { SlideGeneratorService } from '../services/slide-generator.service';

@Component({
  selector: 'app-people-update-slide',
  standalone: true,
  template: ''
})
export class PeopleUpdateSlideComponent {
  @Input() leavePlanData: LeavePlanData[] = [];
  @Input() teamActionsData: TeamActionData[] = [];
  @Input() currentMonth: string = '';
  @Input() currentYear: string = '';

  constructor(private slideGenerator: SlideGeneratorService) {}

  generateSlide(): string {
    const header = this.slideGenerator.generateHeader(
      'People Update',
      'Holiday Plan & Action Items',
      this.currentMonth,
      this.currentYear
    );

    const content = `
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
    `;

    const footer = this.slideGenerator.generateFooter();
    return this.slideGenerator.wrapSlideContent(header + content + footer);
  }
}
