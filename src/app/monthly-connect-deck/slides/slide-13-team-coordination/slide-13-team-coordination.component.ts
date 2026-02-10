import { Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MonthlyData } from '../../monthly-connect.models';

@Component({
  selector: 'app-slide-13-team-coordination',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-13-team-coordination.component.html',
  styleUrls: ['./slide-13-team-coordination.component.css']
})
export class Slide13TeamCoordinationComponent {
  @Input() data: MonthlyData | null = null;

  get uniqueTeamMembersCount(): number {
    return 3; 
  }

  get totalFte(): number {
    return this.uniqueTeamMembersCount;
  }

  get riskDays(): number {
    return 1.0;
  }

  get riskPercentage(): number {
    if (this.totalFte === 0) return 0;
    return this.riskDays / this.totalFte;
  }

  get upcomingLeaveCount(): number {
    return this.data?.leavePlan?.length || 0;
  }

  get openActionItemsCount(): number {
    return this.data?.teamActions?.length || 0;
  }

  get currentDate(): string {
    const now = new Date();
    return now.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }
}
