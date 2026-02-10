import { Component, Input, OnInit, OnChanges, SimpleChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MonthlyData, DefectItem } from '../../monthly-connect.models';

@Component({
  selector: 'app-slide-10-defect-backlogs',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-10-defect-backlogs.component.html',
  styleUrls: ['./slide-10-defect-backlogs.component.css']
})
export class Slide10DefectBacklogsComponent implements OnInit, OnChanges {
  @Input() data: MonthlyData | null = null;

  totalBacklog = 0;
  newCount = 0;
  inProgressCount = 0;
  
  p2Count = 0;
  p3Count = 0;
  p4Count = 0;

  backlogItems: DefectItem[] = [];
  reportMonth = '';
  reportYear: number | string = '';

  ngOnInit() {
    this.updateData();
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['data']) {
      this.updateData();
    }
  }

  private updateData() {
    if (this.data) {
      this.backlogItems = this.data.backlogItems || [];
      this.reportMonth = this.data.generalInfo?.month || 'JAN';
      this.reportYear = this.data.generalInfo?.year || 2026;
      this.calculateMetrics();
    }
  }

  calculateMetrics() {
    this.totalBacklog = this.backlogItems.length;
    this.newCount = this.backlogItems.filter(d => 
      ['New', 'To Do'].includes(d.state)
    ).length;
    this.inProgressCount = this.backlogItems.filter(d => 
      ['In Progress', 'Active', 'Development'].includes(d.state)
    ).length;

    this.p2Count = this.backlogItems.filter(d => d.priority === 2).length;
    this.p3Count = this.backlogItems.filter(d => d.priority === 3).length;
    this.p4Count = this.backlogItems.filter(d => d.priority === 4).length;
  }

  getPriColor(pri: any): string {
    const p = parseInt(pri);
    switch (p) {
      case 2: return '#ef4444'; // Red
      case 3: return '#f59e0b'; // Orange
      case 4: return '#3b82f6'; // Blue
      default: return '#64748b';
    }
  }

  getStateBadgeClass(state: string): string {
    const s = state?.toUpperCase();
    if (s === 'NEW' || s === 'TO DO') return 'badge-new';
    if (s === 'IN PROGRESS' || s === 'ACTIVE') return 'badge-progress';
    return 'badge-other';
  }

  getInitials(name: string): string {
    if (!name) return 'UN';
    const parts = name.split('<')[0].trim().split(' ');
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return parts[0][0].toUpperCase();
  }

  getDisplayName(name: string): string {
    if (!name) return 'Unassigned';
    return name.split('<')[0].trim();
  }

  getTags(tags: string): string[] {
    if (!tags) return [];
    return tags.split(';').map(t => t.trim()).filter(t => t.length > 0);
  }

  getProgressWidth(): string {
    if (this.totalBacklog === 0) return '0%';
    return ((this.inProgressCount / this.totalBacklog) * 100) + '%';
  }
}
