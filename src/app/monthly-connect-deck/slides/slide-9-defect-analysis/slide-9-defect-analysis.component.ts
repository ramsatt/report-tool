import { Component, Input, OnInit, OnChanges, SimpleChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MonthlyData, DefectItem } from '../../monthly-connect.models';

@Component({
  selector: 'app-slide-9-defect-analysis',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-9-defect-analysis.component.html',
  styleUrls: ['./slide-9-defect-analysis.component.css']
})
export class Slide9DefectAnalysisComponent implements OnInit, OnChanges {
  @Input() data: MonthlyData | null = null;

  totalFound = 0;
  resolvedCount = 0;
  resolutionRate = 0;
  
  p2Count = 0;
  p3Count = 0;
  p4Count = 0;

  codingErrorCount = 0;
  dataConfigErrorCount = 0;

  defects: DefectItem[] = [];
  sprintRange = '';

  ngOnInit() {
    this.updateData();
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['data']) {
      this.updateData();
    }
  }

  private updateData() {
    if (this.data && this.data.defectAnalysis) {
      this.defects = this.data.defectAnalysis;
      this.calculateMetrics();
      this.calculateSprintRange();
    }
  }

  calculateSprintRange() {
    const sprints = this.defects
      .map(d => {
        const match = d.iteration?.match(/Sprint (\d+)/);
        return match ? parseInt(match[1]) : null;
      })
      .filter(s => s !== null) as number[];

    if (sprints.length > 0) {
      const min = Math.min(...sprints);
      const max = Math.max(...sprints);
      // This is a bit naive because it doesn't account for year rollover (e.g. 23 to 02)
      // But based on the sample "Sprint 23 — Sprint 02", let's just use what we found
      // or if we see both high and low numbers, we assume a rollover.
      const hasHigh = sprints.some(s => s > 20);
      const hasLow = sprints.some(s => s < 5);
      
      if (hasHigh && hasLow) {
        // Assume rollover: min is the high one, max is the low one
        const highSprints = sprints.filter(s => s > 20);
        const lowSprints = sprints.filter(s => s < 5);
        this.sprintRange = `Sprint ${Math.min(...highSprints)} — Sprint ${Math.max(...lowSprints).toString().padStart(2, '0')}`;
      } else {
        this.sprintRange = `Sprint ${min.toString().padStart(2, '0')} — Sprint ${max.toString().padStart(2, '0')}`;
      }
    } else {
      this.sprintRange = 'Sprint 23 — Sprint 02'; // Default
    }
  }

  calculateMetrics() {
    this.totalFound = this.defects.length;
    this.resolvedCount = this.defects.filter(d => 
      ['Resolved', 'Closed', 'Verified'].includes(d.state)
    ).length;
    
    this.resolutionRate = this.totalFound > 0 ? Math.round((this.resolvedCount / this.totalFound) * 100) : 0;

    this.p2Count = this.defects.filter(d => d.priority === 2).length;
    this.p3Count = this.defects.filter(d => d.priority === 3).length;
    this.p4Count = this.defects.filter(d => d.priority === 4).length;

    this.codingErrorCount = this.defects.filter(d => 
      d.rootCause?.toLowerCase().includes('coding') || d.rootCause?.toLowerCase().includes('logic')
    ).length;
    
    this.dataConfigErrorCount = this.defects.filter(d => 
      d.rootCause?.toLowerCase().includes('data') || d.rootCause?.toLowerCase().includes('config')
    ).length;
  }

  getPriColor(pri: number): string {
    switch (pri) {
      case 2: return 'bg-red-500';
      case 3: return 'bg-orange-500';
      case 4: return 'bg-blue-500';
      default: return 'bg-slate-400';
    }
  }

  getStateClasses(state: string): string {
    const s = state?.toUpperCase();
    if (s === 'RESOLVED') return 'bg-sky-100 text-sky-600 border-sky-200';
    if (s === 'CLOSED' || s === 'VERIFIED') return 'bg-emerald-100 text-emerald-600 border-emerald-200';
    return 'bg-amber-100 text-amber-600 border-amber-200';
  }

  getInitials(name: string): string {
    if (!name) return '??';
    const parts = name.split('<')[0].trim().split(' ');
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return parts[0][0].toUpperCase();
  }

  getDisplayName(name: string): string {
    if (!name) return '';
    return name.split('<')[0].trim();
  }

  getTags(tags: string): string[] {
    if (!tags) return [];
    return tags.split(';').map(t => t.trim()).filter(t => t.length > 0);
  }
}
