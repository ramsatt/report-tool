import { Component, Input, OnInit, OnChanges, SimpleChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MonthlyData } from '../../monthly-connect.models';
import { NgxEchartsModule } from 'ngx-echarts';

@Component({
  selector: 'app-slide-12-roadmap',
  standalone: true,
  imports: [CommonModule, NgxEchartsModule],
  templateUrl: './slide-12-roadmap.component.html',
  styleUrls: ['./slide-12-roadmap.component.css']
})
export class Slide12RoadmapComponent implements OnInit, OnChanges {
  @Input() data: MonthlyData | null = null;
  
  roadmapItems: any[] = [];
  kpiStats = {
    strategicTasks: 0,
    securityCritical: 0,
    completionRate: 0,
    performanceWins: 0
  };

  chartOption: any;

  ngOnInit() {
    this.processData();
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['data']) {
      this.processData();
      this.updateChart();
    }
  }

  private processData() {
    if (!this.data || !this.data.roadmap) return;

    this.roadmapItems = this.data.roadmap.map(item => ({
      ...item,
      // Normalize data for display
      sno: item.sno ? String(item.sno).replace('#', '') : '',
      priorityClass: this.getPriorityClass(item.priority),
      accentColor: this.getAccentColor(item.type),
      statusClass: this.getStatusClass(item.status)
    }));

    // Calculate KPIs
    this.kpiStats.strategicTasks = this.roadmapItems.length;
    this.kpiStats.securityCritical = this.roadmapItems.filter(i => i.type?.toLowerCase().includes('security')).length;
    
    const completed = this.roadmapItems.filter(i => i.status?.toLowerCase() === 'completed' || i.status?.toLowerCase() === 'done').length;
    this.kpiStats.completionRate = this.roadmapItems.length > 0 ? Math.round((completed / this.roadmapItems.length) * 100) : 0;
    
    this.kpiStats.performanceWins = this.roadmapItems.filter(i => i.type?.toLowerCase().includes('performance') && (i.status?.toLowerCase() === 'completed' || i.status?.toLowerCase() === 'done')).length;
  
    this.updateChart();
  }

  private updateChart() {
    if (!this.roadmapItems.length) return;

    const typeCounts: {[key: string]: number} = {};
    const typeMeta: {[key: string]: string} = {};

    this.roadmapItems.forEach(item => {
      const type = item.type || 'Other';
      typeCounts[type] = (typeCounts[type] || 0) + 1;
      typeMeta[type] = this.getAccentColor(type);
    });

    const chartData = Object.keys(typeCounts).map(key => ({
        name: key,
        value: typeCounts[key],
        itemStyle: { color: typeMeta[key] }
    }));

    this.chartOption = {
        tooltip: {
            trigger: 'item'
        },
        legend: {
            orient: 'horizontal',
            bottom: '0%',
            left: 'center',
            itemWidth: 8,
            itemHeight: 8,
            textStyle: { fontSize: 8, color: '#64748b', fontWeight: 'bold' },
            itemGap: 10
        },
        series: [
            {
                name: 'Backlog Distribution',
                type: 'pie',
                radius: ['35%', '55%'],
                center: ['50%', '35%'],
                avoidLabelOverlap: false,
                label: {
                    show: false,
                    position: 'center'
                },
                emphasis: {
                    label: {
                        show: true,
                        fontSize: 10,
                        fontWeight: 'bold',
                        color: '#1e293b'
                    }
                },
                labelLine: {
                    show: false
                },
                data: chartData
            }
        ]
    };
  }

  getPriorityClass(priority: string): string {
    if (!priority) return '';
    const p = priority.toLowerCase();
    if (p.includes('high') || p.includes('critical')) return 'priority-high';
    return 'priority-low';
  }

  getAccentColor(type: string): string {
    if (!type) return '#64748b';
    const t = type.toLowerCase();
    if (t.includes('security')) return '#ff6361';
    if (t.includes('performance')) return '#003f5c';
    if (t.includes('optimization')) return '#ffa600';
    return '#64748b';
  }
  
  getErrorClass(type: string): string {
      // Helper to return bg class similar to template
       if (!type) return 'bg-slate-500';
        const t = type.toLowerCase();
        if (t.includes('security')) return 'bg-[#ff6361]';
        if (t.includes('performance')) return 'bg-[#003f5c]';
        if (t.includes('optimization')) return 'bg-[#ffa600]';
        return 'bg-slate-500';
  }

  getStatusClass(status: string): string {
    if (!status) return 'status-tbd';
    const s = status.toLowerCase();
    if (s === 'done' || s === 'completed') return 'status-done';
    return 'status-tbd';
  }
}
