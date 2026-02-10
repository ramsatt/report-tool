import { Component, Input, OnInit, OnChanges, SimpleChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgxEchartsModule } from 'ngx-echarts';
import { MonthlyData, DefectItem } from '../../monthly-connect.models';
import { SPRINT_DATA_2025, SPRINT_DATA_2026 } from '../../../monthly-report/const/sprint-constants';

@Component({
  selector: 'app-slide-11-defect-metrics',
  standalone: true,
  imports: [CommonModule, NgxEchartsModule],
  templateUrl: './slide-11-defect-metrics.component.html',
  styleUrls: ['./slide-11-defect-metrics.component.css']
})
export class Slide11DefectMetricsComponent implements OnInit, OnChanges {
  @Input() data: MonthlyData | null = null;

  totalVolume = 0;
  highestResolutionSprint = '';
  avgCarryOver = 0;
  velocityRatio = 1.0;
  stabilityIndexValue = 0;
  primaryRootCause = '';
  snapshotDate = '';

  sprintMetrics: any[] = [];
  rootCauseStats: any[] = [];
  
  // ECharts Options
  throughputOption: any;
  rcOption: any;

  ngOnInit() {
    this.updateData();
    const now = new Date();
    this.snapshotDate = now.toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' });
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['data']) {
      this.updateData();
    }
  }

  private updateData() {
    if (this.data && this.data.defectMetrics) {
      this.processData(this.data.defectMetrics);
    }
  }

  private processData(defects: DefectItem[]) {
    this.totalVolume = defects.length;
    if (defects.length === 0) return;

    const allSprints = [...SPRINT_DATA_2025, ...SPRINT_DATA_2026];
    const parseDate = (dStr: string) => dStr ? new Date(dStr).getTime() : null;

    const todayTime = new Date().getTime();

    // Calculate metrics for ALL potential sprints
    const allMetrics = allSprints.map(s => {
      const start = new Date(s.startDate).getTime();
      const end = new Date(s.endDate);
      end.setHours(23, 59, 59, 999);
      const endTime = end.getTime();
      const yearSuffix = s.startDate.split('-')[0].slice(-2);

      // Created in this sprint (based on created date)
      const created = defects.filter(d => {
        const time = parseDate(d.createdDate);
        return time && time >= start && time <= endTime;
      }).length;

      // Resolved in this sprint (based on resolved date)
      const resolved = defects.filter(d => {
        const time = parseDate(d.resolvedDate);
        return time && time >= start && time <= endTime;
      }).length;

      // Carry-over 
      const carryOver = defects.filter(d => {
        const cTime = parseDate(d.createdDate);
        const rTime = parseDate(d.resolvedDate);
        return cTime && cTime < start && (!rTime || rTime >= start);
      }).length;

      const activeAtEnd = defects.filter(d => {
        const cTime = parseDate(d.createdDate);
        const rTime = parseDate(d.resolvedDate);
        return cTime && cTime <= endTime && (!rTime || rTime > endTime);
      }).length;

      const throughput = created > 0 ? (resolved / created) * 100 : (resolved > 0 ? Infinity : 0);

      return {
        name: `${s.name.replace('Sprint ', 'S')} '${yearSuffix}`,
        sortKey: start,
        created,
        resolved,
        carryOver,
        activeAtEnd,
        throughputRate: throughput === Infinity ? '∞' : Math.round(throughput) + '%'
      };
    });

    const sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
    const sixMonthsAgoTime = sixMonthsAgo.getTime();

    // 1. Filter out future sprints
    // 2. Filter by 6-month window
    // 3. Filter out sprints with no throughput activity
    this.sprintMetrics = allMetrics
        .filter(m => m.sortKey <= todayTime && m.sortKey >= sixMonthsAgoTime)
        .filter(m => m.created > 0 || m.resolved > 0)
        .sort((a, b) => a.sortKey - b.sortKey);

    // Executive KPIs
    const maxResolvedMetric = [...this.sprintMetrics].sort((a,b) => b.resolved - a.resolved)[0];
    this.highestResolutionSprint = maxResolvedMetric?.resolved > 0 ? maxResolvedMetric.name : 'N/A';

    const sumCarry = this.sprintMetrics.reduce((acc, m) => acc + m.carryOver, 0);
    this.avgCarryOver = Math.round(sumCarry / this.sprintMetrics.length);

    const totalCreated = defects.length;
    const totalResolvedCount = defects.filter(d => d.resolvedDate).length;
    this.velocityRatio = totalCreated > 0 ? parseFloat((totalResolvedCount / totalCreated).toFixed(1)) : 0;
    this.stabilityIndexValue = totalCreated > 0 ? Math.round((totalResolvedCount / totalCreated) * 100) : 0;

    // Root Cause Stats
    const rcCounts: { [key: string]: number } = {};
    defects.forEach(d => {
      const rc = this.simplifyRootCause(d.rootCause || 'Other');
      rcCounts[rc] = (rcCounts[rc] || 0) + 1;
    });

    const colors = ['#0284c7', '#d97706', '#059669', '#64748b', '#8b5cf6', '#ec4899'];
    this.rootCauseStats = Object.keys(rcCounts).map((rc, i) => ({
      name: rc,
      value: rcCounts[rc],
      itemStyle: { color: colors[i % colors.length] }
    })).sort((a, b) => b.value - a.value);

    this.primaryRootCause = this.rootCauseStats[0]?.name || 'N/A';

    this.initChartOptions();
  }

  private initChartOptions() {
    // Throughput Chart (Bar)
    this.throughputOption = {
      grid: {
        top: '15%',
        bottom: '15%',
        left: '5%',
        right: '5%',
        containLabel: true
      },
      legend: {
        show: false
      },
      xAxis: {
        type: 'category',
        data: this.sprintMetrics.map(m => m.name),
        axisLabel: { color: '#64748b', fontSize: 10, fontWeight: 'bold' },
        axisLine: { show: false },
        axisTick: { show: false }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: '#64748b', fontSize: 9 },
        splitLine: { lineStyle: { color: 'rgba(0,0,0,0.05)' } }
      },
      series: [
        {
          name: 'Created',
          type: 'bar',
          data: this.sprintMetrics.map(m => m.created),
          itemStyle: { color: '#0284c7', borderRadius: [4, 4, 0, 0] },
          barWidth: '25%'
        },
        {
          name: 'Resolved',
          type: 'bar',
          data: this.sprintMetrics.map(m => m.resolved),
          itemStyle: { color: '#059669', borderRadius: [4, 4, 0, 0] },
          barWidth: '25%'
        }
      ]
    };

    // Root Cause Chart (Doughnut)
    this.rcOption = {
      tooltip: { trigger: 'item' },
      legend: {
        orient: 'horizontal',
        bottom: '0%',
        left: 'center',
        itemWidth: 8,
        itemHeight: 8,
        textStyle: { fontSize: 9, color: '#64748b', fontWeight: 'bold' }
      },
      series: [
        {
          name: 'Root Cause',
          type: 'pie',
          radius: ['45%', '70%'],
          center: ['50%', '42%'],
          avoidLabelOverlap: false,
          label: { show: false },
          emphasis: {
            label: { show: true, fontSize: 10, fontWeight: 'bold' }
          },
          labelLine: { show: false },
          data: this.rootCauseStats
        }
      ]
    };
  }

  public simplifyRootCause(rc: string): string {
    rc = rc.toLowerCase();
    if (rc.includes('coding') || rc.includes('logic')) return 'Coding';
    if (rc.includes('data')) return 'Data';
    if (rc.includes('config')) return 'Config';
    if (rc.includes('user')) return 'User';
    if (rc.includes('int') || rc.includes('environment')) return 'Environment';
    return 'Other';
  }
}
