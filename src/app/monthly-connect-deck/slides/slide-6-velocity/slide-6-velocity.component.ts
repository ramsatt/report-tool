import { Component, Input, OnInit, OnChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgxEchartsModule } from 'ngx-echarts';
import { EChartsOption } from 'echarts';
import { GeneralInfo, VelocityStrategy } from '../../monthly-connect.models';

interface VelocityDisplay extends VelocityStrategy {
    period: string;
    statusClass: string;
    displayStatus: string;
    monthRowSpan: number;
}

interface VelocitySummary {
    avgVelocity: number;
    totalSprints: number;
    description: string;
}



@Component({
    selector: 'app-slide-6-velocity',
    standalone: true,
    imports: [CommonModule, NgxEchartsModule],
    templateUrl: './slide-6-velocity.component.html',
    styleUrls: ['./slide-6-velocity.component.css']
})
export class Slide6VelocityComponent implements OnInit, OnChanges {
    @Input() generalInfo: GeneralInfo | null = null;
    @Input() velocityData: VelocityStrategy[] = [];
    
    velocityMetrics: VelocityDisplay[] = [];
    summary: VelocitySummary = {
        avgVelocity: 0,
        totalSprints: 0,
        description: 'Strategic efficiency calculated across last 6 monts.'
    };
    

    
    chartOption: EChartsOption = {};

    ngOnInit() {
        this.processData();
        this.updateChartOption();
    }

    ngOnChanges() {
        this.processData();
        this.updateChartOption();
    }

    private processData() {
        if (!this.velocityData || this.velocityData.length === 0) return;

        // Sort by date (oldest first)
        const sorted = [...this.velocityData].sort((a, b) => {
            const dateA = this.parseExcelDate(a.month);
            const dateB = this.parseExcelDate(b.month);
            return dateA.getTime() - dateB.getTime();
        });

        // Limit to last 18 entries (approx 6-9 months) to ensure it fits the slide without scroll
        const limited = sorted.slice(-18);

        let totalVelocity = 0;
        
        this.velocityMetrics = limited.map(v => {
            const pct = v.deliveryPct || 0;
            totalVelocity += pct;
            
            let statusClass = 'status-verified';
            let displayStatus = 'Verified';
            
            if (pct >= 100) {
                statusClass = 'status-verified';
                displayStatus = 'Verified';
            } else if (pct >= 90) {
                statusClass = 'status-review';
                displayStatus = 'Review';
            } else {
                statusClass = 'status-partial';
                displayStatus = 'Partial';
            }

            const period = this.formatPeriod(v.month);

            return {
                ...v,
                period,
                statusClass,
                displayStatus,
                plannedDeployment: this.formatDeploymentDate(v.plannedDeployment),
                actualDeployment: this.formatDeploymentDate(v.actualDeployment),
                monthRowSpan: 0 // Will be set below
            };
        });

        // Calculate RowSpans for months
        for (let i = 0; i < this.velocityMetrics.length; i++) {
            if (i > 0 && this.velocityMetrics[i].period === this.velocityMetrics[i - 1].period) {
                this.velocityMetrics[i].monthRowSpan = -1; // Flag to skip rendering
            } else {
                let span = 1;
                for (let j = i + 1; j < this.velocityMetrics.length; j++) {
                    if (this.velocityMetrics[j].period === this.velocityMetrics[i].period) {
                        span++;
                    } else {
                        break;
                    }
                }
                this.velocityMetrics[i].monthRowSpan = span;
            }
        }

        this.summary.totalSprints = this.velocityMetrics.length;
        this.summary.avgVelocity = this.summary.totalSprints > 0 
            ? Math.round((totalVelocity / this.summary.totalSprints) * 10) / 10 
            : 0;
    }

    private formatDeploymentDate(dateStr: string | undefined): string {
        if (!dateStr || dateStr === '') return '-';
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr;
        
        const day = date.getDate().toString().padStart(2, '0');
        const month = date.toLocaleDateString('en-US', { month: 'short' }).toUpperCase();
        return `${day} ${month}`;
    }

    private parseExcelDate(excelDate: string | number): Date {
        if (!excelDate) return new Date();
        if (typeof excelDate === 'number') {
            // Excel date serial number
            const epoch = new Date(1899, 11, 30);
            return new Date(epoch.getTime() + excelDate * 86400000);
        }
        
        // Handle strings like "Aug2025" or "Sep 2025"
        const str = String(excelDate).trim();
        const normalized = str.replace(/^([a-zA-Z]{3})(\d{4})$/, '$1 $2');
        const date = new Date(normalized);
        
        return isNaN(date.getTime()) ? new Date() : date;
    }

    private formatPeriod(excelDate: string | number): string {
        const date = this.parseExcelDate(excelDate);
        const month = date.toLocaleDateString('en-US', { month: 'short' }).toUpperCase();
        const year = date.getFullYear().toString().slice(-2);
        return `${month} '${year}`;
    }

    private updateChartOption() {
        if (!this.velocityMetrics.length) return;

        const labels = this.velocityMetrics.map(v => v.sprint.replace('Sprint ', 'S').replace('​', ''));
        const commitData = this.velocityMetrics.map(v => v.committed || 0);
        const deliverData = this.velocityMetrics.map(v => v.delivered || 0);

        this.chartOption = {
            grid: {
                top: 10,
                right: 10,
                bottom: 25,
                left: 35
            },
            tooltip: {
                trigger: 'axis',
                axisPointer: {
                    type: 'shadow'
                },
                backgroundColor: 'rgba(15, 23, 42, 0.9)',
                borderColor: 'transparent',
                textStyle: {
                    color: '#fff',
                    fontSize: 11
                },
                padding: 8
            },
            xAxis: {
                type: 'category',
                data: labels,
                axisLine: {
                    show: false
                },
                axisTick: {
                    show: false
                },
                axisLabel: {
                    color: '#475569', // Deep slate for X-axis labels (Sprints)
                    fontSize: 9,
                    fontWeight: 800,
                    margin: 8
                }
            },
            yAxis: {
                type: 'value',
                splitLine: {
                    lineStyle: {
                        color: '#e2e8f0',
                        type: 'dashed'
                    }
                },
                axisLine: {
                    show: false
                },
                axisTick: {
                    show: false
                },
                axisLabel: {
                    color: '#64748b', // Solid color for Y-axis labels (Values)
                    fontSize: 9,
                    fontWeight: 700,
                    margin: 8
                }
            },
            series: [
                {
                    name: 'Commit',
                    type: 'bar',
                    data: commitData,
                    itemStyle: {
                        color: '#94a3b8',
                        borderRadius: [3, 3, 0, 0]
                    },
                    barMaxWidth: 40
                },
                {
                    name: 'Deliver',
                    type: 'bar',
                    data: deliverData,
                    itemStyle: {
                        color: '#0891b2',
                        borderRadius: [3, 3, 0, 0]
                    },
                    barMaxWidth: 40
                }
            ]
        };
    }
}
