import { Component, Input, OnInit, OnChanges } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgxEchartsModule } from 'ngx-echarts';
import { EChartsOption } from 'echarts';
import { GeneralInfo, VelocityStrategy } from '../../monthly-connect.models';

interface VelocityDisplay extends VelocityStrategy {
    period: string;
    statusClass: string;
    displayStatus: string;
}

interface VelocitySummary {
    avgVelocity: number;
    totalSprints: number;
    description: string;
}

interface ExecutionNote {
    text: string;
    bulletColor: string;
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
        description: 'Strategic efficiency calculated across last 14 release cycles.'
    };
    
    executionNotes: ExecutionNote[] = [
        { text: 'Q4-25 transition stabilized during current stream.', bulletColor: '#818cf8' },
        { text: 'Production hardening achieved 100% completion in recent milestones.', bulletColor: '#22d3ee' },
        { text: 'Sprint 26 duplicate work issues monitored for root cause.', bulletColor: '#cbd5e1' }
    ];
    
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

        let totalVelocity = 0;
        
        this.velocityMetrics = sorted.map(v => {
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
                displayStatus
            };
        });

        this.summary.totalSprints = this.velocityMetrics.length;
        this.summary.avgVelocity = this.summary.totalSprints > 0 
            ? Math.round((totalVelocity / this.summary.totalSprints) * 10) / 10 
            : 0;
    }

    private parseExcelDate(excelDate: string | number): Date {
        if (typeof excelDate === 'number') {
            // Excel date serial number
            const epoch = new Date(1899, 11, 30);
            return new Date(epoch.getTime() + excelDate * 86400000);
        }
        return new Date(excelDate);
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
                    color: '#94a3b8',
                    fontSize: 9,
                    fontWeight: 600,
                    margin: 8
                }
            },
            yAxis: {
                type: 'value',
                splitLine: {
                    lineStyle: {
                        color: '#f1f5f9',
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
                    color: '#cbd5e1',
                    fontSize: 9,
                    margin: 8
                }
            },
            series: [
                {
                    name: 'Commit',
                    type: 'bar',
                    data: commitData,
                    itemStyle: {
                        color: '#cbd5e1',
                        borderRadius: [3, 3, 0, 0]
                    },
                    barMaxWidth: 40
                },
                {
                    name: 'Deliver',
                    type: 'bar',
                    data: deliverData,
                    itemStyle: {
                        color: '#22d3ee',
                        borderRadius: [3, 3, 0, 0]
                    },
                    barMaxWidth: 40
                }
            ]
        };
    }
}
