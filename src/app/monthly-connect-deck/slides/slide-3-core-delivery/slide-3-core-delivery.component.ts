import { Component, Input, OnInit, OnChanges, SecurityContext } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { DeliveryMetric, GeneralInfo } from '../../monthly-connect.models';
import { SPRINT_DATA_2025, SPRINT_DATA_2026, SprintInfo } from '../../../monthly-report/const/sprint-constants';

interface DeliverySummary {
    totalCapacity: number;
    totalCommitted: number;
    totalDelivered: number;
    totalSpilled: number;
    progress: number;
    status: string;
}

interface DisplayMetric extends DeliveryMetric {
    featuresList: string[];
    capacity: number;
    spilled: number;
    statusClass: string;
    displayStatus: string;
    iconHtml: SafeHtml;
    sprintDuration: string;
}

@Component({
    selector: 'app-slide-3-core-delivery',
    standalone: true,
    imports: [CommonModule],
    templateUrl: './slide-3-core-delivery.component.html',
    styleUrls: ['./slide-3-core-delivery.component.css']
})
export class Slide3CoreDeliveryComponent implements OnInit, OnChanges {
    @Input() generalInfo: GeneralInfo | null = null;
    @Input() deliveryMetrics: DeliveryMetric[] = [];

    coreMetrics: DisplayMetric[] = [];
    summary: DeliverySummary = {
        totalCapacity: 0,
        totalCommitted: 0,
        totalDelivered: 0,
        totalSpilled: 0,
        progress: 0,
        status: 'On Track'
    };

    private iconMap: { [key: string]: string } = {
        'cpu': '<svg width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><path d="M9 9h6v6H9z"/><path d="M9 1V4"/><path d="M15 1V4"/><path d="M9 20V23"/><path d="M15 20V23"/><path d="M20 9H23"/><path d="M20 14H23"/><path d="M1 9H4"/><path d="M1 14H4"/></svg>',
        'shield-check': '<svg width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M9 12l2 2 4-4"/></svg>',
        'database': '<svg width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>',
        'server': '<svg width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="8" rx="2" ry="2"/><rect x="2" y="14" width="20" height="8" rx="2" ry="2"/><line x1="6" y1="6" x2="6.01" y2="6"/><line x1="6" y1="18" x2="6.01" y2="18"/></svg>'
    };

    constructor(private sanitizer: DomSanitizer) { }

    ngOnInit() {
        this.processData();
    }

    ngOnChanges() {
        this.processData();
    }

    private processData() {
        if (!this.deliveryMetrics) return;

        // Filter for Core Platform
        const rawMetrics = this.deliveryMetrics.filter(m => m.stream === 'Core Platform');
        
        this.summary = {
            totalCapacity: 0,
            totalCommitted: 0,
            totalDelivered: 0,
            totalSpilled: 0,
            progress: 0,
            status: 'On Track'
        };

        this.coreMetrics = rawMetrics.map((m, index) => {
            const spelled = (m.committed || 0) - (m.delivered || 0);
            const spilled = spelled > 0 ? spelled : 0;
            const capacity = 14; // Fixed capacity per sprint
            
            this.summary.totalCapacity += capacity;
            this.summary.totalCommitted += m.committed || 0;
            this.summary.totalDelivered += m.delivered || 0;
            this.summary.totalSpilled += spilled;

            // Parse features
            let featuresList: string[] = [];
            if (m.features) {
                featuresList = m.features.split(/[,\n]/).map(f => f.trim()).filter(f => f.length > 0);
            }

            const pct = capacity > 0 ? ((m.delivered || 0) / capacity) * 100 : 0;
            let statusClass = 'badge-warning';
            let displayStatus = m.deploymentStatus;

            if (pct >= 100) {
                statusClass = 'badge-success';
                if (!displayStatus) displayStatus = 'Deployed';
            } else if (pct === 0) {
                 statusClass = 'badge-secondary'; // New gray/neutral class
                 if (!displayStatus) displayStatus = 'Planned';
            } else if (pct < 100) {
                statusClass = 'badge-warning';
                if (!displayStatus) displayStatus = 'Partial';
            }

            const iconKey = 'cpu';
            const iconHtml = this.sanitizer.bypassSecurityTrustHtml(this.iconMap[iconKey]);

            let sprintDuration = m.month || ''; 
            const yearStr = this.generalInfo?.year || '2026';
            const year = parseInt(yearStr.toString(), 10);
            
            // Default to 2026 if not 2025
            const sprintData = year === 2025 ? SPRINT_DATA_2025 : SPRINT_DATA_2026;
            
            // Find sprint info - normalize 'Sprint 01' vs 'Sprint 1'
            const sprintNum = m.sprint.replace(/[^0-9]/g, '');
            const targetSprintName = `Sprint ${sprintNum.padStart(2, '0')}`;
            
            const sprintInfo = sprintData.find(s => s.name.toLowerCase() === targetSprintName.toLowerCase());

            if (sprintInfo) {
                 // Format: Jan 7 - Jan 20
                 const parseDate = (dStr: string) => {
                    const parts = dStr.split('-');
                    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
                 };
                 const start = parseDate(sprintInfo.startDate);
                 const end = parseDate(sprintInfo.endDate);

                 const fmt = (d: Date) => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                 sprintDuration = `${fmt(start)} - ${fmt(end)}`;
            }

            return {
                ...m,
                featuresList,
                capacity: capacity, 
                spilled,
                statusClass,
                displayStatus: displayStatus === 'Deployed' ? 'Success' : (displayStatus || 'Planned'), // Map Deployed -> Success to match screenshot? Or keep Deployed. Screenshot says SUCCESS.
                iconHtml,
                sprintDuration,
                deliveryPct: pct // Override with calculated percentage
            };
        });



        if (this.summary.totalCapacity > 0) {
            this.summary.progress = Math.round((this.summary.totalDelivered / this.summary.totalCapacity) * 100);
        } else {
            this.summary.progress = 0;
        }

        if (this.summary.progress >= 90) this.summary.status = 'On Track';
        else if (this.summary.progress >= 70) this.summary.status = 'At Risk';
        else this.summary.status = 'Delayed';
    }
}
