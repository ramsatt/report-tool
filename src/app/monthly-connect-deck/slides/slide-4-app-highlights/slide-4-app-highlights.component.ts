import { Component, Input, OnInit, OnChanges, SecurityContext } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { GeneralInfo } from '../../monthly-connect.models';

@Component({
  selector: 'app-slide-4-app-highlights',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-4-app-highlights.component.html',
  styleUrls: ['./slide-4-app-highlights.component.css']
})
export class Slide4AppHighlightsComponent implements OnInit, OnChanges {
  @Input() generalInfo: GeneralInfo | null = null;
  
  // Default icons from sample - reused from Slide 2 for consistency
  icons = [
    `<svg width="24" height="24" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M5 13l4 4L19 7" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    `<svg width="24" height="24" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    `<svg width="24" height="24" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M13 10V3L4 14h7v7l9-11h-7z" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    // Adding a fourth icon for variety if needed
    `<svg width="24" height="24" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`
  ];

  highlights: Array<{ title: string, desc: string, icon: SafeHtml }> = [];

  constructor(private sanitizer: DomSanitizer) {}

  ngOnInit() {
    this.processHighlights();
  }

  ngOnChanges() {
    this.processHighlights();
  }

  private processHighlights() {
    if (!this.generalInfo?.appHighlights) {
      this.highlights = [];
      return;
    }
    
    this.highlights = (this.generalInfo.appHighlights || []).map((hl, index) => {
      let title = '';
      let desc = hl;

      // Smart title extraction
      const separator = hl.indexOf(':');
      const dash = hl.indexOf(' - ');
      
      if (separator > 0 && separator < 50) {
        title = hl.substring(0, separator).trim();
        desc = hl.substring(separator + 1).trim();
      } else if (dash > 0 && dash < 50) {
        title = hl.substring(0, dash).trim();
        desc = hl.substring(dash + 3).trim();
      } else {
        // Advanced Keyword matching for Title Generation - reused from Slide 2
        const lower = hl.toLowerCase();
        if (lower.includes('deploy') || lower.includes('release') || lower.includes('post-deployment') || lower.includes('production') || lower.includes('go-live')) title = 'Deployment Success';
        else if (lower.includes('security') || lower.includes('vulnerability') || lower.includes('compliance') || lower.includes('code ql') || lower.includes('codeql') || lower.includes('scanning')) title = 'Security Update';
        else if (lower.includes('performance') || lower.includes('optimization') || lower.includes('latency') || lower.includes('speed')) title = 'Performance Win';
        else if (lower.includes('bug') || lower.includes('fix') || lower.includes('defect') || lower.includes('remediation')) title = 'Defect Resolution';
        else if (lower.includes('feature') || lower.includes('enhancement') || lower.includes('ux') || lower.includes('user experience')) title = 'UX Enhancement';
        else if (lower.includes('test') || lower.includes('validation') || lower.includes('qa') || lower.includes('uat')) title = 'Testing Milestone';
        else if (lower.includes('migrat') || lower.includes('transition')) title = 'Migration Update';
        else if (lower.includes('api') || lower.includes('backend') || lower.includes('service')) title = 'Backend Update';
        else if (lower.includes('ui') || lower.includes('frontend') || lower.includes('interface')) title = 'UI Update';
        else {
            // If no specific keyword matches, use a generic but professional title
            title = 'Application Highlight';
        }
      }

      return {
        title: title,
        desc: desc,
        icon: this.sanitizer.bypassSecurityTrustHtml(this.icons[index % this.icons.length])
      };
    });
  }

  getIcon(index: number): string {
    return this.icons[index % this.icons.length];
  }
}
