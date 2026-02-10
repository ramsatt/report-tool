import { Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GeneralInfo, Feedback } from '../../monthly-connect.models';

@Component({
  selector: 'app-slide-8-feedback',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-8-feedback.component.html',
  styleUrls: ['./slide-8-feedback.component.css']
})
export class Slide8FeedbackComponent {
  @Input() generalInfo: GeneralInfo | null = null;
  @Input() feedbackData: Feedback[] = [];

  get totalItems(): string {
    return (this.feedbackData?.length || 0).toString().padStart(2, '0');
  }

  get ongoingCount(): string {
    return (this.feedbackData?.filter(item => {
      const s = item.status?.trim().toUpperCase() || '';
      return s.includes('ONGOING');
    }).length || 0).toString().padStart(2, '0');
  }

  get closedCount(): string {
    return (this.feedbackData?.filter(item => {
      const s = item.status?.trim().toUpperCase() || '';
      return s.includes('COMPLETED') || s.includes('CLOSED');
    }).length || 0).toString().padStart(2, '0');
  }

  getAvatar(owner: string): string {
    if (!owner) return '?';
    return owner.trim().charAt(0).toUpperCase();
  }

  getStatusClass(status: string): string {
    const s = status?.trim().toUpperCase() || '';
    if (s.includes('COMPLETED') || s.includes('CLOSED')) return 'badge-completed';
    if (s.includes('ONGOING')) return 'badge-ongoing';
    return '';
  }
}
