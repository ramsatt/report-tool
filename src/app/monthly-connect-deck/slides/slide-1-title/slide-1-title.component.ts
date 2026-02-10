import { Component, Input } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GeneralInfo } from '../../monthly-connect.models';

@Component({
  selector: 'app-slide-1-title',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './slide-1-title.component.html',
  styleUrls: ['./slide-1-title.component.css']
})
export class Slide1TitleComponent {
  @Input() generalInfo: GeneralInfo | null = null;
  
  get generatedDate(): string {
    return new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
  }
}
