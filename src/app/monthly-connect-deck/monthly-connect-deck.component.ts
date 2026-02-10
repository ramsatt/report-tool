import { Component, OnInit, ElementRef, ViewChild, AfterViewInit, HostListener, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MonthlyConnectService } from './monthly-connect.service';
import { MonthlyData } from './monthly-connect.models';
import { Slide1TitleComponent } from './slides/slide-1-title/slide-1-title.component';
import { Slide2HighlightsComponent } from './slides/slide-2-highlights/slide-2-highlights.component';
import { Slide3CoreDeliveryComponent } from './slides/slide-3-core-delivery/slide-3-core-delivery.component';
import PptxGenJS from 'pptxgenjs';
import * as htmlToImage from 'html-to-image';
import { Slide4AppHighlightsComponent } from './slides/slide-4-app-highlights/slide-4-app-highlights.component';
import { Slide5AppDeliveryComponent } from './slides/slide-5-app-delivery/slide-5-app-delivery.component';
import { Slide6VelocityComponent } from './slides/slide-6-velocity/slide-6-velocity.component';
import { Slide7MigrationComponent } from './slides/slide-7-migration/slide-7-migration.component';
import { Slide8FeedbackComponent } from './slides/slide-8-feedback/slide-8-feedback.component';
import { Slide9DefectAnalysisComponent } from './slides/slide-9-defect-analysis/slide-9-defect-analysis.component';
import { Slide10DefectBacklogsComponent } from './slides/slide-10-defect-backlogs/slide-10-defect-backlogs.component';
import { Slide11DefectMetricsComponent } from './slides/slide-11-defect-metrics/slide-11-defect-metrics.component';
import { Slide12RoadmapComponent } from './slides/slide-12-roadmap/slide-12-roadmap.component';
import { Slide13TeamCoordinationComponent } from './slides/slide-13-team-coordination/slide-13-team-coordination.component';
import { Slide14ThankYouComponent } from './slides/slide-14-thank-you/slide-14-thank-you.component';



@Component({
  selector: 'app-monthly-connect-deck',
  standalone: true,
  imports: [CommonModule, Slide1TitleComponent, Slide2HighlightsComponent, Slide3CoreDeliveryComponent, Slide4AppHighlightsComponent, Slide5AppDeliveryComponent, Slide6VelocityComponent, Slide7MigrationComponent, Slide8FeedbackComponent, Slide9DefectAnalysisComponent, Slide10DefectBacklogsComponent, Slide11DefectMetricsComponent, Slide12RoadmapComponent, Slide13TeamCoordinationComponent, Slide14ThankYouComponent],



  providers: [MonthlyConnectService],
  templateUrl: './monthly-connect-deck.component.html',
  styleUrls: ['./monthly-connect-deck.component.css']
})
export class MonthlyConnectDeckComponent implements OnInit, AfterViewInit {
  slides: string[] = ['Title Slide', 'Highlights', 'Delivery Metrics', 'App Highlights', 'App Delivery', 'Velocity', 'Migration Status', 'Feedback', 'Defects', 'Backlog', 'Defect Metrics', 'Roadmap', 'Team Coordination', 'Thank You'];

  currentSlideIndex = 0;
  
  // Data Store
  monthlyData: MonthlyData | null = null;
  dataLoaded = false;
  
  // Scaling
  scale = 1;
  @ViewChild('deckMainArea') deckMainArea!: ElementRef;
  @ViewChild('exportContainer') exportContainer!: ElementRef;
  @ViewChild('consoleWindow') consoleWindow!: ElementRef;

  isExporting = false;

  // Export Progress State
  exportProgress = 0;
  exportLogs: string[] = [];
  hasExportError = false;

  constructor(
    private monthlyConnectService: MonthlyConnectService,
    private cdr: ChangeDetectorRef
  ) {}

  addExportLog(msg: string) {
    const timestamp = new Date().toLocaleTimeString();
    this.exportLogs.push(`[${timestamp}] ${msg}`);
    this.cdr.detectChanges(); // Force UI update

    // Auto-scroll to bottom of console window
    if (this.consoleWindow) {
      setTimeout(() => {
        try {
          const el = this.consoleWindow.nativeElement;
          el.scrollTop = el.scrollHeight;
        } catch (e) {
          // Ignore scroll errors
        }
      }, 0);
    }
  }

  delay(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  async exportToPptx() {
    if (!this.monthlyData || !this.exportContainer) return;

    this.isExporting = true;
    this.exportProgress = 0;
    this.exportLogs = [];
    this.hasExportError = false;
    this.addExportLog('Starting PPTX export process...');
    
    try {
      const pres = new PptxGenJS();
      pres.layout = 'LAYOUT_16x9'; 

      // Get all slide elements from the export container
      const slideElements = this.exportContainer.nativeElement.querySelectorAll('.export-slide');
      const totalSlides = slideElements.length;
      
      this.addExportLog(`Found ${totalSlides} slides to export.`);

      // Wait a moment for the container to be rendered if it was just shown/moved
      await this.delay(500);

      for (let i = 0; i < slideElements.length; i++) {
        const slideEl = slideElements[i] as HTMLElement;
        this.addExportLog(`Processing slide ${i + 1} of ${totalSlides}...`);
        
        // Small delay to let UI update and browser catch up
        // Increased to 500ms to ensure paint cycle completes before heavy lifting
        await this.delay(500);

        try {
          // Capture slide as image with a timeout race
          const capturePromise = htmlToImage.toPng(slideEl, { 
            quality: 1.0, 
            pixelRatio: 2, 
            width: 1000,
            height: 562.5,
            skipAutoScale: true, // Prevent scaling issues
            cacheBust: true, // Avoid caching issues
          });

          const timeoutPromise = new Promise<string>((_, reject) => 
            setTimeout(() => reject(new Error(`Timeout capturing slide ${i+1}`)), 15000)
          );

          const dataUrl = await Promise.race([capturePromise, timeoutPromise]);

          // Add to PPTX
          const slide = pres.addSlide();
          slide.addImage({ data: dataUrl, x: 0, y: 0, w: '100%', h: '100%' });

          // Update progress (cap at 90% until save is complete)
          this.exportProgress = Math.round(((i + 1) / totalSlides) * 90);
        
        } catch (slideErr) {
          console.error(`Error processing slide ${i+1}:`, slideErr);
          this.addExportLog(`ERROR on slide ${i+1}: ${slideErr instanceof Error ? slideErr.message : String(slideErr)}`);
          // Continue to next slide instead of failing completely?
          // For now, let's keep going but log the error visibly
        }
      }

      this.addExportLog('All slides processed. Generating PowerPoint file...');
      await this.delay(100);

      // Save the file
      const fileName = `Monthly_Connect_${this.monthlyData.generalInfo.month}_${this.monthlyData.generalInfo.year}.pptx`;
      await pres.writeFile({ fileName });

      this.exportProgress = 100;
      this.addExportLog(`Successfully saved: ${fileName}`);
      this.addExportLog('Export complete!');

    } catch (err) {
      console.error('Error generating PPTX:', err);
      this.hasExportError = true;
      this.addExportLog(`FATAL ERROR: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  closeExportModal() {
    this.isExporting = false;
  }

  ngOnInit() {
    this.monthlyConnectService.monthlyData$.subscribe(data => {
      this.monthlyData = data;
      this.dataLoaded = !!data;
      // Give time for ngIf to render the deck content
      setTimeout(() => this.updateScale(), 0);
    });
  }

  ngAfterViewInit() {
    setTimeout(() => this.updateScale(), 0);
  }

  toggleFullScreen() {
    const elem = this.deckMainArea.nativeElement;
    
    if (!document.fullscreenElement) {
      if (elem.requestFullscreen) {
        elem.requestFullscreen();
      } else if ((elem as any).webkitRequestFullscreen) { /* Safari */
        (elem as any).webkitRequestFullscreen();
      } else if ((elem as any).msRequestFullscreen) { /* IE11 */
        (elem as any).msRequestFullscreen();
      }
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      } else if ((document as any).webkitExitFullscreen) { /* Safari */
        (document as any).webkitExitFullscreen();
      } else if ((document as any).msExitFullscreen) { /* IE11 */
        (document as any).msExitFullscreen();
      }
    }
  }

  @HostListener('document:fullscreenchange')
  @HostListener('document:webkitfullscreenchange')
  @HostListener('document:mozfullscreenchange')
  @HostListener('document:MSFullscreenChange')
  onFullScreenChange() {
    setTimeout(() => this.updateScale(), 100);
  }

  @HostListener('document:keydown', ['$event'])
  handleKeyboardEvent(event: KeyboardEvent) {
    // Only navigate if not editing text (not applicable here but good practice)
    if (document.activeElement?.tagName === 'INPUT' || document.activeElement?.tagName === 'TEXTAREA') return;

    if (event.key === 'ArrowRight' || event.code === 'Space') {
      this.nextSlide();
    } else if (event.key === 'ArrowLeft') {
      this.prevSlide();
    } else if (event.key === 'Escape') {
      // Escape usually exits fullscreen by default, but we can handle custom logic if needed
    }
  }

  @HostListener('window:resize')
  onResize() {
    this.updateScale();
  }

  updateScale() {
    if (!this.deckMainArea) return;
    
    const container = this.deckMainArea.nativeElement;
    if (container.offsetWidth === 0) return; // Not visible yet

    const isFullScreen = !!document.fullscreenElement;
    const padding = isFullScreen ? 0 : 40;
    const availableWidth = container.offsetWidth - padding;
    const availableHeight = container.offsetHeight - padding;
    
    const baseWidth = 1000;
    const baseHeight = 562.5;

    const scaleX = availableWidth / baseWidth;
    const scaleY = availableHeight / baseHeight;

    this.scale = Math.min(scaleX, scaleY);
  }

  nextSlide() {
    if (this.currentSlideIndex < this.slides.length - 1) {
      this.currentSlideIndex++;
    }
  }

  prevSlide() {
    if (this.currentSlideIndex > 0) {
      this.currentSlideIndex--;
    }
  }

  // --- File Handling & Processing ---

  onFileChange(evt: any) {
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');

    this.monthlyConnectService.processExcelFile(target.files[0])
      .then(() => {
        // Reset input so same file can be selected again if needed
        (evt.target as HTMLInputElement).value = '';
      })
      .catch(err => console.error('Error processing file:', err));
  }

  resetData() {
    this.monthlyConnectService.resetData();
  }
}
