import { Component, OnInit, AfterViewInit, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';


declare var lucide: any;
declare var XLSX: any;
declare var html2canvas: any;
declare var html2pdf: any;
declare var jspdf: any;

@Component({
  selector: 'app-sprint-planning',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './sprint-planning.component.html',
  styleUrls: ['./sprint-planning.component.css']
})
export class SprintPlanningComponent implements OnInit, AfterViewInit {
  // State
  activeTab: 'input' | 'preview' = 'input';
  previewSubTab: 'visual' | 'code' = 'visual';
  
  sprintData = {
    number: 'Sprint 26',
    startDate: 'Dec 24',
    endDate: 'Jan 06'
  };

  orgMap: { [key: string]: string } = {
    'Manikandan Kasi': 'CAT',
    'Varsha Julakanti': 'Cognizant',
    'Sathish Kumar Ramalingam': 'Cognizant',
    'Parkavi R': 'Cognizant'
  };

  workItems: any[] = [];
  
  // Stats
  totalPoints = 0;
  totalCount = 0;

  // Modal State
  modalVisible = false;
  modalTitle = '';
  modalMsg = '';
  modalCallback: (() => void) | null = null;

  // Preview content
  previewHtmlSafe: SafeHtml | null = null;
  previewCode = '';

  constructor(private sanitizer: DomSanitizer, private cd: ChangeDetectorRef) {}

  ngOnInit() {
    this.calculateStats();
  }

  ngAfterViewInit() {
    this.initLucide();
  }

  initLucide() {
    if (typeof lucide !== 'undefined') {
      lucide.createIcons();
    }
  }

  switchTab(tab: 'input' | 'preview') {
    this.activeTab = tab;
    if (tab === 'preview') {
      this.updatePreviews();
      this.togglePreviewSubTab('visual');
    }
  }

  togglePreviewSubTab(tab: 'visual' | 'code') {
    this.previewSubTab = tab;
  }

  updateMeta() {
    // In Angular with ngModel, sprintData is already updated.
  }

  handleFileUpload(event: any) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt: any) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      
      const newItems = data.map((row: any) => {
        return {
            id: row['Work Item ID'] || row['ID'] || '',
            type: row['Work Item Type'] || row['Type'] || 'USER STORY',
            title: row['Work Item'] || row['Title'] || '',
            scope: row['Scope'] || '',
            resource: row['Resource'] || '',
            points: Number(row['Story Points']) || 0
        };
      }).filter((item:any) => item.id || item.title);

      if (newItems.length > 0) {
           this.workItems = newItems;
           this.calculateStats();
           setTimeout(() => this.initLucide(), 0); // Re-init icons for new rows
           this.cd.detectChanges(); // Force UI update
      } else {
          alert('No valid items found. Please check Excel headers: Work Item ID, Work Item Type, Work Item, Scope, Resource, Story Points');
      }
    };

    reader.readAsBinaryString(file);
    // Clearing the input value allows re-uploading the same file
    event.target.value = '';
  }

  clearData() {
    this.showModal('System Purge', 'Are you sure you want to clear all data? This action is irreversible.', () => {
      this.workItems = [];
      const fileInput = document.getElementById('excel-upload') as HTMLInputElement;
      if(fileInput) fileInput.value = '';
      this.calculateStats();
    });
  }

  addRow() {
    this.workItems.push({ id: '', type: 'BUG', title: '', scope: '', resource: 'Manikandan Kasi', points: 0 });
    this.calculateStats();
    setTimeout(() => this.initLucide(), 0);
  }

  removeRow(idx: number) {
    if(confirm('Delete row?')) {
      this.workItems.splice(idx, 1);
      this.calculateStats();
    }
  }

  onItemChange() {
    this.calculateStats();
  }

  calculateStats() {
    this.totalPoints = this.workItems.reduce((sum, i) => sum + (Number(i.points) || 0), 0);
    this.totalCount = this.workItems.length;
  }

  // --- Modal Logic ---
  showModal(title: string, msg: string, callback: () => void) {
    this.modalTitle = title;
    this.modalMsg = msg;
    this.modalCallback = callback;
    this.modalVisible = true;
  }

  closeModal() {
    this.modalVisible = false;
    this.modalCallback = null;
  }

  confirmModal() {
    if (this.modalCallback) {
      this.modalCallback();
    }
    this.closeModal();
  }

  exportWord() {
    const logoImg = document.querySelector('.logo-preview') as HTMLImageElement;
    let logoSrc = 'assets/cat-logo.png';
    if (logoImg) {
        const b64 = this.getBase64Image(logoImg);
        if (b64) logoSrc = b64;
    }

    let htmlContent = this.generateHTML(logoSrc);
    
    // Optimizations for Word:
    // 1. Remove box-shadow (often renders as black artifacts)
    htmlContent = htmlContent.replace(/box-shadow:.*?;/g, '');
    // 2. Change fixed width 800 to 100% to fit page margins
    htmlContent = htmlContent.replace('width="800"', 'width="100%"');
    
    const preHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>Export HTML To Doc</title>
        <style>
          @page {
            size: 21cm 29.7cm;  /* A4 Portrait */
            margin: 1cm 1.5cm 1cm 1.5cm; /* Margins */
            mso-page-orientation: portrait;
          }
          body { font-family: Arial, sans-serif; }
          table { width: 100% !important; border-collapse: collapse; }
        </style>
        <!--[if gte mso 9]>
        <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
        </xml>
        <![endif]-->
      </head>
      <body>
    `;
    const postHtml = "</body></html>";
    const html = preHtml + htmlContent + postHtml;

    const blob = new Blob(['\ufeff', html], {
        type: 'application/msword'
    });
    
    const url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);
    
    // Create download link
    const link = document.createElement('a');
    link.href = url;
    link.download = `${this.sprintData.number}_Planning_Report.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // --- Helpers ---for Base64 Logo ---
  getBase64Image(img: HTMLImageElement): string | null {
      try {
          if (!img.complete || img.naturalWidth === 0) return null;
          const canvas = document.createElement("canvas");
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          const ctx = canvas.getContext("2d");
          if(ctx) {
              ctx.drawImage(img, 0, 0);
              return canvas.toDataURL("image/png");
          }
          return null;
      } catch(e) {
          console.warn('CORS/Taint error for logo', e);
          return null;
      }
  }

  // --- Generation Logic ---
  getOrg(name: string): string {
    return this.orgMap[name] || 'Cognizant';
  }

  generateHTML(logoSrc: string = 'assets/cat-logo.png'): string {
    const totalPoints = this.totalPoints;
    
    // Build Summary Table
    const summaryMap: any = {};
    this.workItems.forEach(i => {
        if(!summaryMap[i.resource]) summaryMap[i.resource] = { name: i.resource, items: 0, points: 0 };
        summaryMap[i.resource].items++;
        summaryMap[i.resource].points += (Number(i.points)||0);
    });
    
    const summaryRows = Object.values(summaryMap).map((s:any, idx:number) => `
        <tr style="background-color: ${idx % 2 === 0 ? '#ffffff' : '#f4f4f4'};">
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; font-family:'Roboto', Arial, sans-serif; font-size:14px; color:#333; text-align:center;">${idx+1}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; font-family:'Roboto', Arial, sans-serif; font-size:14px; font-weight:700; color:#333;">${s.name}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; font-family:'Roboto', Arial, sans-serif; font-size:14px; color:#555;">${this.getOrg(s.name)}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; font-family:'Roboto', Arial, sans-serif; font-size:14px; text-align:center; color:#333;">${s.items}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; font-family:'Roboto', Arial, sans-serif; font-size:14px; text-align:center; font-weight:900; color:#000;">${s.points}</td>
        </tr>
    `).join('');

    const detailRows = this.workItems.map((item, idx) => {
        const bg = idx % 2 === 0 ? '#ffffff' : '#f4f4f4';
        return `
        <tr>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:14px; color:#333; vertical-align:top;">${item.id}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:12px; font-weight:700; color:#666; vertical-align:top; text-transform:uppercase;">${item.type}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:14px; font-weight:500; color:#000; vertical-align:top;">${item.title}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:14px; color:#444; line-height:1.5; vertical-align:top;">${(item.scope||'').replace(/\n/g, '<br>')}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:14px; font-weight:700; color:#333; white-space:nowrap; vertical-align:top;">${item.resource}</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:14px; text-align:center; color:#777; vertical-align:top;">NA</td>
            <td style="padding:16px 20px; border-bottom:1px solid #e0e0e0; background:${bg}; font-family:'Roboto', Arial, sans-serif; font-size:14px; text-align:center; font-weight:900; color:#000; vertical-align:top;">${item.points}</td>
        </tr>
        `;
    }).join('');

    return `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <link href="https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700;900&family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
</head>
<body style="margin:0; padding:0; background-color:#eeeeee; font-family:'Roboto', Arial, sans-serif;">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td align="center" style="padding: 40px 0;">
            <!-- MAIN CONTAINER -->
            <table border="0" cellpadding="0" cellspacing="0" width="960" style="background-color:#ffffff; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                
                <!-- CAT BRAND HEADER -->
                <tr>
                    <td bgcolor="#000000" style="padding: 0;">
                        <!-- Top Yellow Strip -->
                        <div style="height: 4px; background-color: #FFCD11; width: 100%;"></div>
                        
                        <table width="100%" cellpadding="0" cellspacing="0">
                            <tr>
                                <td width="60" valign="middle" style="padding: 35px 20px 35px 40px;">
                                    <!-- CAT Logo -->
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td bgcolor="#ffffff" style="background-color: #ffffff; padding: 10px;">
                                                <img src="${logoSrc}" width="160" style="width:160px; display:block;" alt="CAT">
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                                <td valign="middle" style="padding: 35px 0;">
                                    <h1 style="margin:0; font-family:'Roboto Condensed', sans-serif; font-size: 32px; font-weight: 900; text-transform: uppercase; color: #ffffff; letter-spacing: 1px; line-height:1;">
                                        <font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">DQME</span></font> <font color="#FFCD11" style="color:#FFCD11;"><span style="color:#FFCD11;">Telematics Portal</span></font>
                                    </h1>
                                    <div style="font-family:'Roboto Condensed', sans-serif; font-size: 14px; font-weight: 700; text-transform: uppercase; color: #999; letter-spacing: 1px; margin-top:5px;">
                                        ${this.sprintData.number} Planning Report
                                    </div>
                                </td>
                                <td align="right" valign="middle" style="padding: 35px 40px 35px 0;">
                                    <div style="background-color:#212121; padding: 10px 20px; font-family:'Roboto', sans-serif; font-size: 14px; font-weight:700; color: #ffffff; display:inline-block;">
                                        <font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">${this.sprintData.startDate} — ${this.sprintData.endDate}</span></font>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>

                <!-- CONTENT AREA -->
                <tr>
                    <td style="padding: 50px 40px 40px 40px;">
                        <p style="font-family:'Roboto', sans-serif; font-size:16px; line-height:1.6; color:#333; margin-top:0;">
                            <strong>Hello Team,</strong><br><br>
                            Following our planning session for <strong>${this.sprintData.number}</strong>, we have finalized the estimation for the committed work items. 
                            The team has committed to a total of <strong style="background-color:#FFCD11; padding:0 4px;">${totalPoints} Story Points</strong>.
                        </p>
                        
                        <div style="height: 30px;"></div>

                        <!-- SECTION TITLE -->
                        <table width="100%" cellpadding="0" cellspacing="0">
                            <tr>
                                <td style="border-left: 6px solid #FFCD11; padding-left: 15px;">
                                    <h3 style="margin:0; font-family:'Roboto Condensed', sans-serif; font-size:20px; font-weight:700; text-transform:uppercase; color:#000;">
                                        Estimation Summary
                                    </h3>
                                </td>
                            </tr>
                        </table>
                        
                        <div style="height: 15px;"></div>

                        <!-- SUMMARY TABLE -->
                        <table width="100%" cellspacing="0" cellpadding="0" style="border-top: 2px solid #000;">
                            <tr bgcolor="#212121">
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; text-align:center;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">#</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Resource</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Organization</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; text-align:center;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Work Items</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; text-align:center;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Story Points</span></font></th>
                            </tr>
                            ${summaryRows}
                            <tr bgcolor="#000000">
                                <td colspan="4" style="padding:12px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:13px; font-weight:700; text-transform:uppercase; text-align:right;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Total Commitment</span></font></td>
                                <td style="padding:12px 14px; color:#FFCD11; font-family:'Roboto', sans-serif; font-size:16px; font-weight:900; text-align:center;"><font color="#FFCD11" style="color:#FFCD11;"><span style="color:#FFCD11;">${totalPoints}</span></font></td>
                            </tr>
                        </table>
                        
                         <div style="height: 20px;"></div>

                        <p style="font-family:'Roboto', sans-serif; font-size:14px; color:#333;">
                            There are currently <strong>no significant dependencies or risk identified</strong>, and the team is ready to commence the sprint.
                        </p>
                        
                        <div style="height: 30px;"></div>

                        <!-- SECTION TITLE -->
                        <table width="100%" cellpadding="0" cellspacing="0">
                            <tr>
                                <td style="border-left: 6px solid #FFCD11; padding-left: 15px;">
                                    <h3 style="margin:0; font-family:'Roboto Condensed', sans-serif; font-size:20px; font-weight:700; text-transform:uppercase; color:#000;">
                                        Detailed Scope & Estimation
                                    </h3>
                                </td>
                            </tr>
                        </table>
                        
                         <div style="height: 15px;"></div>

                        <!-- DETAIL TABLE -->
                        <table width="100%" cellspacing="0" cellpadding="0" style="border-top: 2px solid #000;">
                            <tr bgcolor="#212121">
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">ID</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Type</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Work Item</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Scope</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:left;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Resource</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:center;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Risk</span></font></th>
                                <th style="padding:10px 14px; color:#ffffff; font-family:'Roboto Condensed', sans-serif; font-size:12px; font-weight:700; text-transform:uppercase; text-align:center;"><font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">Pts</span></font></th>
                            </tr>
                            ${detailRows}
                        </table>
                    </td>
                </tr>

                <!-- CAT FOOTER -->
                <tr>
                    <td bgcolor="#212121" style="padding: 30px 40px;">
                        <table width="100%" cellpadding="0" cellspacing="0">
                            <tr>
                                <td style="font-family:'Roboto Condensed', sans-serif; font-size: 14px; font-weight: 700; color: #ffffff; text-transform: uppercase;">
                                    <font color="#ffffff" style="color:#ffffff;"><span style="color:#ffffff;">DQME Telematics Portal</span></font>
                                </td>
                                <td align="right" style="font-family:'Roboto', sans-serif; font-size: 12px; color: #999999;">
                                    <font color="#999999" style="color:#999999;"><span style="color:#999999;">© ${new Date().getFullYear()} Confidential Internal Report</span></font>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
</body>
</html>`;
  }

  updatePreviews() {
      // Try to get base64 logo from the sidebar if available
      const logoImg = document.querySelector('.logo-preview') as HTMLImageElement;
      let logoSrc = 'assets/cat-logo.png';
      if (logoImg) {
          const b64 = this.getBase64Image(logoImg);
          if (b64) logoSrc = b64;
      }

      this.previewHtmlSafe = this.sanitizer.bypassSecurityTrustHtml(this.generateHTML(logoSrc));
  }

  copyVisualReport() {
    // Generate visual report (body content only)
    const logoImg = document.querySelector('.logo-preview') as HTMLImageElement;
    let logoSrc = 'assets/cat-logo.png';
    if (logoImg) {
        const b64 = this.getBase64Image(logoImg);
        if (b64) logoSrc = b64;
    }
    
    const html = this.generateHTML(logoSrc);
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const bodyContent = doc.body.innerHTML;

    const tempDiv = document.createElement('div');
    tempDiv.style.position = 'fixed'; 
    tempDiv.style.opacity = '0.01'; 
    tempDiv.innerHTML = bodyContent;
    document.body.appendChild(tempDiv);
    
    const range = document.createRange();
    range.selectNode(tempDiv);
    window.getSelection()?.removeAllRanges();
    window.getSelection()?.addRange(range);
    
    try {
        const successful = document.execCommand('copy');
        if(successful) {
            this.showModal('System Notification', 'Report Copied! Paste into Outlook (Ctrl+V).', () => {});
        } else throw new Error("Copy failed");
    } catch(e) {
        this.showModal('Operation Failed', 'Rich Text Copy Failed.', () => {});
    }
    
    window.getSelection()?.removeAllRanges();
    document.body.removeChild(tempDiv);
  }

  exportImage() {
      const renderTank = document.createElement('div');
      // Visible Overlay Strategy
      const loader = document.createElement('div');
      loader.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: #222; z-index: 99990; display: flex; align-items: center; justify-content: center; flex-direction: column; opacity: 0.9;';
      loader.innerHTML = '<div style="font-family: sans-serif; font-weight: bold; font-size: 24px; color: #fff;">Rendering Image...</div>';
      document.body.appendChild(loader);

      renderTank.style.cssText = 'position: absolute; top: 20px; left: 50%; transform: translateX(-50%); width: 960px; z-index: 99991; background: #ffffff; display: block; overflow: visible; opacity: 1;';
      
      const logoImg = document.querySelector('.logo-preview') as HTMLImageElement;
      let logoSrc = 'assets/cat-logo.png';
      if (logoImg) {
          const b64 = this.getBase64Image(logoImg);
          if (b64) logoSrc = b64;
      }
      
      renderTank.innerHTML = this.generateHTML(logoSrc);
      
      // Need to append to body to render
      document.body.appendChild(renderTank);

      // Wait for render
      setTimeout(() => {
           // Re-parse to get just the content logic we want if generateHTML returns full doc
           const parser = new DOMParser();
           const doc = parser.parseFromString(this.generateHTML(logoSrc), 'text/html');
           const content = doc.body.firstElementChild; 
           if(content) {
             renderTank.innerHTML = '';
             renderTank.appendChild(content);
           }

           html2canvas(renderTank, { scale: 2, useCORS: true, logging: false }).then((canvas: any) => {
               const link = document.createElement('a');
               link.download = `Estimation_${this.sprintData.number.replace(/\s+/g,'_')}.png`;
               link.href = canvas.toDataURL();
               link.click();
               
               // Cleanup
               document.body.removeChild(renderTank);
               if(document.body.contains(loader)) document.body.removeChild(loader);
           }).catch((err: any) => {
               this.showModal('Export Failed', err.message, () => {});
               document.body.removeChild(renderTank);
               if(document.body.contains(loader)) document.body.removeChild(loader);
           });
      }, 500);
  }

  exportPdf() {
       const renderTank = document.createElement('div');

       const loader = document.createElement('div');
       loader.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: #222; z-index: 99990; display: flex; align-items: center; justify-content: center; flex-direction: column; opacity: 0.9;';
       loader.innerHTML = '<div style="font-family: sans-serif; font-weight: bold; font-size: 24px; color: #fff;">Rendering PDF...</div>';
       document.body.appendChild(loader);

       renderTank.style.cssText = 'position: absolute; top: 20px; left: 50%; transform: translateX(-50%); width: 960px; z-index: 99991; background: #ffffff; display: block; overflow: visible; opacity: 1;';
       
       const logoImg = document.querySelector('.logo-preview') as HTMLImageElement;
       let logoSrc = 'assets/cat-logo.png';
       if (logoImg) {
           const b64 = this.getBase64Image(logoImg);
           if (b64) logoSrc = b64;
       }

       const parser = new DOMParser();
       const doc = parser.parseFromString(this.generateHTML(logoSrc), 'text/html');
       const content = doc.body.firstElementChild;
       if(content) {
         renderTank.appendChild(content);
       }
       document.body.appendChild(renderTank);
       
       // Scroll top
       window.scrollTo(0,0);
       
       setTimeout(() => {
           html2canvas(renderTank, {
               scale: 2,
               useCORS: true,
               scrollY: 0,
               logging: false,
               windowWidth: 1200
           }).then((canvas: any) => {
               try {
                   const imgData = canvas.toDataURL('image/png');
                   const { jsPDF } = jspdf; 
                   const pdf = new jsPDF('p', 'mm', 'a4');
                   const pdfWidth = pdf.internal.pageSize.getWidth();
                   const imgProps = pdf.getImageProperties(imgData);
                   const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
                   
                   pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);
                   pdf.save(`Estimation_${this.sprintData.number.replace(/\s+/g,'_')}.pdf`);
                   
                   document.body.removeChild(renderTank);
                   if(document.body.contains(loader)) document.body.removeChild(loader);
               } catch(e) {
                   throw e;
               }
           }).catch((err: any) => {
               this.showModal('Export Failed', 'PDF Error: ' + err.message, () => {});
               document.body.removeChild(renderTank);
               if(document.body.contains(loader)) document.body.removeChild(loader);
           });
       }, 800);
  }
}
