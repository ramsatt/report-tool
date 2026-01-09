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
  selector: 'app-sprint-closure',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './sprint-closure.component.html',
  styleUrls: ['./sprint-closure.component.css']
})
export class SprintClosureComponent implements OnInit, AfterViewInit {
  // State
  activeTab: 'input' | 'preview' = 'input';
  previewSubTab: 'visual' | 'code' = 'visual';

  sprintData = {
    number: 'Sprint 25',
    startDate: 'Dec 10',
    endDate: 'Dec 23'
  };

  workItems: any[] = [];
  
  // Stats
  totalPoints = 0;
  totalCount = 0;
  donePoints = 0;
  spillPoints = 0;
  yieldExp = '0%';
  delPointsDisplay = '0/0';
  yieldValue = 0;

  // Modal State
  modalVisible = false;
  modalTitle = '';
  modalMsg = '';
  modalIsConfirm = false;
  modalCallback: (() => void) | null = null;
  modalOkText = 'ACKNOWLEDGE';

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
        let status = row['Status'] || row['State'] || 'Completed';
        if (status.toString().replace(/\s/g, '').toLowerCase() === 'inprogress') {
            status = 'In-Progress';
        }
        
        return {
            id: row['Work ID'] || row['ID'] || row['Ref ID'] || '',
            type: row['Work Item Type'] || row['Type'] || 'USER STORY',
            title: row['Title'] || row['Operation Title'] || row['Description'] || '',
            points: Number(row['Story Points']) || Number(row['Points']) || 0,
            status: status,
            devOverview: row['Development Overview'] || row['Overview'] || 'Migration',
            sprint: row['Sprint'] || ''
        };
      }).filter((item:any) => item.title || item.id);

      this.workItems = newItems;

      if (this.workItems.length > 0 && this.workItems[0].sprint) {
         this.sprintData.number = 'Sprint ' + this.workItems[0].sprint;
      }
      
      this.calculateStats();

      setTimeout(() => this.initLucide(), 0);
      this.cd.detectChanges(); // Force UI update
    };

    reader.readAsBinaryString(file);
    // Clear input to allow re-upload 
    event.target.value = '';
  }

  clearData() {
    this.showModal('System Purge', 'Are you sure you want to clear all data? This action is irreversible.', true, () => {
      this.workItems = [];
      const fileInput = document.getElementById('excel-upload') as HTMLInputElement;
      if(fileInput) fileInput.value = '';
      this.calculateStats();
    });
  }

  addRow() {
    this.workItems.push({ 
        id: '', type: 'USER STORY', title: '', points: 0, 
        devOverview: 'Migration', status: 'Completed' 
    });
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
    this.totalCount = this.workItems.length;
    this.totalPoints = this.workItems.reduce((sum, i) => sum + (Number(i.points) || 0), 0);
    this.donePoints = this.workItems.filter(i => i.status === 'Completed').reduce((sum, i) => sum + (Number(i.points) || 0), 0);
    
    this.delPointsDisplay = `${this.donePoints}/${this.totalPoints}`;
    const pct = this.totalPoints > 0 ? Math.round((this.donePoints / this.totalPoints) * 100) : 0;
    this.yieldExp = `${pct}%`;
    this.yieldValue = pct;
  }

  // --- Modal Logic ---
  showModal(title: string, msg: string, isConfirm: boolean, callback: () => void) {
    this.modalTitle = title;
    this.modalMsg = msg;
    this.modalIsConfirm = isConfirm;
    this.modalCallback = callback;
    this.modalOkText = isConfirm ? 'CONFIRM TERM' : 'ACKNOWLEDGE';
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

  // --- Helper for Base64 Logo ---
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
  formatDateForEmail(dateStr: string): string {
    if (!dateStr) return '';
    // Typically dateStr is like "Dec 24" or "2024-12-24" depending on input. 
    // If user typed "Dec 24", we assume it's display ready or we can parse it.
    // The previous code did: const [y, m, d] = dateStr.split('-');
    // But input type is text in one place and date in another? 
    // Actually current input is type="text" in sprint_closure.html (Line 61 is not visible but sprint_planning has type="text").
    // Wait, let's assume it's just text string as it appears in the UI screenshot "Dec 10".
    return dateStr; 
  }

  generateHTML(logoSrc: string = 'assets/cat-logo.png'): string {
    const totalPoints = this.totalPoints;
    const donePoints = this.donePoints;
    const spillPoints = totalPoints - donePoints;
    const percentage = totalPoints > 0 ? ((donePoints / totalPoints) * 100).toFixed(2) : '0.00';

    // Generate Rows
    const rowsHtml = this.workItems.map((item, idx) => {
        const isEven = idx % 2 === 0;
        const bg = isEven ? '#ffffff' : '#f4f4f4';
        const isDone = item.status === 'Completed';
        
        // Styling specific to status
        const statusColor = isDone ? '#000000' : '#888888';
        const statusBg = isDone ? '#FFCD11' : '#e5e5e5';
        const statusFontWeight = isDone ? 'bold' : 'normal';

        return `
        <tr>
            <td style="padding: 12px; font-family: Arial, sans-serif; font-size: 13px; color: #000; border-bottom: 1px solid #ddd; background-color: ${bg}; white-space: nowrap;">${item.id}</td>
            <td style="padding: 12px; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold;color: #666; text-transform: uppercase; border-bottom: 1px solid #ddd; background-color: ${bg}; white-space: nowrap;">${item.type}</td>
            <td style="padding: 12px; font-family: Arial, sans-serif; font-size: 13px; color: #000; border-bottom: 1px solid #ddd; background-color: ${bg};">${item.title}</td>
            <td style="padding: 12px; font-family: Arial, sans-serif; font-size: 13px; color: #000; text-align: center; border-bottom: 1px solid #ddd; background-color: ${bg}; font-weight: bold; white-space: nowrap;">${item.points}</td>
            <td style="padding: 12px; font-family: Arial, sans-serif; font-size: 13px; color: #444; border-bottom: 1px solid #ddd; background-color: ${bg};">${item.devOverview}</td>
            <td style="padding: 12px; font-family: Arial, sans-serif; border-bottom: 1px solid #ddd; background-color: ${statusBg}; color: ${statusColor}; text-align: center; vertical-align: middle; font-weight: ${statusFontWeight}; font-size: 11px; text-transform: uppercase; white-space: nowrap;">
                ${item.status}
            </td>
        </tr>
        `;
    }).join('');

    return `<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!--[if !mso]><!-->
    <link href="https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700;900&family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
    <!--<![endif]-->
    <style>
        table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
        img { -ms-interpolation-mode: bicubic; }
    </style>
</head>
<body style="margin: 0; padding: 0; background-color: #eeeeee; font-family: Arial, sans-serif;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family: Arial, sans-serif;">
        <tr>
            <td align="center" style="padding: 20px 0;">
                <!-- MAIN CONTAINER -->
                <table border="0" cellpadding="0" cellspacing="0" width="800" style="background-color: #ffffff; border-top: 8px solid #FFCD11; box-shadow: 0 5px 15px rgba(0,0,0,0.15); font-family: Arial, sans-serif;">
                    
                    <!-- HEADER -->
                    <tr>
                        <td bgcolor="#111111" style="background-color: #111111; padding: 30px 40px; color: #ffffff;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family: Arial, sans-serif; color: #ffffff;">
                                <tr>
                                    <td align="left" valign="middle" style="color: #ffffff;">
                                        <table border="0" cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif;">
                                            <tr>
                                                <td valign="middle" style="padding-right: 20px;">
                                                    <table border="0" cellpadding="0" cellspacing="0">
                                                        <tr>
                                                            <td bgcolor="#ffffff" style="background-color: #ffffff; padding: 5px;">
                                                                <img src="${logoSrc}" width="100" height="50" style="height: 50px; width: auto; display: block; border: 0;" alt="CAT" border="0">
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td valign="middle" style="color: #ffffff;">
                                                    <font face="Arial, sans-serif" color="#ffffff" style="font-size: 28px; line-height: 100%; font-weight: bold; font-style: italic; color: #ffffff;">
                                                        <font color="#FFCD11" style="color:#FFCD11"><span style="color:#FFCD11">DQME</span></font> <span style="color:#ffffff">Telematics Portal</span>
                                                    </font>
                                                    <br>
                                                    <font face="Arial, sans-serif" color="#999999" style="font-size: 12px; font-weight: bold; text-transform: uppercase; letter-spacing: 2px;">
                                                        Sprint Closure Report
                                                    </font>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="right" valign="middle" style="color: #ffffff;">
                                        <font face="Arial, sans-serif" color="#ffffff" style="font-size: 32px; font-weight: bold; display: block;">
                                            ${this.sprintData.number}
                                        </font>
                                        <br>
                                        <font face="Arial, sans-serif" color="#FFCD11" style="font-size: 14px; font-weight: bold; letter-spacing: 1px;">
                                            ${this.formatDateForEmail(this.sprintData.startDate)} - ${this.formatDateForEmail(this.sprintData.endDate)}
                                        </font>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- INTRO TEXT -->
                    <tr>
                        <td style="padding: 40px 40px 0 40px; background-color: #ffffff; font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333333;">
                            <font face="Arial, sans-serif" color="#333333">
                            Hi Mahesh,<br><br>
                            Please find the closure summary for <strong>${this.sprintData.number}</strong> below. The team has successfully delivered <strong>${donePoints}</strong> story points out of the committed <strong>${totalPoints}</strong>. ${spillPoints > 0 ? `The remaining <strong>${spillPoints}</strong> points have been carried over to the next sprint.` : ''}<br><br>
                            Could you please review and approve the metrics? Your confirmation would be appreciated to ensure alignment on the deliverables.
                            </font>
                        </td>
                    </tr>

                    <!-- CLOSURE SUMMARY (Metric Blocks) -->
                    <tr>
                        <td style="padding: 40px 40px 20px 40px; background-color: #ffffff;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="font-family: Arial, sans-serif;">
                                <tr>
                                    <!-- Block 1: Committed -->
                                    <td width="23%" valign="top" bgcolor="#111111" style="background-color: #111111; padding: 20px 15px; border-left: 6px solid #FFCD11; color: #ffffff;">
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr><td style="padding-bottom: 5px; color: #ffffff;"><font face="Arial Black, Arial, sans-serif" color="#FFCD11" style="font-size: 32px; line-height: 100%; font-weight: 900; color: #FFCD11;"><span style="color: #FFCD11;">${totalPoints}</span></font></td></tr>
                                            <tr><td style="color: #ffffff;"><font face="Arial, sans-serif" color="#ffffff" style="font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; color: #ffffff;"><span style="color: #ffffff;">Committed SP</span></font></td></tr>
                                        </table>
                                    </td>
                                    
                                    <td width="2%">&nbsp;</td>

                                    <!-- Block 2: Completed -->
                                    <td width="23%" valign="top" bgcolor="#111111" style="background-color: #111111; padding: 20px 15px; border-left: 6px solid #FFCD11; color: #ffffff;">
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr><td style="padding-bottom: 5px; color: #ffffff;"><font face="Arial Black, Arial, sans-serif" color="#FFCD11" style="font-size: 32px; line-height: 100%; font-weight: 900; color: #FFCD11;"><span style="color: #FFCD11;">${donePoints}</span></font></td></tr>
                                            <tr><td style="color: #ffffff;"><font face="Arial, sans-serif" color="#ffffff" style="font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; color: #ffffff;"><span style="color: #ffffff;">Completed SP</span></font></td></tr>
                                        </table>
                                    </td>

                                    <td width="2%">&nbsp;</td>

                                    <!-- Block 3: Spilled -->
                                    <td width="23%" valign="top" bgcolor="#111111" style="background-color: #111111; padding: 20px 15px; border-left: 6px solid #FFCD11; color: #ffffff;">
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr><td style="padding-bottom: 5px; color: #ffffff;"><font face="Arial Black, Arial, sans-serif" color="${spillPoints > 0 ? '#ff4444' : '#FFCD11'}" style="font-size: 32px; line-height: 100%; font-weight: 900; color: ${spillPoints > 0 ? '#ff4444' : '#FFCD11'};"><span style="color: ${spillPoints > 0 ? '#ff4444' : '#FFCD11'};">${spillPoints}</span></font></td></tr>
                                            <tr><td style="color: #ffffff;"><font face="Arial, sans-serif" color="#ffffff" style="font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; color: #ffffff;"><span style="color: #ffffff;">Spilled SP</span></font></td></tr>
                                        </table>
                                    </td>

                                    <td width="2%">&nbsp;</td>

                                    <!-- Block 4: Progress -->
                                    <td width="23%" valign="top" bgcolor="#111111" style="background-color: #111111; padding: 20px 15px; border-left: 6px solid #FFCD11; color: #ffffff;">
                                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                            <tr><td style="padding-bottom: 5px; color: #ffffff;"><font face="Arial Black, Arial, sans-serif" color="#ffffff" style="font-size: 24px; line-height: 100%; font-weight: 900; color: #ffffff;"><span style="color: #ffffff;">${percentage}%</span></font></td></tr>
                                            <tr><td style="padding-bottom: 8px; color: #ffffff;"><font face="Arial, sans-serif" color="#999999" style="font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; color: #999999;"><span style="color: #999999;">Completion</span></font></td></tr>
                                            <tr>
                                                <td>
                                                    <table border="0" cellpadding="0" cellspacing="0" width="100%" height="6" bgcolor="#333333" style="background-color: #333333;">
                                                        <tr>
                                                            <td width="${percentage}%" bgcolor="#FFCD11" height="6" style="font-size:1px; line-height:1px; background-color: #FFCD11;">&nbsp;</td>
                                                            <td width="${100 - Number(percentage)}%" height="6" style="font-size:1px; line-height:1px;">&nbsp;</td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- DETAILS TABLE -->
                    <tr>
                        <td style="padding: 0 40px 40px 40px;">
                            <h3 style="font-family: 'Arial Black', Arial, sans-serif; font-size: 16px; text-transform: uppercase; margin-bottom: 20px; border-left: 6px solid #FFCD11; padding-left: 10px;">Work Item Breakdown</h3>
                            
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; border: 1px solid #e5e5e5; font-family: Arial, sans-serif;">
                                <thead>
                                    <tr bgcolor="#000000">
                                        <th align="left" style="padding: 12px; color: #FFCD11; font-family: Arial, sans-serif; font-size: 11px; text-transform: uppercase; white-space: nowrap;"><font color="#FFCD11">Work ID</font></th>
                                        <th align="left" style="padding: 12px; color: #FFCD11; font-family: Arial, sans-serif; font-size: 11px; text-transform: uppercase; white-space: nowrap;"><font color="#FFCD11">Type</font></th>
                                        <th align="left" style="padding: 12px; color: #FFCD11; font-family: Arial, sans-serif; font-size: 11px; text-transform: uppercase;"><font color="#FFCD11">Title</font></th>
                                        <th align="center" style="padding: 12px; color: #FFCD11; font-family: Arial, sans-serif; font-size: 11px; text-transform: uppercase; white-space: nowrap;"><font color="#FFCD11">Pts</font></th>
                                        <th align="left" style="padding: 12px; color: #FFCD11; font-family: Arial, sans-serif; font-size: 11px; text-transform: uppercase;"><font color="#FFCD11">Overview</font></th>
                                        <th align="center" style="padding: 12px; color: #FFCD11; font-family: Arial, sans-serif; font-size: 11px; text-transform: uppercase; white-space: nowrap;"><font color="#FFCD11">Status</font></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${rowsHtml || '<tr><td colspan="6" align="center" style="padding:30px; color:#999; font-style:italic;">No items data available.</td></tr>'}
                                </tbody>
                            </table>
                        </td>
                    </tr>

                    <!-- FOOTER -->
                    <tr>
                        <td bgcolor="#212121" style="padding: 30px 40px; background-color: #212121;">
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="font-family:'Roboto Condensed', sans-serif; font-size: 14px; font-weight: 700; color: #ffffff; text-transform: uppercase;">
                                        <font color="#ffffff" style="color:#ffffff"><span style="color:#ffffff">DQME Telematics Portal</span></font>
                                    </td>
                                    <td align="right" style="font-family:'Roboto', sans-serif; font-size: 12px; color: #999999;">
                                        <font color="#999999">© ${new Date().getFullYear()} Confidential Internal Report</font>
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
    // Extract JUST the body content
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const bodyContent = doc.body.innerHTML;

    // Create a temporary container
    const tempDiv = document.createElement('div');
    tempDiv.style.position = 'fixed';
    tempDiv.style.left = '0';
    tempDiv.style.top = '0';
    tempDiv.style.opacity = '0.01';
    tempDiv.style.pointerEvents = 'none';
    tempDiv.style.zIndex = '-9999';
    tempDiv.innerHTML = bodyContent;
    document.body.appendChild(tempDiv);
    
    const range = document.createRange();
    range.selectNode(tempDiv);
    window.getSelection()?.removeAllRanges();
    window.getSelection()?.addRange(range);
    
    try {
        const successful = document.execCommand('copy');
        if(successful) {
            this.showModal('System Notification', 'Report Copied! You can now paste (Ctrl+V) directly into Outlook.', false, () => {});
        } else {
            throw new Error('execCommand returned false');
        }
    } catch (err) {
         this.showModal('Operation Failed', 'Copy failed. Please try selecting the preview manually and copying.', false, () => {});
    }
    
    window.getSelection()?.removeAllRanges();
    document.body.removeChild(tempDiv);
  }

  exportImage() {
      const renderTank = document.createElement('div');
      
      const loader = document.createElement('div');
      loader.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: #222; z-index: 99990; display: flex; align-items: center; justify-content: center; flex-direction: column; opacity: 0.9;';
      loader.innerHTML = '<div style="font-family: sans-serif; font-weight: bold; font-size: 24px; color: #fff;">Rendering Image...</div>';
      document.body.appendChild(loader);

      renderTank.style.cssText = 'position: absolute; top: 20px; left: 50%; transform: translateX(-50%); width: 800px; z-index: 99991; background: #ffffff; display: block; overflow: visible; opacity: 1;';
      
      const logoImg = document.querySelector('.logo-preview') as HTMLImageElement;
      let logoSrc = 'assets/cat-logo.png';
      if (logoImg) {
          const b64 = this.getBase64Image(logoImg);
          if (b64) logoSrc = b64;
      }
      
      renderTank.innerHTML = this.generateHTML(logoSrc);
      
      document.body.appendChild(renderTank);

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
               link.download = `Sprint_Report_${this.sprintData.number.replace(/\s+/g,'_')}.png`;
               link.href = canvas.toDataURL();
               link.click();
               
               // Cleanup
               document.body.removeChild(renderTank);
               if(document.body.contains(loader)) document.body.removeChild(loader);
           }).catch((err: any) => {
               this.showModal('Export Failed', err.message, false, () => {});
               document.body.removeChild(renderTank);
               if(document.body.contains(loader)) document.body.removeChild(loader);
           });
      }, 500);
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
    link.download = `Sprint_Report_${this.sprintData.number.replace(/\s+/g,'_')}.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  exportPdf() {
       const renderTank = document.createElement('div');

       const loader = document.createElement('div');
       loader.id = 'pdf-loader';
       loader.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: #222; z-index: 99990; display: flex; align-items: center; justify-content: center; flex-direction: column; opacity: 0.9;';
       loader.innerHTML = '<div style="font-family: sans-serif; font-weight: bold; font-size: 24px; color: #fff;">Generating PDF...</div><div style="margin-top: 10px; font-size: 14px; color: #ccc;">Using High-Fidelity Capture</div>';
       document.body.appendChild(loader);

       renderTank.style.cssText = 'position: absolute; top: 20px; left: 50%; transform: translateX(-50%); width: 800px; z-index: 99991; background: #ffffff; display: block; overflow: visible; opacity: 1; box-shadow: 0 0 20px rgba(0,0,0,0.5);';
       
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
       
       window.scrollTo(0,0);
       
       setTimeout(() => {
           html2canvas(renderTank, {
               scale: 2,
               useCORS: true,
               scrollY: 0,
               logging: false,
               windowWidth: 1024
           }).then((canvas: any) => {
               try {
                   const imgData = canvas.toDataURL('image/png');
                   const { jsPDF } = jspdf; 
                   const pdf = new jsPDF('p', 'mm', 'a4');
                   const pdfWidth = pdf.internal.pageSize.getWidth();
                   const imgProps = pdf.getImageProperties(imgData);
                   const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
                   
                   pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);
                   pdf.save(`Sprint_Report_${this.sprintData.number.replace(/\s+/g,'_')}.pdf`);
                   
                   document.body.removeChild(renderTank);
                   if(document.body.contains(loader)) document.body.removeChild(loader);
               } catch(e) {
                   throw e;
               }
           }).catch((err: any) => {
               this.showModal('Export Failed', 'PDF Error: ' + err.message, false, () => {});
               document.body.removeChild(renderTank);
               if(document.body.contains(loader)) document.body.removeChild(loader);
           });
       }, 800);
  }
}
