# Export Feature Quick Reference

## Available Export Formats

The Monthly Report Generator now supports **4 export formats**:

### 1. 📄 PDF Export
**Button:** Red "PDF" button  
**Icon:** File text icon  
**Output:** `Monthly_Report_[Month]_[Year].pdf`  
**Description:** Generates a PDF with high-quality images of all slides  
**Use Case:** For printing, email distribution, or archiving  
**Editable:** ❌ No (images only)

---

### 2. 🖼️ Image Export  
**Button:** Blue "IMG" button  
**Icon:** Image icon  
**Output:** Multiple PNG files: `Slide_01_[Month]_[Year].png`, `Slide_02_[Month]_[Year].png`, etc.  
**Description:** Exports each slide as a separate high-resolution PNG image  
**Use Case:** For social media, presentations, or embedding in documents  
**Editable:** ❌ No (images only)

---

### 3. 📊 PPTX (IMG) Export
**Button:** Orange "PPTX (IMG)" button  
**Icon:** Monitor icon  
**Output:** `Monthly_Report_[Month]_[Year].pptx`  
**Description:** PowerPoint file with embedded slide images  
**Use Case:** When you need PowerPoint but don't plan to edit content  
**Editable:** ❌ No (images embedded in slides)

---

### 4. ✨ PPTX (CONTENT) Export **[NEW]**
**Button:** Purple "PPTX (CONTENT)" button  
**Icon:** File type icon  
**Output:** `Monthly_Report_Content_[Month]_[Year].pptx`  
**Description:** PowerPoint file with **fully editable** text, tables, and shapes  
**Use Case:** When you need to customize the report after export  
**Editable:** ✅ Yes (all content is editable)

---

## When to Use Each Format

| Format | Best For | Pros | Cons |
|--------|----------|------|------|
| **PDF** | Final distribution, printing, archiving | Universal format, preserves layout | Not editable |
| **IMG** | Social media, embedding, presentations | High quality, flexible use | Multiple files |
| **PPTX (IMG)** | Quick PowerPoint for viewing | Fast generation, preserves exact design | Not editable |
| **PPTX (CONTENT)** | Customizable presentations | Fully editable, flexible | Slightly larger file size |

---

## Export Process

### Step 1: Prepare Your Data
- Enter data manually or upload Excel template
- Review in "Preview Deck" tab

### Step 2: Choose Export Format
- Click the appropriate button based on your needs
- Wait for the loading overlay (progress will be shown)

### Step 3: Download
- File downloads automatically
- Check your Downloads folder

### Step 4: Use Your Report
- PDF: Open in any PDF reader
- IMG: Files numbered sequentially (Slide_01, Slide_02, etc.)
- PPTX (IMG): Open in PowerPoint, view slides
- PPTX (CONTENT): Open in PowerPoint, edit as needed!

---

## PPTX (CONTENT) - What's Editable?

When you export with **PPTX (CONTENT)**, you can edit:

✅ **All text** - Titles, descriptions, bullet points  
✅ **Tables** - Migration status, feedback items, defect backlog  
✅ **Colors** - Change any color scheme  
✅ **Fonts** - Modify font styles and sizes  
✅ **Layout** - Rearrange elements on slides  
✅ **Data** - Update metrics, add/remove items  
✅ **Branding** - Customize company logos and names  

❌ **Not editable:**  
- Background images (if any)
- Embedded photos (logos are text-based)

---

## Troubleshooting

### Export takes too long
- **Solution:** Check your internet connection
- **Tip:** PPTX (CONTENT) is faster than image-based exports

### File doesn't open
- **PDF:** Ensure you have a PDF reader installed
- **PPTX:** Requires Microsoft PowerPoint or compatible software (Google Slides, LibreOffice)
- **IMG:** Should open in any image viewer

### Content looks different than preview
- **PDF/IMG/PPTX(IMG):** Should match preview exactly
- **PPTX(CONTENT):** May have slight formatting differences due to editable nature

### Download doesn't start
- **Check:** Browser popup blocker settings
- **Solution:** Allow downloads from this site

---

## Technical Details

### Export Process Time
- **PDF:** ~5-15 seconds (depends on slide count)
- **IMG:** ~10-20 seconds (one file per slide)
- **PPTX (IMG):** ~5-15 seconds
- **PPTX (CONTENT):** ~2-5 seconds ⚡ (fastest!)

### File Sizes (approximate)
- **PDF:** 2-5 MB
- **IMG:** 500KB-1MB per slide
- **PPTX (IMG):** 3-6 MB
- **PPTX (CONTENT):** 50-200 KB (smallest!)

### Supported Browsers
- ✅ Chrome (recommended)
- ✅ Edge
- ✅ Firefox
- ✅ Safari

---

## Pro Tips

1. **Use PPTX (CONTENT)** when you need flexibility
2. **Use PDF** for final, polished distribution
3. **Use IMG** when you need individual slides
4. **Preview before exporting** to ensure data is correct
5. **Keep backups** of your exported files
6. **Test edits** in PowerPoint after export

---

**Need Help?** Refer to the full documentation in `PPTX_CONTENT_EXPORT_SUMMARY.md`
