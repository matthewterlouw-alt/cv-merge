const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', '*');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    const { cvData, filename } = req.body;

    if (!cvData) {
      return res.status(400).json({ error: 'Missing cvData' });
    }

    const templateUrl = 'https://docs.google.com/document/d/1yXWnobc7FhdvwbYupZHNjHWBqJ612vqdbEbwph_IP2Y/export?format=docx';
    console.log('Fetching template from:', templateUrl);
    
    const templateResponse = await fetch(templateUrl, {
      redirect: 'follow'
    });
    
    console.log('Response status:', templateResponse.status);
    console.log('Response headers:', Object.fromEntries(templateResponse.headers.entries()));
    
    if (!templateResponse.ok) {
      const text = await templateResponse.text();
      console.log('Error response:', text.substring(0, 500));
      return res.status(400).json({ error: 'Failed to fetch template', status: templateResponse.status });
    }
    
    const arrayBuffer = await templateResponse.arrayBuffer();
    console.log('ArrayBuffer size:', arrayBuffer.byteLength);
    
    const bytes = Buffer.from(arrayBuffer);

    const zip = new PizZip(bytes);
    const doc = new Docxtemplater(zip, { 
      paragraphLoop: true, 
      linebreaks: true,
      delimiters: { start: '{{', end: '}}' }
    });
    
    // FIXED: Now includes all template placeholders correctly
    doc.render({
      // Contact info (fixed)
      NAME: cvData.NAME || "",
      LOCATION: cvData.LOCATION || "",
      PHONE: cvData.PHONE || "",
      EMAIL: cvData.EMAIL || "",
      LINKEDIN: cvData.LINKEDIN || "",
      
      // Target title (from vacancy)
      TARGET_TITLE: cvData.TARGET_TITLE || "",
      
      // Section labels (for language switching)
      SECTION_LABELS: cvData.SECTION_LABELS || {},
      
      // FIXED: Template uses SUMMARY_TEXT, not SUMMARY
      SUMMARY_TEXT: cvData.SUMMARY_TEXT || "",
      
      // Skills array
      SKILLS: (cvData.SKILLS || []).map(s => ({
        CATEGORY: s.CATEGORY || "",
        CONTENT: s.CONTENT || ""
      })),
      
      // FIXED: Work blocks now include COMPANY_CONTEXT and ROLE_SCOPE
      WORK_BLOCKS: (cvData.WORK_BLOCKS || []).map(block => ({
        BLOCK_HEADER: block.BLOCK_HEADER || "",
        COMPANY_CONTEXT: block.COMPANY_CONTEXT || "",  // ADDED
        ENTRIES: (block.ENTRIES || []).map(entry => ({ 
          ENTRY_HEADER: entry.ENTRY_HEADER || "", 
          ROLE_SCOPE: entry.ROLE_SCOPE || "",  // ADDED
          BULLETS: entry.BULLETS || [] 
        }))
      })),
      
      // Certifications array
      CERTIFICATIONS: cvData.CERTIFICATIONS || [],
      
      // Awards array
      AWARDS: cvData.AWARDS || [],
      
      // Languages array
      LANGUAGES: cvData.LANGUAGES || [],
      
      // Leadership array
      LEADERSHIP: (cvData.LEADERSHIP || []).map(item => ({ 
        HEADER: item.HEADER || "", 
        DETAIL: item.DETAIL || "" 
      }))
    });

    const output = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const base64 = output.toString('base64');
    
    return res.status(200).json({ docxBase64: base64, filename: filename || 'cv.docx' });
  } catch (error) {
    console.error('Merge error:', error);
    return res.status(500).json({ error: error.message });
  }
};
