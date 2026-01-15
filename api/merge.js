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
    
    // FIXED: Flatten section labels to avoid nested object issues
    const sectionLabels = cvData.SECTION_LABELS || {};
    
    doc.render({
      // Contact info (fixed)
      NAME: cvData.NAME || "",
      LOCATION: cvData.LOCATION || "",
      PHONE: cvData.PHONE || "",
      EMAIL: cvData.EMAIL || "",
      LINKEDIN: cvData.LINKEDIN || "",
      
      // Target title
      TARGET_TITLE: cvData.TARGET_TITLE || "",
      
      // FIXED: Flattened section labels (no more nested object)
      LABEL_SUMMARY: sectionLabels.SUMMARY || "SUMMARY",
      LABEL_SKILLS: sectionLabels.SKILLS || "SKILLS",
      LABEL_WORK_EXPERIENCE: sectionLabels.WORK_EXPERIENCE || "WORK EXPERIENCE",
      LABEL_CERTIFICATIONS: sectionLabels.CERTIFICATIONS || "CERTIFICATIONS",
      LABEL_AWARDS: sectionLabels.AWARDS || "AWARDS",
      LABEL_LANGUAGES: sectionLabels.LANGUAGES || "LANGUAGES",
      LABEL_LEADERSHIP: sectionLabels.LEADERSHIP || "LEADERSHIP AND ACTIVITIES",
      
      // Summary text
      SUMMARY_TEXT: cvData.SUMMARY_TEXT || "",
      
      // Skills array
      SKILLS: (cvData.SKILLS || []).map(s => ({
        CATEGORY: s.CATEGORY || "",
        CONTENT: s.CONTENT || ""
      })),
      
      // Work entries - flat structure, each role is standalone
      WORK_ENTRIES: (cvData.WORK_ENTRIES || []).map(entry => ({
      ROLE_LINE: entry.ROLE_LINE || "",
      BULLETS: entry.BULLETS || []
      })),
      
      // Simple arrays
      CERTIFICATIONS: cvData.CERTIFICATIONS || [],
      AWARDS: cvData.AWARDS || [],
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
