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

    const templateResponse = await fetch('https://docs.google.com/document/d/1yXWnobc7FhdvwbYupZHNjHWBqJ612vqdbEbwph_IP2Y/export?format=docx');
    if (!templateResponse.ok) {
      return res.status(400).json({ error: 'Failed to fetch template' });
    }
    
    const arrayBuffer = await templateResponse.arrayBuffer();
    const bytes = Buffer.from(arrayBuffer);

    const zip = new PizZip(bytes);
    const doc = new Docxtemplater(zip, { 
      paragraphLoop: true, 
      linebreaks: true,
      delimiters: { start: '{{', end: '}}' }
    });
    
    doc.render({
      NAME: cvData.NAME || "",
      LOCATION: cvData.LOCATION || "",
      PHONE: cvData.PHONE || "",
      EMAIL: cvData.EMAIL || "",
      LINKEDIN: cvData.LINKEDIN || "",
      TARGET_TITLE: cvData.TARGET_TITLE || "",
      SECTION_LABELS: cvData.SECTION_LABELS || {},
      SUMMARY: cvData.SUMMARY || "",
      SKILLS: cvData.SKILLS || [],
      WORK_BLOCKS: (cvData.WORK_BLOCKS || []).map(b => ({
        BLOCK_HEADER: b.BLOCK_HEADER || "",
        ENTRIES: (b.ENTRIES || []).map(e => ({ 
          ENTRY_HEADER: e.ENTRY_HEADER || "", 
          BULLETS: e.BULLETS || [] 
        }))
      })),
      EDUCATION: cvData.EDUCATION || [],
      LEADERSHIP: (cvData.LEADERSHIP || []).map(i => ({ 
        HEADER: i.HEADER || "", 
        DETAIL: i.DETAIL || "" 
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
