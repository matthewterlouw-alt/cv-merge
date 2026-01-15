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
    const { cvData, templateBase64, filename } = req.body;

    if (!cvData || !templateBase64) {
      return res.status(400).json({ error: 'Missing cvData or templateBase64' });
    }

    // Debug: log what we received
    console.log('templateBase64 length:', templateBase64?.length);
    console.log('templateBase64 starts with:', templateBase64?.substring(0, 100));

    // Strip data URL prefix if present
    let base64Data = templateBase64;
    if (base64Data.includes(',')) {
      base64Data = base64Data.split(',')[1];
    }

    const bytes = Buffer.from(base64Data, 'base64');
    console.log('Buffer length:', bytes.length);
    console.log('First bytes:', bytes.slice(0, 4).toString('hex'));

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
