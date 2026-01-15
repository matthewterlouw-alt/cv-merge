const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', '*');
  
  if (req.method === 'OPTIONS') return res.status(200).end();
  
  const { cvData, templateBase64, filename } = req.body;
  
  const bytes = Buffer.from(templateBase64, 'base64');
  const zip = new PizZip(bytes);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  
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
      ENTRIES: (b.ENTRIES || []).map(e => ({ ENTRY_HEADER: e.ENTRY_HEADER || "", BULLETS: e.BULLETS || [] }))
    })),
    EDUCATION: cvData.EDUCATION || [],
    LEADERSHIP: (cvData.LEADERSHIP || []).map(i => ({ HEADER: i.HEADER || "", DETAIL: i.DETAIL || "" }))
  });

  const output = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  const base64 = output.toString('base64');
  
  res.json({ docxBase64: base64, filename });
};
