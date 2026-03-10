import express from 'express';
import cors from 'cors';
import multer from 'multer';
import xlsx from 'xlsx';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

function numberToChinese(num) {
  const digits = ['', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
  const units = ['', '拾', '佰', '仟', '万', '拾', '佰', '仟', '亿'];

  if (num === 0) return '零';

  const parts = num.toString().split('.');
  let integerPart = parseInt(parts[0]);
  let decimalPart = parts[1] ? parseInt(parts[1].padEnd(2, '0').substring(0, 2)) : 0;
  let result = '';

  if (integerPart > 0) {
    const str = integerPart.toString();
    let zeroFlag = false;
    for (let i = 0; i < str.length; i++) {
      const n = parseInt(str[str.length - 1 - i]);
      const unit = units[i];
      if (n === 0) {
        if (!zeroFlag && i > 0 && i % 4 !== 0) result = '零' + result;
        zeroFlag = true;
      } else {
        result = digits[n] + unit + result;
        zeroFlag = false;
      }
    }
    result = result.replace(/^零+/, '');
  }

  if (decimalPart > 0) {
    let decimalStr = '';
    const decimalDigits = decimalPart.toString().padStart(2, '0');
    for (let i = 0; i < decimalDigits.length; i++) {
      const n = parseInt(decimalDigits[i]);
      if (n !== 0) decimalStr += digits[n] + (i === 0 ? '角' : '分');
    }
    result += decimalStr;
  }

  return result || '零';
}

function parseExcel(buffer) {
  try {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonData.length < 2) {
      throw new Error('Excel文件至少需要包含表头和一行数据');
    }

    const hasComplexFormat = jsonData.some((row) => {
      if (!row) return false;
      return String(row[0]).includes('患儿姓名') ||
             String(row[1]).includes('患儿姓名') ||
             String(row[1]).includes('黄奕程');
    });

    if (hasComplexFormat) {
      return parseComplexFormat(jsonData);
    } else {
      return parseSimpleFormat(jsonData);
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function parseSimpleFormat(jsonData) {
  const FIELD_MAPPING = {
    '患儿姓名': 'patientName',
    '住院号': 'hospitalId',
    '身份证号': 'idCard',
    '合作医院': 'hospitalName',
    '秘书处建议资助金额': 'amount'
  };

  const headers = jsonData[0].map(h => String(h).trim());
  const columnIndices = {};
  headers.forEach((header, index) => {
    if (FIELD_MAPPING[header]) columnIndices[FIELD_MAPPING[header]] = index;
  });

  const data = [];
  for (let i = 1; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || row.length === 0) continue;

    const item = {};
    Object.keys(columnIndices).forEach(key => {
      let value = row[columnIndices[key]];
      if (key === 'amount' && value !== undefined) {
        const num = parseFloat(String(value).replace(/,/g, ''));
        if (!isNaN(num)) value = num.toLocaleString('zh-CN');
      }
      item[key] = String(value || '');
    });

    if (item.patientName) data.push(item);
  }

  return { success: true, data, total: data.length };
}

function parseComplexFormat(jsonData) {
  const data = [];
  let patientName = '', hospitalId = '', idCard = '', hospitalName = '', amount = '';

  if (jsonData[3] && jsonData[3][1]) patientName = String(jsonData[3][1]).trim();
  if (jsonData[8] && jsonData[8][12]) hospitalId = String(jsonData[8][12]).trim();
  if (jsonData[9] && jsonData[9][12]) idCard = String(jsonData[9][12]).trim();
  if (jsonData[6] && jsonData[6][2]) hospitalName = String(jsonData[6][2]).trim();
  if (!hospitalName && jsonData[7] && jsonData[7][2]) hospitalName = String(jsonData[7][2]).trim();
  if (jsonData[40] && jsonData[40][12]) {
    const value = String(jsonData[40][12]).trim();
    const numMatch = value.match(/(\d+)/);
    if (numMatch) amount = parseInt(numMatch[1]).toLocaleString('zh-CN');
  }

  if (patientName) {
    data.push({ patientName, hospitalId, idCard, hospitalName, amount });
  }

  return { success: true, data, total: data.length };
}

function convertHospitalName(name) {
  if (!name) return '';
  if (name.includes('妇儿')) return '广州医科大学附属妇女儿童医疗中心';
  if (name.includes('珠江医院') || name.includes('珠江')) return '南方医科大学珠江医院';
  return name;
}

async function generateWordDocument(patientData) {
  const date = new Date().toLocaleDateString('zh-CN');
  const hospitalName = convertHospitalName(patientData.hospitalName);
  const amountNum = parseFloat(String(patientData.amount).replace(/,/g, '')) || 0;
  const chineseAmount = numberToChinese(amountNum);

  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          children: [new TextRun({ text: "广州市易娱公益基金会儿童大病救助项目", size: 22 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 }
        }),
        new Paragraph({
          children: [new TextRun({ text: "救助通知书", bold: true, size: 36 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 }
        }),
        new Paragraph({
          children: [new TextRun({
            text: `今收悉${hospitalName}医院患儿（姓名：${patientData.patientName || ''}；住院号：${patientData.hospitalId || ''}；身份证号：${patientData.idCard || ''}）的救助申请资料，经基金会审核，同意救助其医疗费用人民币${patientData.amount || '0'}元（大写：人民币${chineseAmount}元整），救助款项仅用于该患者院内治疗费用。`
          })],
          spacing: { after: 400 }
        }),
        new Paragraph({
          children: [new TextRun({ text: "广州市易娱公益基金会", size: 22 })],
          alignment: AlignmentType.RIGHT,
          spacing: { after: 100 }
        }),
        new Paragraph({
          children: [new TextRun({ text: date, size: 22 })],
          alignment: AlignmentType.RIGHT
        })
      ]
    }]
  });

  return Packer.toBuffer(doc);
}

app.post('/api/parse-excel', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, error: '请上传Excel文件' });
  const result = parseExcel(req.file.buffer);
  if (!result.success) return res.status(400).json(result);
  res.json(result);
});

app.post('/api/generate-documents', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, error: '请上传Excel文件' });

  try {
    const parseResult = parseExcel(req.file.buffer);
    if (!parseResult.success) return res.status(400).json(parseResult);
    const patients = parseResult.data;
    if (patients.length === 0) return res.status(400).json({ success: false, error: '没有有效的数据行' });

    const zip = new JSZip();
    for (const patient of patients) {
      const docBuffer = await generateWordDocument(patient);
      const safeName = String(patient.patientName).replace(/[<>:"/\\|?*]/g, '').trim();
      zip.file(`${safeName}救助通知书.docx`, docBuffer);
    }

    const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="救助通知书_${date}.zip"`);
    res.send(zipBuffer);
  } catch (error) {
    res.status(500).json({ success: false, error: '生成文档时发生错误: ' + error.message });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Export for Vercel serverless
export default app;

// Local development server
if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}
