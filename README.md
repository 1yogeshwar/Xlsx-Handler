# Xlsx-Handler


```javascript
const XLSX = require('xlsx');

const importExcel = (fileBuffer) => {
  // Read from buffer instead of file path
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  // Validate headers exist
  if (rawData.length === 0) {
    throw new Error('Excel file is empty');
  }

  const validatedData = rawData.map((row, index) => {
    let remark = 'success';
    const remarks = [];
    
    // Validate date format (dd-mm-yyyy or dd/mm/yyyy or dd.mm.yyyy)
    const dateRegex = /^(\d{2})[-\/.](\d{2})[-\/.](\d{4})$/;
    if (!row.Date || !dateRegex.test(String(row.Date).trim())) {
      remarks.push('Invalid date format');
    }
    
    // Check for null/empty values
    const fields = [
      { name: 'Tournament', value: row.Tournament },
      { name: 'Home Team', value: row['Home Team'] },
      { name: 'Home Goals', value: row['Home Goals'] },
      { name: 'Away Goals', value: row['Away Goals'] },
      { name: 'Away Team', value: row['Away Team'] },
      { name: 'Win Conditions', value: row['Win Conditions'] },
      { name: 'Home Ground', value: row['Home Ground'] }
    ];
    
    fields.forEach(field => {
      if (field.value === null || field.value === undefined || field.value === '') {
        remarks.push(`${field.name} is null/empty`);
      }
    });
    
    // Check for 0 values in goals
    if (row['Home Goals'] === 0) remarks.push('Home Goals is 0');
    if (row['Away Goals'] === 0) remarks.push('Away Goals is 0');
    
    if (remarks.length > 0) {
      remark = `failed: ${remarks.join(', ')}`;
    }
    
    return {
      ID: row.ID,
      Tournament: row.Tournament,
      Date: row.Date,
      'Home Team': row['Home Team'],
      'Home Goals': row['Home Goals'],
      'Away Goals': row['Away Goals'],
      'Away Team': row['Away Team'],
      'Win Conditions': row['Win Conditions'],
      'Home Ground': row['Home Ground'],
      Remark: remark
    };
  });

  return validatedData;
};

const exportToExcel = (data) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Validated Data');
  
  return workbook;
};

module.exports = { importExcel, exportToExcel };
```










**Controller (handles POST with file upload):**

```javascript
const { importExcel, exportToExcel } = require('./excelService');
const XLSX = require('xlsx');

const uploadAndValidate = async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    // Read file buffer from multer
    const fileBuffer = req.file.buffer;
    
    // Validate and get data
    const validatedData = importExcel(fileBuffer);
    
    // Send back validated data
    res.json({
      total: validatedData.length,
      success: validatedData.filter(r => r.Remark === 'success').length,
      failed: validatedData.filter(r => r.Remark.startsWith('failed')).length,
      data: validatedData
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

const downloadValidatedExcel = (req, res) => {
  try {
    // Get validated data from request body (sent from frontend)
    const validatedData = req.body.data;
    
    if (!validatedData || validatedData.length === 0) {
      return res.status(400).json({ error: 'No data to export' });
    }
    
    const workbook = exportToExcel(validatedData);
    
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Disposition', 'attachment; filename=validated_data.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};

module.exports = { uploadAndValidate, downloadValidatedExcel };
```








**Multer setup (in your route file):**

```javascript
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() }); // Store in memory

router.post('/upload', upload.single('file'), uploadAndValidate);
router.post('/download', downloadValidatedExcel);
```
