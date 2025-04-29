// js/file-to-excel.js

document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('fileInput');
  const convertBtn = document.getElementById('convertBtn');
  const statusDiv = document.getElementById('status');

  convertBtn.addEventListener('click', async () => {
    const files = fileInput.files;
    if (!files.length) {
      statusDiv.textContent = 'Please select at least one file.';
      return;
    }
    statusDiv.textContent = 'Processing files...';

    try {
      const allTexts = [];

      for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        let text = '';

        if (ext === 'txt') {
          text = await readTextFile(file);
        } else if (ext === 'docx') {
          text = await readDocxFile(file);
        } else if (ext === 'pdf') {
          text = await readPdfFile(file);
        } else {
          statusDiv.textContent = `Unsupported file type: ${file.name}`;
          return;
        }

        allTexts.push({ fileName: file.name, content: text });
      }

      // Convert allTexts to Excel
      const wb = XLSX.utils.book_new();

      allTexts.forEach(({ fileName, content }) => {
        // Split content by lines for rows
        const rows = content.split(/\r?\n/).map(line => [line]);
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(fileName));
      });

      // Generate Excel file and trigger download
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);

      const a = document.createElement('a');
      a.href = url;
      a.download = 'converted_files.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      statusDiv.textContent = 'Excel file generated and downloaded successfully!';
    } catch (error) {
      console.error(error);
      statusDiv.textContent = 'Error processing files. See console for details.';
    }
  });

  function readTextFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => resolve(e.target.result);
      reader.onerror = e => reject(e);
      reader.readAsText(file);
    });
  }

  function readDocxFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        const arrayBuffer = e.target.result;
        mammoth.extractRawText({ arrayBuffer: arrayBuffer })
          .then(result => resolve(result.value))
          .catch(err => reject(err));
      };
      reader.onerror = e => reject(e);
      reader.readAsArrayBuffer(file);
    });
  }

  async function readPdfFile(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let text = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const strings = content.items.map(item => item.str);
      text += strings.join(' ') + '\n';
    }
    return text;
  }

  function sanitizeSheetName(name) {
    // Excel sheet names max length 31, no special chars like : \ / ? * [ ]
    return name.substring(0, 31).replace(/[:\\/?*\[\]]/g, '_');
  }
});
