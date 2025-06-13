import * as fs from 'fs';
import * as path from 'path';
import * as readline from 'readline';
import * as XLSX from 'xlsx';

interface RowData {
  id: string;
  provenance?: string;
  task?: string;
  difficulty?: string;
  language?: string;
  scope?: string;
  depth?: string;
}

async function parseJsonl(filePath: string): Promise<RowData[]> {
  const rows: RowData[] = [];

  if (!fs.existsSync(filePath)) {
    console.warn(`File not found: ${filePath}`);
    return rows;
  }

  const fileStream = fs.createReadStream(filePath);
  const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });

  for await (const line of rl) {
    try {
      const json = JSON.parse(line);
      const meta = json.metadata || {};

      rows.push({
        id: json.id || '',
        provenance: meta.provenance || '',
        task: meta.task || '',
        difficulty: meta.difficulty || '',
        language: meta.language || '',
        scope: meta.scope || '',
        depth: meta.depth || ''
      });
    } catch (error) {
      console.warn(`Skipping malformed line:\n${line}\nError: ${error}`);
    }
  }

  return rows;
}

async function convertJsonlFilesToExcel(directory: string) {
  const validatePath = path.join(directory, 'validate.jsonl');
  const trainPath = path.join(directory, 'train.jsonl');
  const outputExcelPath = path.join(directory, 'output.xlsx');

  const files = [
    { name: 'validate', path: validatePath },
    { name: 'train', path: trainPath }
  ];

  const workbook = XLSX.utils.book_new();
  let hasSheets = false;

  for (const file of files) {
    const data = await parseJsonl(file.path);
    if (data.length === 0) {
      console.warn(`No valid rows found in: ${file.path}`);
      continue;
    }

    const sheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, sheet, file.name);
    console.log(`Sheet added: ${file.name}`);
    hasSheets = true;
  }

  if (hasSheets) {
    XLSX.writeFile(workbook, outputExcelPath);
    console.log(`Excel file created: ${outputExcelPath}`);
  } else {
    console.error('No sheets were created. Excel file not saved.');
  }
}

// Get directory from command line argument
const directoryPath = process.argv[2];
if (!directoryPath) {
  console.error('[ERROR] Please provide a directory path as an argument.');
  process.exit(1);
}

convertJsonlFilesToExcel(directoryPath).catch(err => {
  console.error(`Failed to convert files: ${err}`);
});
