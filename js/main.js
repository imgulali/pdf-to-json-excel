import fs from "fs";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { getDocument } from "pdfjs-dist";
import { utils, writeFile } from "xlsx";

const __dirname = dirname(fileURLToPath(import.meta.url));

function formatName(name) {
  return name
    .trim()
    .replace(/\s+/g, " ")
    .split(" ")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(" ");
}

function validatePdfFile(pdfFilePath) {
  if (!fs.existsSync(pdfFilePath)) {
    throw new Error(`PDF file "${pdfFilePath}" not found.`);
  }
}

function validateFileName(fileName, extension) {
  const validExtensions = [".json", ".xlsx"];
  const fileExtension = fileName.slice(fileName.lastIndexOf("."));

  if (!validExtensions.includes(fileExtension)) {
    throw new Error(
      `Invalid file extension for ${fileName}. Expected ${extension} file.`
    );
  }
}

async function main(pdfFileName, jsonFileName, excelFileName) {
  try {
    const io = join(__dirname, "io");

    const pdfFilePath = join(io, pdfFileName);

    validatePdfFile(pdfFilePath);

    const pdf = await getDocument(pdfFilePath).promise;
    let text = "";

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      text += content.items.map((item) => item.str).join(" ") + "\n";
    }

    const matches = [...text.matchAll(/(F\d{10})\s+([A-Z\s]+)/g)];
    if (matches.length === 0) {
      return console.log("No student data found in the PDF.");
    }

    const students = matches.map(([_, studentId, name]) => ({
      studentId,
      name: formatName(name),
    }));

    validateFileName(jsonFileName, ".json");
    validateFileName(excelFileName, ".xlsx");

    const jsonFilePath = join(io, jsonFileName || "data.json");
    const excelFilePath = join(io, excelFileName || "students.xlsx");

    fs.writeFileSync(jsonFilePath, JSON.stringify(students, null, 2), "utf8");

    const wb = utils.book_new();
    const ws = utils.json_to_sheet(students);

    utils.book_append_sheet(wb, ws, "Students");
    writeFile(wb, excelFilePath);
  } catch (error) {
    console.error(`Error: ${error.message}`);
  }
}

const args = process.argv.slice(2);
let pdfFileName = "";
let jsonFileName = "data.json";
let excelFileName = "students.xlsx";

for (let i = 0; i < args.length; i++) {
  if (args[i] === "--json" && args[i + 1]) {
    jsonFileName = args[i + 1];
    i++;
  } else if (args[i] === "--excel" && args[i + 1]) {
    excelFileName = args[i + 1];
    i++;
  } else if (!pdfFileName) {
    pdfFileName = args[i];
  }
}

main(pdfFileName, jsonFileName, excelFileName);
