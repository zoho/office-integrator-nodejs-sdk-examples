import { exec } from 'child_process';

// List all the Node.js files you want to run in an array
const files = ["delete_pdf_document.js", "delete_pdf_session.js", "edit_pdf.js", "get_pdf_document_detail.js", "get_pdf_session_detail.js"];

// Loop through the array and run each file using the 'node' command
var i = 1;
files.forEach(file => {
  exec(`node ${file}`, (error, stdout, stderr) => {
    if (error) {
      console.error(`Error running ${file}: ${error}`);
      return;
    }
    console.log(`${i++} - ${file} output: ${stdout}`);
  });
});
