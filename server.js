// // Backend (Node.js)
// const express = require("express");
// const multer = require("multer");
// const xlsx = require("xlsx");
// const cors = require("cors");
// const fs = require("fs");
// const path = require("path");

// const app = express();
// const port = 5000;

// // Middleware
// app.use(cors());
// app.use(express.json());

// // File upload setup
// const upload = multer({ dest: "uploads/" });

// // Endpoint to handle file upload and merging
// app.post("/merge", upload.array("files"), (req, res) => {
//   const files = req.files;
//   const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : [];

//   if (!files || files.length === 0) {
//     return res.status(400).json({ error: "No files uploaded." });
//   }

//   try {
//     const mergedData = [];

//     files.forEach((file, index) => {
//       const workbook = xlsx.readFile(file.path);
//       const sheetName = workbook.SheetNames[0]; // Assume the first sheet
//       const sheet = workbook.Sheets[sheetName];
//       const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

//       const startRow = rowConfig[index]?.startRow || 1; // Default to row 1

//       for (let i = startRow - 1; i < data.length; i++) {
//         mergedData.push(data[i]);
//       }

//       // Clean up temporary file
//       fs.unlinkSync(file.path);
//     });

//     // Create a new workbook and sheet for merged data
//     const newWorkbook = xlsx.utils.book_new();
//     const newSheet = xlsx.utils.aoa_to_sheet(mergedData);
//     xlsx.utils.book_append_sheet(newWorkbook, newSheet, "MergedSheet");

//     // Save merged file
//     const outputPath = path.join(__dirname, "merged.xlsx");
//     xlsx.writeFile(newWorkbook, outputPath);

//     return res.download(outputPath, "merged.xlsx", () => {
//       fs.unlinkSync(outputPath); // Clean up merged file after download
//     });
//   } catch (error) {
//     console.error(error);
//     return res.status(500).json({ error: "An error occurred while merging files." });
//   }
// });

// app.listen(port, () => {
//   console.log(`Server running on http://localhost:${port}`);
// });

// // Backend (Node.js)
// const express = require('express')
// const multer = require('multer')
// const xlsx = require('xlsx')
// const cors = require('cors')
// const fs = require('fs')
// const path = require('path')

// const app = express()
// const port = 5000

// // Middleware
// app.use(cors())
// app.use(express.json())

// // File upload setup
// const upload = multer({ dest: 'uploads/' })

// // Endpoint to handle file upload and merging
// app.post('/merge', upload.array('files'), (req, res) => {
//   const files = req.files
//   const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : []

//   if (!files || files.length === 0) {
//     return res.status(400).json({ error: 'No files uploaded.' })
//   }

//   try {
//     const mergedData = []

//     files.forEach((file, index) => {
//       const workbook = xlsx.readFile(file.path)
//       const sheetName = workbook.SheetNames[0] // Assume the first sheet
//       const sheet = workbook.Sheets[sheetName]
//       const data = xlsx.utils.sheet_to_json(sheet, { header: 1 })

//       const startRow = rowConfig[index]?.startRow || 1 // Default to row 1
//       const endColumn = rowConfig[index]?.endColumn || 'Z' // Default to column Z

//       // Convert endColumn to column index (Excel A = 1, B = 2, ..., Z = 26)
//       const endColumnIndex = xlsx.utils.decode_col(endColumn.toUpperCase())

//       for (let i = startRow - 1; i < data.length; i++) {
//         const rowData = data[i].slice(0, endColumnIndex + 1) // Slice up to the end column
//         mergedData.push(rowData)
//       }

//       // Clean up temporary file
//       fs.unlinkSync(file.path)
//     })

//     // Create a new workbook and sheet for merged data
//     const newWorkbook = xlsx.utils.book_new()
//     const newSheet = xlsx.utils.aoa_to_sheet(mergedData)
//     xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'MergedSheet')

//     // Save merged file
//     const outputPath = path.join(__dirname, 'merged.xlsx')
//     xlsx.writeFile(newWorkbook, outputPath)

//     return res.download(outputPath, 'merged.xlsx', () => {
//       fs.unlinkSync(outputPath) // Clean up merged file after download
//     })
//   } catch (error) {
//     console.error(error)
//     return res
//       .status(500)
//       .json({ error: 'An error occurred while merging files.' })
//   }
// })

// app.listen(port, () => {
//   console.log(`Server running on http://localhost:${port}`)
// })

// Backend (Node.js)
// const express = require('express')
// const multer = require('multer')
// const xlsx = require('xlsx')
// const cors = require('cors')
// const fs = require('fs')
// const path = require('path')

// const app = express()
// const port = 5000

// // Middleware
// app.use(cors())
// app.use(express.json())

// // File upload setup
// const upload = multer({ dest: 'uploads/' })

// // Endpoint to handle file upload and merging
// app.post('/merge', upload.array('files'), (req, res) => {
//   const files = req.files
//   const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : []

//   if (!files || files.length === 0) {
//     return res.status(400).json({ error: 'No files uploaded.' })
//   }

//   try {
//     const mergedData = []

//     files.forEach((file, index) => {
//       const workbook = xlsx.readFile(file.path)
//       const sheetName = workbook.SheetNames[0] // Assume the first sheet
//       const sheet = workbook.Sheets[sheetName]
//       const data = xlsx.utils.sheet_to_json(sheet, { header: 1 })

//       const startRow = rowConfig[index]?.startRow || 1 // Default to row 1
//       const endRow = rowConfig[index]?.endRow || data.length // Default to the last row

//       // Loop over the specified row range
//       for (let i = startRow - 1; i < endRow; i++) {
//         mergedData.push(data[i])
//       }

//       // Clean up temporary file
//       fs.unlinkSync(file.path)
//     })

//     // Create a new workbook and sheet for merged data
//     const newWorkbook = xlsx.utils.book_new()
//     const newSheet = xlsx.utils.aoa_to_sheet(mergedData)
//     xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'MergedSheet')

//     // Save merged file
//     const outputPath = path.join(__dirname, 'merged.xlsx')
//     xlsx.writeFile(newWorkbook, outputPath)

//     return res.download(outputPath, 'merged.xlsx', () => {
//       fs.unlinkSync(outputPath) // Clean up merged file after download
//     })
//   } catch (error) {
//     console.error(error)
//     return res
//       .status(500)
//       .json({ error: 'An error occurred while merging files.' })
//   }
// })

// app.listen(port, () => {
//   console.log(`Server running on http://localhost:${port}`)
// })
//Finallllllllllllllllllllllllllllllllllllllllllll
// const express = require('express')
// const multer = require('multer')
// const xlsx = require('xlsx')
// const cors = require('cors')
// const fs = require('fs')
// const path = require('path')

// const app = express()
// const port = 5000

// // Middleware
// app.use(cors())
// app.use(express.json())

// // File upload setup
// const upload = multer({ dest: 'uploads/' })

// // Function to convert column letter to index
// function columnToIndex(col) {
//   let index = 0
//   for (let i = 0; i < col.length; i++) {
//     index = index * 26 + col.charCodeAt(i) - 65 + 1
//   }
//   return index - 1 // Adjust for zero-based index
// }

// // Endpoint to handle file upload and merging
// app.post('/merge', upload.array('files'), (req, res) => {
//   const files = req.files
//   const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : []

//   if (!files || files.length === 0) {
//     return res.status(400).json({ error: 'No files uploaded.' })
//   }

//   try {
//     const mergedData = []

//     files.forEach((file, index) => {
//       const workbook = xlsx.readFile(file.path)
//       const sheetName = workbook.SheetNames[0] // Assume the first sheet
//       const sheet = workbook.Sheets[sheetName]
//       const data = xlsx.utils.sheet_to_json(sheet, { header: 1 })

//       const startRow = rowConfig[index]?.startRow || 1 // Default to row 1
//       const endRow = rowConfig[index]?.endRow || data.length // Default to the last row
//       const startCol = rowConfig[index]?.startColumn || 'A' // Default to column A
//       const endCol = rowConfig[index]?.endColumn || 'Z' // Default to column Z

//       // Convert columns to indices
//       const startColIndex = columnToIndex(startCol)
//       const endColIndex = columnToIndex(endCol)

//       // Loop over the specified row range and column range
//       for (let i = startRow - 1; i < endRow; i++) {
//         const row = data[i]?.slice(startColIndex, endColIndex + 1) // Slice by columns

//         // Filter out empty rows
//         if (
//           row &&
//           row.some((cell) => cell !== undefined && cell !== null && cell !== '')
//         ) {
//           mergedData.push(row)
//         }
//       }

//       // Clean up temporary file
//       fs.unlinkSync(file.path)
//     })

//     // Create a new workbook and sheet for merged data
//     const newWorkbook = xlsx.utils.book_new()
//     const newSheet = xlsx.utils.aoa_to_sheet(mergedData)
//     xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'MergedSheet')

//     // Save merged file
//     const outputPath = path.join(__dirname, 'merged.xlsx')
//     xlsx.writeFile(newWorkbook, outputPath)

//     return res.download(outputPath, 'merged.xlsx', () => {
//       fs.unlinkSync(outputPath) // Clean up merged file after download
//     })
//   } catch (error) {
//     console.error(error)
//     return res
//       .status(500)
//       .json({ error: 'An error occurred while merging files.' })
//   }
// })

// app.listen(port, () => {
//   console.log(`Server running on http://localhost:${port}`)
// })

// const express = require('express')
// const multer = require('multer')
// const xlsx = require('xlsx')
// const cors = require('cors')
// const fs = require('fs')
// const path = require('path')

// const app = express()
// const port = 5000

// // Middleware
// app.use(cors())
// app.use(express.json())

// // File upload setup
// const upload = multer({ dest: 'uploads/' })

// // Function to convert column letter to index
// function columnToIndex(col) {
//   let index = 0
//   for (let i = 0; i < col.length; i++) {
//     index = index * 26 + col.charCodeAt(i) - 65 + 1
//   }
//   return index - 1 // Adjust for zero-based index
// }

// // Endpoint to handle file upload and merging
// app.post('/merge', upload.array('files'), (req, res) => {
//   const files = req.files
//   const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : []

//   if (!files || files.length === 0) {
//     return res.status(400).json({ error: 'No files uploaded.' })
//   }

//   try {
//     const mergedData = []

//     files.forEach((file, index) => {
//       const workbook = xlsx.readFile(file.path)
//       const sheetName = workbook.SheetNames[0] // Assume the first sheet
//       const sheet = workbook.Sheets[sheetName]
//       const data = xlsx.utils.sheet_to_json(sheet, { header: 1 })

//       const startRow = rowConfig[index]?.startRow || 1 // Default to row 1
//       const endRow = rowConfig[index]?.endRow || data.length // Default to the last row
//       const startCol = rowConfig[index]?.startColumn || 'A' // Default to column A
//       const endCol = rowConfig[index]?.endColumn || 'Z' // Default to column Z

//       // Convert columns to indices
//       const startColIndex = columnToIndex(startCol)
//       const endColIndex = columnToIndex(endCol)

//       // Loop over the specified row range and column range
//       for (let i = startRow - 1; i < endRow; i++) {
//         const row = data[i]?.slice(startColIndex, endColIndex + 1) // Slice by columns

//         // Filter out empty rows (skip completely empty rows)
//         if (
//           row &&
//           row.some((cell) => cell !== undefined && cell !== null && cell !== '')
//         ) {
//           mergedData.push(row)
//         }
//       }

//       // Clean up temporary file
//       fs.unlinkSync(file.path)
//     })

//     // Create a new workbook and sheet for merged data
//     const newWorkbook = xlsx.utils.book_new()
//     const newSheet = xlsx.utils.aoa_to_sheet(mergedData)
//     xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'MergedSheet')

//     // Save merged file
//     const outputPath = path.join(__dirname, 'merged.xlsx')
//     xlsx.writeFile(newWorkbook, outputPath)

//     return res.download(outputPath, 'merged.xlsx', () => {
//       fs.unlinkSync(outputPath) // Clean up merged file after download
//     })
//   } catch (error) {
//     console.error(error)
//     return res
//       .status(500)
//       .json({ error: 'An error occurred while merging files.' })
//   }
// })

// app.listen(port, () => {
//   console.log(`Server running on http://localhost:${port}`)
// })

// server.js (backend code)
const express = require('express')
const multer = require('multer')
const xlsx = require('xlsx')
const cors = require('cors')
const fs = require('fs')
const path = require('path')

const app = express()
app.use(cors())
app.use(express.json())

// Use the /tmp directory on Vercel for file uploads
const upload = multer({ dest: '/tmp/uploads/' })

app.post('/merge', upload.array('files'), (req, res) => {
  const files = req.files
  const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : []

  if (!files || files.length === 0) {
    return res.status(400).json({ error: 'No files uploaded.' })
  }

  try {
    const mergedData = []

    files.forEach((file, index) => {
      const workbook = xlsx.readFile(file.path)
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      const data = xlsx.utils.sheet_to_json(sheet, { header: 1 })

      const startRow = rowConfig[index]?.startRow || 1

      for (let i = startRow - 1; i < data.length; i++) {
        mergedData.push(data[i])
      }

      fs.unlinkSync(file.path)
    })

    const newWorkbook = xlsx.utils.book_new()
    const newSheet = xlsx.utils.aoa_to_sheet(mergedData)
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'MergedSheet')

    const outputPath = path.join(__dirname, 'merged.xlsx')
    xlsx.writeFile(newWorkbook, outputPath)

    return res.download(outputPath, 'merged.xlsx', () => {
      fs.unlinkSync(outputPath)
    })
  } catch (error) {
    console.error(error)
    return res
      .status(500)
      .json({ error: 'An error occurred while merging files.' })
  }
})

module.exports = app
