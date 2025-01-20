const express = require('express')
const multer = require('multer')
const xlsx = require('xlsx')
const fs = require('fs')
const path = require('path')
const cors = require('cors')
const stream = require('stream')

const app = express()
const port = 5000

// Middleware
app.use(cors())
app.use(express.json())

// In-memory file storage instead of saving to disk
const storage = multer.memoryStorage()
const upload = multer({ storage: storage })

// Function to convert column letter to index
function columnToIndex(col) {
  let index = 0
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + col.charCodeAt(i) - 65 + 1
  }
  return index - 1 // Adjust for zero-based index
}

// Helper to create the merged file stream
function createMergedStream(mergedData) {
  const readable = new stream.Readable()
  readable._read = () => {} // No-op
  mergedData.forEach((row) => {
    readable.push(JSON.stringify(row) + '\n')
  })
  readable.push(null) // End the stream
  return readable
}

// Endpoint to handle file upload and merging
app.post('/merge', upload.array('files'), (req, res) => {
  const files = req.files
  const rowConfig = req.body.rowConfig ? JSON.parse(req.body.rowConfig) : {}

  if (!files || files.length === 0) {
    return res.status(400).json({ error: 'No files uploaded.' })
  }

  try {
    const mergedData = []
    const tempFilePath = path.join(__dirname, 'merged_temp.xlsx')

    // Open a write stream for the final merged data
    const writeStream = fs.createWriteStream(tempFilePath)

    files.forEach((file) => {
      const workbook = xlsx.read(file.buffer)
      const sheetName = workbook.SheetNames[0] // Assume the first sheet
      const sheet = workbook.Sheets[sheetName]
      const data = xlsx.utils.sheet_to_json(sheet, { header: 1 })

      const startRow = rowConfig.startRow || 1
      const endRow = rowConfig.endRow || data.length
      const startCol = rowConfig.startColumn || 'A'
      const endCol = rowConfig.endColumn || 'Z'

      // Convert columns to indices
      const startColIndex = columnToIndex(startCol)
      const endColIndex = columnToIndex(endCol)

      // Loop over the specified row range and column range
      for (let i = startRow - 1; i < endRow; i++) {
        const row = data[i]?.slice(startColIndex, endColIndex + 1) // Slice by columns

        // Filter out empty rows
        if (
          row &&
          row.some((cell) => cell !== undefined && cell !== null && cell !== '')
        ) {
          mergedData.push(row)
        }
      }
    })

    // Create a readable stream from mergedData
    const mergedStream = createMergedStream(mergedData)

    // Pipe the merged data to the write stream
    mergedStream.pipe(writeStream)

    // Once writing is done, send the result
    writeStream.on('finish', () => {
      res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx')
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      )

      // Read the merged file and send it as a response
      fs.createReadStream(tempFilePath).pipe(res)
    })
  } catch (error) {
    console.error(error)
    return res
      .status(500)
      .json({ error: 'An error occurred while merging files.' })
  }
})

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`)
})
