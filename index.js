const { YoutubeTranscript } = require("youtube-transcript");
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const xlsx = require("xlsx");

// Function to process each URL and fetch the transcript
async function processUrlAndGetTranscript(url) {
  try {
    const transcript = await YoutubeTranscript.fetchTranscript(url, { lang: "en" });
    return transcript.map(item => new Paragraph({ children: [new TextRun(item.text)] }));
  } catch (error) {
    console.error(`Error fetching transcript for URL: ${url}`, error);
    return []; // Return an empty array to avoid breaking the document structure
  }
}

// Function to process all sheets asynchronously
async function transcriptUrlsFromAllSheets(filePath) {
  const workbook = xlsx.readFile(filePath);
  
  for (const sheetName of workbook.SheetNames) {
    console.log(`Processing sheet: ${sheetName}`);
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    console.log('data', data);
    const doc = new Document({
        sections: [{
          properties: {},
          children: [
            // You can add Paragraphs or other elements here
          ],
        }],
      });

    let paragraphs = [];

    // Process each row asynchronously
    for (const row of data) {
      if (row[0]) {
        const httpsIndex = row[0].indexOf("https");
        if (httpsIndex !== -1) {
          const url = row[0].substring(httpsIndex); // Assuming the entire cell is the URL
          console.log('url', url);
          const transcriptParagraphs = await processUrlAndGetTranscript(url);
          paragraphs = paragraphs.concat(transcriptParagraphs);
        }
      }
    }

    // Add fetched paragraphs to the document
    if (paragraphs.length > 0) {
      doc.addSection({ properties: {}, children: paragraphs });
    }

    // Pack and save the document
    try {
      const buffer = await Packer.toBuffer(doc);
      fs.writeFileSync(`${sheetName}.docx`, buffer);
      console.log(`Document saved: ${sheetName}.docx`);
    } catch (error) {
      console.error(`Error saving document for sheet: ${sheetName}`, error);
    }
  }
}

transcriptUrlsFromAllSheets("./cm_videos.xlsx");
