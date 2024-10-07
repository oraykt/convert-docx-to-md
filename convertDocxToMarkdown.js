const fs = require('fs');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const gfm = require('turndown-plugin-gfm').gfm;
const path = require('path');

// Initialize Turndown Service and enable GFM plugin
const turndownService = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced'
});
turndownService.use(gfm);  // Enable GitHub Flavored Markdown (GFM)

// Function to ensure a directory exists
function ensureDirSync(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        console.log(`Directory created: ${dirPath}`);
    }
}

// Function to extract base64 images from HTML and save them to files
async function extractAndSaveImages(htmlContent, imagesDir, outputMarkdown) {
    const base64ImageRegex = /<img[^>]+src="data:image\/(png|jpeg|jpg);base64,([^"]+)"[^>]*>/g;
    let match;
    let imageIndex = 1;

    // Process each base64 image
    while ((match = base64ImageRegex.exec(htmlContent)) !== null) {
        const [imageTag, extension, base64Data] = match;

        // Create the image file path
        const imageName = `image_${imageIndex}.${extension}`;
        const imagePath = path.join(imagesDir, imageName);

        // Convert base64 to binary data and save the image
        const imageBuffer = Buffer.from(base64Data, 'base64');
        await fs.promises.writeFile(imagePath, imageBuffer);
        console.log(`Image saved: ${imagePath}`);

        // Replace the base64 image in the HTML with the Markdown reference
        const relativeImagePath = path.relative(path.dirname(outputMarkdown), imagePath);
        htmlContent = htmlContent.replace(imageTag, `![Image](${relativeImagePath})`);

        imageIndex++;
    }

    return htmlContent;
}

// Function to convert HTML tables to Markdown
function convertHtmlTableToMarkdown(html) {
    return html.replace(/<table>(.*?)<\/table>/gs, (match) => {
        let table = match
            .replace(/<thead>/g, '')
            .replace(/<\/thead>/g, '')
            .replace(/<tbody>/g, '')
            .replace(/<\/tbody>/g, '');

        const rows = table.match(/<tr>(.*?)<\/tr>/gs);
        if (!rows || rows.length === 0) return ''; // Return empty string if no rows are found

        let markdownTable = '';

        // Process each row
        rows.forEach((row, rowIndex) => {
            const columns = row.match(/<t[dh]>(.*?)<\/t[dh]>/gs);
            if (!columns) return; // Skip if no columns are found

            // Extract the text content from each column and add padding
            const columnText = columns.map(col => {
                // Add space before and after text in each cell
                return `${col.replace(/<\/?[^>]+(>|$)/g, '').trim()}`;
            });

            // Join columns with ' | ' and add new line
            markdownTable += '| ' + columnText.join(' | ') + ' |\r\n';

            // Add separator after the header row
            if (rowIndex === 0) {
                markdownTable += '| ' + columnText.map(() => '---').join(' | ') + ' |\r\n';
            }
        });

        // Ensure we add a newline after the table
        return `\r\n${markdownTable}\r\n`;
    });
}



// Function to fix table rows that are merged into one line by checking for '| |'
function fixMarkdownTableFormatting(markdownContent) {
    // Look for occurrences of "| |" and replace them with newlines between rows
    return markdownContent.replace(/\|\s+\|\s+/g, '|\n| ');
}


// Function to convert DOCX to Markdown and handle images & tables
async function convertDocxToMarkdown(inputFilePath, outputFilePath, imagesDir) {
    try {
        console.log('Starting DOCX to Markdown conversion...');
        ensureDirSync(imagesDir);

        // Read the DOCX file
        const data = await fs.promises.readFile(inputFilePath);
        console.log('DOCX file successfully read.');

        // Convert DOCX to HTML with mammoth
        const result = await mammoth.convertToHtml({ buffer: data });
        console.log('Mammoth conversion completed.');

        // Extract base64 images and save them as files
        let htmlContent = result.value;
        htmlContent = await extractAndSaveImages(htmlContent, imagesDir, outputFilePath);

        // Convert HTML tables to Markdown
        htmlContent = convertHtmlTableToMarkdown(htmlContent);

        // Convert remaining HTML content to Markdown
        let markdownContent = turndownService.turndown(htmlContent);

        // Post-process markdown to ensure proper table formatting
        markdownContent = fixMarkdownTableFormatting(markdownContent);

        // Save the Markdown content to a file with correct newlines
        await fs.promises.writeFile(outputFilePath, markdownContent, { encoding: 'utf-8', flag: 'w' });
        console.log(`Conversion successful! Markdown saved to ${outputFilePath}`);
    } catch (error) {
        console.error('Error converting DOCX to Markdown:', error);
    }
}


// Example usage
const inputDocx = __dirname + '/Test.docx';             // Path to the DOCX file
const outputMarkdown = __dirname + '/Test.md';           // Path to the output Markdown file
const imagesDirectory = __dirname + '/images';           // Directory to save images

convertDocxToMarkdown(inputDocx, outputMarkdown, imagesDirectory);
