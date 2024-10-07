import mammoth from 'mammoth';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import JSZip from 'jszip';

const turndownService = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced'
});
turndownService.use(gfm);

// Function to convert DOCX to Markdown and create a ZIP with Markdown and images
export async function convertDocxToZip(file) {
    try {
        const zip = new JSZip();  // Create a new JSZip instance
        const imagesFolder = zip.folder('images');  // Create a folder for images in the zip

        // Convert DOCX to HTML with mammoth
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        let htmlContent = result.value;

        // Replace base64 images with placeholders and add images to the zip
        htmlContent = await extractImagesAndReplaceWithPlaceholders(htmlContent, imagesFolder);

        // Convert HTML tables to Markdown
        htmlContent = convertHtmlTableToMarkdown(htmlContent);

        // Convert remaining HTML content to Markdown
        let markdownContent = turndownService.turndown(htmlContent);

        // Post-process markdown to ensure proper table formatting
        markdownContent = fixMarkdownTableFormatting(markdownContent);

        // Add Markdown file to the zip
        zip.file('converted.md', markdownContent);

        // Generate the ZIP file
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        return zipBlob;

    } catch (error) {
        console.error('Error converting DOCX to Markdown and ZIP:', error);
        return null;
    }
}

// Function to extract base64 images from HTML, save them to the zip, and replace them with placeholders in the HTML
async function extractImagesAndReplaceWithPlaceholders(htmlContent, imagesFolder) {
    const base64ImageRegex = /<img[^>]+src="data:image\/(png|jpeg|jpg);base64,([^"]+)"[^>]*>/g;
    let match;
    let imageIndex = 1;

    // Process each base64 image
    while ((match = base64ImageRegex.exec(htmlContent)) !== null) {
        const [imageTag, extension, base64Data] = match;
        const imageName = `image_${imageIndex}.${extension}`;

        // Convert base64 string to binary data using atob() and ArrayBuffer in the browser
        const binaryData = atob(base64Data);
        const arrayBuffer = new Uint8Array(binaryData.length);
        for (let i = 0; i < binaryData.length; i++) {
            arrayBuffer[i] = binaryData.charCodeAt(i);
        }

        // Add image to the images folder in the ZIP
        imagesFolder.file(imageName, arrayBuffer);

        // Replace the base64 image in the HTML with a Markdown reference to the image
        const relativeImagePath = `images/${imageName}`;
        htmlContent = htmlContent.replace(imageTag, `<img alt="${imageName}" src="${relativeImagePath}">`);

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
            const columnText = columns.map(col => `${col.replace(/<\/?[^>]+(>|$)/g, '').trim()}`);

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
