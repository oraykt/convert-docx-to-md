import mammoth from 'mammoth';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';

const turndownService = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced'
});
turndownService.use(gfm);

// Function to convert DOCX to Markdown in the browser
export async function convertDocxToMarkdown(file) {
    try {
        // Convert DOCX to HTML with mammoth
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        let htmlContent = result.value;

        // Replace base64 images with empty Markdown image placeholders
        htmlContent = replaceImagesWithPlaceholders(htmlContent);

        // Convert tables to Markdown
        htmlContent = convertHtmlTableToMarkdown(htmlContent);

        // Convert remaining HTML content to Markdown
        let markdownContent = turndownService.turndown(htmlContent);

        // Post-process markdown to ensure proper table formatting
        markdownContent = fixMarkdownTableFormatting(markdownContent);

        return markdownContent;
    } catch (error) {
        console.error('Error converting DOCX to Markdown:', error);
        return null;
    }
}

// Function to replace base64 images with "![IMAGE]()"
function replaceImagesWithPlaceholders(htmlContent) {
    const base64ImageRegex = /<img[^>]+src="data:image\/(png|jpeg|jpg);base64,([^"]+)"[^>]*>/g;
    let imageIndex = 1;

    // Replace each base64 image with a placeholder like ![IMAGE]() for user to fill in later
    return htmlContent.replace(base64ImageRegex, (match, extension, base64Data) => {
        const imagePlaceholder = `![IMAGE ${imageIndex}]()`;
        imageIndex++;
        return imagePlaceholder;
    });
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
