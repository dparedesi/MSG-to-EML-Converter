export class HTMLFormatter {
  
  /**
   * Convert plain text with structured patterns to HTML
   */
  public static convertTextToHTML(plainText: string): string {
    if (!plainText || plainText.trim().length === 0) {
      return plainText;
    }

    const lines = plainText.split('\n');
    const htmlLines: string[] = [];
    let inList = false;
    let listType: 'ul' | 'ol' | null = null;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmedLine = line.trim();

      // Skip empty lines but preserve spacing in HTML
      if (trimmedLine === '') {
        if (inList) {
          htmlLines.push(`</${listType}>`);
          inList = false;
          listType = null;
        }
        htmlLines.push('<br>');
        continue;
      }

      // Detect bullet points (tab + asterisk or bullet)
      if (this.isBulletPoint(line)) {
        if (!inList || listType !== 'ul') {
          if (inList) htmlLines.push(`</${listType}>`);
          htmlLines.push('<ul>');
          inList = true;
          listType = 'ul';
        }
        const content = this.formatLineContent(line.replace(/^\s*[*•-]\s*/, ''));
        htmlLines.push(`<li>${content}</li>`);
        continue;
      }

      // Detect numbered lists
      if (this.isNumberedItem(line)) {
        if (!inList || listType !== 'ol') {
          if (inList) htmlLines.push(`</${listType}>`);
          htmlLines.push('<ol>');
          inList = true;
          listType = 'ol';
        }
        const content = this.formatLineContent(line.replace(/^\s*\d+\.\s*/, ''));
        htmlLines.push(`<li>${content}</li>`);
        continue;
      }

      // Close any open list
      if (inList) {
        htmlLines.push(`</${listType}>`);
        inList = false;
        listType = null;
      }

      // Detect headings (lines in ALL CAPS or with specific patterns)
      if (this.isHeading(trimmedLine)) {
        htmlLines.push(`<h3><strong>${this.formatLineContent(trimmedLine)}</strong></h3>`);
        continue;
      }

      // Detect sub-headings
      if (this.isSubHeading(trimmedLine)) {
        htmlLines.push(`<h4>${this.formatLineContent(trimmedLine)}</h4>`);
        continue;
      }

      // Regular paragraph
      const formattedContent = this.formatLineContent(trimmedLine);
      if (formattedContent.length > 0) {
        htmlLines.push(`<p>${formattedContent}</p>`);
      }
    }

    // Close any remaining open list
    if (inList) {
      htmlLines.push(`</${listType}>`);
    }

    return this.wrapInHTML(htmlLines.join('\n'));
  }

  private static isBulletPoint(line: string): boolean {
    return /^\s*[*•-]\s/.test(line);
  }

  private static isNumberedItem(line: string): boolean {
    return /^\s*\d+\.\s/.test(line);
  }

  private static isHeading(line: string): boolean {
    // All caps lines, or lines with specific patterns
    return /^[A-Z\s\*\-]{3,}$/.test(line) || 
           line.includes('DO NOT FORWARD') ||
           line.includes('Meeting Title:') ||
           line.includes('Purpose:');
  }

  private static isSubHeading(line: string): boolean {
    const subHeadingPatterns = [
      'Contacts', 'Invitees', 'Required:', 'Optional:', 'PIN', 'Chime Link:',
      'Context', 'EA Meeting Contact:', 'L8 Owner:', 'Doc Owner:', 'Note Taker:'
    ];
    return subHeadingPatterns.some(pattern => line.includes(pattern));
  }

  private static formatLineContent(content: string): string {
    if (!content) return '';

    // First escape HTML entities to prevent issues
    content = content
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    // Convert email links (after escaping, so we look for &lt;mailto:)
    content = content.replace(
      /&lt;mailto:([^&]+)&gt;/g, 
      '<a href="mailto:$1">$1</a>'
    );

    // Convert web links (simple URLs without < >)
    content = content.replace(
      /https?:\/\/[^\s]+/g,
      '<a href="$&">$&</a>'
    );

    // Bold text patterns (text between asterisks) - but not if already inside HTML tags
    content = content.replace(
      /\*([^*<>]+)\*/g,
      '<strong>$1</strong>'
    );

    return content;
  }

  private static wrapInHTML(content: string): string {
    return `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Calibri, Arial, sans-serif; font-size: 14px; line-height: 1.6; margin: 20px; }
        h3 { color: #1f497d; margin-top: 20px; margin-bottom: 10px; }
        h4 { color: #365f91; margin-top: 15px; margin-bottom: 8px; }
        ul, ol { margin: 10px 0; padding-left: 30px; }
        li { margin: 5px 0; }
        p { margin: 8px 0; }
        a { color: #0563c1; text-decoration: underline; }
        .meeting-info { background-color: #f8f9fa; padding: 15px; border-left: 4px solid #1f497d; margin: 15px 0; }
    </style>
</head>
<body>
${content}
</body>
</html>`;
  }
}
