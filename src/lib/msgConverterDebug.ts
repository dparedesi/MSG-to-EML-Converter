import * as MsgReaderModule from '@kenjiuno/msgreader';
import { HTMLFormatter } from './htmlFormatter';

export interface ConversionLog {
  message: string;
  timestamp: Date;
  level: 'info' | 'warn' | 'error';
}

export interface ConversionResult {
  success: boolean;
  emlData?: Uint8Array;
  filename?: string;
  logs: ConversionLog[];
  error?: string;
  debugOutput?: string; // Add debug output
}

// Correct interface based on actual @kenjiuno/msgreader structure
interface MSGData {
  // Basic properties
  subject?: string;
  body?: string;
  bodyHTML?: string;
  
  // Sender information - correct property names from analysis
  senderName?: string;
  senderSmtpAddress?: string;
  senderEmail?: string;
  
  // Date information
  creationTime?: string | Date;
  lastModificationTime?: string | Date;
  clientSubmitTime?: string | Date;
  messageDeliveryTime?: string | Date;
  
  // Message identifiers
  messageId?: string;
  
  // Recipients - correct structure from analysis
  recipients?: Array<{
    name?: string;
    email?: string;
    smtpAddress?: string;
    recipType?: string; // "to", "cc", "bcc" as strings, not numbers!
  }>;
  
  // Attachments - correct structure from analysis
  attachments?: Array<{
    name?: string;
    fileName?: string;
    fileNameShort?: string;
    data?: Uint8Array | ArrayBuffer | string | Buffer;
    contentLength?: number;
    extension?: string;
    msg?: MSGData; // For nested MSG files
  }>;
  
  // Additional metadata
  messageClass?: string;
}

export class MSGToEMLConverterDebug {
  private logs: ConversionLog[] = [];

  private log(message: string, level: 'info' | 'warn' | 'error' = 'info') {
    this.logs.push({
      message,
      timestamp: new Date(),
      level
    });
  }

  private sanitizeFilename(filename: string, defaultName: string = 'unnamed_file'): string {
    if (!filename) return defaultName;
    
    const baseName = filename.split(/[/\\]/).pop() || defaultName;
    const sanitized = baseName
      .replace(/[\\/*?:"<>|]/g, '_')
      .replace(/[\s_]+/g, '_')
      .replace(/^_+|_+$/g, '');
    
    return sanitized || defaultName;
  }

  private extractSenderInfo(msgData: MSGData): { name: string; email: string } {
    // Use correct property names from analysis
    const senderName = msgData.senderName || '';
    const senderEmail = msgData.senderSmtpAddress || msgData.senderEmail || '';

    this.log(`Extracted sender - Name: '${senderName}', Email: '${senderEmail}'`);
    return { name: senderName.trim(), email: senderEmail.trim() };
  }

  private parseDate(dateValue: string | Date | undefined): string {
    if (!dateValue) return new Date().toUTCString();
    
    try {
      let date: Date;
      
      if (dateValue instanceof Date) {
        date = dateValue;
      } else if (typeof dateValue === 'string') {
        date = new Date(dateValue);
        
        if (isNaN(date.getTime())) {
          const numValue = parseInt(dateValue, 10);
          if (!isNaN(numValue)) {
            date = new Date((numValue / 10000) - 11644473600000);
          } else {
            date = new Date();
          }
        }
      } else {
        date = new Date();
      }

      return date.toUTCString();
    } catch (error) {
      this.log(`Warning: Failed to parse date '${dateValue}': ${error}`, 'warn');
      return new Date().toUTCString();
    }
  }

  private processRecipients(msgData: MSGData): { toAddresses: string[], ccAddresses: string[], bccAddresses: string[] } {
    const toAddresses: string[] = [];
    const ccAddresses: string[] = [];
    const bccAddresses: string[] = [];

    if (!msgData.recipients || !Array.isArray(msgData.recipients)) {
      this.log('No recipients found in MSG data');
      return { toAddresses, ccAddresses, bccAddresses };
    }

    this.log(`Processing ${msgData.recipients.length} recipients`);

    msgData.recipients.forEach((recipient, index) => {
      try {
        const name = recipient.name || '';
        const email = recipient.smtpAddress || recipient.email || '';
        const type = recipient.recipType;

        this.log(`Recipient ${index}: name='${name}', email='${email}', type='${type}'`);

        if (!email && !name) {
          this.log(`Warning: Recipient ${index} has no name or email`, 'warn');
          return;
        }

        // Clean up the name - handle "LastName, FirstName" format properly
        let cleanName = name.trim();
        
        // Remove trailing semicolons but preserve commas that are part of "LastName, FirstName" format
        if (cleanName.endsWith(';')) {
          cleanName = cleanName.replace(/;+$/, '').trim();
        }
        
        // Check if this looks like "LastName, FirstName" format and preserve it
        const namePattern = /^[^,]+,\s*[^,]+$/;
        if (!namePattern.test(cleanName)) {
          // If it's not "LastName, FirstName" format, remove trailing commas too
          cleanName = cleanName.replace(/[,]+$/, '').trim();
        }
        
        let formattedAddress = '';
        if (email && email.includes('@')) {
          formattedAddress = cleanName ? `"${cleanName}" <${email}>` : email;
        } else if (cleanName) {
          formattedAddress = cleanName;
          this.log(`Warning: Recipient '${cleanName}' has no valid email address`, 'warn');
        }

        if (formattedAddress) {
          // recipType values from analysis: "to", "cc" (strings, not numbers)
          if (type === 'to') {
            toAddresses.push(formattedAddress);
            this.log(`Added TO: ${formattedAddress}`);
          } else if (type === 'cc') {
            ccAddresses.push(formattedAddress);
            this.log(`Added CC: ${formattedAddress}`);
          } else if (type === 'bcc') {
            bccAddresses.push(formattedAddress);
            this.log(`Added BCC: ${formattedAddress}`);
          } else {
            toAddresses.push(formattedAddress);
            this.log(`Added TO (default): ${formattedAddress} (unknown type: ${type})`);
          }
        }
      } catch (error) {
        this.log(`Error processing recipient ${index}: ${error}`, 'error');
      }
    });

    this.log(`Final recipients - TO: ${toAddresses.length}, CC: ${ccAddresses.length}, BCC: ${bccAddresses.length}`);
    return { toAddresses, ccAddresses, bccAddresses };
  }

  private hasStructuredContent(body: string): boolean {
    if (!body) return false;
    
    // Look for patterns that indicate structured content
    const structuredPatterns = [
      /^\s*[*•-]\s/m,           // Bullet points
      /^\s*\d+\.\s/m,           // Numbered lists
      /<mailto:[^>]+>/,         // Email links
      /https?:\/\/[^\s]+/,      // Web links
      /^\s*[A-Z\s]{3,}:\s/m,    // Headers like "Meeting Title:"
      /DO NOT FORWARD/i,        // Warning messages
      /Contacts|Invitees|Required:|Optional:/  // Section headers
    ];
    
    return structuredPatterns.some(pattern => pattern.test(body));
  }

  // Create EML manually using RFC 2822 format instead of relying on emailjs-mime-builder
  private buildEMLManually(msgData: MSGData): string {
    const lines: string[] = [];
    
    // Add headers
    if (msgData.subject) {
      lines.push(`Subject: ${msgData.subject}`);
    }

    // Handle sender
    const sender = this.extractSenderInfo(msgData);
    if (sender.email || sender.name) {
      const fromHeader = sender.email ? 
        (sender.name ? `"${sender.name}" <${sender.email}>` : sender.email) :
        sender.name;
      lines.push(`From: ${fromHeader}`);
    }

    // Handle recipients
    const recipients = this.processRecipients(msgData);
    if (recipients.toAddresses.length > 0) {
      lines.push(`To: ${recipients.toAddresses.join(', ')}`);
    }
    if (recipients.ccAddresses.length > 0) {
      lines.push(`Cc: ${recipients.ccAddresses.join(', ')}`);
    }
    if (recipients.bccAddresses.length > 0) {
      lines.push(`Bcc: ${recipients.bccAddresses.join(', ')}`);
    }

    // Handle date
    const dateValue = msgData.creationTime || msgData.messageDeliveryTime;
    if (dateValue) {
      lines.push(`Date: ${this.parseDate(dateValue)}`);
    }

    // Handle Message-ID
    const messageId = msgData.messageId;
    if (messageId) {
      lines.push(`Message-ID: ${messageId}`);
    }

    // Add MIME headers
    lines.push('MIME-Version: 1.0');
    
    // Enhanced HTML processing - convert structured text to HTML if no HTML exists
    let processedHTML = msgData.bodyHTML;
    if (!processedHTML && msgData.body && this.hasStructuredContent(msgData.body)) {
      this.log('Detected structured content in plain text - converting to HTML');
      processedHTML = HTMLFormatter.convertTextToHTML(msgData.body);
      this.log(`Generated HTML from plain text (${processedHTML.length} chars)`);
    }
    
    const hasHtmlAndText = processedHTML && msgData.body;
    const hasAttachments = msgData.attachments && msgData.attachments.length > 0;
    
    if (hasAttachments) {
      lines.push('Content-Type: multipart/mixed; boundary="boundary-main"');
      lines.push('');
      lines.push('--boundary-main');
      
      if (hasHtmlAndText) {
        lines.push('Content-Type: multipart/alternative; boundary="boundary-alt"');
        lines.push('');
        lines.push('--boundary-alt');
        lines.push('Content-Type: text/plain; charset=utf-8');
        lines.push('');
        lines.push(msgData.body || '');
        lines.push('');
        lines.push('--boundary-alt');
        lines.push('Content-Type: text/html; charset=utf-8');
        lines.push('');
        lines.push(processedHTML || '');
        lines.push('');
        lines.push('--boundary-alt--');
      } else if (processedHTML) {
        lines.push('Content-Type: text/html; charset=utf-8');
        lines.push('');
        lines.push(processedHTML);
      } else {
        lines.push('Content-Type: text/plain; charset=utf-8');
        lines.push('');
        lines.push(msgData.body || '');
      }
      
      // Add attachments (simplified for debug)
      msgData.attachments?.forEach((attachment, index) => {
        lines.push('');
        lines.push('--boundary-main');
        lines.push(`Content-Type: application/octet-stream`);
        lines.push(`Content-Disposition: attachment; filename="${attachment.fileName || attachment.name || `attachment_${index}`}"`);
        lines.push('Content-Transfer-Encoding: base64');
        lines.push('');
        lines.push('[ATTACHMENT DATA PLACEHOLDER]');
      });
      
      lines.push('');
      lines.push('--boundary-main--');
    } else if (hasHtmlAndText) {
      lines.push('Content-Type: multipart/alternative; boundary="boundary-alt"');
      lines.push('');
      lines.push('--boundary-alt');
      lines.push('Content-Type: text/plain; charset=utf-8');
      lines.push('');
      lines.push(msgData.body || '');
      lines.push('');
      lines.push('--boundary-alt');
      lines.push('Content-Type: text/html; charset=utf-8');
      lines.push('');
      lines.push(processedHTML || '');
      lines.push('');
      lines.push('--boundary-alt--');
    } else if (processedHTML) {
      lines.push('Content-Type: text/html; charset=utf-8');
      lines.push('');
      lines.push(processedHTML);
    } else {
      lines.push('Content-Type: text/plain; charset=utf-8');
      lines.push('');
      lines.push(msgData.body || '');
    }

    return lines.join('\r\n');
  }

  public async convertMSGToEML(msgFile: File): Promise<ConversionResult> {
    this.logs = []; // Reset logs
    
    try {
      this.log(`Starting conversion of MSG file: '${msgFile.name}'`);
      
      // Read the MSG file
      const arrayBuffer = await msgFile.arrayBuffer();
      this.log(`MSG file read: ${arrayBuffer.byteLength} bytes`);
      
      const msgReader = new (MsgReaderModule as { default: new (buffer: ArrayBuffer) => { getFileData: () => unknown } }).default(arrayBuffer);
      const msgData = msgReader.getFileData() as MSGData;
      
      if (!msgData) {
        throw new Error('Failed to parse MSG file data - getFileData() returned null/undefined');
      }

      this.log('MSG file parsed successfully');
      this.log(`Raw MSG data keys: ${Object.keys(msgData).join(', ')}`);
      
      // Log key fields that exist
      this.log(`senderName: '${msgData.senderName || 'N/A'}'`);
      this.log(`senderSmtpAddress: '${msgData.senderSmtpAddress || 'N/A'}'`);
      this.log(`subject: '${msgData.subject || 'N/A'}'`);
      this.log(`recipients count: ${msgData.recipients?.length || 0}`);
      
      // Log the ENTIRE raw MSG data structure for debugging
      this.log('=== FULL MSG DATA DUMP ===');
      this.log(JSON.stringify(msgData, null, 2));
      this.log('=== END MSG DATA DUMP ===');
      
      // Log each recipient in detail
      if (msgData.recipients) {
        msgData.recipients.forEach((recipient, index) => {
          this.log(`=== RECIPIENT ${index} DETAIL ===`);
          this.log(`Raw recipient object: ${JSON.stringify(recipient, null, 2)}`);
        });
      }
      
      // Build EML manually for debugging
      const emlContent = this.buildEMLManually(msgData);
      
      this.log('Generated EML headers (first 500 chars):');
      this.log(emlContent.substring(0, 500));
      
      if (!emlContent || emlContent.length === 0) {
        throw new Error('Generated EML content is empty');
      }
      
      // Convert string to Uint8Array
      const encoder = new TextEncoder();
      const emlData = encoder.encode(emlContent);
      
      const originalName = msgFile.name.replace(/\.msg$/i, '');
      const filename = this.sanitizeFilename(`${originalName}_debug.eml`);
      
      this.log(`Successfully converted MSG to EML. Output size: ${emlData.length} bytes`);
      this.log(`Suggested filename: ${filename}`);
      
      return {
        success: true,
        emlData,
        filename,
        logs: [...this.logs],
        debugOutput: emlContent
      };
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      this.log(`Conversion failed: ${errorMessage}`, 'error');
      
      if (error instanceof Error && error.stack) {
        this.log(`Error stack: ${error.stack}`, 'error');
      }
      
      return {
        success: false,
        error: errorMessage,
        logs: [...this.logs]
      };
    }
  }
}
