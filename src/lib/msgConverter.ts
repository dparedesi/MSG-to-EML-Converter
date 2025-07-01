import MsgReader from '@kenjiuno/msgreader';
import MimeBuilder from 'emailjs-mime-builder';
import { EnhancedAttachmentHandler } from './enhancedAttachmentHandler';

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
}

// Enhanced interface for MSG data based on actual @kenjiuno/msgreader structure
interface MSGData {
  // Basic properties
  subject?: string;
  body?: string;
  bodyHTML?: string;
  
  // Sender information - multiple possible sources
  senderName?: string;
  senderEmailAddress?: string;
  sentByName?: string;
  sentByEmailAddress?: string;
  displayFrom?: string;
  from?: string;
  
  // Date information - multiple formats
  creationTime?: string | Date;
  deliveryTime?: string | Date;
  lastModificationTime?: string | Date;
  messageDeliveryTime?: string | Date;
  
  // Message identifiers
  messageId?: string;
  internetMessageId?: string;
  conversationId?: string;
  
  // Recipients
  recipients?: Array<{
    name?: string;
    email?: string;
    recipientType?: number;
    displayName?: string;
    emailAddress?: string;
  }>;
  
  // Attachments
  attachments?: Array<{
    name?: string;
    longFilename?: string;
    shortFilename?: string;
    filename?: string;
    contentType?: string;
    data?: Uint8Array | ArrayBuffer | string | Buffer;
    attachMethod?: number;
    isEmbedded?: boolean;
    contentId?: string;
    msg?: MSGData; // For nested MSG files
  }>;
  
  // Additional metadata
  importance?: number;
  priority?: number;
  sensitivity?: number;
  categories?: string[];
  messageClass?: string;
}

export class MSGToEMLConverter {
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

  private guessMimeType(filename: string): [string, string] {
    const ext = filename.toLowerCase().split('.').pop();
    const mimeTypes: Record<string, [string, string]> = {
      'txt': ['text', 'plain'],
      'html': ['text', 'html'],
      'htm': ['text', 'html'],
      'xml': ['text', 'xml'],
      'css': ['text', 'css'],
      'js': ['text', 'javascript'],
      'json': ['application', 'json'],
      'pdf': ['application', 'pdf'],
      'doc': ['application', 'msword'],
      'docx': ['application', 'vnd.openxmlformats-officedocument.wordprocessingml.document'],
      'xls': ['application', 'vnd.ms-excel'],
      'xlsx': ['application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
      'ppt': ['application', 'vnd.ms-powerpoint'],
      'pptx': ['application', 'vnd.openxmlformats-officedocument.presentationml.presentation'],
      'rtf': ['application', 'rtf'],
      'zip': ['application', 'zip'],
      'rar': ['application', 'x-rar-compressed'],
      '7z': ['application', 'x-7z-compressed'],
      'tar': ['application', 'x-tar'],
      'gz': ['application', 'gzip'],
      'jpg': ['image', 'jpeg'],
      'jpeg': ['image', 'jpeg'],
      'png': ['image', 'png'],
      'gif': ['image', 'gif'],
      'bmp': ['image', 'bmp'],
      'tiff': ['image', 'tiff'],
      'svg': ['image', 'svg+xml'],
      'ico': ['image', 'x-icon'],
      'mp3': ['audio', 'mpeg'],
      'wav': ['audio', 'wav'],
      'ogg': ['audio', 'ogg'],
      'mp4': ['video', 'mp4'],
      'avi': ['video', 'x-msvideo'],
      'mov': ['video', 'quicktime'],
      'wmv': ['video', 'x-ms-wmv']
    };

    return mimeTypes[ext || ''] || ['application', 'octet-stream'];
  }

  private extractSenderInfo(msgData: MSGData): { name: string; email: string } {
    // Try multiple sender fields in order of preference
    let senderName = '';
    let senderEmail = '';

    // Try various sender name fields
    senderName = msgData.senderName || msgData.sentByName || msgData.displayFrom || '';
    
    // Try various sender email fields
    senderEmail = msgData.senderEmailAddress || msgData.sentByEmailAddress || '';

    // Parse from combined "from" field if available
    if (!senderEmail && msgData.from) {
      const fromMatch = msgData.from.match(/<([^>]+)>/);
      if (fromMatch) {
        senderEmail = fromMatch[1];
        const nameMatch = msgData.from.match(/^([^<]+)</);
        if (nameMatch && !senderName) {
          senderName = nameMatch[1].trim().replace(/^["']|["']$/g, '');
        }
      } else if (msgData.from.includes('@')) {
        senderEmail = msgData.from;
      } else {
        senderName = msgData.from;
      }
    }

    return { name: senderName.trim(), email: senderEmail.trim() };
  }

  private parseDate(dateValue: string | Date | undefined): string {
    if (!dateValue) return new Date().toUTCString();
    
    try {
      let date: Date;
      
      if (dateValue instanceof Date) {
        date = dateValue;
      } else if (typeof dateValue === 'string') {
        // Handle various date formats
        date = new Date(dateValue);
        
        // If parsing failed, try to handle Windows FILETIME format
        if (isNaN(date.getTime())) {
          // FILETIME is 64-bit value representing 100-nanosecond intervals since January 1, 1601
          const numValue = parseInt(dateValue, 10);
          if (!isNaN(numValue)) {
            // Convert FILETIME to JavaScript Date
            date = new Date((numValue / 10000) - 11644473600000);
          } else {
            // Fallback to current date
            date = new Date();
          }
        }
      } else {
        date = new Date();
      }

      // Format as RFC 2822 date
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
      return { toAddresses, ccAddresses, bccAddresses };
    }

    msgData.recipients.forEach((recipient, index) => {
      try {
        const name = recipient.name || recipient.displayName || '';
        const email = recipient.email || recipient.emailAddress || '';
        const type = recipient.recipientType;

        if (!email && !name) {
          this.log(`Warning: Recipient ${index} has no name or email`, 'warn');
          return;
        }

        let formattedAddress = '';
        if (email && email.includes('@')) {
          formattedAddress = name ? `"${name}" <${email}>` : email;
        } else if (name) {
          formattedAddress = name;
          this.log(`Warning: Recipient '${name}' has no valid email address`, 'warn');
        }

        if (formattedAddress) {
          // MAPI recipient types: 1=TO, 2=CC, 3=BCC
          if (type === 1 || String(type) === '1') {
            toAddresses.push(formattedAddress);
          } else if (type === 2 || String(type) === '2') {
            ccAddresses.push(formattedAddress);
          } else if (type === 3 || String(type) === '3') {
            bccAddresses.push(formattedAddress);
          } else {
            // Default to TO if type is unclear
            toAddresses.push(formattedAddress);
            this.log(`Warning: Unknown recipient type '${type}' for '${formattedAddress}', defaulting to TO`, 'warn');
          }
        }
      } catch (error) {
        this.log(`Error processing recipient ${index}: ${error}`, 'error');
      }
    });

    return { toAddresses, ccAddresses, bccAddresses };
  }

  private async buildEMLFromMSG(msgData: MSGData, nestingLevel: number = 0, msgReader?: unknown): Promise<MimeBuilder> {
    const subject = msgData.subject || 'No Subject';
    this.log(`${'  '.repeat(nestingLevel)}Processing MSG (Subject: '${subject}')`);

    // Use enhanced attachment handler to extract attachments
    const attachmentHandler = new EnhancedAttachmentHandler();
    const attachmentResult = await attachmentHandler.extractAttachments(msgData, msgReader);
    
    // Log attachment handler results
    attachmentResult.logs.forEach(log => this.log(`${'  '.repeat(nestingLevel + 1)}${log}`));
    attachmentResult.errors.forEach(error => this.log(`${'  '.repeat(nestingLevel + 1)}${error}`, 'error'));

    const hasAttachments = attachmentResult.attachments.length > 0;
    const hasHtmlAndText = msgData.bodyHTML && msgData.body;
    
    let rootBuilder: MimeBuilder;
    
    if (hasAttachments) {
      rootBuilder = new MimeBuilder('multipart/mixed');
    } else if (hasHtmlAndText) {
      rootBuilder = new MimeBuilder('multipart/alternative');
    } else if (msgData.bodyHTML) {
      rootBuilder = new MimeBuilder('text/html; charset=utf-8');
    } else {
      rootBuilder = new MimeBuilder('text/plain; charset=utf-8');
    }

    // Set standard headers
    rootBuilder.setHeader('MIME-Version', '1.0');
    
    if (msgData.subject) {
      // Handle Unicode subjects properly
      rootBuilder.setHeader('Subject', msgData.subject);
    }

    // Handle sender
    const sender = this.extractSenderInfo(msgData);
    if (sender.email || sender.name) {
      const fromHeader = sender.email ? 
        (sender.name ? `"${sender.name}" <${sender.email}>` : sender.email) :
        sender.name;
      rootBuilder.setHeader('From', fromHeader);
    }

    // Handle recipients
    const recipients = this.processRecipients(msgData);
    if (recipients.toAddresses.length > 0) {
      rootBuilder.setHeader('To', recipients.toAddresses.join(', '));
    }
    if (recipients.ccAddresses.length > 0) {
      rootBuilder.setHeader('Cc', recipients.ccAddresses.join(', '));
    }
    if (recipients.bccAddresses.length > 0) {
      rootBuilder.setHeader('Bcc', recipients.bccAddresses.join(', '));
    }

    // Handle date - try multiple date fields
    const dateValue = msgData.creationTime || msgData.deliveryTime || msgData.messageDeliveryTime;
    if (dateValue) {
      rootBuilder.setHeader('Date', this.parseDate(dateValue));
    }

    // Handle Message-ID
    const messageId = msgData.internetMessageId || msgData.messageId;
    if (messageId) {
      rootBuilder.setHeader('Message-ID', messageId);
    }

    // Handle priority/importance
    if (msgData.importance !== undefined) {
      const priority = msgData.importance === 2 ? 'high' : msgData.importance === 0 ? 'low' : 'normal';
      rootBuilder.setHeader('X-Priority', priority);
    }

    // Handle body content
    if (!hasAttachments && !hasHtmlAndText) {
      // Simple single-part message
      const content = msgData.bodyHTML || msgData.body || '';
      rootBuilder.setContent(content);
    } else {
      // Complex multipart message
      let bodyBuilder: MimeBuilder;
      
      if (hasHtmlAndText) {
        // Create multipart/alternative for HTML and text
        bodyBuilder = new MimeBuilder('multipart/alternative');
        bodyBuilder.createChild('text/plain; charset=utf-8').setContent(msgData.body || '');
        bodyBuilder.createChild('text/html; charset=utf-8').setContent(msgData.bodyHTML || '');
      } else if (msgData.bodyHTML) {
        bodyBuilder = new MimeBuilder('text/html; charset=utf-8').setContent(msgData.bodyHTML);
      } else {
        bodyBuilder = new MimeBuilder('text/plain; charset=utf-8').setContent(msgData.body || '');
      }

      if (hasAttachments) {
        rootBuilder.appendChild(bodyBuilder);
      } else {
        rootBuilder = bodyBuilder;
      }
    }

    // Handle attachments using enhanced handler
    if (hasAttachments) {
      const enhancedHandler = new EnhancedAttachmentHandler();
      await enhancedHandler.buildAttachmentsForEML(
        attachmentResult.attachments, 
        rootBuilder, 
        async (nestedMsgData) => await this.buildEMLFromMSG(nestedMsgData as MSGData, nestingLevel + 1)
      );
      
      // Log additional attachment handler results
      enhancedHandler.getLogs().forEach(log => this.log(`${'  '.repeat(nestingLevel + 1)}${log}`));
      enhancedHandler.getErrors().forEach(error => this.log(`${'  '.repeat(nestingLevel + 1)}${error}`, 'error'));
    }

    return rootBuilder;
  }


  public async convertMSGToEML(msgFile: File): Promise<ConversionResult> {
    this.logs = []; // Reset logs
    
    try {
      this.log(`Starting conversion of MSG file: '${msgFile.name}'`);
      
      // Read the MSG file
      const arrayBuffer = await msgFile.arrayBuffer();
      this.log(`MSG file read: ${arrayBuffer.byteLength} bytes`);
      
      const msgReader = new MsgReader(arrayBuffer);
      const msgData = msgReader.getFileData() as MSGData;
      
      if (!msgData) {
        throw new Error('Failed to parse MSG file data - getFileData() returned null/undefined');
      }

      this.log('MSG file parsed successfully');
      this.log(`Found subject: ${msgData.subject || 'N/A'}`);
      this.log(`Found ${msgData.recipients?.length || 0} recipients`);
      this.log(`Found ${msgData.attachments?.length || 0} attachments`);
      this.log(`Body text length: ${msgData.body?.length || 0} chars`);
      this.log(`Body HTML length: ${msgData.bodyHTML?.length || 0} chars`);
      
      // Build EML from MSG data
      const emlBuilder = await this.buildEMLFromMSG(msgData, 0, msgReader);
      const emlContent = emlBuilder.build();
      
      if (!emlContent || emlContent.length === 0) {
        throw new Error('Generated EML content is empty');
      }
      
      // Convert string to Uint8Array
      const encoder = new TextEncoder();
      const emlData = encoder.encode(emlContent);
      
      const originalName = msgFile.name.replace(/\.msg$/i, '');
      const filename = this.sanitizeFilename(`${originalName}.eml`);
      
      this.log(`Successfully converted MSG to EML. Output size: ${emlData.length} bytes`);
      this.log(`Suggested filename: ${filename}`);
      
      return {
        success: true,
        emlData,
        filename,
        logs: [...this.logs]
      };
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      this.log(`Conversion failed: ${errorMessage}`, 'error');
      
      // Log the error stack for debugging
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
