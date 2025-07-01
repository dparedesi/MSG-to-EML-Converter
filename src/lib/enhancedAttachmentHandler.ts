import MsgReader from '@kenjiuno/msgreader';
import MimeBuilder from 'emailjs-mime-builder';

interface MSGAttachment {
  name?: string;
  longFilename?: string;
  shortFilename?: string;
  filename?: string;
  contentType?: string;
  data?: Uint8Array | ArrayBuffer | string | Buffer;
  dataId?: string | number;
  contentId?: string;
  pidContentId?: string;
  isEmbedded?: boolean;
  attachmentHidden?: boolean;
  msg?: unknown;
  innerMsgContentFields?: unknown;
  getAttachment?: () => Uint8Array | ArrayBuffer;
  [key: string]: unknown;
}

interface MSGReader {
  getFileData(): unknown;
  getAttachment(dataId: string | number): Uint8Array | ArrayBuffer | null;
}

export interface AttachmentInfo {
  name: string;
  contentType: string;
  data: Uint8Array;
  isEmbedded: boolean;
  contentId?: string;
  isNestedMsg: boolean;
  nestedMsgData?: unknown;
}

export interface AttachmentExtractionResult {
  attachments: AttachmentInfo[];
  logs: string[];
  errors: string[];
}

export class EnhancedAttachmentHandler {
  private logs: string[] = [];
  private errors: string[] = [];

  private log(message: string, isError: boolean = false) {
    if (isError) {
      this.errors.push(message);
    } else {
      this.logs.push(message);
    }
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

  private guessMimeType(filename: string): string {
    const ext = filename.toLowerCase().split('.').pop();
    const mimeTypes: Record<string, string> = {
      'txt': 'text/plain',
      'html': 'text/html',
      'htm': 'text/html',
      'xml': 'text/xml',
      'css': 'text/css',
      'js': 'text/javascript',
      'json': 'application/json',
      'pdf': 'application/pdf',
      'doc': 'application/msword',
      'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'xls': 'application/vnd.ms-excel',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'ppt': 'application/vnd.ms-powerpoint',
      'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'rtf': 'application/rtf',
      'zip': 'application/zip',
      'rar': 'application/x-rar-compressed',
      '7z': 'application/x-7z-compressed',
      'tar': 'application/x-tar',
      'gz': 'application/gzip',
      'jpg': 'image/jpeg',
      'jpeg': 'image/jpeg',
      'png': 'image/png',
      'gif': 'image/gif',
      'bmp': 'image/bmp',
      'tiff': 'image/tiff',
      'svg': 'image/svg+xml',
      'ico': 'image/x-icon',
      'mp3': 'audio/mpeg',
      'wav': 'audio/wav',
      'ogg': 'audio/ogg',
      'mp4': 'video/mp4',
      'avi': 'video/x-msvideo',
      'mov': 'video/quicktime',
      'wmv': 'video/x-ms-wmv',
      'msg': 'application/vnd.ms-outlook'
    };

    return mimeTypes[ext || ''] || 'application/octet-stream';
  }

  private extractAttachmentData(attachment: MSGAttachment, msgReader?: MSGReader): Uint8Array | null {
    // First try to get data using msgReader.getAttachment(dataId) for regular attachments
    if (attachment.dataId && msgReader) {
      try {
        const attachmentData = msgReader.getAttachment(attachment.dataId);
        if (attachmentData) {
          if (attachmentData instanceof Uint8Array) {
            return attachmentData;
          } else if (attachmentData instanceof ArrayBuffer) {
            return new Uint8Array(attachmentData);
          } else if (Buffer.isBuffer && Buffer.isBuffer(attachmentData)) {
            return new Uint8Array(attachmentData);
          }
        }
      } catch (error) {
        this.log(`Failed to get attachment data via msgReader.getAttachment(${attachment.dataId}): ${error}`, true);
      }
    }

    // Try multiple ways to get attachment data
    const possibleDataFields = ['data', 'content', 'body', 'dataBody', 'attachData', 'innerMsgContent'];
    
    for (const field of possibleDataFields) {
      if (attachment[field]) {
        const data = attachment[field];
        
        if (data instanceof Uint8Array) {
          return data;
        } else if (data instanceof ArrayBuffer) {
          return new Uint8Array(data);
        } else if (Buffer.isBuffer && Buffer.isBuffer(data)) {
          return new Uint8Array(data);
        } else if (typeof data === 'string') {
          // Handle base64 encoded data
          try {
            const binaryString = atob(data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes;
          } catch {
            // If not base64, treat as text
            return new TextEncoder().encode(data);
          }
        }
      }
    }

    // Try to access data through getAttachment method if available
    if (typeof attachment.getAttachment === 'function') {
      try {
        const attachmentData = attachment.getAttachment();
        if (attachmentData && (attachmentData instanceof Uint8Array || attachmentData instanceof ArrayBuffer)) {
          return attachmentData instanceof Uint8Array ? attachmentData : new Uint8Array(attachmentData);
        }
      } catch (error) {
        this.log(`Failed to get attachment via getAttachment(): ${error}`, true);
      }
    }

    return null;
  }

  private async tryParseAsNestedMsg(data: Uint8Array): Promise<unknown> {
    try {
      const msgReader = new MsgReader(data.buffer as ArrayBuffer);
      const nestedMsgData = msgReader.getFileData();
      return nestedMsgData;
    } catch (error) {
      this.log(`Failed to parse data as nested MSG: ${error}`);
      return undefined;
    }
  }

  public async extractAttachments(msgData: unknown, msgReader?: unknown): Promise<AttachmentExtractionResult> {
    this.logs = [];
    this.errors = [];
    
    const attachments: AttachmentInfo[] = [];

    // Type guard to check if msgData has attachments
    const typedMsgData = msgData as { attachments?: unknown[] };
    if (!typedMsgData.attachments || !Array.isArray(typedMsgData.attachments)) {
      this.log('No attachments found in MSG data');
      return { attachments, logs: this.logs, errors: this.errors };
    }

    this.log(`Processing ${typedMsgData.attachments.length} attachments`);

    for (let i = 0; i < typedMsgData.attachments.length; i++) {
      const attachment = typedMsgData.attachments[i] as MSGAttachment;
      this.log(`\n--- Processing Attachment ${i + 1} ---`);

      try {
        // Get attachment name
        const rawName = attachment.name || attachment.longFilename || 
                       attachment.shortFilename || attachment.filename || 
                       `attachment_${i + 1}`;
        const filename = this.sanitizeFilename(rawName);
        
        this.log(`Attachment name: ${filename}`);

        // Check if this is already a nested MSG object (parsed by the library)
        if (attachment.msg) {
          this.log(`✓ Nested MSG object detected: ${filename}`);
          attachments.push({
            name: filename.endsWith('.eml') ? filename : `${filename}.eml`,
            contentType: 'message/rfc822',
            data: new Uint8Array(0), // Will be populated when building EML
            isEmbedded: false,
            isNestedMsg: true,
            nestedMsgData: attachment.msg
          });
          continue;
        }

        // Check if this is a nested MSG with innerMsgContentFields (new library format)
        if (attachment.innerMsgContentFields) {
          this.log(`✓ Nested MSG with innerMsgContentFields detected: ${filename}`);
          attachments.push({
            name: filename.endsWith('.eml') ? filename : `${filename.replace(/\.msg$/i, '')}.eml`,
            contentType: 'message/rfc822',
            data: new Uint8Array(0), // Will be populated when building EML
            isEmbedded: false,
            isNestedMsg: true,
            nestedMsgData: attachment.innerMsgContentFields
          });
          continue;
        }

        // Extract attachment data
        const data = this.extractAttachmentData(attachment, msgReader as MSGReader);
        
        if (!data || data.length === 0) {
          this.log(`⚠️  No data found for attachment: ${filename}`, true);
          continue;
        }

        this.log(`✓ Extracted ${data.length} bytes of data`);

        // Determine content type
        let contentType = attachment.contentType || 
                         (typeof attachment.attachMimeTag === 'string' ? attachment.attachMimeTag : '') || 
                         this.guessMimeType(filename);
        
        // Check if this might be a nested MSG file by extension or content
        const isLikelyMsg = filename.toLowerCase().endsWith('.msg') ||
                           contentType === 'application/vnd.ms-outlook' ||
                           rawName.includes('Re:') || rawName.includes('RE:') ||
                           rawName.includes('Fwd:') || rawName.includes('FW:');

        let nestedMsgData = null;
        let isNestedMsg = false;

        if (isLikelyMsg) {
          this.log(`Checking if attachment is a nested MSG file...`);
          nestedMsgData = await this.tryParseAsNestedMsg(data);
          if (nestedMsgData) {
            this.log(`✓ Successfully parsed as nested MSG!`);
            isNestedMsg = true;
            contentType = 'message/rfc822';
          } else {
            this.log(`Not a valid MSG file, treating as regular attachment`);
          }
        }

        // Determine if embedded (look for pidContentId which indicates inline images)
        const isEmbedded = attachment.isEmbedded || 
                          (attachment.contentId && attachment.contentId.length > 0) ||
                          (attachment.pidContentId && attachment.pidContentId.length > 0) ||
                          attachment.attachmentHidden === true;

        attachments.push({
          name: isNestedMsg ? (filename.endsWith('.eml') ? filename : `${filename.replace(/\.msg$/i, '')}.eml`) : filename,
          contentType,
          data,
          isEmbedded,
          contentId: attachment.contentId || attachment.pidContentId,
          isNestedMsg,
          nestedMsgData
        });

        this.log(`✓ Successfully processed: ${filename} (${contentType}, ${data.length} bytes)`);

      } catch (error) {
        this.log(`❌ Error processing attachment ${i + 1}: ${error}`, true);
      }
    }

    this.log(`\n✅ Successfully processed ${attachments.length} out of ${typedMsgData.attachments.length} attachments`);
    
    return { attachments, logs: this.logs, errors: this.errors };
  }

  public async buildAttachmentsForEML(attachmentInfos: AttachmentInfo[], parentBuilder: MimeBuilder, 
                               emlBuilderFunction: (msgData: unknown) => Promise<MimeBuilder>): Promise<void> {
    
    for (let index = 0; index < attachmentInfos.length; index++) {
      const attachmentInfo = attachmentInfos[index];
      try {
        if (attachmentInfo.isNestedMsg && attachmentInfo.nestedMsgData) {
          // Handle nested MSG as EML attachment
          this.log(`Building nested EML for: ${attachmentInfo.name}`);
          
          const nestedEmlBuilder = await emlBuilderFunction(attachmentInfo.nestedMsgData);
          const nestedEmlContent = nestedEmlBuilder.build();
          
          const messageBuilder = parentBuilder.createChild('message/rfc822');
          messageBuilder.setContent(nestedEmlContent);
          messageBuilder.setHeader('Content-Disposition', `attachment; filename="${attachmentInfo.name}"`);
          
          this.log(`✓ Added nested EML: ${attachmentInfo.name}`);
          
        } else if (attachmentInfo.data && attachmentInfo.data.length > 0) {
          // Handle regular attachment
          const attachmentBuilder = parentBuilder.createChild(attachmentInfo.contentType);
          attachmentBuilder.setContent(attachmentInfo.data);
          
          // Set proper disposition
          if (attachmentInfo.isEmbedded && attachmentInfo.contentId) {
            attachmentBuilder.setHeader('Content-Disposition', `inline; filename="${attachmentInfo.name}"`);
            attachmentBuilder.setHeader('Content-ID', `<${attachmentInfo.contentId}>`);
          } else {
            attachmentBuilder.setHeader('Content-Disposition', `attachment; filename="${attachmentInfo.name}"`);
          }
          
          // Set encoding for binary files
          if (!attachmentInfo.contentType.startsWith('text/')) {
            attachmentBuilder.setHeader('Content-Transfer-Encoding', 'base64');
          }
          
          this.log(`✓ Added regular attachment: ${attachmentInfo.name} (${attachmentInfo.contentType})`);
          
        } else {
          this.log(`⚠️  Skipping attachment with no data: ${attachmentInfo.name}`, true);
        }
        
      } catch (error) {
        this.log(`❌ Error building attachment ${index + 1}: ${error}`, true);
      }
    }
  }

  public getLogs(): string[] {
    return [...this.logs];
  }

  public getErrors(): string[] {
    return [...this.errors];
  }
}
