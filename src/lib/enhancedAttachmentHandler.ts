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

  private rawMsgBuffer?: Uint8Array;

  public setRawMsgBuffer(buffer: Uint8Array) {
    this.rawMsgBuffer = buffer;
  }

  private extractAttachmentFromRawBuffer(attachment: MSGAttachment): Uint8Array | null {
    if (!this.rawMsgBuffer || !attachment.name || typeof attachment.contentLength !== 'number') {
      return null;
    }

    this.log(`🔍 Raw binary search for: ${attachment.name}`);
    
    try {
      // Convert to Buffer for easier searching
      const msgBuffer = Buffer.from(this.rawMsgBuffer);
      const attachmentName = attachment.name;
      const expectedSize = attachment.contentLength as number;
      
      // Search for attachment name in the MSG file
      let nameIndex = -1;
      let nameBufferLength = 0;
      
      // Try UTF-8 encoding first
      const nameBufferUtf8 = Buffer.from(attachmentName, 'utf8');
      nameIndex = msgBuffer.indexOf(nameBufferUtf8);
      if (nameIndex !== -1) {
        nameBufferLength = nameBufferUtf8.length;
      } else {
        // Try UTF-16LE encoding
        const nameBufferUtf16 = Buffer.from(attachmentName, 'utf16le');
        nameIndex = msgBuffer.indexOf(nameBufferUtf16);
        if (nameIndex !== -1) {
          nameBufferLength = nameBufferUtf16.length;
        }
      }
      
      if (nameIndex === -1) {
        this.log(`❌ Could not find attachment name in MSG buffer`);
        return null;
      }
      
      this.log(`📍 Found attachment name at offset: ${nameIndex}`);
      
      // Define file signatures to look for
      const fileSignatures: Record<string, Buffer> = {
        'png': Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]),
        'jpg': Buffer.from([0xFF, 0xD8, 0xFF]),
        'jpeg': Buffer.from([0xFF, 0xD8, 0xFF]),
        'pdf': Buffer.from([0x25, 0x50, 0x44, 0x46]),
        'zip': Buffer.from([0x50, 0x4B, 0x03, 0x04]),
        'docx': Buffer.from([0x50, 0x4B, 0x03, 0x04]),
        'xlsx': Buffer.from([0x50, 0x4B, 0x03, 0x04]),
        'pptx': Buffer.from([0x50, 0x4B, 0x03, 0x04])
      };
      
      // Determine which signature to look for
      const fileExt = attachmentName.toLowerCase().split('.').pop() || '';
      const signature = fileSignatures[fileExt];
      
      if (!signature) {
        this.log(`⚠️ Unknown file type: ${fileExt}, trying generic approach`);
        return null;
      }
      
      this.log(`🔍 Looking for ${fileExt.toUpperCase()} signature: ${Array.from(signature).map(b => b.toString(16).padStart(2, '0')).join(' ')}`);
      
      // Search for the file signature near the attachment name
      // Try multiple locations around the name reference
      const searchRanges = [
        { start: Math.max(0, nameIndex - 10000), end: Math.min(msgBuffer.length, nameIndex + 50000) },
        { start: Math.max(0, nameIndex + nameBufferLength), end: Math.min(msgBuffer.length, nameIndex + nameBufferLength + 100000) }
      ];
      
      for (const range of searchRanges) {
        const searchBuffer = msgBuffer.slice(range.start, range.end);
        let sigIndex = searchBuffer.indexOf(signature);
        
        while (sigIndex !== -1) {
          const absoluteOffset = range.start + sigIndex;
          this.log(`🎯 Found ${fileExt.toUpperCase()} signature at offset: ${absoluteOffset}`);
          
          // Extract data of expected size starting from signature
          const candidateData = msgBuffer.slice(absoluteOffset, absoluteOffset + expectedSize);
          
          if (candidateData.length === expectedSize) {
            // Verify this is likely the correct file by checking if it ends properly
            let isValidFile = true;
            
            // Additional validation for specific file types
            if (fileExt === 'png') {
              // PNG files should end with IEND chunk
              const iendSignature = Buffer.from([0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82]);
              const lastBytes = candidateData.slice(-8);
              isValidFile = lastBytes.equals(iendSignature);
              if (!isValidFile) {
                // Try looking for IEND anywhere in the last 50 bytes
                const searchArea = candidateData.slice(-50);
                isValidFile = searchArea.indexOf(iendSignature) !== -1;
              }
            } else if (fileExt === 'docx' || fileExt === 'xlsx' || fileExt === 'pptx') {
              // ZIP files should have proper structure
              isValidFile = candidateData.length > 100; // Basic size check
            }
            
            if (isValidFile) {
              this.log(`✅ Successfully extracted ${expectedSize} bytes for ${attachmentName}`);
              return new Uint8Array(candidateData);
            } else {
              this.log(`⚠️ File validation failed, continuing search...`);
            }
          } else {
            this.log(`⚠️ Size mismatch: found ${candidateData.length}, expected ${expectedSize}`);
          }
          
          // Look for next occurrence
          sigIndex = searchBuffer.indexOf(signature, sigIndex + 1);
        }
      }
      
      this.log(`❌ Could not locate valid ${fileExt} data in MSG buffer`);
      return null;
      
    } catch (error) {
      this.log(`❌ Raw buffer extraction failed: ${error}`, true);
      return null;
    }
  }

  private extractAttachmentData(attachment: MSGAttachment, msgReader?: MSGReader): Uint8Array | null {
    // CRITICAL WORKAROUND: @kenjiuno/msgreader has systemic bugs across versions where 
    // getAttachment(dataId) fails and attachment binary data is completely missing from 
    // all internal data structures. We implement multiple fallback strategies:
    
    this.log(`Attempting to extract attachment data for: ${attachment.name || 'unnamed'}`);
    this.log(`Expected size: ${attachment.contentLength || 'unknown'} bytes`);

    // Strategy 1: Try standard msgReader.getAttachment (expected to fail due to library bugs)
    if (attachment.dataId && msgReader) {
      try {
        const attachmentData = msgReader.getAttachment(attachment.dataId);
        if (attachmentData && (attachmentData instanceof Uint8Array || attachmentData instanceof ArrayBuffer || (Buffer.isBuffer && Buffer.isBuffer(attachmentData)))) {
          const data = attachmentData instanceof Uint8Array ? attachmentData : 
                      attachmentData instanceof ArrayBuffer ? new Uint8Array(attachmentData) :
                      new Uint8Array(attachmentData);
          if (data.length > 0) {
            this.log(`✅ Extracted via msgReader.getAttachment(): ${data.length} bytes`);
            return data;
          }
        }
      } catch (error) {
        this.log(`❌ msgReader.getAttachment(${attachment.dataId}) failed: ${error}`, true);
      }
    }

    // Strategy 2: Check msgReader's internal data structures
    if (attachment.dataId && msgReader && (msgReader as unknown as { fieldsData?: Record<string, unknown> }).fieldsData) {
      const fieldsData = (msgReader as unknown as { fieldsData: Record<string, unknown> }).fieldsData;
      
      // Try direct access
      if (fieldsData[attachment.dataId]) {
        const data = fieldsData[attachment.dataId];
        if (data instanceof Uint8Array || data instanceof ArrayBuffer || (Buffer.isBuffer && Buffer.isBuffer(data))) {
          const extractedData = data instanceof Uint8Array ? data : 
                               data instanceof ArrayBuffer ? new Uint8Array(data) :
                               new Uint8Array(data);
          if (extractedData.length > 0) {
            this.log(`✅ Extracted via fieldsData[${attachment.dataId}]: ${extractedData.length} bytes`);
            return extractedData;
          }
        }
      }
      
      // Try nearby keys (library might use different IDs)
      const allKeys = Object.keys(fieldsData).map(k => parseInt(k)).filter(k => !isNaN(k));
      const dataIdNum = typeof attachment.dataId === 'number' ? attachment.dataId : parseInt(String(attachment.dataId));
      
      if (!isNaN(dataIdNum)) {
        const nearbyKeys = allKeys.filter(k => Math.abs(k - dataIdNum) < 50);
        for (const key of nearbyKeys) {
          const data = fieldsData[key];
          if (data && (data instanceof Uint8Array || data instanceof ArrayBuffer || (Buffer.isBuffer && Buffer.isBuffer(data)))) {
            const dataSize = data instanceof ArrayBuffer ? data.byteLength : data.length;
            if (dataSize === attachment.contentLength) {
              const extractedData = data instanceof Uint8Array ? data : 
                                  data instanceof ArrayBuffer ? new Uint8Array(data) :
                                  new Uint8Array(data);
              this.log(`✅ Found matching data at fieldsData[${key}]: ${extractedData.length} bytes`);
              return extractedData;
            }
          }
        }
      }
    }

    // Strategy 3: Check attachment object properties
    const possibleDataFields = ['data', 'content', 'body', 'dataBody', 'attachData', 'innerMsgContent', 'attachmentData'];
    for (const field of possibleDataFields) {
      const data = attachment[field];
      if (data && (data instanceof Uint8Array || data instanceof ArrayBuffer || (Buffer.isBuffer && Buffer.isBuffer(data)))) {
        const extractedData = data instanceof Uint8Array ? data : 
                             data instanceof ArrayBuffer ? new Uint8Array(data) :
                             new Uint8Array(data);
        if (extractedData.length > 0) {
          this.log(`✅ Found data in attachment.${field}: ${extractedData.length} bytes`);
          return extractedData;
        }
      }
    }

    // Strategy 4: Try attachment's getAttachment method
    if (typeof attachment.getAttachment === 'function') {
      try {
        const attachmentData = attachment.getAttachment();
        if (attachmentData && (attachmentData instanceof Uint8Array || attachmentData instanceof ArrayBuffer)) {
          const data = attachmentData instanceof Uint8Array ? attachmentData : new Uint8Array(attachmentData);
          if (data.length > 0) {
            this.log(`✅ Extracted via attachment.getAttachment(): ${data.length} bytes`);
            return data;
          }
        }
      } catch (error) {
        this.log(`❌ attachment.getAttachment() failed: ${error}`, true);
      }
    }

    // Strategy 5: ULTIMATE FALLBACK - Raw binary search in MSG buffer
    // This bypasses the broken library completely and searches for attachment data directly
    this.log(`🚨 All standard methods failed, attempting raw binary extraction...`);
    const rawExtractedData = this.extractAttachmentFromRawBuffer(attachment);
    if (rawExtractedData) {
      return rawExtractedData;
    }

    this.log(`❌ COMPLETE FAILURE: Could not extract attachment data for: ${attachment.name || 'unnamed'}`, true);
    this.log(`This appears to be a critical bug in @kenjiuno/msgreader library`, true);
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
