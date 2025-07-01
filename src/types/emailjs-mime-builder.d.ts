declare module 'emailjs-mime-builder' {
  export default class MimeBuilder {
    constructor(contentType?: string);
    setHeader(key: string, value: string): MimeBuilder;
    setContent(content: string | Uint8Array): MimeBuilder;
    createChild(contentType: string): MimeBuilder;
    appendChild(child: MimeBuilder): MimeBuilder;
    build(): string;
  }
}
