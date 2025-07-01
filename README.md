# MSG to EML Converter ✉️

A modern web application to convert Microsoft Outlook `.msg` files into standard `.eml` files. Built with Next.js, TypeScript, and Tailwind CSS for seamless deployment on Vercel.

## Features

- **🔄 MSG to EML Conversion**: Convert Outlook `.msg` files to standard `.eml` format
- **📎 Nested Attachments**: Handles nested `.msg` attachments recursively
- **🎨 Modern UI**: Clean, responsive interface with drag-and-drop support
- **⚡ Client-Side Processing**: All processing happens in the browser - no server required
- **📱 Responsive Design**: Works perfectly on desktop and mobile devices
- **🚀 Vercel Ready**: Optimized for deployment on Vercel platform

## How It Works

1. **Upload**: Drag and drop or browse for your `.msg` file
2. **Convert**: The app processes the file client-side using JavaScript libraries
3. **Download**: Get your converted `.eml` file instantly

## Tech Stack

- **Frontend**: Next.js 15 with App Router
- **Language**: TypeScript
- **Styling**: Tailwind CSS
- **Icons**: Lucide React
- **MSG Parsing**: @kenjiuno/msgreader
- **EML Generation**: emailjs-mime-builder
- **File Handling**: file-saver

## Development

### Prerequisites

- Node.js 18 or higher
- npm or yarn

### Installation

```bash
# Clone the repository
git clone <your-repo-url>
cd MSG-to-EML-Converter

# Install dependencies
npm install

# Run development server
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

### Build for Production

```bash
# Build the application
npm run build

# Start production server
npm start
```

## Deployment on Vercel

1. **Push to GitHub**: Push your code to a GitHub repository

2. **Connect to Vercel**: 
   - Go to [vercel.com](https://vercel.com)
   - Import your GitHub repository
   - Vercel will automatically detect it's a Next.js project

3. **Deploy**: 
   - Vercel will build and deploy automatically
   - Your app will be available at `https://your-app-name.vercel.app`

### Environment Configuration

No environment variables are required as this is a client-side only application.

## Browser Compatibility

- Chrome 88+
- Firefox 85+
- Safari 14+
- Edge 88+

## File Size Limits

The application can handle MSG files up to the browser's memory limit (typically several hundred MB for modern browsers).

## Security

- All processing happens client-side
- No files are uploaded to servers
- No data is stored or transmitted

## Supported Features

### MSG File Support
- ✅ Subject, From, To, CC headers
- ✅ Plain text and HTML body content
- ✅ File attachments
- ✅ Nested MSG attachments (recursive conversion)
- ✅ Date/time information
- ✅ Message IDs

### EML Output
- ✅ RFC 2822 compliant format
- ✅ Proper MIME structure
- ✅ Embedded attachments as message/rfc822
- ✅ Base64 encoding for binary attachments

## Known Limitations

- Very large files (>100MB) may cause browser memory issues
- Some advanced Outlook-specific features may not be preserved
- Requires modern browser with File API support

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [Next.js](https://nextjs.org/)
- MSG parsing by [@kenjiuno/msgreader](https://www.npmjs.com/package/@kenjiuno/msgreader)
- EML generation by [emailjs-mime-builder](https://www.npmjs.com/package/emailjs-mime-builder)
- Icons by [Lucide](https://lucide.dev/)
