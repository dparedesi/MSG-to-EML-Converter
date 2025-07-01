import type { Metadata } from 'next';
import { Inter } from 'next/font/google';
import './globals.css';

const inter = Inter({ subsets: ['latin'] });

export const metadata: Metadata = {
  title: 'MSG to EML Converter',
  description: 'Convert Microsoft Outlook .msg files to standard .eml format with support for nested attachments and embedded messages.',
  keywords: ['msg', 'eml', 'email', 'converter', 'outlook', 'microsoft'],
  authors: [{ name: 'MSG to EML Converter' }],
  openGraph: {
    title: 'MSG to EML Converter',
    description: 'Convert Microsoft Outlook .msg files to standard .eml format',
    type: 'website',
  },
  twitter: {
    card: 'summary_large_image',
    title: 'MSG to EML Converter',
    description: 'Convert Microsoft Outlook .msg files to standard .eml format',
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className={inter.className}>{children}</body>
    </html>
  );
}
