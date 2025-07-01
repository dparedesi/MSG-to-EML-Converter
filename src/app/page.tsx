'use client';

import { useState, useCallback } from 'react';
import { Upload, Download, FileText, AlertCircle, CheckCircle, Loader2 } from 'lucide-react';
import { MSGToEMLConverterDebug as MSGToEMLConverter, ConversionResult, ConversionLog } from '@/lib/msgConverterDebug';
import { saveAs } from 'file-saver';

export default function Home() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionResult, setConversionResult] = useState<ConversionResult | null>(null);
  const [dragActive, setDragActive] = useState(false);

  const converter = new MSGToEMLConverter();

  const handleFileSelect = useCallback((file: File) => {
    if (file.name.toLowerCase().endsWith('.msg')) {
      setSelectedFile(file);
      setConversionResult(null);
    } else {
      alert('Please select a .msg file');
    }
  }, []);

  const handleDrag = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      handleFileSelect(files[0]);
    }
  }, [handleFileSelect]);

  const handleFileInput = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      handleFileSelect(files[0]);
    }
  }, [handleFileSelect]);

  const handleConvert = async () => {
    if (!selectedFile) return;

    setIsConverting(true);
    try {
      const result = await converter.convertMSGToEML(selectedFile);
      setConversionResult(result);
    } catch (error) {
      setConversionResult({
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error occurred',
        logs: []
      });
    }
    setIsConverting(false);
  };

  const handleDownload = () => {
    if (conversionResult?.success && conversionResult.emlData && conversionResult.filename) {
      const blob = new Blob([conversionResult.emlData], { type: 'message/rfc822' });
      saveAs(blob, conversionResult.filename);
    }
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const getLogIcon = (level: ConversionLog['level']) => {
    switch (level) {
      case 'error':
        return <AlertCircle className="w-4 h-4 text-red-500" />;
      case 'warn':
        return <AlertCircle className="w-4 h-4 text-yellow-500" />;
      default:
        return <CheckCircle className="w-4 h-4 text-blue-500" />;
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-8 px-4">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <div className="flex items-center justify-center mb-4">
            <FileText className="w-8 h-8 text-blue-600 mr-2" />
            <h1 className="text-3xl font-bold text-gray-900">MSG to EML Converter</h1>
          </div>
          <p className="text-lg text-gray-600 max-w-2xl mx-auto">
            Convert Microsoft Outlook .msg files to standard .eml format with support for nested attachments and embedded messages.
          </p>
        </div>

        {/* File Upload Area */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <div
            className={`border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
              dragActive
                ? 'border-blue-400 bg-blue-50'
                : 'border-gray-300 hover:border-gray-400'
            }`}
            onDragEnter={handleDrag}
            onDragLeave={handleDrag}
            onDragOver={handleDrag}
            onDrop={handleDrop}
          >
            <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
            <h3 className="text-lg font-semibold text-gray-900 mb-2">
              Upload your .MSG file
            </h3>
            <p className="text-gray-600 mb-4">
              Drag and drop your file here, or click to browse
            </p>
            <input
              type="file"
              accept=".msg"
              onChange={handleFileInput}
              className="hidden"
              id="file-input"
            />
            <label
              htmlFor="file-input"
              className="inline-flex items-center px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 cursor-pointer transition-colors"
            >
              Choose File
            </label>
          </div>

          {/* Selected File Info */}
          {selectedFile && (
            <div className="mt-6 p-4 bg-gray-50 rounded-lg">
              <h4 className="font-semibold text-gray-900 mb-2">Selected File Details:</h4>
              <div className="space-y-1 text-sm text-gray-600">
                <p><span className="font-medium">Name:</span> {selectedFile.name}</p>
                <p><span className="font-medium">Size:</span> {formatFileSize(selectedFile.size)}</p>
                <p><span className="font-medium">Type:</span> {selectedFile.type || 'application/vnd.ms-outlook'}</p>
              </div>
            </div>
          )}
        </div>

        {/* Convert Button */}
        {selectedFile && !conversionResult && (
          <div className="text-center mb-6">
            <button
              onClick={handleConvert}
              disabled={isConverting}
              className="inline-flex items-center px-6 py-3 bg-green-600 text-white font-semibold rounded-lg hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            >
              {isConverting ? (
                <>
                  <Loader2 className="w-5 h-5 mr-2 animate-spin" />
                  Converting...
                </>
              ) : (
                <>
                  <FileText className="w-5 h-5 mr-2" />
                  Convert &apos;{selectedFile.name}&apos; to EML
                </>
              )}
            </button>
          </div>
        )}

        {/* Conversion Results */}
        {conversionResult && (
          <div className="bg-white rounded-lg shadow-lg p-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">
              Conversion Results
            </h3>

            {conversionResult.success ? (
              <div className="space-y-4">
                <div className="flex items-center p-4 bg-green-50 rounded-lg">
                  <CheckCircle className="w-6 h-6 text-green-500 mr-3" />
                  <div>
                    <p className="font-semibold text-green-800">Conversion Successful!</p>
                    <p className="text-green-600">Your EML file is ready for download.</p>
                  </div>
                </div>

                <div className="text-center">
                  <button
                    onClick={handleDownload}
                    className="inline-flex items-center px-6 py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 transition-colors"
                  >
                    <Download className="w-5 h-5 mr-2" />
                    Download {conversionResult.filename}
                  </button>
                </div>
              </div>
            ) : (
              <div className="flex items-center p-4 bg-red-50 rounded-lg">
                <AlertCircle className="w-6 h-6 text-red-500 mr-3" />
                <div>
                  <p className="font-semibold text-red-800">Conversion Failed</p>
                  <p className="text-red-600">{conversionResult.error}</p>
                </div>
              </div>
            )}

            {/* Conversion Logs */}
            {conversionResult.logs.length > 0 && (
              <div className="mt-6">
                <h4 className="font-semibold text-gray-900 mb-3">Conversion Process:</h4>
                <div className="bg-gray-50 rounded-lg p-4 max-h-64 overflow-y-auto">
                  <div className="space-y-2">
                    {conversionResult.logs.map((log, index) => (
                      <div key={index} className="flex items-start space-x-2 text-sm">
                        {getLogIcon(log.level)}
                        <div className="flex-1">
                          <span className="text-gray-600">
                            [{log.timestamp.toLocaleTimeString()}]
                          </span>
                          <span className="ml-2 text-gray-800">{log.message}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* Reset Button */}
            <div className="mt-6 text-center">
              <button
                onClick={() => {
                  setSelectedFile(null);
                  setConversionResult(null);
                }}
                className="text-blue-600 hover:text-blue-800 font-medium"
              >
                Convert Another File
              </button>
            </div>
          </div>
        )}

        {/* Footer */}
        <div className="mt-12 text-center text-gray-600">
          <p className="text-sm">
            Built with Next.js, TypeScript, and Tailwind CSS • 
            Supports nested MSG attachments and embedded messages
          </p>
        </div>
      </div>
    </div>
  );
}
