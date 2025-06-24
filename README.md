# COA Document Processor - Word Add-in

Word Add-in for processing Certificate of Analysis (COA) documents with AI-powered field extraction.

## Features

- Seamless Word integration
- PDF document upload
- AI-powered field extraction
- Multi-language support (English/Chinese)
- Template management for different regions (CN/EU/US)
- Direct insertion into Word documents

## Prerequisites

- Node.js 14+
- Microsoft Word (Desktop version)
- Local HTTPS certificates
- Backend API running on https://localhost:8000

## Installation

1. Clone the repository:
```bash
cd coa-word-addin
```

2. Install dependencies:
```bash
npm install
```

3. Set up HTTPS certificates:
```bash
# Place your certificates in the certs/ directory:
# - certs/localhost.crt
# - certs/localhost.key
```

4. Configure environment:
```bash
cp .env.example .env
# Edit .env if needed
```

## Development

### Start Development Server

```bash
npm run dev
```

This will start the webpack dev server on https://localhost:3000

### Load Add-in in Word

1. Open Word
2. Go to Insert > Office Add-ins > Manage My Add-ins
3. Click "Upload My Add-in"
4. Browse and select `manifest.xml`
5. The add-in will appear in the Home tab

### Debugging

For debugging in Word Desktop:
```bash
npm run start:desktop
```

## Building for Production

```bash
npm run build
```

This creates optimized files in the `dist/` directory.

## Project Structure

```
coa-word-addin/
├── src/
│   ├── taskpane/          # Main UI components
│   │   ├── components/    # React components
│   │   ├── index.html     # Task pane HTML
│   │   └── index.tsx      # React entry point
│   ├── commands/          # Office command functions
│   └── types/             # TypeScript definitions
├── assets/                # Icons and images
├── certs/                 # HTTPS certificates
├── manifest.xml           # Office Add-in manifest
├── webpack.config.js      # Webpack configuration
└── package.json          # NPM dependencies
```

## Usage

1. **Select Compound**: Choose from BGB-21447, BGB-16673, or BGB-43395
2. **Select Template**: Choose regional template (CN/EU/US)
3. **Upload PDF**: Select COA PDF document to process
4. **Review Results**: Check and edit extracted fields
5. **Insert to Word**: Insert data into your document

## Components

- **CompoundSelector**: Dropdown for compound selection
- **TemplateSelector**: Regional template selection
- **FileUploader**: PDF upload with validation
- **ResultDisplay**: Show and edit extracted data
- **App**: Main application container

## Configuration

### manifest.xml
- Add-in metadata and permissions
- Button configuration
- Resource URLs

### webpack.config.js
- HTTPS certificate configuration
- Development server settings
- Build optimization

## Browser Compatibility

The add-in uses IE11 rendering engine in Word Desktop. Features:
- Polyfills included for IE11 compatibility
- No Edge WebView2 support required
- Fluent UI components work in IE11

## Troubleshooting

### Certificate Issues
- Ensure certificates are properly placed in `certs/` directory
- Certificates must be trusted by your system
- Use the same certificates registered with Office

### Loading Issues
- Clear Office cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
- Check console for errors (F12 Developer Tools)
- Verify backend API is running

### IE11 Compatibility
- Use babel polyfills
- Avoid modern JavaScript features not supported
- Test thoroughly in Word Desktop

## Security Notes

- Always use HTTPS in development and production
- Never commit certificates to version control
- Implement proper authentication before deployment

## API Integration

The add-in communicates with the backend API at:
- Development: `https://localhost:8000`
- Production: Configure in environment variables

Ensure CORS is properly configured on the backend.

## Future Enhancements

- [ ] Batch document processing
- [ ] Custom field mapping
- [ ] Export functionality
- [ ] Multi-language UI
- [ ] Advanced template editor
