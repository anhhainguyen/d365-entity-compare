# D365 Entity Compare

Chrome extension for comparing Microsoft Dynamics 365 Finance & Operations Data Entities across two environments.

## Features

- Compare available entities between a source and target environment
- Group entities by application module
- Compare selected OData-backed entity records
- Export an HTML comparison report
- Reuse the current authenticated D365 browser session

## Local Development

1. Open Chrome and go to `chrome://extensions`
2. Enable Developer mode
3. Click Load unpacked
4. Select this folder

## Files

- `manifest.json`: Chrome extension manifest
- `popup.html`: extension UI
- `popup.js`: compare logic
- `content.js`: fetches data from the active D365 tab using the browser session
- `background.js`: opens the full-page view
- `icons/`: Chrome extension icons
- `privacy.html`: privacy policy page for Chrome Web Store submission

## Privacy

The extension reads data from user-selected D365 environments only to perform comparisons inside the browser. It stores saved profiles and selections locally in the browser.

See `privacy.html` for the full privacy policy.

## License

This project is provided as-is unless a separate license is added.