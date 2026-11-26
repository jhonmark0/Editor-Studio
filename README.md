# Editor Studio

Editor Studio is a local photo, PDF, Word, and video editor for Windows and modern web browsers. Files stay on your device; the app makes no AI or API calls and does not upload your media.

## Install

Download and run `Editor-Studio-Setup-v2.0.1.exe` from the [latest GitHub release](https://github.com/jhonmark0/Editor-Studio/releases/latest). Setup installs only for the Windows account running it; it does not offer an all-users option or require administrator access. It can create Start Menu and Desktop shortcuts.

## Use in a browser

Open [Editor Studio on the web](https://jhonmark0.github.io/Editor-Studio/) on a desktop, tablet, or phone. No installation or account is required, and photo, PDF, Word, and video editing runs inside the browser tab without uploading your files. Video performance and file-size limits depend on the device; use the Windows app for larger jobs.

## Build from source

Requirements: Windows 10 or later and Node.js 22.12 or later.

```powershell
npm.cmd install
npm.cmd run dist:win
```

Run `npm.cmd run build:web` to build the responsive browser workspace into `docs`. GitHub Pages publishes that folder from the `main` branch.

The x64 Windows installer and update metadata are written to the `release` folder. This test build uses a local self-signed certificate. Windows does not trust it until a tester imports the public certificate; use a trusted code-signing certificate for public distribution.

## Publish updates

Packaged builds check the latest public release in the configured GitHub repository after startup. When a newer version is available, users can download it or choose **Later**, then restart to install it. Update releases must include the matching Windows installer and `latest.yml`, and must use the same signing identity as the installed app.

For controlled testing, a tester may import `Editor Studio Local Testing.cer` into the current user's **Trusted Root Certification Authorities** store. Verify its thumbprint is `FBCEE81BEE1EAEA1CCE523B44E219EA52D7BCF74` before trusting it. The self-signed certificate is not suitable for public distribution; use a trusted code-signing certificate. Never upload or embed a signing certificate's private key.

## Use

Drop files onto the source-file box or use the file picker, choose an output folder, and assign a task.

- **Photos:** batch resize, rotate, flip, convert to JPG/PNG/WebP/PDF, grayscale, or crop/resize for a passport-style print.
- **PDFs:** merge PDFs, combine images and PDFs in selection order, split one PDF into page files, or convert selectable-text PDFs to Word `.docx`. The Word conversion keeps each page's layout and artwork in place and makes selectable text editable; scanned/image-only PDFs require OCR first.
- **Word:** combine two or more `.docx` files in the order you arrange, or convert one Word document to PDF while preserving its page layout and artwork. Word files are never mixed into PDF merges.
- **Videos:** trim a clip, convert it to MP4, or extract its audio to M4A.

The original files are not overwritten. Passport-style processing only crops and resizes; it cannot change clothing or certify that a photo meets official requirements.
