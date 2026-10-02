# Secure Excel Column Cleaner

## What changed?

The Excel-processing logic is now on the Node.js server.

The browser receives only the UI and a small upload/download script. It does NOT receive the Excel-processing algorithm.

The six requested operations are performed server-side:

1. Enable editing / remove worksheet protection where present
2. Delete C-I
3. Delete H-K
4. Delete K onwards
5. Unmerge D and E
6. Delete E and resize D
7. Rebuild the requested final columns when all required headers are detected

Required final columns:

- Sr No.
- Student Id
- Student Name
- Attendance Date
- Batch Details
- Module Name
- Module Sequence
- Session Name

## Run locally

Install Node.js, then open a terminal in this folder:

```bash
npm install
npm start
```

Open:

http://localhost:5000

## Deploy

This project includes `vercel.json` for Vercel deployment.

You can also run the same project on a normal Node.js hosting provider.

## Important security note

No website can truly hide HTML/CSS/JavaScript that must be delivered to a browser.

This version protects the important Excel-processing logic by keeping it on the server.

The frontend blocks Ctrl+U, F12, right-click and common DevTools shortcuts as a basic deterrent. These controls are NOT security and can be bypassed.

For real protection, keep business logic in `server.js` and do not put the Excel-processing algorithm back into the frontend.
