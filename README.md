# whiteboard-bulk-export

A userscript that exports all your boards from Microsoft Whiteboard.

Each board can be downloaded as a PNG image, a ZIP archive, or both, using this filename format:

```text
<title>__<uuid>.png
<title>__<uuid>.zip
```

## Install

1. Install a userscript manager such as Tampermonkey.
2. Install [`whiteboard-bulk-export.user.js`](https://raw.githubusercontent.com/wyh2001/whiteboard-bulk-export/main/whiteboard-bulk-export.user.js).

## Use

1. Open the Whiteboard gallery.
2. Click **Export all whiteboards**.
3. Confirm the prompt and keep the tab open until the export finishes.

Your browser may ask for permission to download multiple files.

## Notes

- Future UI changes may break the script.
- Opening boards may change their order.
- The script uses UI automation to operate Whiteboard's existing export controls; it does not make direct API requests to Whiteboard, use undocumented endpoints, or upload board data elsewhere.
- This project is unofficial and is not affiliated with Microsoft.

## AI Disclosure

Built with AI assistance (Codex).

## License

Licensed under the [MIT License](LICENSE).
